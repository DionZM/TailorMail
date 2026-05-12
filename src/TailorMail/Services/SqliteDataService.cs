using System.Text.Json;
using Microsoft.Data.Sqlite;
using TailorMail.Models;

namespace TailorMail.Services;

public class SqliteDataService : IDataService
{
    private static readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = false,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
    };

    // R-05: Instance field with lock instead of static flag
    private readonly object _dedupLock = new();
    private bool _dedupDone;

    private readonly string _dbPath;
    private readonly string _connectionString;

    public SqliteDataService()
    {
        var dataDir = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data");
        System.IO.Directory.CreateDirectory(dataDir);
        _dbPath = System.IO.Path.Combine(dataDir, "tailormail.db");
        // P-02: Enable WAL mode and busy_timeout for better concurrency
        _connectionString = $"Data Source={_dbPath}";
        InitializeDatabase();
        MigrateFromJsonIfNeeded();
    }

    private SqliteConnection CreateConnection() => new(_connectionString);

    private void InitializeDatabase()
    {
        using var conn = CreateConnection();
        conn.Open();

        // P-02: Enable WAL mode and set busy timeout
        using (var pragma = conn.CreateCommand())
        {
            pragma.CommandText = "PRAGMA journal_mode=WAL; PRAGMA busy_timeout=5000;";
            pragma.ExecuteNonQuery();
        }

        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            CREATE TABLE IF NOT EXISTS RecipientGroups (
                Id TEXT PRIMARY KEY,
                Name TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS Recipients (
                Id TEXT PRIMARY KEY,
                GroupId TEXT NOT NULL,
                Name TEXT NOT NULL DEFAULT '',
                ShortName TEXT NOT NULL DEFAULT '',
                ToEmails TEXT NOT NULL DEFAULT '',
                CcEmails TEXT NOT NULL DEFAULT '',
                BccEmails TEXT NOT NULL DEFAULT '',
                Remark TEXT NOT NULL DEFAULT '',
                IsSelected INTEGER NOT NULL DEFAULT 0,
                VariablesJson TEXT NOT NULL DEFAULT '{}',
                SortOrder INTEGER NOT NULL DEFAULT 0,
                FOREIGN KEY (GroupId) REFERENCES RecipientGroups(Id)
            );

            CREATE INDEX IF NOT EXISTS IX_Recipients_GroupId ON Recipients(GroupId);

            CREATE TABLE IF NOT EXISTS AppSettings (
                Id INTEGER PRIMARY KEY CHECK (Id = 1),
                SettingsJson TEXT NOT NULL DEFAULT '{}'
            );

            CREATE TABLE IF NOT EXISTS AttachmentConfig (
                Id INTEGER PRIMARY KEY CHECK (Id = 1),
                ConfigJson TEXT NOT NULL DEFAULT '{}'
            );

            CREATE TABLE IF NOT EXISTS MailTemplates (
                Id TEXT PRIMARY KEY,
                Name TEXT NOT NULL DEFAULT '',
                Subject TEXT NOT NULL DEFAULT '',
                Body TEXT NOT NULL DEFAULT '',
                CreatedAt TEXT NOT NULL DEFAULT ''
            );
            """;
        cmd.ExecuteNonQuery();

        DeduplicateDefaultGroups(conn);
    }

    private void DeduplicateDefaultGroups(SqliteConnection conn)
    {
        // H-03: Entire dedup logic inside lock to prevent TOCTOU race
        lock (_dedupLock)
        {
            if (_dedupDone) return;

            // E-02: Transaction protection + exception handling
            using var tx = conn.BeginTransaction();
            try
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = """
                    SELECT Id FROM RecipientGroups WHERE Name = '默认分组' ORDER BY rowid
                    """;
                cmd.Transaction = tx;
                var ids = new List<string>();
                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                        ids.Add(reader.GetString(0));
                }

                if (ids.Count <= 1)
                {
                    tx.Commit();
                    _dedupDone = true;
                    return;
                }

                var keepId = ids[0];
                for (int i = 1; i < ids.Count; i++)
                {
                    using var mergeCmd = conn.CreateCommand();
                    mergeCmd.CommandText = """
                        UPDATE Recipients SET GroupId = @keepId WHERE GroupId = @removeId;
                        DELETE FROM RecipientGroups WHERE Id = @removeId
                        """;
                    mergeCmd.Transaction = tx;
                    mergeCmd.Parameters.AddWithValue("@keepId", keepId);
                    mergeCmd.Parameters.AddWithValue("@removeId", ids[i]);
                    mergeCmd.ExecuteNonQuery();
                }

                tx.Commit();
                _dedupDone = true;
            }
            catch (Exception ex)
            {
                try { tx.Rollback(); } catch { }
                AppLogger.Error("DeduplicateDefaultGroups 失败", ex);
            }
        }
    }

    private void MigrateFromJsonIfNeeded()
    {
        var dataDir = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data");
        var migrationFlag = System.IO.Path.Combine(dataDir, ".migrated_to_sqlite");

        if (System.IO.File.Exists(migrationFlag)) return;

        var jsonService = new JsonDataService();

        // E-01: Log migration failures instead of silently swallowing
        try
        {
            var groups = jsonService.LoadRecipientGroups();
            if (groups.Count > 0 && groups.Any(g => g.Recipients.Count > 0))
            {
                SaveRecipientGroups(groups);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("迁移收件人分组失败", ex);
        }

        try
        {
            var settings = jsonService.LoadSettings();
            if (!string.IsNullOrEmpty(settings.LastSubject) || !string.IsNullOrEmpty(settings.Smtp.Host))
            {
                SaveSettings(settings);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("迁移应用设置失败", ex);
        }

        try
        {
            var config = jsonService.LoadAttachmentConfig();
            if (config.CommonAttachments.Count > 0 || config.RecipientAttachments.Count > 0)
            {
                SaveAttachmentConfig(config);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("迁移附件配置失败", ex);
        }

        try
        {
            var templates = jsonService.LoadTemplates();
            if (templates.Count > 0)
            {
                SaveTemplates(templates);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("迁移邮件模板失败", ex);
        }

        System.IO.File.WriteAllText(migrationFlag, DateTime.Now.ToString("O"));
    }

    public List<RecipientGroup> LoadRecipientGroups()
    {
        using var conn = CreateConnection();
        conn.Open();

        var groups = new List<RecipientGroup>();
        var groupMap = new Dictionary<string, RecipientGroup>();

        // PERF-3: Single query with JOIN instead of N+1 queries
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = """
                SELECT r.Id, r.GroupId, r.Name, r.ShortName, r.ToEmails, r.CcEmails, r.BccEmails, r.Remark, r.IsSelected, r.VariablesJson, r.SortOrder,
                       g.Name AS GroupName
                FROM Recipients r
                INNER JOIN RecipientGroups g ON r.GroupId = g.Id
                ORDER BY g.rowid, r.SortOrder, r.rowid
                """;
            using var reader = cmd.ExecuteReader();

            // PERF-9: Use named ordinals instead of hardcoded indices
            var idOrdinal = reader.GetOrdinal("Id");
            var groupIdOrdinal = reader.GetOrdinal("GroupId");
            var nameOrdinal = reader.GetOrdinal("Name");
            var shortNameOrdinal = reader.GetOrdinal("ShortName");
            var toEmailsOrdinal = reader.GetOrdinal("ToEmails");
            var ccEmailsOrdinal = reader.GetOrdinal("CcEmails");
            var bccEmailsOrdinal = reader.GetOrdinal("BccEmails");
            var remarkOrdinal = reader.GetOrdinal("Remark");
            var isSelectedOrdinal = reader.GetOrdinal("IsSelected");
            var variablesOrdinal = reader.GetOrdinal("VariablesJson");
            var groupNameOrdinal = reader.GetOrdinal("GroupName");

            while (reader.Read())
            {
                var groupId = reader.GetString(groupIdOrdinal);

                if (!groupMap.TryGetValue(groupId, out var group))
                {
                    group = new RecipientGroup
                    {
                        Id = groupId,
                        Name = reader.GetString(groupNameOrdinal)
                    };
                    groupMap[groupId] = group;
                    groups.Add(group);
                }

                var variablesJson = reader.GetString(variablesOrdinal);
                Dictionary<string, string> variables = [];
                if (!string.IsNullOrEmpty(variablesJson) && variablesJson != "{}")
                {
                    try
                    {
                        variables = JsonSerializer.Deserialize<Dictionary<string, string>>(variablesJson, _jsonOptions) ?? [];
                    }
                    catch (Exception ex)
                    {
                        AppLogger.Warning($"解析收件人变量JSON失败: {ex.Message}");
                    }
                }

                group.Recipients.Add(new Recipient
                {
                    Id = reader.GetString(idOrdinal),
                    Name = reader.GetString(nameOrdinal),
                    ShortName = reader.GetString(shortNameOrdinal),
                    ToEmails = reader.GetString(toEmailsOrdinal),
                    CcEmails = reader.GetString(ccEmailsOrdinal),
                    BccEmails = reader.GetString(bccEmailsOrdinal),
                    Remark = reader.GetString(remarkOrdinal),
                    IsSelected = reader.GetInt32(isSelectedOrdinal) == 1,
                    Variables = variables
                });
            }
        }

        // Also load groups that have no recipients
        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT Id, Name FROM RecipientGroups WHERE Id NOT IN (SELECT DISTINCT GroupId FROM Recipients) ORDER BY rowid";
            using var reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                groups.Add(new RecipientGroup
                {
                    Id = reader.GetString(0),
                    Name = reader.GetString(1)
                });
            }
        }

        if (groups.Count == 0)
        {
            var defaultGroup = new RecipientGroup { Name = "默认分组" };
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "INSERT INTO RecipientGroups (Id, Name) VALUES (@id, @name)";
                cmd.Parameters.AddWithValue("@id", defaultGroup.Id);
                cmd.Parameters.AddWithValue("@name", defaultGroup.Name);
                cmd.ExecuteNonQuery();
            }
            groups.Add(defaultGroup);
        }

        return groups;
    }

    // P-01: Async overload
    public async Task<List<RecipientGroup>> LoadRecipientGroupsAsync()
    {
        await using var conn = CreateConnection();
        await conn.OpenAsync();

        var groups = new List<RecipientGroup>();
        var groupMap = new Dictionary<string, RecipientGroup>();

        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = """
                SELECT r.Id, r.GroupId, r.Name, r.ShortName, r.ToEmails, r.CcEmails, r.BccEmails, r.Remark, r.IsSelected, r.VariablesJson, r.SortOrder,
                       g.Name AS GroupName
                FROM Recipients r
                INNER JOIN RecipientGroups g ON r.GroupId = g.Id
                ORDER BY g.rowid, r.SortOrder, r.rowid
                """;
            await using var reader = await cmd.ExecuteReaderAsync();

            var idOrdinal = reader.GetOrdinal("Id");
            var groupIdOrdinal = reader.GetOrdinal("GroupId");
            var nameOrdinal = reader.GetOrdinal("Name");
            var shortNameOrdinal = reader.GetOrdinal("ShortName");
            var toEmailsOrdinal = reader.GetOrdinal("ToEmails");
            var ccEmailsOrdinal = reader.GetOrdinal("CcEmails");
            var bccEmailsOrdinal = reader.GetOrdinal("BccEmails");
            var remarkOrdinal = reader.GetOrdinal("Remark");
            var isSelectedOrdinal = reader.GetOrdinal("IsSelected");
            var variablesOrdinal = reader.GetOrdinal("VariablesJson");
            var groupNameOrdinal = reader.GetOrdinal("GroupName");

            while (await reader.ReadAsync())
            {
                var groupId = reader.GetString(groupIdOrdinal);

                if (!groupMap.TryGetValue(groupId, out var group))
                {
                    group = new RecipientGroup
                    {
                        Id = groupId,
                        Name = reader.GetString(groupNameOrdinal)
                    };
                    groupMap[groupId] = group;
                    groups.Add(group);
                }

                var variablesJson = reader.GetString(variablesOrdinal);
                Dictionary<string, string> variables = [];
                if (!string.IsNullOrEmpty(variablesJson) && variablesJson != "{}")
                {
                    try
                    {
                        variables = JsonSerializer.Deserialize<Dictionary<string, string>>(variablesJson, _jsonOptions) ?? [];
                    }
                    catch (Exception ex)
                    {
                        AppLogger.Warning($"解析收件人变量JSON失败: {ex.Message}");
                    }
                }

                group.Recipients.Add(new Recipient
                {
                    Id = reader.GetString(idOrdinal),
                    Name = reader.GetString(nameOrdinal),
                    ShortName = reader.GetString(shortNameOrdinal),
                    ToEmails = reader.GetString(toEmailsOrdinal),
                    CcEmails = reader.GetString(ccEmailsOrdinal),
                    BccEmails = reader.GetString(bccEmailsOrdinal),
                    Remark = reader.GetString(remarkOrdinal),
                    IsSelected = reader.GetInt32(isSelectedOrdinal) == 1,
                    Variables = variables
                });
            }
        }

        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT Id, Name FROM RecipientGroups WHERE Id NOT IN (SELECT DISTINCT GroupId FROM Recipients) ORDER BY rowid";
            await using var reader = await cmd.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                groups.Add(new RecipientGroup
                {
                    Id = reader.GetString(0),
                    Name = reader.GetString(1)
                });
            }
        }

        if (groups.Count == 0)
        {
            var defaultGroup = new RecipientGroup { Name = "默认分组" };
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "INSERT INTO RecipientGroups (Id, Name) VALUES (@id, @name)";
                cmd.Parameters.AddWithValue("@id", defaultGroup.Id);
                cmd.Parameters.AddWithValue("@name", defaultGroup.Name);
                await cmd.ExecuteNonQueryAsync();
            }
            groups.Add(defaultGroup);
        }

        return groups;
    }

    public void SaveRecipientGroups(List<RecipientGroup> groups)
    {
        using var conn = CreateConnection();
        conn.Open();
        using var tx = conn.BeginTransaction();

        try
        {
            // P-03: Use UPSERT (INSERT OR REPLACE) instead of DELETE + re-insert
            using (var groupCmd = conn.CreateCommand())
            {
                groupCmd.CommandText = """
                    INSERT OR REPLACE INTO RecipientGroups (Id, Name) VALUES (@id, @name)
                    """;
                groupCmd.Transaction = tx;
                var gIdParam = groupCmd.Parameters.Add("@id", SqliteType.Text);
                var gNameParam = groupCmd.Parameters.Add("@name", SqliteType.Text);

                foreach (var group in groups)
                {
                    gIdParam.Value = group.Id;
                    gNameParam.Value = group.Name;
                    groupCmd.ExecuteNonQuery();
                }
            }

            // P-03: Collect all current recipient IDs for cleanup
            var currentRecipientIds = new HashSet<string>();
            foreach (var group in groups)
                foreach (var r in group.Recipients)
                    currentRecipientIds.Add(r.Id);

            // H-10: Guard against empty list deleting all data
            if (currentRecipientIds.Count == 0)
            {
                // No recipients to save — skip deletion to prevent accidental data loss
            }
            else
            {
                // P-03: Delete recipients that are no longer in any group
                using (var cmd = conn.CreateCommand())
                {
                    var placeholders = string.Join(",", currentRecipientIds.Select((_, i) => $"@delId{i}"));
                    cmd.CommandText = $"DELETE FROM Recipients WHERE Id NOT IN ({placeholders})";
                    cmd.Transaction = tx;
                    int idx = 0;
                    foreach (var id in currentRecipientIds)
                        cmd.Parameters.AddWithValue($"@delId{idx++}", id);
                    cmd.ExecuteNonQuery();
                }
            }

            // Delete groups that are no longer present
            if (groups.Count > 0)
            {
                using (var cmd = conn.CreateCommand())
                {
                    var placeholders = string.Join(",", groups.Select((_, i) => $"@g{i}"));
                    cmd.CommandText = $"DELETE FROM RecipientGroups WHERE Id NOT IN ({placeholders})";
                    cmd.Transaction = tx;
                    for (int i = 0; i < groups.Count; i++)
                        cmd.Parameters.AddWithValue($"@g{i}", groups[i].Id);
                    cmd.ExecuteNonQuery();
                }
            }
            else
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "DELETE FROM RecipientGroups";
                    cmd.Transaction = tx;
                    cmd.ExecuteNonQuery();
                }
            }

            // P-03: UPSERT recipients
            using (var recCmd = conn.CreateCommand())
            {
                recCmd.CommandText = """
                    INSERT OR REPLACE INTO Recipients (Id, GroupId, Name, ShortName, ToEmails, CcEmails, BccEmails, Remark, IsSelected, VariablesJson, SortOrder)
                    VALUES (@id, @groupId, @name, @shortName, @toEmails, @ccEmails, @bccEmails, @remark, @isSelected, @variablesJson, @sortOrder)
                    """;
                recCmd.Transaction = tx;
                var idParam = recCmd.Parameters.Add("@id", SqliteType.Text);
                var groupIdParam = recCmd.Parameters.Add("@groupId", SqliteType.Text);
                var nameParam = recCmd.Parameters.Add("@name", SqliteType.Text);
                var shortNameParam = recCmd.Parameters.Add("@shortName", SqliteType.Text);
                var toEmailsParam = recCmd.Parameters.Add("@toEmails", SqliteType.Text);
                var ccEmailsParam = recCmd.Parameters.Add("@ccEmails", SqliteType.Text);
                var bccEmailsParam = recCmd.Parameters.Add("@bccEmails", SqliteType.Text);
                var remarkParam = recCmd.Parameters.Add("@remark", SqliteType.Text);
                var isSelectedParam = recCmd.Parameters.Add("@isSelected", SqliteType.Integer);
                var variablesParam = recCmd.Parameters.Add("@variablesJson", SqliteType.Text);
                var sortOrderParam = recCmd.Parameters.Add("@sortOrder", SqliteType.Integer);

                foreach (var group in groups)
                {
                    for (int i = 0; i < group.Recipients.Count; i++)
                    {
                        var r = group.Recipients[i];
                        idParam.Value = r.Id;
                        groupIdParam.Value = group.Id;
                        nameParam.Value = r.Name ?? "";
                        shortNameParam.Value = r.ShortName ?? "";
                        toEmailsParam.Value = r.ToEmails ?? "";
                        ccEmailsParam.Value = r.CcEmails ?? "";
                        bccEmailsParam.Value = r.BccEmails ?? "";
                        remarkParam.Value = r.Remark ?? "";
                        isSelectedParam.Value = r.IsSelected ? 1 : 0;
                        variablesParam.Value = JsonSerializer.Serialize(r.Variables, _jsonOptions);
                        sortOrderParam.Value = i;
                        recCmd.ExecuteNonQuery();
                    }
                }
            }

            tx.Commit();
        }
        catch
        {
            tx.Rollback();
            throw;
        }
    }

    public void SaveRecipientVariables(IEnumerable<Recipient> recipients)
    {
        using var conn = CreateConnection();
        conn.Open();
        using var tx = conn.BeginTransaction();

        try
        {
            using var cmd = conn.CreateCommand();
            cmd.Transaction = tx;
            cmd.CommandText = "UPDATE Recipients SET VariablesJson = @variablesJson WHERE Id = @id";
            var idParam = cmd.Parameters.Add("@id", SqliteType.Text);
            var variablesParam = cmd.Parameters.Add("@variablesJson", SqliteType.Text);

            foreach (var recipient in recipients)
            {
                idParam.Value = recipient.Id;
                variablesParam.Value = JsonSerializer.Serialize(recipient.Variables, _jsonOptions);
                cmd.ExecuteNonQuery();
            }

            tx.Commit();
        }
        catch
        {
            tx.Rollback();
            throw;
        }
    }

    public AppSettings LoadSettings()
    {
        using var conn = CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT SettingsJson FROM AppSettings WHERE Id = 1";
        var json = cmd.ExecuteScalar() as string;

        if (string.IsNullOrEmpty(json)) return new AppSettings();
        return JsonSerializer.Deserialize<AppSettings>(json, _jsonOptions) ?? new AppSettings();
    }

    // P-01: Async overload
    public async Task<AppSettings> LoadSettingsAsync()
    {
        await using var conn = CreateConnection();
        await conn.OpenAsync();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT SettingsJson FROM AppSettings WHERE Id = 1";
        var json = await cmd.ExecuteScalarAsync() as string;

        if (string.IsNullOrEmpty(json)) return new AppSettings();
        return JsonSerializer.Deserialize<AppSettings>(json, _jsonOptions) ?? new AppSettings();
    }

    public void SaveSettings(AppSettings settings)
    {
        using var conn = CreateConnection();
        conn.Open();

        var json = JsonSerializer.Serialize(settings, _jsonOptions);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO AppSettings (Id, SettingsJson) VALUES (1, @json)
            ON CONFLICT(Id) DO UPDATE SET SettingsJson = @json
            """;
        cmd.Parameters.AddWithValue("@json", json);
        cmd.ExecuteNonQuery();
    }

    public AttachmentConfig LoadAttachmentConfig()
    {
        using var conn = CreateConnection();
        conn.Open();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT ConfigJson FROM AttachmentConfig WHERE Id = 1";
        var json = cmd.ExecuteScalar() as string;

        if (string.IsNullOrEmpty(json)) return new AttachmentConfig();
        return JsonSerializer.Deserialize<AttachmentConfig>(json, _jsonOptions) ?? new AttachmentConfig();
    }

    public void SaveAttachmentConfig(AttachmentConfig config)
    {
        using var conn = CreateConnection();
        conn.Open();

        var json = JsonSerializer.Serialize(config, _jsonOptions);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO AttachmentConfig (Id, ConfigJson) VALUES (1, @json)
            ON CONFLICT(Id) DO UPDATE SET ConfigJson = @json
            """;
        cmd.Parameters.AddWithValue("@json", json);
        cmd.ExecuteNonQuery();
    }

    public List<MailTemplate> LoadTemplates()
    {
        using var conn = CreateConnection();
        conn.Open();

        var templates = new List<MailTemplate>();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT Id, Name, Subject, Body, CreatedAt FROM MailTemplates ORDER BY rowid";
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            templates.Add(new MailTemplate
            {
                Id = reader.GetString(0),
                Name = reader.GetString(1),
                Subject = reader.GetString(2),
                Body = reader.GetString(3),
                CreatedAt = DateTime.TryParse(reader.GetString(4), out var dt) ? dt : DateTime.Now
            });
        }

        return templates;
    }

    public void SaveTemplates(List<MailTemplate> templates)
    {
        using var conn = CreateConnection();
        conn.Open();
        using var tx = conn.BeginTransaction();

        try
        {
            // P-03: UPSERT instead of DELETE + INSERT
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = """
                    INSERT OR REPLACE INTO MailTemplates (Id, Name, Subject, Body, CreatedAt)
                    VALUES (@id, @name, @subject, @body, @createdAt)
                    """;
                cmd.Transaction = tx;
                var idParam = cmd.Parameters.Add("@id", SqliteType.Text);
                var nameParam = cmd.Parameters.Add("@name", SqliteType.Text);
                var subjectParam = cmd.Parameters.Add("@subject", SqliteType.Text);
                var bodyParam = cmd.Parameters.Add("@body", SqliteType.Text);
                var createdAtParam = cmd.Parameters.Add("@createdAt", SqliteType.Text);

                foreach (var t in templates)
                {
                    idParam.Value = t.Id;
                    nameParam.Value = t.Name;
                    subjectParam.Value = t.Subject;
                    bodyParam.Value = t.Body;
                    createdAtParam.Value = t.CreatedAt.ToString("O");
                    cmd.ExecuteNonQuery();
                }
            }

            // Delete templates no longer in the list
            if (templates.Count > 0)
            {
                using var cmd = conn.CreateCommand();
                var placeholders = string.Join(",", templates.Select((_, i) => $"@t{i}"));
                cmd.CommandText = $"DELETE FROM MailTemplates WHERE Id NOT IN ({placeholders})";
                cmd.Transaction = tx;
                for (int i = 0; i < templates.Count; i++)
                    cmd.Parameters.AddWithValue($"@t{i}", templates[i].Id);
                cmd.ExecuteNonQuery();
            }
            else
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = "DELETE FROM MailTemplates";
                cmd.Transaction = tx;
                cmd.ExecuteNonQuery();
            }

            tx.Commit();
        }
        catch
        {
            tx.Rollback();
            throw;
        }
    }
}
