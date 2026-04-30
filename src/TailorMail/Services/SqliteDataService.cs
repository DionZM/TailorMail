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

    private readonly string _dbPath;
    private readonly string _connectionString;

    public SqliteDataService()
    {
        var dataDir = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data");
        System.IO.Directory.CreateDirectory(dataDir);
        _dbPath = System.IO.Path.Combine(dataDir, "tailormail.db");
        _connectionString = $"Data Source={_dbPath}";
        InitializeDatabase();
        MigrateFromJsonIfNeeded();
    }

    private SqliteConnection CreateConnection() => new(_connectionString);

    private void InitializeDatabase()
    {
        using var conn = CreateConnection();
        conn.Open();

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
    }

    private void MigrateFromJsonIfNeeded()
    {
        var dataDir = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data");
        var migrationFlag = System.IO.Path.Combine(dataDir, ".migrated_to_sqlite");

        if (System.IO.File.Exists(migrationFlag)) return;

        var jsonService = new JsonDataService();

        try
        {
            var groups = jsonService.LoadRecipientGroups();
            if (groups.Count > 0 && groups.Any(g => g.Recipients.Count > 0))
            {
                SaveRecipientGroups(groups);
            }
        }
        catch { }

        try
        {
            var settings = jsonService.LoadSettings();
            if (!string.IsNullOrEmpty(settings.LastSubject) || !string.IsNullOrEmpty(settings.Smtp.Host))
            {
                SaveSettings(settings);
            }
        }
        catch { }

        try
        {
            var config = jsonService.LoadAttachmentConfig();
            if (config.CommonAttachments.Count > 0 || config.RecipientAttachments.Count > 0)
            {
                SaveAttachmentConfig(config);
            }
        }
        catch { }

        try
        {
            var templates = jsonService.LoadTemplates();
            if (templates.Count > 0)
            {
                SaveTemplates(templates);
            }
        }
        catch { }

        System.IO.File.WriteAllText(migrationFlag, DateTime.Now.ToString("O"));
    }

    public List<RecipientGroup> LoadRecipientGroups()
    {
        using var conn = CreateConnection();
        conn.Open();

        var groups = new List<RecipientGroup>();

        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "SELECT Id, Name FROM RecipientGroups ORDER BY rowid";
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

        foreach (var group in groups)
        {
            group.Recipients = LoadRecipientsByGroup(conn, group.Id);
        }

        if (groups.Count == 0)
        {
            groups.Add(new RecipientGroup { Name = "默认分组" });
        }

        return groups;
    }

    private List<Recipient> LoadRecipientsByGroup(SqliteConnection conn, string groupId)
    {
        var recipients = new List<Recipient>();

        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            SELECT Id, Name, ShortName, ToEmails, CcEmails, BccEmails, Remark, IsSelected, VariablesJson
            FROM Recipients WHERE GroupId = @groupId ORDER BY SortOrder, rowid
            """;
        cmd.Parameters.AddWithValue("@groupId", groupId);

        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            var variablesJson = reader.GetString(8);
            Dictionary<string, string> variables = [];
            if (!string.IsNullOrEmpty(variablesJson) && variablesJson != "{}")
            {
                try
                {
                    variables = JsonSerializer.Deserialize<Dictionary<string, string>>(variablesJson, _jsonOptions) ?? [];
                }
                catch { }
            }

            recipients.Add(new Recipient
            {
                Id = reader.GetString(0),
                Name = reader.GetString(1),
                ShortName = reader.GetString(2),
                ToEmails = reader.GetString(3),
                CcEmails = reader.GetString(4),
                BccEmails = reader.GetString(5),
                Remark = reader.GetString(6),
                IsSelected = reader.GetInt32(7) == 1,
                Variables = variables
            });
        }

        return recipients;
    }

    public void SaveRecipientGroups(List<RecipientGroup> groups)
    {
        using var conn = CreateConnection();
        conn.Open();
        using var tx = conn.BeginTransaction();

        try
        {
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "DELETE FROM Recipients; DELETE FROM RecipientGroups";
                cmd.Transaction = tx;
                cmd.ExecuteNonQuery();
            }

            foreach (var group in groups)
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "INSERT INTO RecipientGroups (Id, Name) VALUES (@id, @name)";
                    cmd.Transaction = tx;
                    cmd.Parameters.AddWithValue("@id", group.Id);
                    cmd.Parameters.AddWithValue("@name", group.Name);
                    cmd.ExecuteNonQuery();
                }

                for (int i = 0; i < group.Recipients.Count; i++)
                {
                    var r = group.Recipients[i];
                    using var cmd = conn.CreateCommand();
                    cmd.CommandText = """
                        INSERT INTO Recipients (Id, GroupId, Name, ShortName, ToEmails, CcEmails, BccEmails, Remark, IsSelected, VariablesJson, SortOrder)
                        VALUES (@id, @groupId, @name, @shortName, @toEmails, @ccEmails, @bccEmails, @remark, @isSelected, @variablesJson, @sortOrder)
                        """;
                    cmd.Transaction = tx;
                    cmd.Parameters.AddWithValue("@id", r.Id);
                    cmd.Parameters.AddWithValue("@groupId", group.Id);
                    cmd.Parameters.AddWithValue("@name", r.Name ?? "");
                    cmd.Parameters.AddWithValue("@shortName", r.ShortName ?? "");
                    cmd.Parameters.AddWithValue("@toEmails", r.ToEmails ?? "");
                    cmd.Parameters.AddWithValue("@ccEmails", r.CcEmails ?? "");
                    cmd.Parameters.AddWithValue("@bccEmails", r.BccEmails ?? "");
                    cmd.Parameters.AddWithValue("@remark", r.Remark ?? "");
                    cmd.Parameters.AddWithValue("@isSelected", r.IsSelected ? 1 : 0);
                    cmd.Parameters.AddWithValue("@variablesJson", JsonSerializer.Serialize(r.Variables, _jsonOptions));
                    cmd.Parameters.AddWithValue("@sortOrder", i);
                    cmd.ExecuteNonQuery();
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
            using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = "DELETE FROM MailTemplates";
                cmd.Transaction = tx;
                cmd.ExecuteNonQuery();
            }

            foreach (var t in templates)
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = """
                    INSERT INTO MailTemplates (Id, Name, Subject, Body, CreatedAt)
                    VALUES (@id, @name, @subject, @body, @createdAt)
                    """;
                cmd.Transaction = tx;
                cmd.Parameters.AddWithValue("@id", t.Id);
                cmd.Parameters.AddWithValue("@name", t.Name);
                cmd.Parameters.AddWithValue("@subject", t.Subject);
                cmd.Parameters.AddWithValue("@body", t.Body);
                cmd.Parameters.AddWithValue("@createdAt", t.CreatedAt.ToString("O"));
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
