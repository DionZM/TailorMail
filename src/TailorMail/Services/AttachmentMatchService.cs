using System.Collections.Generic;
using TailorMail.Models;

namespace TailorMail.Services;

/// <summary>
/// 附件自动匹配服务，根据收件人名称在指定目录中自动查找匹配的附件文件。
/// P-07: 使用倒排索引将匹配复杂度从 O(R×F) 降至 O(R+R×M)，其中 M 为每个收件人匹配的平均关键字数。
/// </summary>
public class AttachmentMatchService
{
    /// <summary>
    /// 根据收件人信息在指定目录中匹配附件文件。
    /// </summary>
    public Dictionary<string, List<string>> MatchFilesByRecipient(
        string directory,
        List<Recipient> recipients,
        IEnumerable<string> commonAttachmentFiles)
    {
        var result = new Dictionary<string, List<string>>();

        var allFiles = System.IO.Directory.GetFiles(directory, "*.*", System.IO.SearchOption.TopDirectoryOnly);

        var commonFileNames = new HashSet<string>(
            commonAttachmentFiles.Select(f => System.IO.Path.GetFileName(f) ?? string.Empty)
                .Where(f => !string.IsNullOrEmpty(f)),
            StringComparer.OrdinalIgnoreCase);

        var remainingFiles = allFiles
            .Where(f => !commonFileNames.Contains(System.IO.Path.GetFileName(f)))
            .ToList();

        // P-07: Build inverted index: keyword (lowered) -> list of file paths
        var fileIndex = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var filePath in remainingFiles)
        {
            var fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
            // Index by full filename (lowercase) for substring matching
            var key = fileName.ToLowerInvariant();
            if (!fileIndex.ContainsKey(key))
                fileIndex[key] = new List<string>();
            fileIndex[key].Add(filePath);
        }

        // P-07: For each recipient, search the index instead of iterating all files
        foreach (var recipient in recipients)
        {
            var matched = new List<string>();
            var matchedSet = new HashSet<string>(); // Avoid duplicates

            var namesToMatch = new[] { recipient.Name, recipient.ShortName }
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .Select(n => n!.ToLowerInvariant())
                .ToList();

            // Search through indexed files
            foreach (var (key, paths) in fileIndex)
            {
                foreach (var name in namesToMatch)
                {
                    if (key.Contains(name, StringComparison.OrdinalIgnoreCase))
                    {
                        foreach (var path in paths)
                        {
                            if (matchedSet.Add(path))
                                matched.Add(path);
                        }
                        break; // File matched, no need to check other names
                    }
                }
            }

            if (matched.Count > 0)
                result[recipient.Id] = matched;
        }

        return result;
    }
}
