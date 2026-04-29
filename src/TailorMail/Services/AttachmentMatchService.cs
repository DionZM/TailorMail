﻿using System.Collections.Generic;
using TailorMail.Models;

namespace TailorMail.Services;

/// <summary>
/// 附件自动匹配服务，根据收件人名称在指定目录中自动查找匹配的附件文件。
/// 匹配规则：文件名（不含扩展名）包含收件人的名称或简称（不区分大小写），
/// 且该文件不在公共附件列表中，则视为该收件人的专属附件。
/// </summary>
public class AttachmentMatchService
{
    /// <summary>
    /// 根据收件人信息在指定目录中匹配附件文件。
    /// </summary>
    /// <param name="directory">待搜索的目录路径，仅搜索该目录下的顶层文件（不递归子目录）。</param>
    /// <param name="recipients">收件人列表，使用每个收件人的 Name 和 ShortName 进行匹配。</param>
    /// <param name="commonAttachmentFiles">公共附件文件路径列表，这些文件会被排除在匹配结果之外。</param>
    /// <returns>
    /// 字典，键为收件人 ID，值为该收件人匹配到的文件路径列表。
    /// 仅包含有匹配结果的收件人。
    /// </returns>
    public Dictionary<string, List<string>> MatchFilesByRecipient(
        string directory,
        List<Recipient> recipients,
        IEnumerable<string> commonAttachmentFiles)
    {
        var result = new Dictionary<string, List<string>>();

        // 获取目录下所有文件（仅顶层，不递归）
        var allFiles = System.IO.Directory.GetFiles(directory, "*.*", System.IO.SearchOption.TopDirectoryOnly);

        // 提取公共附件的文件名集合，用于排除
        var commonFileNames = new HashSet<string>(
            commonAttachmentFiles.Select(f => System.IO.Path.GetFileName(f) ?? string.Empty)
                .Where(f => !string.IsNullOrEmpty(f)),
            StringComparer.OrdinalIgnoreCase);

        // 过滤掉公共附件，得到待匹配的文件列表
        var remainingFiles = allFiles
            .Where(f => !commonFileNames.Contains(System.IO.Path.GetFileName(f)))
            .ToList();

        // 遍历每个收件人，用名称和简称匹配文件名
        foreach (var recipient in recipients)
        {
            var matched = new List<string>();

            // 收集该收件人所有可用的匹配关键字（名称和简称，过滤空白）
            var namesToMatch = new[] { recipient.Name, recipient.ShortName }
                .Where(n => !string.IsNullOrWhiteSpace(n))
                .ToList();

            // 检查每个文件：文件名（不含扩展名）是否包含收件人的名称或简称
            foreach (var filePath in remainingFiles)
            {
                var fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);
                if (namesToMatch.Any(name => fileName.Contains(name, StringComparison.OrdinalIgnoreCase)))
                {
                    matched.Add(filePath);
                }
            }

            if (matched.Count > 0)
                result[recipient.Id] = matched;
        }

        return result;
    }
}
