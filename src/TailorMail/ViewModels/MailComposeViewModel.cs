﻿using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using TailorMail.Models;
using TailorMail.Services;

namespace TailorMail.ViewModels;

/// <summary>
/// 邮件撰写视图模型，负责邮件主题和正文的编辑、模板的保存/加载/删除以及自动保存功能。
/// 邮件正文使用 XAML 格式保存以保留富文本编辑器的格式信息。
/// </summary>
public partial class MailComposeViewModel : ObservableObject
{
    private readonly IDataService _dataService;

    /// <summary>
    /// 获取或设置邮件主题。
    /// </summary>
    [ObservableProperty]
    private string _subject = string.Empty;

    /// <summary>
    /// 获取或设置邮件正文（纯文本或 HTML 格式，用于回退显示）。
    /// </summary>
    [ObservableProperty]
    private string _body = string.Empty;

    /// <summary>
    /// 获取或设置邮件正文的 XAML 序列化内容，用于 WPF 富文本编辑器状态保存和恢复。
    /// </summary>
    [ObservableProperty]
    private string _bodyXaml = string.Empty;

    /// <summary>
    /// 获取或设置邮件模板列表。
    /// </summary>
    [ObservableProperty]
    private ObservableCollection<MailTemplate> _templates = [];

    /// <summary>
    /// 获取或设置当前选中的模板。
    /// </summary>
    [ObservableProperty]
    private MailTemplate? _selectedTemplate;

    /// <summary>
    /// 获取或设置保存模板时的模板名称输入。
    /// </summary>
    [ObservableProperty]
    private string _templateName = string.Empty;

    public MailComposeViewModel(IDataService dataService)
    {
        _dataService = dataService;
        LoadLastContent();
        LoadTemplates();
    }

    /// <summary>
    /// 从应用设置中加载上次编辑的邮件内容，实现会话恢复。
    /// </summary>
    private void LoadLastContent()
    {
        var settings = _dataService.LoadSettings();
        Subject = settings.LastSubject;
        Body = settings.LastBody;
        BodyXaml = settings.LastBodyXaml;
    }

    /// <summary>
    /// 从数据服务加载邮件模板列表。
    /// </summary>
    private void LoadTemplates()
    {
        Templates = new ObservableCollection<MailTemplate>(_dataService.LoadTemplates());
    }

    /// <summary>
    /// 将当前邮件内容保存为新模板。模板中保存的是 XAML 格式的正文。
    /// </summary>
    [RelayCommand]
    private void SaveAsTemplate()
    {
        if (string.IsNullOrWhiteSpace(TemplateName)) return;
        var template = new MailTemplate { Name = TemplateName, Subject = Subject, Body = BodyXaml };
        var list = _dataService.LoadTemplates();
        list.Add(template);
        _dataService.SaveTemplates(list);
        LoadTemplates();
        TemplateName = string.Empty;
    }

    /// <summary>
    /// 加载选中的模板内容到编辑器。
    /// </summary>
    [RelayCommand]
    private void LoadTemplate()
    {
        if (SelectedTemplate == null) return;
        Subject = SelectedTemplate.Subject;
        BodyXaml = SelectedTemplate.Body;
    }

    /// <summary>
    /// 删除选中的模板。
    /// </summary>
    [RelayCommand]
    private void DeleteTemplate()
    {
        if (SelectedTemplate == null) return;
        var list = _dataService.LoadTemplates();
        list.RemoveAll(t => t.Id == SelectedTemplate.Id);
        _dataService.SaveTemplates(list);
        LoadTemplates();
    }

    /// <summary>
    /// 保存当前邮件内容到应用设置，用于下次启动时恢复。
    /// </summary>
    public void SaveCurrent()
    {
        var settings = _dataService.LoadSettings();
        settings.LastSubject = Subject;
        settings.LastBody = Body;
        settings.LastBodyXaml = BodyXaml;
        _dataService.SaveSettings(settings);
    }
}
