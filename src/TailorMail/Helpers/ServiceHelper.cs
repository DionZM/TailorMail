﻿using Microsoft.Extensions.DependencyInjection;

namespace TailorMail.Helpers;

/// <summary>
/// 服务定位器辅助类，提供从全局 DI 容器获取服务实例的快捷方法。
/// 用于在无法通过构造函数注入获取服务的场景（如 ViewModel 内部创建其他 ViewModel）。
/// </summary>
public static class ServiceHelper
{
    /// <summary>
    /// 从全局 DI 容器获取指定类型的服务实例。
    /// </summary>
    /// <typeparam name="T">要获取的服务类型。</typeparam>
    /// <returns>服务实例。</returns>
    public static T GetRequiredService<T>() where T : notnull
    {
        return (T)App.Services.GetRequiredService(typeof(T));
    }
}
