using System.Runtime.InteropServices;

namespace TailorMail.Helpers;

/// <summary>
/// 文件夹选择器，通过 Windows Shell COM 接口（IFileOpenDialog）打开原生文件夹选择对话框。
/// 相比 <see cref="System.Windows.Forms.FolderBrowserDialog"/>，提供更现代的文件夹选择体验。
/// </summary>
public static class FolderPicker
{
    /// <summary>
    /// 打开文件夹选择对话框。
    /// </summary>
    /// <param name="description">对话框标题（可选）。</param>
    /// <returns>用户选择的文件夹路径；若取消或失败则返回 null。</returns>
    public static string? PickFolder(string? description = null)
    {
        // IFileOpenDialog 的 CLSID
        var clsid = new Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7");
        // IFileOpenDialog 的 IID
        var iid = new Guid("d57c7288-d4ad-4768-be02-9d969532d960");
        var hr = CoCreateInstance(ref clsid, IntPtr.Zero, 1, ref iid, out var punk);
        if (hr != 0 || punk == IntPtr.Zero)
            return null;

        try
        {
            var dialog = (IFileOpenDialog)Marshal.GetObjectForIUnknown(punk);
            // 设置为文件夹选择模式，仅允许选择文件系统路径
            dialog.SetOptions(FOS.FOS_PICKFOLDERS | FOS.FOS_FORCEFILESYSTEM);

            if (!string.IsNullOrEmpty(description))
                dialog.SetTitle(description!);

            hr = dialog.Show(IntPtr.Zero);
            if (hr != 0)
                return null;

            // 获取用户选择的结果
            dialog.GetResult(out var item);
            item.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, out var path);
            Marshal.ReleaseComObject(item);
            Marshal.ReleaseComObject(dialog);
            return path;
        }
        finally
        {
            Marshal.Release(punk);
        }
    }

    /// <summary>
    /// COM 对象创建函数，用于创建 IFileOpenDialog 实例。
    /// </summary>
    [DllImport("ole32.dll")]
    private static extern int CoCreateInstance(
        ref Guid rclsid, IntPtr pUnkOuter, uint dwClsContext,
        ref Guid riid, out IntPtr ppv);

    /// <summary>
    /// Windows Shell 文件打开对话框 COM 接口。
    /// </summary>
    [ComImport]
    [Guid("d57c7288-d4ad-4768-be02-9d969532d960")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileOpenDialog
    {
        [PreserveSig] int Show(IntPtr parent);
        void SetFileTypes(uint cFileTypes, IntPtr rgFilterSpec);
        void SetFileTypeIndex(uint iFileType);
        void GetFileTypeIndex(out uint piFileType);
        void Advise(IntPtr pfde, out uint pdwCookie);
        void Unadvise(uint dwCookie);
        void SetOptions(FOS fos);
        void GetOptions(out FOS pfos);
        void SetDefaultFolder(IShellItem psi);
        void SetFolder(IShellItem psi);
        void GetFolder(out IShellItem ppsi);
        void GetCurrentSelection(out IShellItem ppsi);
        void SetFileName([MarshalAs(UnmanagedType.LPWStr)] string pszName);
        void GetFileName([MarshalAs(UnmanagedType.LPWStr)] out string pszName);
        void SetTitle([MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel([MarshalAs(UnmanagedType.LPWStr)] string pszText);
        void SetFileNameLabel([MarshalAs(UnmanagedType.LPWStr)] string pszLabel);
        void GetResult(out IShellItem ppsi);
        void AddPlace(IShellItem psi, int fdap);
        void SetDefaultExtension([MarshalAs(UnmanagedType.LPWStr)] string pszDefaultExtension);
        void Close(int hr);
        void SetClientGuid(ref Guid guid);
        void ClearClientData();
        void SetFilter(IntPtr pFilter);
        void GetResults(out IntPtr ppenum);
        void GetSelectedItems(out IntPtr ppsai);
    }

    /// <summary>
    /// Windows Shell 项 COM 接口，用于获取文件夹路径等信息。
    /// </summary>
    [ComImport]
    [Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem
    {
        void BindToHandler(IntPtr pbc, ref Guid bhid, ref Guid riid, out IntPtr ppv);
        void GetParent(out IShellItem ppsi);
        void GetDisplayName(SIGDN sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes(uint sfgaoMask, out uint psfgaoAttribs);
        void Compare(IShellItem psi, uint hint, out int piOrder);
    }

    /// <summary>
    /// Shell 项显示名称格式枚举。
    /// </summary>
    private enum SIGDN : uint
    {
        /// <summary>文件系统路径格式。</summary>
        SIGDN_FILESYSPATH = 0x80058000,
    }

    /// <summary>
    /// 文件对话框选项标志枚举。
    /// </summary>
    [Flags]
    private enum FOS : uint
    {
        /// <summary>选择文件夹模式。</summary>
        FOS_PICKFOLDERS = 0x00000020,

        /// <summary>仅允许选择文件系统路径。</summary>
        FOS_FORCEFILESYSTEM = 0x00000040,
    }
}
