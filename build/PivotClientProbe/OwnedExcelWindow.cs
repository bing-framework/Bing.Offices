using System;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;

namespace Bing.Offices.PivotClientProbe;

/// <summary>
/// 提供按进程标识查找 Excel 原生窗口对象的功能。
/// </summary>
public static class OwnedExcelWindow
{
    /// <summary>
    /// 表示原生 Office 对象模型的窗口对象标识。
    /// </summary>
    private const uint ObjIdNativeOm = 0xFFFFFFF0;

    /// <summary>
    /// 表示 Excel 工作簿窗口的 Win32 窗口类名。
    /// </summary>
    private const string ExcelWindowClassName = "EXCEL7";

    /// <summary>
    /// 表示 IDispatch 自动化接口的固定标识。
    /// </summary>
    private static readonly Guid DispatchInterfaceId = new("00020400-0000-0000-C000-000000000046");

    /// <summary>
    /// 按指定进程标识查找 Excel 窗口并获取其原生对象模型对象。
    /// </summary>
    /// <param name="processId">要查找的 Excel 进程标识。</param>
    /// <returns>找到的原生 Window COM 对象；尚无对应窗口时返回 <see langword="null"/>。</returns>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="processId"/>不是正数。</exception>
    /// <exception cref="COMException">获取原生对象模型对象失败。</exception>
    public static object TryGetWindow(int processId)
    {
        if (processId <= 0)
            throw new ArgumentOutOfRangeException(nameof(processId));

        var targetProcessId = checked((uint)processId);
        object nativeWindow = null;
        Exception callbackException = null;

        bool VisitWindow(IntPtr window, IntPtr lParam)
        {
            try
            {
                if (GetWindowThreadProcessId(window, out var windowProcessId) == 0
                    || windowProcessId != targetProcessId)
                    return true;

                var className = new StringBuilder(256);
                if (GetClassName(window, className, className.Capacity) == 0
                    || !string.Equals(className.ToString(), ExcelWindowClassName, StringComparison.Ordinal))
                    return true;

                var interfaceId = DispatchInterfaceId;
                var result = AccessibleObjectFromWindow(window, ObjIdNativeOm, ref interfaceId, out nativeWindow);
                Marshal.ThrowExceptionForHR(result);
                return false;
            }
            catch (Exception exception)
            {
                callbackException = exception;
                return false;
            }
        }

        bool VisitTopLevelWindow(IntPtr window, IntPtr lParam)
        {
            try
            {
                if (GetWindowThreadProcessId(window, out var windowProcessId) != 0
                    && windowProcessId == targetProcessId
                    && VisitWindow(window, IntPtr.Zero))
                    EnumChildWindows(window, VisitWindow, IntPtr.Zero);
            }
            catch (Exception exception)
            {
                callbackException = exception;
            }

            return nativeWindow == null && callbackException == null;
        }

        EnumWindows(VisitTopLevelWindow, IntPtr.Zero);
        if (callbackException != null)
            ExceptionDispatchInfo.Capture(callbackException).Throw();
        return nativeWindow;
    }

    /// <summary>
    /// 接收 EnumWindows 和 EnumChildWindows 的窗口枚举回调。
    /// </summary>
    /// <param name="window">当前窗口句柄。</param>
    /// <param name="lParam">传递给枚举函数的用户数据。</param>
    /// <returns>继续枚举时返回 true。</returns>
    [return: MarshalAs(UnmanagedType.Bool)]
    private delegate bool WindowEnumerationCallback(IntPtr window, IntPtr lParam);

    /// <summary>
    /// 枚举桌面上的顶层窗口。
    /// </summary>
    /// <param name="callback">用于处理每个顶层窗口的回调。</param>
    /// <param name="lParam">传递给回调的用户数据。</param>
    /// <returns>回调未中止枚举时返回非零值。</returns>
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(WindowEnumerationCallback callback, IntPtr lParam);

    /// <summary>
    /// 枚举指定顶层窗口的所有子窗口。
    /// </summary>
    /// <param name="parentWindow">要枚举其子窗口的顶层窗口句柄。</param>
    /// <param name="callback">用于处理每个子窗口的回调。</param>
    /// <param name="lParam">传递给回调的用户数据。</param>
    /// <returns>回调未中止枚举时返回非零值。</returns>
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumChildWindows(IntPtr parentWindow, WindowEnumerationCallback callback, IntPtr lParam);

    /// <summary>
    /// 获取指定窗口所属线程和进程的标识。
    /// </summary>
    /// <param name="window">要查询的窗口句柄。</param>
    /// <param name="processId">接收窗口所属进程标识的位置。</param>
    /// <returns>窗口所属线程标识。</returns>
    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    /// <summary>
    /// 获取指定窗口的类名。
    /// </summary>
    /// <param name="window">要查询的窗口句柄。</param>
    /// <param name="className">接收窗口类名的缓冲区。</param>
    /// <param name="maximumCount">缓冲区的字符容量。</param>
    /// <returns>写入缓冲区的字符数；失败时返回零。</returns>
    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int GetClassName(IntPtr window, StringBuilder className, int maximumCount);

    /// <summary>
    /// 通过辅助功能对象接口获取窗口的原生对象模型对象。
    /// </summary>
    /// <param name="window">包含目标对象模型的窗口句柄。</param>
    /// <param name="objectId">要请求的窗口对象标识。</param>
    /// <param name="interfaceId">所请求 COM 接口的标识。</param>
    /// <param name="accessibleObject">接收返回的 COM 对象。</param>
    /// <returns>HRESULT；失败值由托管调用方转换为 <see cref="COMException"/>。</returns>
    [DllImport("oleacc.dll", PreserveSig = true)]
    private static extern int AccessibleObjectFromWindow(IntPtr window, uint objectId, ref Guid interfaceId,
        [MarshalAs(UnmanagedType.Interface)] out object accessibleObject);
}
