using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using Microsoft.Win32.SafeHandles;

if (!OperatingSystem.IsWindows() || args.Length < 5)
{
    Console.Error.WriteLine(
        "用法: Bing.Offices.WindowsJobRunner <输出 JSON> <内存 GiB> <CPU 等效核心数> <可执行文件> [参数...]。");
    return 2;
}

var artifactPath = Path.GetFullPath(args[0]);
if (!int.TryParse(args[1], out var memoryGiB) || memoryGiB < 1)
    throw new ArgumentOutOfRangeException(nameof(args), "内存上限必须是正整数 GiB。");
if (!int.TryParse(args[2], out var cpuCount) || cpuCount < 1)
    throw new ArgumentOutOfRangeException(nameof(args), "CPU 等效核心数必须是正整数。");

var childPath = args[3];
var childArguments = args.Skip(4).ToArray();
Directory.CreateDirectory(Path.GetDirectoryName(artifactPath)!);
var timeoutMs = ReadPositiveEnvironmentInt("BING_OFFICES_JOB_TIMEOUT_MS");
var isolateTemp = ReadBooleanEnvironment("BING_OFFICES_JOB_ISOLATE_TEMP");
var isolatedTempDirectory = isolateTemp
    ? Path.Combine(Path.GetTempPath(), $"bing-offices-job-{Guid.NewGuid():N}")
    : null;
if (isolatedTempDirectory != null)
    Directory.CreateDirectory(isolatedTempDirectory);

using var job = NativeMethods.CreateJobObject(IntPtr.Zero, null);
if (job.IsInvalid)
    throw new Win32Exception(Marshal.GetLastWin32Error(), "无法创建 Windows Job Object。");

var hostProcessorCount = Environment.ProcessorCount;
var cpuRate = Math.Clamp((uint)Math.Ceiling(cpuCount * 10000d / hostProcessorCount), 1u, 10000u);
var memoryLimitBytes = checked((ulong)memoryGiB * 1024UL * 1024UL * 1024UL);
NativeMethods.SetJobInformation(job, new NativeMethods.JobObjectExtendedLimitInformation
{
    BasicLimitInformation = new NativeMethods.JobObjectBasicLimitInformation
    {
        LimitFlags = NativeMethods.JobObjectLimitJobMemory | NativeMethods.JobObjectLimitProcessMemory
    },
    JobMemoryLimit = new UIntPtr(memoryLimitBytes),
    ProcessMemoryLimit = new UIntPtr(memoryLimitBytes)
}, NativeMethods.JobObjectExtendedLimitInformationClass);
NativeMethods.SetJobInformation(job, new NativeMethods.JobObjectCpuRateControlInformation
{
    ControlFlags = NativeMethods.JobObjectCpuRateControlEnable | NativeMethods.JobObjectCpuRateControlHardCap,
    CpuRate = cpuRate
}, NativeMethods.JobObjectCpuRateControlInformationClass);

var startInfo = new ProcessStartInfo
{
    FileName = childPath,
    UseShellExecute = false,
    RedirectStandardOutput = true,
    RedirectStandardError = true,
    CreateNoWindow = true
};
if (isolatedTempDirectory != null)
{
    startInfo.Environment["TEMP"] = isolatedTempDirectory;
    startInfo.Environment["TMP"] = isolatedTempDirectory;
}
foreach (var argument in childArguments)
    startInfo.ArgumentList.Add(argument);

using var process = Process.Start(startInfo)
    ?? throw new InvalidOperationException("无法启动受限子进程。");
if (!NativeMethods.AssignProcessToJobObject(job, process.Handle))
    throw new Win32Exception(Marshal.GetLastWin32Error(), "无法将子进程加入 Windows Job Object。");

var outputTask = process.StandardOutput.ReadToEndAsync();
var errorTask = process.StandardError.ReadToEndAsync();
var timedOut = false;
var exitTask = process.WaitForExitAsync();
if (timeoutMs.HasValue)
{
    var completedTask = await Task.WhenAny(exitTask, Task.Delay(timeoutMs.Value)).ConfigureAwait(false);
    if (completedTask != exitTask)
    {
        timedOut = true;
        try
        {
            process.Kill(entireProcessTree: true);
        }
        catch (Exception exception) when (exception is InvalidOperationException or Win32Exception)
        {
            // 子进程可能在超时竞争窗口中已退出；仍等待退出任务收集完整日志。
        }
    }
}
await exitTask.ConfigureAwait(false);
var output = await outputTask.ConfigureAwait(false);
var error = await errorTask.ConfigureAwait(false);
process.Refresh();

var tempFilesAfterTermination = 0;
long tempBytesAfterTermination = 0;
var tempCleanupStatus = isolatedTempDirectory == null ? "not-enabled" : "pending";
if (isolatedTempDirectory != null)
{
    (tempFilesAfterTermination, tempBytesAfterTermination) = CountFiles(isolatedTempDirectory);
    try
    {
        Directory.Delete(isolatedTempDirectory, recursive: true);
        tempCleanupStatus = "deleted";
    }
    catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
    {
        tempCleanupStatus = $"failed:{exception.GetType().Name}";
    }
}

var peakWorkingSetBytes = 0L;
try
{
    peakWorkingSetBytes = process.PeakWorkingSet64;
}
catch (InvalidOperationException)
{
    // 快速退出的子进程可能已经释放进程查询句柄；此时保留零值而不丢失退出码和日志。
}
var jobLimits = NativeMethods.GetJobInformation<NativeMethods.JobObjectExtendedLimitInformation>(job,
    NativeMethods.JobObjectExtendedLimitInformationClass);

var result = new
{
    schema = 1,
    kind = "windows-job-resource-runner",
    generatedUtc = DateTimeOffset.UtcNow,
    status = timedOut ? "runner-timeout" : process.ExitCode == 0 ? "passed" : "child-failed",
    childPath,
    childArguments,
    exitCode = process.ExitCode,
    hostProcessorCount,
    requestedCpuEquivalent = cpuCount,
    cpuRateUnitsPer10000 = cpuRate,
    requestedMemoryLimitBytes = memoryLimitBytes,
    assignedToJob = true,
    childExitCode = process.ExitCode,
    runnerTimedOut = timedOut,
    timeoutMs,
    isolatedTempDirectory,
    tempFilesAfterTermination,
    tempBytesAfterTermination,
    tempCleanupStatus,
    peakWorkingSetBytes,
    peakProcessMemoryBytes = jobLimits.PeakProcessMemoryUsed.ToUInt64(),
    peakJobMemoryBytes = jobLimits.PeakJobMemoryUsed.ToUInt64(),
    stdout = output,
    stderr = error
};
File.WriteAllText(artifactPath, JsonSerializer.Serialize(result), new UTF8Encoding(false));
Console.WriteLine(JsonSerializer.Serialize(result));
return timedOut ? 124 : process.ExitCode;

static int? ReadPositiveEnvironmentInt(string name)
{
    var value = Environment.GetEnvironmentVariable(name);
    return int.TryParse(value, out var parsed) && parsed > 0 ? parsed : null;
}

static bool ReadBooleanEnvironment(string name)
{
    var value = Environment.GetEnvironmentVariable(name);
    return string.Equals(value, "1", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
        || string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase);
}

static (int fileCount, long totalBytes) CountFiles(string directory)
{
    if (!Directory.Exists(directory))
        return (0, 0);

    var files = Directory.EnumerateFiles(directory, "*", SearchOption.AllDirectories);
    var count = 0;
    long bytes = 0;
    foreach (var file in files)
    {
        count++;
        try
        {
            bytes += new FileInfo(file).Length;
        }
        catch (IOException)
        {
            // 文件可能在子进程终止时刚被删除，保留可复核的文件计数。
        }
    }
    return (count, bytes);
}

/// <summary>
/// 封装 Windows Job Object 的原生调用和资源限制结构。
/// </summary>
internal static class NativeMethods
{
    /// <summary>
    /// Job Object 扩展限制信息类编号。
    /// </summary>
    internal const int JobObjectExtendedLimitInformationClass = 9;

    /// <summary>
    /// Job Object CPU 速率控制信息类编号。
    /// </summary>
    internal const int JobObjectCpuRateControlInformationClass = 15;

    /// <summary>
    /// 限制子进程内存的标志位。
    /// </summary>
    internal const uint JobObjectLimitProcessMemory = 0x100;

    /// <summary>
    /// 限制 Job Object 总内存的标志位。
    /// </summary>
    internal const uint JobObjectLimitJobMemory = 0x200;

    /// <summary>
    /// 启用 CPU 速率控制的标志位。
    /// </summary>
    internal const uint JobObjectCpuRateControlEnable = 0x1;

    /// <summary>
    /// 启用 CPU 硬上限的标志位。
    /// </summary>
    internal const uint JobObjectCpuRateControlHardCap = 0x4;

    /// <summary>
    /// 创建 Windows Job Object。
    /// </summary>
    /// <param name="attributes">原生安全属性；此运行器传入空指针。</param>
    /// <param name="name">Job Object 名称；null 表示不命名。</param>
    /// <returns>新建 Job Object 的安全句柄。</returns>
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    internal static extern SafeFileHandle CreateJobObject(IntPtr attributes, string? name);

    /// <summary>
    /// 将进程加入指定 Job Object。
    /// </summary>
    /// <param name="job">目标 Job Object 句柄。</param>
    /// <param name="process">目标进程句柄。</param>
    /// <returns>成功加入时为 true，否则为 false。</returns>
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool AssignProcessToJobObject(SafeFileHandle job, IntPtr process);

    /// <summary>
    /// 设置 Job Object 的资源限制信息。
    /// </summary>
    /// <param name="job">目标 Job Object 句柄。</param>
    /// <param name="infoClass">信息类编号。</param>
    /// <param name="information">指向信息结构的非托管指针。</param>
    /// <param name="informationLength">信息结构长度。</param>
    /// <returns>设置成功时为 true，否则为 false。</returns>
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetInformationJobObject(SafeFileHandle job, int infoClass,
        IntPtr information, uint informationLength);

    /// <summary>
    /// 查询 Job Object 的资源统计信息。
    /// </summary>
    /// <param name="job">目标 Job Object 句柄。</param>
    /// <param name="infoClass">信息类编号。</param>
    /// <param name="information">接收信息结构的非托管指针。</param>
    /// <param name="informationLength">信息结构缓冲区长度。</param>
    /// <param name="returnLength">实际写入的信息长度。</param>
    /// <returns>查询成功时为 true，否则为 false。</returns>
    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool QueryInformationJobObject(SafeFileHandle job, int infoClass,
        IntPtr information, uint informationLength, out uint returnLength);

    /// <summary>
    /// 将托管结构写入 Job Object 信息类。
    /// </summary>
    /// <typeparam name="T">要写入的资源限制结构类型。</typeparam>
    /// <param name="job">目标 Job Object 句柄。</param>
    /// <param name="information">要写入的信息结构。</param>
    /// <param name="infoClass">信息类编号。</param>
    internal static void SetJobInformation<T>(SafeFileHandle job, T information, int infoClass)
        where T : struct
    {
        var size = Marshal.SizeOf<T>();
        var pointer = Marshal.AllocHGlobal(size);
        try
        {
            Marshal.StructureToPtr(information, pointer, false);
            if (!SetInformationJobObject(job, infoClass, pointer, checked((uint)size)))
                throw new Win32Exception(Marshal.GetLastWin32Error(),
                    $"无法设置 Windows Job Object 信息类 {infoClass}。");
        }
        finally
        {
            Marshal.DestroyStructure<T>(pointer);
            Marshal.FreeHGlobal(pointer);
        }
    }

    /// <summary>
    /// 从 Job Object 信息类读取托管结构。
    /// </summary>
    /// <typeparam name="T">要读取的资源统计结构类型。</typeparam>
    /// <param name="job">目标 Job Object 句柄。</param>
    /// <param name="infoClass">信息类编号。</param>
    /// <returns>读取到的信息结构。</returns>
    internal static T GetJobInformation<T>(SafeFileHandle job, int infoClass)
        where T : struct
    {
        var size = Marshal.SizeOf<T>();
        var pointer = Marshal.AllocHGlobal(size);
        try
        {
            if (!QueryInformationJobObject(job, infoClass, pointer, checked((uint)size), out _))
                throw new Win32Exception(Marshal.GetLastWin32Error(),
                    $"无法读取 Windows Job Object 信息类 {infoClass}。");
            return Marshal.PtrToStructure<T>(pointer);
        }
        finally
        {
            Marshal.FreeHGlobal(pointer);
        }
    }

    /// <summary>
    /// 描述 Job Object 的基础资源限制信息。
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    internal struct JobObjectBasicLimitInformation
    {
        /// <summary>
        /// 每个进程的用户态时间限制。
        /// </summary>
        internal long PerProcessUserTimeLimit;
        /// <summary>
        /// Job Object 的用户态时间限制。
        /// </summary>
        internal long PerJobUserTimeLimit;
        /// <summary>
        /// 已启用的限制标志位。
        /// </summary>
        internal uint LimitFlags;
        /// <summary>
        /// 最小工作集大小。
        /// </summary>
        internal UIntPtr MinimumWorkingSetSize;
        /// <summary>
        /// 最大工作集大小。
        /// </summary>
        internal UIntPtr MaximumWorkingSetSize;
        /// <summary>
        /// Job Object 允许的活动进程数。
        /// </summary>
        internal uint ActiveProcessLimit;
        /// <summary>
        /// 进程处理器亲和性掩码。
        /// </summary>
        internal UIntPtr Affinity;
        /// <summary>
        /// 进程优先级类别。
        /// </summary>
        internal uint PriorityClass;
        /// <summary>
        /// 调度类别。
        /// </summary>
        internal uint SchedulingClass;
    }

    /// <summary>
    /// 描述 Job Object 的 IO 计数器。
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    internal struct IoCounters
    {
        /// <summary>
        /// 读取操作次数。
        /// </summary>
        internal ulong ReadOperationCount;
        /// <summary>
        /// 写入操作次数。
        /// </summary>
        internal ulong WriteOperationCount;
        /// <summary>
        /// 其他操作次数。
        /// </summary>
        internal ulong OtherOperationCount;
        /// <summary>
        /// 读取传输字节数。
        /// </summary>
        internal ulong ReadTransferCount;
        /// <summary>
        /// 写入传输字节数。
        /// </summary>
        internal ulong WriteTransferCount;
        /// <summary>
        /// 其他传输字节数。
        /// </summary>
        internal ulong OtherTransferCount;
    }

    /// <summary>
    /// 描述 Job Object 的扩展限制和峰值内存信息。
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    internal struct JobObjectExtendedLimitInformation
    {
        /// <summary>
        /// 基础资源限制信息。
        /// </summary>
        internal JobObjectBasicLimitInformation BasicLimitInformation;
        /// <summary>
        /// Job Object 的 IO 计数器。
        /// </summary>
        internal IoCounters IoInfo;
        /// <summary>
        /// 每个进程的内存上限。
        /// </summary>
        internal UIntPtr ProcessMemoryLimit;
        /// <summary>
        /// Job Object 的内存上限。
        /// </summary>
        internal UIntPtr JobMemoryLimit;
        /// <summary>
        /// 进程峰值内存使用量。
        /// </summary>
        internal UIntPtr PeakProcessMemoryUsed;
        /// <summary>
        /// Job Object 峰值内存使用量。
        /// </summary>
        internal UIntPtr PeakJobMemoryUsed;
    }

    /// <summary>
    /// 描述 Job Object 的 CPU 速率控制信息。
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    internal struct JobObjectCpuRateControlInformation
    {
        /// <summary>
        /// CPU 速率控制标志位。
        /// </summary>
        internal uint ControlFlags;
        /// <summary>
        /// 以万分比表示的 CPU 速率上限。
        /// </summary>
        internal uint CpuRate;
    }
}
