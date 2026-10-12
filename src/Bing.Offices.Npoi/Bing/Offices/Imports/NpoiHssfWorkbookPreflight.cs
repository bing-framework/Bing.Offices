using System;
using System.IO;
using System.Threading;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using NPOI.HSSF.EventUserModel;
using NPOI.HSSF.Record;
using NPOI.POIFS.FileSystem;

namespace Bing.Offices.Imports;

/// <summary>
/// 在创建 HSSF 工作簿对象模型前检查工作簿结构预算。
/// </summary>
internal static class NpoiHssfWorkbookPreflight
{
    /// <summary>
    /// HSSF OLE2 容器的文件签名。
    /// </summary>
    private static readonly byte[] Ole2Signature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

    /// <summary>
    /// 按配置的工作表、列和物理单元格限制扫描 HSSF 记录。
    /// </summary>
    /// <param name="source">可定位的工作簿流。</param>
    /// <param name="limits">工作簿资源限制。</param>
    /// <param name="cancellationToken">用于取消记录扫描的令牌。</param>
    internal static void Validate(Stream source, ExcelResourceLimits limits,
        CancellationToken cancellationToken)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));
        if (limits == null || !source.CanRead || !source.CanSeek
            || source.Length < Ole2Signature.Length
            || (!limits.MaxSheets.HasValue && !limits.MaxColumnsPerSheet.HasValue
                && !limits.MaxCells.HasValue))
            return;

        try
        {
            source.Position = 0;
            if (!HasOle2Signature(source))
                return;

            source.Position = 0;
            var listener = new Listener(limits, cancellationToken);
            var request = new HSSFRequest();
            request.AddListenerForAllRecords(listener);
            var fileSystem = new POIFSFileSystem(new NpoiNonDisposingStream(source,
                cancellationToken, flushUnderlying: false));
            try
            {
                new HSSFEventFactory().ProcessWorkbookEvents(request, fileSystem);
            }
            finally
            {
                fileSystem.Close();
            }
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch (BingOfficesException)
        {
            throw;
        }
        catch (Exception exception) when (exception is not OutOfMemoryException
            && exception is not StackOverflowException)
        {
            throw new BingOfficesImportException("HSSF BIFF 资源预检失败。", exception, "NPOI",
                BingOfficesStage.Preflight);
        }
        finally
        {
            source.Position = 0;
        }
    }

    /// <summary>
    /// 检查输入是否以 OLE2 容器签名开头。
    /// </summary>
    private static bool HasOle2Signature(Stream source)
    {
        var signature = new byte[Ole2Signature.Length];
        var totalRead = 0;
        while (totalRead < signature.Length)
        {
            var read = source.Read(signature, totalRead, signature.Length - totalRead);
            if (read == 0)
                return false;
            totalRead += read;
        }

        for (var index = 0; index < Ole2Signature.Length; index++)
        {
            if (signature[index] != Ole2Signature[index])
                return false;
        }
        return true;
    }

    /// <summary>
    /// 收集 HSSF 记录并检查物理工作簿大小。
    /// </summary>
    private sealed class Listener : IHSSFListener
    {
        /// <summary>
        /// 当前工作簿使用的资源限制。
        /// </summary>
        private readonly ExcelResourceLimits _limits;
        /// <summary>
        /// 用于中断 HSSF 记录扫描的取消令牌。
        /// </summary>
        private readonly CancellationToken _cancellationToken;
        /// <summary>
        /// 已读取的工作表数量。
        /// </summary>
        private int _sheetCount;
        /// <summary>
        /// 已读取的物理单元格数量。
        /// </summary>
        private long _cellCount;
        /// <summary>
        /// 指示当前记录是否属于工作表区域。
        /// </summary>
        private bool _inWorksheet;

        /// <summary>
        /// 初始化一个 <see cref="Listener" /> 类型的实例。
        /// </summary>
        /// <param name="limits">工作簿资源限制。</param>
        /// <param name="cancellationToken">用于取消记录扫描的令牌。</param>
        internal Listener(ExcelResourceLimits limits, CancellationToken cancellationToken)
        {
            _limits = limits;
            _cancellationToken = cancellationToken;
        }

        /// <inheritdoc />
        public void ProcessRecord(Record record)
        {
            _cancellationToken.ThrowIfCancellationRequested();
            if (record is BoundSheetRecord)
            {
                _sheetCount++;
                if (_limits.MaxSheets.HasValue && _sheetCount > _limits.MaxSheets.Value)
                    Resource($"HSSF 工作表数量超过限制: {_limits.MaxSheets.Value}");
                return;
            }

            if (record is BOFRecord bof)
            {
                if (bof.Type == BOFRecordType.Worksheet)
                    _inWorksheet = true;
                return;
            }
            if (record is EOFRecord)
            {
                _inWorksheet = false;
                return;
            }
            if (!_inWorksheet)
                return;

            if (record is MulRKRecord multipleNumbers)
            {
                CheckColumn(multipleNumbers.LastColumn);
                CheckCells(multipleNumbers.NumColumns);
                return;
            }
            if (record is MulBlankRecord multipleBlanks)
            {
                CheckColumn(multipleBlanks.LastColumn);
                CheckCells(multipleBlanks.NumColumns);
                return;
            }
            if (record is CellValueRecordInterface cell)
            {
                CheckColumn(cell.Column);
                CheckCells(1);
            }
        }

        /// <summary>
        /// 检查单元格所在列是否超过限制。
        /// </summary>
        private void CheckColumn(int zeroBasedColumn)
        {
            if (_limits.MaxColumnsPerSheet.HasValue
                && zeroBasedColumn + 1 > _limits.MaxColumnsPerSheet.Value)
                Resource($"HSSF Sheet 最大使用列超过限制: {_limits.MaxColumnsPerSheet.Value}");
        }

        /// <summary>
        /// 累计物理单元格并检查总量限制。
        /// </summary>
        private void CheckCells(int count)
        {
            if (_limits.MaxCells.HasValue && count > _limits.MaxCells.Value - _cellCount)
                Resource($"HSSF 物理单元格数量超过限制: {_limits.MaxCells.Value}");
            _cellCount += count;
        }

        /// <summary>
        /// 抛出 HSSF 资源限制异常。
        /// </summary>
        private static void Resource(string message) =>
            throw new BingOfficesResourceLimitException(message, provider: "NPOI",
                operation: BingOfficesOperation.Import, stage: BingOfficesStage.Preflight);
    }
}
