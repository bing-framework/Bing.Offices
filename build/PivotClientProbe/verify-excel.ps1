param(
    [Parameter(Mandatory = $true)]
    [string]$WorkbookPath,

    [Parameter(Mandatory = $true)]
    [string]$ResultPath,

    [string]$OwnerPath,
    [ValidateRange(0, 600)][int]$FaultAfterOwnerSeconds = 0,
    [int]$OwnedProcessId = 0,
    [string]$OwnedStartTimeUtc,
    [ValidateRange(0, 600)][int]$FaultBeforeAttachSeconds = 0
)

$ErrorActionPreference = 'Stop'
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)

$script:ComObjects = [System.Collections.Generic.List[object]]::new()
$script:OwnedWorkbooks = [System.Collections.Generic.List[object]]::new()
$script:ExcelApplication = $null
$script:OwnsApplication = $false
$script:Report = [ordered]@{
    Status = 'Running'
    WorkbookPath = [System.IO.Path]::GetFullPath($WorkbookPath)
    ResultPath = [System.IO.Path]::GetFullPath($ResultPath)
    CopyPath = $null
    Excel = $null
    Connection = $null
    OwnerProcessId = $null
    OwnedExcel = $false
    RepairMode = $null
    RefreshOnFileOpen = $null
    InitialSource = $null
    InitialPivot = $null
    InitialObservation = $null
    DisplayAfterSourceEditBeforeRefresh = $null
    RefreshedPivot = $null
    ReopenedSource = $null
    ReopenedPivot = $null
    CleanupErrors = [System.Collections.Generic.List[string]]::new()
    Error = $null
}
$script:Failure = $null

function Add-ComObject {
    <# 记录由本脚本创建或取得的 COM 对象，供 finally 释放。 #>
    param([object]$Value)

    if ($null -ne $Value) {
        foreach ($tracked in $script:ComObjects) {
            if ([object]::ReferenceEquals($tracked, $Value)) {
                return ,$Value
            }
        }
        $script:ComObjects.Add($Value)
    }
    return ,$Value
}

function Write-Utf8Json {
    <# 将进程归属状态写为 UTF-8 无 BOM JSON。 #>
    param([string]$Path, [object]$Value)

    $json = ConvertTo-Json -InputObject $Value -Depth 16
    $json = $json.Replace("`r`n", "`n").Replace("`r", "`n")
    [System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($Path),
        $json + "`n", [System.Text.UTF8Encoding]::new($false))
}

function Get-ComRange {
    <# 获取并跟踪 Excel 单元格区域对象。 #>
    param(
        [Parameter(Mandatory = $true)][object]$Sheet,
        [Parameter(Mandatory = $true)][string]$Address
    )

    return ,(Add-ComObject $Sheet.Range($Address))
}

function ConvertTo-Matrix {
    <# 将 Excel 返回的二维 COM 数组转换为可写入 JSON 的矩阵。 #>
    param([object]$Value)

    if ($null -eq $Value) {
        return ,@()
    }
    if ($Value -isnot [System.Array]) {
        return ,@(@($Value))
    }

    $rowCount = $Value.GetLength(0)
    $columnCount = $Value.GetLength(1)
    $rows = [object[]]::new($rowCount)
    for ($row = 0; $row -lt $rowCount; $row++) {
        $cells = [object[]]::new($columnCount)
        for ($column = 0; $column -lt $columnCount; $column++) {
            $cells[$column] = $Value.GetValue($row + $Value.GetLowerBound(0),
                $column + $Value.GetLowerBound(1))
        }
        $rows[$row] = $cells
    }
    return ,$rows
}

function Assert-EqualValue {
    <# 比较 Excel 单元格值并在不匹配时停止验收。 #>
    param(
        [object]$Actual,
        [object]$Expected,
        [string]$Context
    )

    if ($Expected -is [ValueType]) {
        if ([double]$Actual -ne [double]$Expected) {
            throw "$Context 期望 $Expected，实际为 $Actual。"
        }
        return
    }

    if ([string]$Actual -cne [string]$Expected) {
        throw "$Context 期望 '$Expected'，实际为 '$Actual'。"
    }
}

function Assert-SourceMatrix {
    <# 校验完整来源表区域，确保行序和值均符合夹具。 #>
    param(
        [object[]]$Matrix,
        [double]$NorthAmount,
        [string]$Context
    )

    $expected = [System.Collections.Generic.List[object]]::new()
    $expected.Add([object[]]@('Warehouse', 'Amount'))
    $expected.Add([object[]]@('North', $NorthAmount))
    $expected.Add([object[]]@('South', 20))
    $expected.Add([object[]]@('North', 30))
    if ($Matrix.Count -ne $expected.Count) {
        throw "$Context 行数期望 $($expected.Count)，实际为 $($Matrix.Count)。"
    }
    for ($row = 0; $row -lt $expected.Count; $row++) {
        if ($Matrix[$row].Count -ne 2) {
            throw "$Context 第 $($row + 1) 行列数不是 2。"
        }
        for ($column = 0; $column -lt 2; $column++) {
            Assert-EqualValue $Matrix[$row][$column] $expected[$row][$column] `
                "$Context 第 $($row + 1) 行第 $($column + 1) 列"
        }
    }
}

function Get-PivotSnapshot {
    <# 读取透视表来源、字段角色和 TableRange2 完整矩阵，并跟踪 COM 对象。 #>
    param(
        [Parameter(Mandatory = $true)][object]$Sheet,
        [Parameter(Mandatory = $true)][string]$PivotName
    )

    $pivotTables = Add-ComObject $Sheet.PivotTables()
    $pivot = Add-ComObject $pivotTables.Item($PivotName)
    $tableRange = Add-ComObject $pivot.TableRange2
    $cache = Add-ComObject $pivot.PivotCache()
    $rowFields = Add-ComObject $pivot.RowFields()
    $rowField = Add-ComObject $rowFields.Item(1)
    $dataFields = Add-ComObject $pivot.DataFields()
    $dataField = Add-ComObject $dataFields.Item(1)
    return [ordered]@{
        Address = [string]$tableRange.Address($false, $false)
        Matrix = ConvertTo-Matrix $tableRange.Value2
        Pivot = $pivot
        Cache = $cache
        CacheSource = [string]$cache.SourceData
        CacheSourceType = [int]$cache.SourceType
        Roles = [ordered]@{
            RowFieldCount = [int]$rowFields.Count
            RowSourceName = [string]$rowField.SourceName
            RowCaption = [string]$rowField.Caption
            DataFieldCount = [int]$dataFields.Count
            DataSourceName = [string]$dataField.SourceName
            DataCaption = [string]$dataField.Caption
            AggregateFunction = [int]$dataField.Function
        }
    }
}

function Wait-InitialPivotSnapshot {
    <# 观察正常打开后的客户端初始化，不主动刷新或修改工作簿。 #>
    param(
        [Parameter(Mandatory = $true)][object]$Sheet,
        [Parameter(Mandatory = $true)][string]$PivotName
    )

    $elapsed = [System.Diagnostics.Stopwatch]::StartNew()
    $observations = [System.Collections.Generic.List[object]]::new()
    do {
        $snapshot = Get-PivotSnapshot $Sheet $PivotName
        $observations.Add([ordered]@{
            ElapsedMilliseconds = $elapsed.Elapsed.TotalMilliseconds
            Ready = [bool]$script:ExcelApplication.Ready
            CalculationState = [int]$script:ExcelApplication.CalculationState
            Address = $snapshot.Address
            CacheSource = $snapshot.CacheSource
            Matrix = $snapshot.Matrix
        })
        if (-not [string]::IsNullOrWhiteSpace($snapshot.CacheSource) -and
            $snapshot.Matrix.Count -gt 0) {
            break
        }
        if ($elapsed.Elapsed.TotalSeconds -ge 30) {
            break
        }
        [System.Threading.Thread]::Sleep(250)
    } while ($true)
    $elapsed.Stop()
    $script:Report.InitialObservation = [ordered]@{
        PassiveOnly = $true
        TimeoutSeconds = 30
        ElapsedMilliseconds = $elapsed.Elapsed.TotalMilliseconds
        Samples = $observations
    }
    return $snapshot
}

function Assert-PivotSnapshot {
    <# 校验来源范围、唯一字段角色、汇总函数和完整透视结果。 #>
    param(
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Snapshot,
        [double]$ExpectedNorth,
        [double]$ExpectedSouth,
        [double]$ExpectedTotal,
        [string]$Context
    )

    $source = $Snapshot.CacheSource -replace '\[[^\]]+\]', '' -replace "'", '' -replace '\$', '' `
        -replace '\s', ''
    $sourceIsExpected = $source -match '(?i)^Data!(?:A1:B4|R1C1:R4C2)$'
    if (-not $sourceIsExpected) {
        throw "$Context 的 PivotCache.SourceData 不是 Data!A1:B4：$($Snapshot.CacheSource)"
    }
    if ($Snapshot.Roles.RowFieldCount -ne 1 -or $Snapshot.Roles.RowSourceName -cne 'Warehouse') {
        throw "$Context 必须恰有 Warehouse 行字段，实际为 $($Snapshot.Roles.RowFieldCount) 个 '$($Snapshot.Roles.RowSourceName)'。"
    }
    if ($Snapshot.Roles.DataFieldCount -ne 1 -or $Snapshot.Roles.DataSourceName -cne 'Amount') {
        throw "$Context 必须恰有 Amount 数据字段，实际为 $($Snapshot.Roles.DataFieldCount) 个 '$($Snapshot.Roles.DataSourceName)'。"
    }
    if ($Snapshot.Roles.AggregateFunction -ne -4157) {
        throw "$Context 的 Amount 汇总函数不是 Sum。"
    }
    Assert-PivotMatrix -Matrix $Snapshot.Matrix -ExpectedNorth $ExpectedNorth `
        -ExpectedSouth $ExpectedSouth -ExpectedTotal $ExpectedTotal `
        -DataCaption $Snapshot.Roles.DataCaption -Context $Context
}

function Assert-PivotMatrix {
    <# 按业务行和总计行角色验证矩阵，不依赖 Excel 客户端本地化标题。 #>
    param(
        [object[]]$Matrix,
        [double]$ExpectedNorth,
        [double]$ExpectedSouth,
        [double]$ExpectedTotal,
        [string]$DataCaption,
        [string]$Context
    )

    if ($Matrix.Count -ne 4) {
        throw "$Context 应包含表头、两条业务行和总计行，且恰有两列。"
    }
    foreach ($row in $Matrix) {
        if ($row.Count -ne 2) {
            throw "$Context 每一行必须恰有两列。"
        }
    }
    if ([string]::IsNullOrWhiteSpace([string]$Matrix[0][0]) -or
        [string]::IsNullOrWhiteSpace([string]$Matrix[0][1])) {
        throw "$Context 的两个表头都必须有值。"
    }
    if ([string]$Matrix[0][1] -cne $DataCaption) {
        throw "$Context 的金额列标题与 PivotField.Caption 不一致。"
    }

    $northRows = [System.Collections.Generic.List[object]]::new()
    $southRows = [System.Collections.Generic.List[object]]::new()
    foreach ($row in $Matrix) {
        if ([string]$row[0] -ceq 'North') {
            $northRows.Add($row)
        }
        elseif ([string]$row[0] -ceq 'South') {
            $southRows.Add($row)
        }
    }
    if ($northRows.Count -ne 1 -or $southRows.Count -ne 1) {
        throw "$Context 未找到唯一的 North 与 South 行。"
    }
    Assert-EqualValue $northRows[0][1] $ExpectedNorth "$Context North"
    Assert-EqualValue $southRows[0][1] $ExpectedSouth "$Context South"

    $totalRow = $Matrix[$Matrix.Count - 1]
    if ([string]::IsNullOrWhiteSpace([string]$totalRow[0])) {
        throw "$Context 总计行标签不能为空。"
    }
    if ([string]$totalRow[0] -ceq 'North' -or [string]$totalRow[0] -ceq 'South') {
        throw "$Context 最后一行不是总计行。"
    }
    Assert-EqualValue $totalRow[1] $ExpectedTotal "$Context 总计"
}

function Write-Report {
    <# 以 UTF-8 无 BOM 和 LF 写入验收结果。 #>
    if ([string]::Equals($script:Report.WorkbookPath, $script:Report.ResultPath,
            [System.StringComparison]::OrdinalIgnoreCase) -or
        [string]::Equals($script:Report.CopyPath, $script:Report.ResultPath,
            [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'ResultPath 不能覆盖输入工作簿或验证副本。'
    }
    $parent = [System.IO.Path]::GetDirectoryName($script:Report.ResultPath)
    if (-not [System.IO.Directory]::Exists($parent)) {
        [System.IO.Directory]::CreateDirectory($parent) | Out-Null
    }
    $json = ConvertTo-Json -InputObject $script:Report -Depth 32
    $json = $json.Replace("`r`n", "`n").Replace("`r", "`n")
    $encoding = [System.Text.UTF8Encoding]::new($false)
    [System.IO.File]::WriteAllText($script:Report.ResultPath, $json + "`n", $encoding)
}

try {
    if (-not [System.IO.File]::Exists($script:Report.WorkbookPath)) {
        throw "工作簿不存在：$($script:Report.WorkbookPath)"
    }
    if ([string]::Equals($script:Report.WorkbookPath, $script:Report.ResultPath,
            [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'ResultPath 不能与 WorkbookPath 相同。'
    }

    $directory = [System.IO.Path]::GetDirectoryName($script:Report.WorkbookPath)
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($script:Report.WorkbookPath)
    $copyPath = [System.IO.Path]::Combine($directory, "$baseName.client-verified.xlsx")
    if ([string]::Equals($copyPath, $script:Report.WorkbookPath,
            [System.StringComparison]::OrdinalIgnoreCase)) {
        throw '客户端验证副本不能覆盖原工作簿。'
    }
    if ([string]::Equals($copyPath, $script:Report.ResultPath,
            [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'ResultPath 不能与客户端验证副本相同。'
    }
    if ([System.IO.File]::Exists($copyPath)) {
        throw "验证副本已存在，拒绝覆盖：$copyPath"
    }
    $script:Report.CopyPath = $copyPath

    $existingExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class BasicPivotWindowOwner {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
}
'@
    if ($OwnedProcessId -gt 0) {
        if ($FaultBeforeAttachSeconds -gt 0) {
            [System.Threading.Thread]::Sleep($FaultBeforeAttachSeconds * 1000)
        }
        $ownedHost = Get-Process -Id $OwnedProcessId -ErrorAction Stop
        if ($ownedHost.ProcessName -ine 'EXCEL' -or
            [Math]::Abs(($ownedHost.StartTime.ToUniversalTime() -
                [DateTime]::Parse($OwnedStartTimeUtc).ToUniversalTime()).TotalMilliseconds) -gt 1) {
            throw 'Parent-owned Excel identity changed.'
        }
        Add-Type -LiteralPath (Join-Path $PSScriptRoot 'OwnedExcelWindow.cs')
        $connectTimer = [System.Diagnostics.Stopwatch]::StartNew()
        do {
            $nativeWindow = Add-ComObject ([Bing.Offices.PivotClientProbe.OwnedExcelWindow]::TryGetWindow($OwnedProcessId))
            if ($null -ne $nativeWindow) { break }
            [System.Threading.Thread]::Sleep(100)
        } while ($connectTimer.Elapsed.TotalSeconds -lt 20)
        if ($null -eq $nativeWindow) { throw 'Parent-owned Excel native window did not become ready.' }
        $script:ExcelApplication = Add-ComObject $nativeWindow.Application
        $script:Report.Connection = 'ParentOwnedNativeWindow'
    }
    else {
        $script:ExcelApplication = Add-ComObject (New-Object -ComObject Excel.Application)
        $script:Report.Connection = 'LegacyDirectCom'
    }
    [uint32]$ownerId = 0
    [void][BasicPivotWindowOwner]::GetWindowThreadProcessId(
        [IntPtr]$script:ExcelApplication.Hwnd, [ref]$ownerId)
    if ($ownerId -eq 0 -or ($OwnedProcessId -gt 0 -and $ownerId -ne $OwnedProcessId) -or
        ($OwnedProcessId -eq 0 -and $ownerId -in $existingExcel)) {
        throw 'Cannot verify ownership of a dedicated Excel process.'
    }
    $owner = Get-Process -Id $ownerId -ErrorAction Stop
    if ($owner.ProcessName -ine 'EXCEL') {
        throw 'The verified application owner is not Excel.'
    }
    $script:OwnsApplication = $true
    $script:Report.OwnerProcessId = $ownerId
    $script:Report.OwnedExcel = $true
    if (-not [string]::IsNullOrWhiteSpace($OwnerPath)) {
        Write-Utf8Json $OwnerPath ([ordered]@{
            ProcessId = $ownerId
            StartTimeUtc = $owner.StartTime.ToUniversalTime().ToString('o')
            Name = $owner.ProcessName
            AcquiredApplication = $true
            Connection = $script:Report.Connection
        })
    }
    if ($FaultAfterOwnerSeconds -gt 0) {
        [System.Threading.Thread]::Sleep($FaultAfterOwnerSeconds * 1000)
    }
    $script:ExcelApplication.Visible = $false
    $script:ExcelApplication.DisplayAlerts = $false
    $script:ExcelApplication.AskToUpdateLinks = $false
    $script:ExcelApplication.EnableEvents = $false
    $script:ExcelApplication.Interactive = $false
    $script:ExcelApplication.ScreenUpdating = $false
    $script:ExcelApplication.AutomationSecurity = 3
    $script:Report.Excel = [ordered]@{
        Version = [string]$script:ExcelApplication.Version
        Build = [string]$script:ExcelApplication.Build
    }

    $script:Report.RepairMode = [ordered]@{
        Parameter = 'Workbooks.Open.CorruptLoad'
        OpenMode = 'xlNormalLoad (default)'
        RepairRequested = $false
        Detail = '以默认正常模式打开；未请求或尝试修复。'
    }

    $workbooks = Add-ComObject $script:ExcelApplication.Workbooks
    if ($OwnedProcessId -gt 0) {
        for ($index = $workbooks.Count; $index -ge 1; $index--) {
            $startupBook = Add-ComObject $workbooks.Item($index)
            if ([string]$startupBook.FullName -ine $script:Report.WorkbookPath -or
                -not $startupBook.ReadOnly) {
                throw 'Parent-owned Excel opened an unexpected startup workbook.'
            }
            $startupBook.Close($false)
        }
    }
    $workbook = Add-ComObject $workbooks.Open($script:Report.WorkbookPath, 0, $true)
    $script:OwnedWorkbooks.Add($workbook)
    if (-not $workbook.ReadOnly) {
        throw '输入工作簿未以只读模式打开。'
    }
    $worksheets = Add-ComObject $workbook.Worksheets
    $dataSheet = Add-ComObject $worksheets.Item('Data')
    $summarySheet = Add-ComObject $worksheets.Item('Summary')
    $sourceRange = Get-ComRange $dataSheet 'A1:B4'
    $initialSource = ConvertTo-Matrix $sourceRange.Value2
    Assert-SourceMatrix -Matrix $initialSource -NorthAmount 10 -Context '首次打开的来源'
    $script:Report.InitialSource = $initialSource

    $initialPivot = Wait-InitialPivotSnapshot $summarySheet 'SalesByWarehouse'
    $script:Report.RefreshOnFileOpen = [bool]$initialPivot.Cache.RefreshOnFileOpen
    $script:Report.InitialPivot = [ordered]@{
        Address = $initialPivot.Address
        CacheSource = $initialPivot.CacheSource
        CacheSourceType = $initialPivot.CacheSourceType
        Roles = $initialPivot.Roles
        Matrix = $initialPivot.Matrix
    }
    Assert-PivotSnapshot $initialPivot 40 20 60 '首次显示的透视表'

    (Get-ComRange $dataSheet 'B2').Value2 = 15
    $beforeRefresh = Get-PivotSnapshot $summarySheet 'SalesByWarehouse'
    $script:Report.DisplayAfterSourceEditBeforeRefresh = [ordered]@{
        Address = $beforeRefresh.Address
        CacheSource = $beforeRefresh.CacheSource
        Roles = $beforeRefresh.Roles
        Matrix = $beforeRefresh.Matrix
    }
    # 刷新前记录客户端实际行为，不要求客户端一定保留旧缓存。
    $sourceChangeVisible = $false
    foreach ($row in $beforeRefresh.Matrix) {
        if ([string]$row[0] -ceq 'North') {
            $sourceChangeVisible = [double]$row[1] -eq 45
        }
    }
    $script:Report.DisplayAfterSourceEditBeforeRefresh.SourceChangeVisible = $sourceChangeVisible
    if ($sourceChangeVisible) {
        Assert-PivotSnapshot $beforeRefresh 45 20 65 '来源修改后、刷新前的透视表'
    }
    else {
        Assert-PivotSnapshot $beforeRefresh 40 20 60 '来源修改后、刷新前的透视表'
    }

    $cache = $initialPivot.Cache
    $cache.Refresh() | Out-Null
    $refreshTableResult = $beforeRefresh.Pivot.RefreshTable()
    $refreshed = Get-PivotSnapshot $summarySheet 'SalesByWarehouse'
    $script:Report.RefreshedPivot = [ordered]@{
        Address = $refreshed.Address
        CacheSource = $refreshed.CacheSource
        Roles = $refreshed.Roles
        Matrix = $refreshed.Matrix
        CacheRefreshReturned = $refreshTableResult
    }
    Assert-PivotSnapshot $refreshed 45 20 65 '显式刷新后的透视表'

    $workbook.SaveAs($copyPath, 51)
    $workbook.Close($false)
    [void]$script:OwnedWorkbooks.Remove($workbook)

    $reopened = Add-ComObject $workbooks.Open($copyPath, 0, $true)
    $script:OwnedWorkbooks.Add($reopened)
    if (-not $reopened.ReadOnly) {
        throw '验证副本未以只读模式打开。'
    }
    $reopenedSheets = Add-ComObject $reopened.Worksheets
    $reopenedDataSheet = Add-ComObject $reopenedSheets.Item('Data')
    $reopenedSummarySheet = Add-ComObject $reopenedSheets.Item('Summary')
    $reopenedSource = ConvertTo-Matrix (Get-ComRange $reopenedDataSheet 'A1:B4').Value2
    Assert-SourceMatrix -Matrix $reopenedSource -NorthAmount 15 -Context '重新打开的来源'
    $script:Report.ReopenedSource = $reopenedSource
    $reopenedPivot = Get-PivotSnapshot $reopenedSummarySheet 'SalesByWarehouse'
    Assert-PivotSnapshot $reopenedPivot 45 20 65 '重新打开后的透视表'
    $script:Report.ReopenedPivot = [ordered]@{
        Address = $reopenedPivot.Address
        CacheSource = $reopenedPivot.CacheSource
        CacheSourceType = $reopenedPivot.CacheSourceType
        Roles = $reopenedPivot.Roles
        Matrix = $reopenedPivot.Matrix
        RefreshOnFileOpen = [bool]$reopenedPivot.Cache.RefreshOnFileOpen
    }
    $script:Report.Status = 'Passed'
}
catch {
    $script:Failure = $_.Exception
    $script:Report.Status = 'Failed'
    $script:Report.Error = $_.Exception.ToString()
}
finally {
    for ($index = $script:OwnedWorkbooks.Count - 1; $index -ge 0; $index--) {
        try {
            $script:OwnedWorkbooks[$index].Close($false)
        }
        catch {
            $script:Report.CleanupErrors.Add($_.Exception.Message)
        }
    }
    if ($script:OwnsApplication -and $null -ne $script:ExcelApplication) {
        try {
            $script:ExcelApplication.Quit()
        }
        catch {
            $script:Report.CleanupErrors.Add($_.Exception.Message)
        }
    }
    for ($index = $script:ComObjects.Count - 1; $index -ge 0; $index--) {
        try {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject(
                $script:ComObjects[$index])
        }
        catch {
            $script:Report.CleanupErrors.Add($_.Exception.Message)
        }
    }
    if ($script:Report.Status -eq 'Passed' -and $script:Report.CleanupErrors.Count -gt 0) {
        $script:Report.Status = 'Failed'
        $script:Failure = [System.Exception]::new('Excel COM 资源清理失败。')
        $script:Report.Error = $script:Failure.Message
    }
    try {
        Write-Report
    }
    catch {
        if ($null -eq $script:Failure) {
            $script:Failure = $_.Exception
        }
    }
}

if ($null -ne $script:Failure) {
    [Console]::Error.WriteLine($script:Failure.ToString())
    exit 1
}

Write-Output "Excel client verification passed. Result: $($script:Report.ResultPath)"
