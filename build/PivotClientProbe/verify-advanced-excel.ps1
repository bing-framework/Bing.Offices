param(
    [Parameter(Mandatory = $true)][string]$WorkbookPath,
    [Parameter(Mandatory = $true)][string]$ResultPath,
    [Parameter(Mandatory = $true)][string]$OwnerPath,
    [ValidateRange(0, 600)][int]$FaultAfterOwnerSeconds = 0,
    [int]$OwnedProcessId = 0,
    [string]$OwnedStartTimeUtc,
    [ValidateRange(0, 600)][int]$FaultBeforeAttachSeconds = 0
)

$ErrorActionPreference = 'Stop'
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$script:Objects = [System.Collections.Generic.List[object]]::new()
$script:OwnedWorkbooks = [System.Collections.Generic.List[object]]::new()
$script:Application = $null
$script:OwnsApplication = $false
$script:Failure = $null
$script:Report = [ordered]@{ Status = 'Running'; WorkbookPath = [IO.Path]::GetFullPath($WorkbookPath)
    Excel = $null; Initial = $null; BeforeRefresh = $null; Refreshed = $null; Reopened = $null
    CleanupErrors = [System.Collections.Generic.List[string]]::new(); Error = $null }

function Track {
    <# 跟踪本次取得的 COM 引用。 #>
    param([object]$Value)
    if ($null -ne $Value) {
        foreach ($item in $script:Objects) { if ([object]::ReferenceEquals($item, $Value)) { return ,$Value } }
        $script:Objects.Add($Value)
    }
    return ,$Value
}

function Write-Utf8Json {
    <# 将验收状态按无 BOM UTF-8 和 LF 保存。 #>
    param([string]$Path, [object]$Value)
    $text = (ConvertTo-Json -InputObject $Value -Depth 40).Replace("`r`n", "`n").Replace("`r", "`n")
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($Path), $text + "`n", [Text.UTF8Encoding]::new($false))
}

function Matrix {
    <# 保留完整二维工作表值。 #>
    param([object]$Value)
    if ($Value -isnot [Array]) { return ,@(@($Value)) }
    $rows = [object[]]::new($Value.GetLength(0))
    for ($r = 0; $r -lt $rows.Length; $r++) {
        $cells = [object[]]::new($Value.GetLength(1))
        for ($c = 0; $c -lt $cells.Length; $c++) {
            $cells[$c] = $Value.GetValue($r + $Value.GetLowerBound(0), $c + $Value.GetLowerBound(1))
        }
        $rows[$r] = $cells
    }
    return ,$rows
}

function Axis-Items {
    <# 读取结果单元格对应的实际行列字段项。 #>
    param([object]$Items)
    $result = [ordered]@{}
    for ($i = 1; $i -le $Items.Count; $i++) {
        $item = Track $Items.Item($i)
        $parent = Track $item.Parent
        $name = [string]$parent.SourceName
        if ($name -in @('Warehouse', 'Category', 'Month', 'Status')) { $result[$name] = [string]$item.Name }
    }
    return $result
}

function Assert-Number {
    <# 比较客户端聚合数值。 #>
    param([object]$Actual, [double]$Expected, [string]$Context)
    if ($null -eq $Actual -or $Actual -is [string] -or [Math]::Abs([double]$Actual - $Expected) -gt 0.0000001) {
        throw "$Context expected=$Expected actual=$Actual"
    }
}

function Snapshot {
    <# 验证完整来源、字段和每个结果单元格。 #>
    param([object]$Book, [bool]$Edited)
    $sheets = Track $Book.Worksheets
    if ($sheets.Count -ne 2) { throw 'Unexpected complete worksheet count.' }
    $data = Track $sheets.Item('Data')
    $summary = Track $sheets.Item('Summary')
    $source = Matrix (Track $data.Range('A1:G8')).Value2
    $expected = @($script:Fixture.expectedSource)
    if ($Edited) { $expected[12] = '15' }
    $flat = @($source | ForEach-Object { foreach ($value in $_) { $value } })
    if ($flat.Count -ne $expected.Count) { throw 'Complete source dimensions differ.' }
    for ($i = 0; $i -lt $flat.Count; $i++) {
        if ([string]$flat[$i] -cne [string]$expected[$i]) { throw "Source cell $i differs." }
    }
    $tables = Track $summary.PivotTables()
    if ($tables.Count -ne 1) { throw 'Expected exactly one advanced pivot.' }
    $pivot = Track $tables.Item('AdvancedWarehouse')
    $rowFields = Track $pivot.RowFields()
    $columnFields = Track $pivot.ColumnFields()
    $values = Track $pivot.DataFields()
    $pages = Track $pivot.PageFields()
    $rowNames = @(for ($i = 1; $i -le $rowFields.Count; $i++) { [string](Track $rowFields.Item($i)).SourceName })
    $dataAxis = Track $pivot.DataPivotField
    $columnNames = @(for ($i = 1; $i -le $columnFields.Count; $i++) {
        $field = Track $columnFields.Item($i)
        if ([string]$field.Name -ceq [string]$dataAxis.Name) { 'Values' } else { [string]$field.SourceName }
    })
    if (($rowNames -join '|') -cne 'Warehouse|Category' -or ($columnNames -join '|') -cne 'Month|Status|Values') {
        throw "Advanced axis order differs: $($rowNames -join '|') / $($columnNames -join '|')"
    }
    if ($values.Count -ne 5 -or $pages.Count -ne 1 -or [string](Track $pages.Item(1)).SourceName -cne 'Inclusion') {
        throw 'Advanced field roles differ.'
    }
    $functions = @(-4157, -4106, -4139, -4136, -4112)
    for ($i = 1; $i -le 5; $i++) {
        $field = Track $values.Item($i)
        if ([string]$field.Name -cne $script:Fixture.aggregations[$i - 1] -or [int]$field.Function -ne $functions[$i - 1]) {
            throw "Measure $i differs."
        }
    }
    if ([int](Track $rowFields.Item(1)).AutoSortOrder -ne 1 -or [int](Track $rowFields.Item(2)).AutoSortOrder -ne 2) {
        throw 'Advanced label sorting differs.'
    }
    $pageItems = Track (Track $pages.Item(1)).PivotItems()
    $script:Report.LastFilterItems = [System.Collections.Generic.List[object]]::new()
    $selected = @(for ($i = 1; $i -le $pageItems.Count; $i++) {
        $item = Track $pageItems.Item($i)
        $script:Report.LastFilterItems.Add([ordered]@{ Name = [string]$item.Name; Visible = [bool]$item.Visible; RawVisible = $item.Visible })
        if ($item.Visible) { [string]$item.Name }
    })
    $script:Report.LastFilterItemCount = [int]$pageItems.Count
    $script:Report.LastDisplayedMatrix = Matrix (Track $pivot.TableRange2).Value2
    $script:Report.LastGrandValues = @(for ($metric = 0; $metric -lt 5; $metric++) {
        $grandCell = Track $pivot.GetPivotData([string]$script:Fixture.aggregations[$metric])
        [ordered]@{ Name = $script:Fixture.aggregations[$metric]; Value = $grandCell.Value2; Text = [string]$grandCell.Text }
    })
    $cache = Track $pivot.PivotCache()
    if ($cache.RecordCount -ne 7 -or [string]$cache.SourceData -notmatch 'Data!R1C1:R8C7$') {
        throw "Full cache source differs: $($cache.SourceData), records=$($cache.RecordCount)"
    }
    $body = Track $pivot.DataBodyRange
    $bodyCells = Track $body.Cells
    $checked = [System.Collections.Generic.List[object]]::new()
    $filtered = @($script:Fixture.sourceRows | Where-Object { $_.Inclusion -in @('Yes', 'Maybe') })
    for ($i = 1; $i -le $bodyCells.Count; $i++) {
        $cell = Track $bodyCells.Item($i)
        $pc = Track $cell.PivotCell
        $rowKeys = Axis-Items (Track $pc.RowItems)
        $columnKeys = Axis-Items (Track $pc.ColumnItems)
        $measure = [string](Track $pc.DataField).Name
        if ($measure -notin $script:Fixture.aggregations) { throw "Unknown result measure: $measure" }
        $matching = @($filtered | Where-Object {
            $record = $_
            $matches = $true
            foreach ($keys in @($rowKeys, $columnKeys)) {
                foreach ($name in $keys.Keys) { if ([string]$record.$name -cne $keys[$name]) { $matches = $false } }
            }
            $matches
        })
        $actual = $cell.Value2
        $expectedValue = $null
        if ($matching.Count -gt 0) {
            $amounts = @($matching | ForEach-Object {
                if ($Edited -and $_.Item -ceq 'A') { 15.0 } else { [double]$_.Amount }
            })
            $stats = $amounts | Measure-Object -Sum -Average -Minimum -Maximum
            $expectedValue = switch ($measure) {
                'Total' { $stats.Sum }; 'Mean' { $stats.Average }; 'Low' { $stats.Minimum }
                'High' { $stats.Maximum }; 'Count' { [double]$matching.Count }
            }
            Assert-Number $actual $expectedValue ([string]$cell.Address())
        }
        elseif ($null -ne $actual) { throw "Empty combination must remain blank: $($cell.Address())=$actual" }
        $script:Report.LastCheckedAddress = [string]$cell.Address()
        $checked.Add([ordered]@{ Address = [string]$cell.Address(); Measure = $measure
            RowItems = $rowKeys; ColumnItems = $columnKeys; Actual = $actual; Expected = $expectedValue })
    }
    if (($selected -join '|') -cne 'Yes|Maybe') { throw "Actual report filter differs: $($selected -join '|')" }
    if ($checked.Count -lt 20) { throw 'Advanced data body is incomplete.' }
    $grandExpected = if ($Edited) { $script:Fixture.expectedRefreshedGrandTotals } else { $script:Fixture.expectedGrandTotals }
    for ($i = 0; $i -lt 5; $i++) {
        $grand = Track $pivot.GetPivotData([string]$script:Fixture.aggregations[$i])
        Assert-Number $grand.Value2 ([double]$grandExpected[$i]) "Grand total $i"
    }
    return [ordered]@{ Matrix = (Matrix (Track $pivot.TableRange2).Value2)
        RowFields = $rowNames; ColumnFields = $columnNames; Selected = $selected
        CacheSource = [string]$cache.SourceData; CacheRecords = [int]$cache.RecordCount
        BodyRows = [int](Track $body.Rows).Count; BodyColumns = [int](Track $body.Columns).Count
        CheckedCells = $checked }
}

try {
    $inputPath = $script:Report.WorkbookPath
    $script:Fixture = [IO.File]::ReadAllText($inputPath + '.probe.json', [Text.Encoding]::UTF8) | ConvertFrom-Json
    if ($script:Fixture.structuralVerification -cne 'Passed') { throw 'Structural fixture verification is missing.' }
    $copyPath = [IO.Path]::Combine([IO.Path]::GetDirectoryName($inputPath), [IO.Path]::GetFileNameWithoutExtension($inputPath) + '.advanced-verified.xlsx')
    if ([IO.File]::Exists($copyPath)) { throw 'Verification copy already exists.' }
    $existingExcel = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class AdvancedPivotWindowOwner {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
}
'@
    if ($OwnedProcessId -gt 0) {
        if ($FaultBeforeAttachSeconds -gt 0) {
            # 父进程已持有专用进程句柄；模拟 COM 连接前阻塞以验证超时清理。
            [Threading.Thread]::Sleep($FaultBeforeAttachSeconds * 1000)
        }
        $ownedHost = Get-Process -Id $OwnedProcessId -ErrorAction Stop
        if ($ownedHost.ProcessName -ine 'EXCEL' -or
            [Math]::Abs(($ownedHost.StartTime.ToUniversalTime() - [DateTime]::Parse($OwnedStartTimeUtc).ToUniversalTime()).TotalMilliseconds) -gt 1) {
            throw 'Parent-owned Excel identity changed.'
        }
        Add-Type -LiteralPath (Join-Path $PSScriptRoot 'OwnedExcelWindow.cs')
        $connectTimer = [Diagnostics.Stopwatch]::StartNew()
        do {
            $nativeWindow = Track ([Bing.Offices.PivotClientProbe.OwnedExcelWindow]::TryGetWindow($OwnedProcessId))
            if ($null -ne $nativeWindow) { break }
            [Threading.Thread]::Sleep(100)
        } while ($connectTimer.Elapsed.TotalSeconds -lt 20)
        if ($null -eq $nativeWindow) { throw 'Parent-owned Excel native window did not become ready.' }
        $script:Application = Track $nativeWindow.Application
        $script:Report.Connection = 'ParentOwnedNativeWindow'
    }
    else {
        $script:Application = Track (New-Object -ComObject Excel.Application)
    }
    [uint32]$ownerId = 0
    [void][AdvancedPivotWindowOwner]::GetWindowThreadProcessId([IntPtr]$script:Application.Hwnd, [ref]$ownerId)
    if ($ownerId -eq 0 -or ($OwnedProcessId -gt 0 -and $ownerId -ne $OwnedProcessId) -or
        ($OwnedProcessId -eq 0 -and $ownerId -in $existingExcel)) {
        throw 'Cannot verify ownership of a dedicated Excel process.'
    }
    $owner = Get-Process -Id $ownerId
    $script:OwnsApplication = $true
    Write-Utf8Json $OwnerPath ([ordered]@{ ProcessId = $ownerId; StartTimeUtc = $owner.StartTime.ToUniversalTime().ToString('o'); Name = $owner.ProcessName; AcquiredApplication = $true; Connection = $script:Report.Connection })
    if ($FaultAfterOwnerSeconds -gt 0) {
        # 仅供超时职责验证：发布实例所有权后模拟阻塞，不读取用户工作簿。
        [Threading.Thread]::Sleep($FaultAfterOwnerSeconds * 1000)
    }
    $script:Application.Visible = $false
    $script:Application.DisplayAlerts = $false
    $script:Application.AskToUpdateLinks = $false
    $script:Application.EnableEvents = $false
    $script:Application.Interactive = $false
    $script:Application.AutomationSecurity = 3
    $script:Report.Excel = [ordered]@{ Version = [string]$script:Application.Version; Build = [string]$script:Application.Build
        OwnerProcessId = $ownerId; RepairRequested = $false }
    $books = Track $script:Application.Workbooks
    if ($OwnedProcessId -gt 0) {
        # 只关闭启动命令打开的任务夹具，再按已设置的安全选项正常打开。
        for ($index = $books.Count; $index -ge 1; $index--) {
            $startupBook = Track $books.Item($index)
            if ([string]$startupBook.FullName -ine $inputPath -or -not $startupBook.ReadOnly) {
                throw 'Parent-owned Excel opened an unexpected startup workbook.'
            }
            $startupBook.Close($false)
        }
    }
    $book = Track $books.Open($inputPath, 0, $true)
    $script:OwnedWorkbooks.Add($book)
    $script:Report.OpenedWorkbook = [ordered]@{ FullName = [string]$book.FullName; ReadOnly = [bool]$book.ReadOnly }
    if (-not $book.ReadOnly) { throw 'Fixture must open read-only.' }
    $initialSummary = Track (Track $book.Worksheets).Item('Summary')
    $initialPivot = Track (Track $initialSummary.PivotTables()).Item('AdvancedWarehouse')
    $observed = [System.Collections.Generic.List[object]]::new()
    $initialTimer = [Diagnostics.Stopwatch]::StartNew()
    do {
        $initialCountCell = Track $initialPivot.GetPivotData('Count')
        $initialCountValue = $initialCountCell.Value2
        $observed.Add([ordered]@{ ElapsedMilliseconds = $initialTimer.Elapsed.TotalMilliseconds
            Ready = $script:Application.Ready; CalculationState = $script:Application.CalculationState
            Count = $initialCountValue })
        if ($null -ne $initialCountValue) { break }
        [Threading.Thread]::Sleep(250)
    } while ($initialTimer.Elapsed.TotalSeconds -lt 30)
    $script:Report.InitialObservation = $observed
    try {
        $script:Report.Initial = Snapshot $book $false
    }
    catch {
        # 初次验收保持失败；显式刷新仅生成诊断证据，不冒充首次显示成功。
        $initialError = $_
        $script:Report.InitialVerificationError = $_.Exception.ToString()
        (Track $initialPivot.PivotCache()).Refresh() | Out-Null
        $initialPivot.RefreshTable() | Out-Null
        $script:Report.DiagnosticAfterExplicitRefresh = @(for ($metric = 0; $metric -lt 5; $metric++) {
            $grandCell = Track $initialPivot.GetPivotData([string]$script:Fixture.aggregations[$metric])
            [ordered]@{ Name = $script:Fixture.aggregations[$metric]; Value = $grandCell.Value2 }
        })
        $diagnosticPath = [IO.Path]::ChangeExtension([IO.Path]::GetFullPath($ResultPath), '.diagnostic.xlsx')
        if ([IO.File]::Exists($diagnosticPath)) { throw 'Diagnostic copy already exists.' }
        $nativeValues = Track $initialPivot.DataFields()
        $nativeCount = Track $nativeValues.Item('Count')
        $nativeCount.Orientation = 0
        $nativeSource = Track $initialPivot.PivotFields('Item')
        $rebuiltCount = Track $initialPivot.AddDataField($nativeSource, 'Count', -4112)
        $initialPivot.RefreshTable() | Out-Null
        $script:Report.DiagnosticNativeRebuiltCount = (Track $initialPivot.GetPivotData('Count')).Value2
        $book.SaveCopyAs($diagnosticPath)
        $script:Report.DiagnosticCopy = $diagnosticPath
        throw $initialError
    }
    $data = Track (Track $book.Worksheets).Item('Data')
    (Track $data.Range('F2')).Value2 = 15
    $script:Report.BeforeRefresh = Matrix (Track (Track (Track $book.Worksheets).Item('Summary')).Range('A1:AB12')).Value2
    $pivot = Track (Track (Track (Track $book.Worksheets).Item('Summary')).PivotTables()).Item('AdvancedWarehouse')
    (Track $pivot.PivotCache()).Refresh() | Out-Null
    $pivot.RefreshTable() | Out-Null
    $script:Report.Refreshed = Snapshot $book $true
    $book.SaveAs($copyPath, 51)
    $book.Close($false)
    [void]$script:OwnedWorkbooks.Remove($book)
    $reopened = Track $books.Open($copyPath, 0, $true)
    $script:OwnedWorkbooks.Add($reopened)
    $script:Report.Reopened = Snapshot $reopened $true
    $script:Report.CopyPath = $copyPath
    $script:Report.Status = 'Passed'
}
catch {
    $script:Failure = $_.Exception
    $script:Report.Status = 'Failed'
    $script:Report.Error = $_.Exception.ToString() + "`n" + $_.ScriptStackTrace
}
finally {
    foreach ($book in $script:OwnedWorkbooks) {
        try { $book.Close($false) } catch { $script:Report.CleanupErrors.Add($_.Exception.Message) }
    }
    if ($script:OwnsApplication) {
        try { $script:Application.Quit() } catch { $script:Report.CleanupErrors.Add($_.Exception.Message) }
    }
    for ($i = $script:Objects.Count - 1; $i -ge 0; $i--) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:Objects[$i]) }
        catch { $script:Report.CleanupErrors.Add($_.Exception.Message) }
    }
    if ($script:Report.CleanupErrors.Count -gt 0 -and $script:Report.Status -ceq 'Passed') {
        $script:Report.Status = 'Failed'
        $script:Failure = [Exception]::new('Excel COM cleanup failed.')
    }
    Write-Utf8Json $ResultPath $script:Report
}
if ($null -ne $script:Failure) { [Console]::Error.WriteLine($script:Failure.ToString()); exit 1 }
