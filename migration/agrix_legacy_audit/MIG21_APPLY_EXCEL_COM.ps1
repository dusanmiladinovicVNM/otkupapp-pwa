param(
    [Parameter(Mandatory=$true)][string]$Workbook,
    [Parameter(Mandatory=$true)][string]$Plan
)
$ErrorActionPreference = 'Stop'
$manifest = Get-Content -LiteralPath $Plan -Raw -Encoding UTF8 | ConvertFrom-Json
$operations = @($manifest.operations)
if ($operations.Count -lt 1) { throw 'No approved operations' }
if (-not (Test-Path -LiteralPath $Workbook)) { throw "Workbook not found: $Workbook" }
$excel = $null; $book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3  # ForceDisable VBA execution during repair
    $book = $excel.Workbooks.Open($Workbook,0,$false)
    if ($book.ReadOnly) { throw 'Target is read-only' }

    function Find-Table([string]$name) {
        foreach ($sheet in $book.Worksheets) {
            foreach ($table in $sheet.ListObjects) {
                if ($table.Name -eq $name) { return $table }
            }
        }
        throw "Missing ListObject: $name"
    }
    function Column-Idx($table,[string]$name) {
        try { return [int]$table.ListColumns.Item($name).Index }
        catch { throw "Missing column $($table.Name).$name" }
    }
    function Text([object]$value) {
        if ($null -eq $value) { return '' }
        return ([string]$value).Trim().ToUpperInvariant()
    }
    function Find-Row($table,[string]$field,[string]$value) {
        $idx=Column-Idx $table $field
        $found=0
        for ($i=1;$i -le $table.ListRows.Count;$i++) {
            if ((Text $table.DataBodyRange.Cells.Item($i,$idx).Value2) -eq $value.ToUpperInvariant()) {
                if ($found -gt 0) { throw "Nonunique $field=$value in $($table.Name)" }
                $found=$i
            }
        }
        return $found
    }
    $opi=Find-Table 'tblOtpremnicaIzvori'
    $otk=Find-Table 'tblOtkup'
    $iOid=Column-Idx $opi 'OtkupID'
    $iPid=Column-Idx $opi 'OtpremnicaID'
    $iLink=Column-Idx $opi 'OtpremnicaIzvorID'
    $iOtkBack=Column-Idx $otk 'OtpremnicaID'
    $iOtkId=Column-Idx $otk 'OtkupID'
    $modified=0
    foreach ($op in $operations) {
        $oid=[string]$op.OtkupID
        $pid=[string]$op.new_OtpremnicaID
        $old=[string]$op.old_OtpremnicaID
        $rid=[string]$op.OtpremnicaIzvorID
        if ($op.action -eq 'MOVE') {
            $lr=Find-Row $opi 'OtpremnicaIzvorID' $rid
            if ($lr -lt 1) { throw "Expected OPI link missing: $rid" }
            if ((Text $opi.DataBodyRange.Cells.Item($lr,$iOid).Value2) -ne (Text $oid)) {throw "Unexpected OTK in $rid"}
            if ((Text $opi.DataBodyRange.Cells.Item($lr,$iPid).Value2) -ne (Text $old)) {throw "Unexpected old OTP in $rid"}
            $opi.DataBodyRange.Cells.Item($lr,$iPid).Value2=$pid
        } elseif ($op.action -eq 'ADD') {
            if ((Find-Row $opi 'OtpremnicaIzvorID' $rid) -ne 0) { throw "New OPI ID collision: $rid" }
            $lr=$opi.ListRows.Add()
            $lr.Range.Cells.Item(1,$iLink).Value2=$rid
            $lr.Range.Cells.Item(1,$iOid).Value2=$oid
            $lr.Range.Cells.Item(1,$iPid).Value2=$pid
            foreach ($extra in @('CreatedAt','ModifiedAt')) {
                try {
                    $col=Column-Idx $opi $extra
                    $lr.Range.Cells.Item(1,$col).Value2=[double]([datetime]::Now.ToOADate())
                } catch { throw "Unable to stamp new link $rid field $extra: $_" }
            }
            foreach ($extra in @('CreatedBy','ModifiedBy')) {
                $col=Column-Idx $opi $extra
                $lr.Range.Cells.Item(1,$col).Value2='MIG21'
            }
        } else { throw "Unsupported operation $($op.action)" }

        $oi=Find-Row $otk 'OtkupID' $oid
        if ($oi -lt 1) { throw "OTK header missing: $oid" }
        $current=Text $otk.DataBodyRange.Cells.Item($oi,$iOtkBack).Value2
        $expected=Text ([string]$op.old_Otkup_Backlink)
        if ($current -ne $expected) { throw "Concurrent change of OtkupID=$oid backlink, expected $expected got $current" }
        if ($current -ne (Text $pid)) { $otk.DataBodyRange.Cells.Item($oi,$iOtkBack).Value2=$pid }
        $modified++
    }
    $book.Save()
    Write-Output "MIG21 Excel COM PASS. Corrected OPI/OTK links: $modified"
} finally {
    if ($null -ne $book) {
        try { $book.Close($false) | Out-Null } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($book)
    }
    if ($null -ne $excel) {
        try { $excel.Quit() | Out-Null } catch {}
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
