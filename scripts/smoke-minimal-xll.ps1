[CmdletBinding()]
param(
    [string]$XllPath
)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrEmpty($XllPath)) {
    $XllPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'target/release/minimal_xll.xll'
}
$xll = (Resolve-Path -LiteralPath $XllPath).Path

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

function Invoke-SmokePass([int]$Pass) {
    $excel = $null
    $books = $null
    $book = $null
    $sheets = $null
    $sheet = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        if (-not $excel.RegisterXLL($xll)) { throw "Excel rejected RegisterXLL($xll)" }

        $books = $excel.Workbooks
        $book = $books.Add()
        $sheets = $book.Worksheets
        $sheet = $sheets.Item(1)
        $sheet.Range('A1').Formula2 = '=RUST.ADD(2,3)'
        $sheet.Range('A2').Formula2 = '=RUST.ECHO("Aé水😀")'
        $sheet.Range('A3').Formula2 = '=RUST.REFERENCE.KIND(A1:B1)'
        $sheet.Range('A4').Formula2 = '=RUST.OPTION.KIND()'
        $sheet.Range('A6').Value2 = 1
        $sheet.Range('B6').Value2 = 'text'
        $sheet.Range('A7').Value2 = $true
        $sheet.Range('B7').Formula2 = '=NA()'
        $sheet.Range('D6').Formula2 = '=RUST.ARRAY.ECHO(A6:B7)'
        $sheet.Range('F1').Formula2 = '=RUST.ADD(ROW(),1)'
        $sheet.Range('F1:F500').FillDown()
        $excel.CalculateFullRebuild()

        $add = $sheet.Range('A1').Value2
        $echo = $sheet.Range('A2').Value2
        $referenceKind = $sheet.Range('A3').Value2
        $optionKind = $sheet.Range('A4').Value2
        $array11 = $sheet.Range('D6').Value2
        $array12 = $sheet.Range('E6').Value2
        $array21 = $sheet.Range('D7').Value2
        $repeatedLast = $sheet.Range('F500').Value2
        if ($add -ne 5) { throw "RUST.ADD returned $add" }
        if ($echo -cne 'Aé水😀') { throw 'RUST.ECHO did not preserve Unicode text' }
        if ($referenceKind -notin @('SRef', 'Ref')) { throw "unexpected reference kind $referenceKind" }
        if ($optionKind -ne 'missing') { throw "unexpected option kind $optionKind" }
        if ($array11 -ne 1 -or $array12 -ne 'text' -or $array21 -ne $true) { throw 'mixed array spill mismatch' }
        if ($repeatedLast -ne 501) { throw "repeated calculation ended at $repeatedLast" }

        $observed = [ordered]@{
            pass = $Pass
            version = $excel.Version
            build = $excel.Build
            operating_system = $excel.OperatingSystem
            mtr_enabled = $excel.MultiThreadedCalculation.Enabled
            mtr_threads = $excel.MultiThreadedCalculation.ThreadCount
            add = $add
            echo_exact = $true
            reference_kind = $referenceKind
            option_kind = $optionKind
            array_1_1 = $array11
            array_1_2 = $array12
            array_2_1 = $array21
            array_2_2 = $sheet.Range('E7').Text
            repeated_last = $repeatedLast
        }
        [pscustomobject]$observed
        [void]$book.Close($false)
        $book = $null
        [void]$excel.Quit()
        $excel = $null
    }
    finally {
        if ($null -ne $book) { [void]$book.Close($false) }
        if ($null -ne $excel) { [void]$excel.Quit() }
        Release-ComObject $sheet
        Release-ComObject $sheets
        Release-ComObject $book
        Release-ComObject $books
        Release-ComObject $excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

Invoke-SmokePass 1 | Format-List
Invoke-SmokePass 2 | Format-List
