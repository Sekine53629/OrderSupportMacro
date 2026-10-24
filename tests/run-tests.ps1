# OrderSupportMacro の計算のテスト。
# 新しい Excel を起動し、modules/*.bas を白紙のブックに組み込んで CalcTests.RunAllTests を実行し、結果を表示する。
# テストのデータはすべて CalcTests の中で作る架空の値で、実データは使わない。
#
# 前提: Excel の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が有効。
# 使い方: pwsh -File tests/run-tests.ps1
# FAIL が 1 件でもあれば終了コード 1 を返す。起動した Excel だけを止める。

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$moduleDir = Join-Path $root 'modules'
$vbextStdModule = 1

Add-Type -Namespace Win32 -Name User32 -MemberDefinition @'
[DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint processId);
'@

$xl = New-Object -ComObject Excel.Application
$excelPid = [uint32]0
[void][Win32.User32]::GetWindowThreadProcessId([IntPtr]$xl.Hwnd, [ref]$excelPid)
$failed = $false
try {
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $book = $xl.Workbooks.Add()
    foreach ($bas in Get-ChildItem $moduleDir -Filter '*.bas' | Sort-Object Name) {
        $code = [System.IO.File]::ReadAllText($bas.FullName, [System.Text.Encoding]::UTF8)
        $code = ($code -split "`r`n" | Where-Object { $_ -notmatch '^Attribute VB_' }) -join "`r`n"
        $comp = $book.VBProject.VBComponents.Add($vbextStdModule)
        $comp.Name = $bas.BaseName
        $comp.CodeModule.AddFromString($code)
    }
    $result = $xl.Run('CalcTests.RunAllTests')
    # 分類レポート (シート・グラフ・PDF)。PDF は tests/out に書く
    $outDir = Join-Path $root 'tests\out'
    New-Item -ItemType Directory -Force $outDir | Out-Null
    $result += $xl.Run('CalcTests.RunReportTest', $outDir)
    Write-Host $result
    if ($result -match 'FAIL') { $failed = $true }
    $book.Close($false)
}
finally {
    $xl.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Seconds 2
    $proc = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
    if ($proc) { Stop-Process -Id $excelPid -Force }
}
if ($failed) { exit 1 }
