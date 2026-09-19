[CmdletBinding()]
param(
    [string] $PublishedExecutable = (Join-Path $PSScriptRoot '..\..\artifacts\publish\win-x64-single-file-rc\ReceiptXcel 收據產生工具.exe'),
    [switch] $LeaveSecondRunOpen
)
$ErrorActionPreference = 'Stop'
$source = (Get-Item -LiteralPath $PublishedExecutable).FullName
if ([IO.Path]::GetFileName($source) -ne 'ReceiptXcel 收據產生工具.exe') {
    throw 'Unexpected single-file executable name.'
}
$directory = Join-Path ([IO.Path]::GetTempPath()) ('ReceiptXcel-single-file-smoke-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $directory | Out-Null
if (@(Get-ChildItem -LiteralPath $directory -Force).Count -ne 0) {
    throw 'Smoke directory must start completely empty.'
}
$target = Join-Path $directory ([IO.Path]::GetFileName($source))
Copy-Item -LiteralPath $source -Destination $target
$entries = @(Get-ChildItem -LiteralPath $directory -Force)
if ($entries.Count -ne 1 -or $entries[0].Name -ne [IO.Path]::GetFileName($target)) {
    throw 'Only the EXE is allowed in the clean distribution directory.'
}
if ((Get-FileHash -LiteralPath $source).Hash -ne (Get-FileHash -LiteralPath $target).Hash) {
    throw 'EXE copy hash mismatch.'
}
$runs = @()
for ($run = 1; $run -le 2; $run++) {
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $process = Start-Process -FilePath $target -WorkingDirectory $directory -WindowStyle Hidden -PassThru
    $ready = $false
    try {
        while ($watch.Elapsed.TotalSeconds -lt 30) {
            $process.Refresh()
            if ($process.HasExited) { throw "Single-file app exited before readiness (exit $($process.ExitCode))." }
            if ($process.MainWindowHandle -ne 0 -and
                $process.MainWindowTitle -eq 'ReceiptXcel｜自行收納款項收據產生工具') {
                $ready = $true
                break
            }
            Start-Sleep -Milliseconds 50
        }
        if (-not $ready) { throw 'No ready MainWindow within 30 seconds.' }
        $watch.Stop()
        # Approximation: process launch to expected main-window handle/title.
        # Actual Loaded, embedded availability and generation need the UI check.
        Start-Sleep -Milliseconds 500
        $process.Refresh()
        if ($process.HasExited) { throw 'App exited immediately after MainWindow appeared.' }
        $runs += [pscustomobject]@{
            Run = $run
            ApproximateStartupMilliseconds = $watch.ElapsedMilliseconds
            ProcessId = $process.Id
            MainWindowTitle = $process.MainWindowTitle
        }
    }
    finally {
        if (-not ($LeaveSecondRunOpen -and $run -eq 2 -and $ready) -and -not $process.HasExited) {
            if (-not $process.CloseMainWindow() -or -not $process.WaitForExit(10000)) {
                throw "Smoke process $($process.Id) needs manual closure."
            }
            if ($process.ExitCode -ne 0) { throw "Abnormal close: $($process.ExitCode)" }
        }
    }
}
if (@(Get-ChildItem -LiteralPath $directory -Force).Count -ne 1) {
    throw 'Runtime added a dependency alongside the distributed EXE.'
}
[pscustomobject]@{
    CleanDirectory = $directory
    Executable = $target
    SizeBytes = (Get-Item -LiteralPath $target).Length
    SHA256 = (Get-FileHash -LiteralPath $target).Hash
    Runs = $runs
    SecondRunLeftOpen = [bool]$LeaveSecondRunOpen
    # Intentionally retain only this unique smoke copy for manual inspection.
} | ConvertTo-Json -Depth 4
