[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$repositoryRoot = $PSScriptRoot
$projectPath = Join-Path $repositoryRoot 'ReportExtract\ReportExtract.csproj'
$distributionRoot = Join-Path $repositoryRoot 'dist'
$portableDirectory = Join-Path $distributionRoot 'ReportExtract-win-x64'
$publishDirectory = Join-Path $portableDirectory 'Backend'
$zipPath = Join-Path $distributionRoot 'ReportExtract-win-x64.zip'
$assetsPath = Join-Path $repositoryRoot 'ReportExtract\obj\project.assets.json'
$nuGetPackagesPath = if ($env:NUGET_PACKAGES) { $env:NUGET_PACKAGES } else { Join-Path $HOME '.nuget\packages' }
$launcherSourcePath = Join-Path $repositoryRoot 'Launcher\ReportExtractLauncher.c'
$launcherPath = Join-Path $portableDirectory 'ReportExtract.exe'
$launcherObjectPath = Join-Path $portableDirectory 'ReportExtractLauncher.obj'

if (-not (Test-Path -LiteralPath $projectPath)) {
    throw "ReportExtract project was not found: $projectPath"
}

if (-not (Test-Path -LiteralPath $launcherSourcePath)) {
    throw "ReportExtract launcher source was not found: $launcherSourcePath"
}

# A runtime-specific restore is needed only when the current assets file or the required .NET 8 runtime packs are absent.
$needsRestore = -not (Test-Path -LiteralPath $assetsPath)
if (-not $needsRestore) {
    $needsRestore = -not (Select-String -LiteralPath $assetsPath -SimpleMatch 'net8.0-windows/win-x64' -Quiet)
}

if (-not $needsRestore) {
    $runtimePackageIds = @(
        'microsoft.netcore.app.runtime.win-x64',
        'microsoft.windowsdesktop.app.runtime.win-x64',
        'microsoft.aspnetcore.app.runtime.win-x64'
    )
    $needsRestore = $runtimePackageIds | Where-Object {
        -not (Get-ChildItem -LiteralPath (Join-Path $nuGetPackagesPath $_) -Directory -Filter '8.*' -ErrorAction SilentlyContinue)
    } | Select-Object -First 1
}

if ($needsRestore) {
    & dotnet restore $projectPath --runtime win-x64
    if ($LASTEXITCODE -ne 0) {
        throw "Restore failed with exit code $LASTEXITCODE."
    }
}

if (Test-Path -LiteralPath $portableDirectory) {
    Remove-Item -LiteralPath $portableDirectory -Recurse -Force
}

if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}

& dotnet publish $projectPath --configuration Release --runtime win-x64 --self-contained true --no-restore --output $publishDirectory
if ($LASTEXITCODE -ne 0) {
    throw "Publish failed with exit code $LASTEXITCODE."
}

function Find-PublishedFile([string]$fileName) {
    return @(Get-ChildItem -LiteralPath $publishDirectory -Recurse -File -Filter $fileName)
}

$requiredBackendFiles = @(
    'ReportExtract.exe',
    'DuckDB.NET.Bindings.dll',
    'DuckDB.NET.Data.dll'
)

foreach ($requiredFile in $requiredBackendFiles) {
    if ((Find-PublishedFile $requiredFile).Count -eq 0) {
        throw "Publish completed without required file '$requiredFile' under $publishDirectory"
    }
}

$duckDbNativeFiles = Find-PublishedFile 'duckdb.dll'
if ($duckDbNativeFiles.Count -eq 0) {
    throw "Publish completed without required file 'duckdb.dll' under $publishDirectory"
}

if ($duckDbNativeFiles.Count -ne 1) {
    $paths = $duckDbNativeFiles.FullName -join [Environment]::NewLine
    throw "Publish found $($duckDbNativeFiles.Count) duckdb.dll files under $publishDirectory; expected exactly one:$([Environment]::NewLine)$paths"
}

$duckDbNativePath = $duckDbNativeFiles[0].FullName

function Find-VisualStudioDeveloperCommand {
    $vsWherePath = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (Test-Path -LiteralPath $vsWherePath) {
        $installationPath = & $vsWherePath -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath
        $developerCommandPath = Join-Path $installationPath 'Common7\Tools\VsDevCmd.bat'
        if ($installationPath -and (Test-Path -LiteralPath $developerCommandPath)) {
            return $developerCommandPath
        }
    }

    throw 'Visual Studio Build Tools with the Desktop development with C++ workload are required to build the portable ReportExtract launcher.'
}

$developerCommandPath = Find-VisualStudioDeveloperCommand
$compileCommand = 'call "{0}" -arch=x64 -host_arch=x64 >nul && cl /nologo /O2 /MT /DUNICODE /D_UNICODE /Fo"{1}" /Fe"{2}" "{3}" /link /SUBSYSTEM:WINDOWS user32.lib' -f $developerCommandPath, $launcherObjectPath, $launcherPath, $launcherSourcePath
& cmd.exe /d /s /c $compileCommand
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $launcherPath)) {
    throw "Launcher build failed with exit code $LASTEXITCODE."
}

Remove-Item -LiteralPath $launcherObjectPath -Force
New-Item -ItemType Directory -Path (Join-Path $portableDirectory 'Data\Temp'), (Join-Path $portableDirectory 'Data\Logs') -Force | Out-Null

if (-not (Test-Path -LiteralPath $launcherPath)) {
    throw "Portable distribution completed without top-level ReportExtract.exe: $launcherPath"
}

if (-not (Test-Path -LiteralPath (Join-Path $publishDirectory 'ReportExtract.exe'))) {
    throw "Portable distribution completed without Backend\\ReportExtract.exe."
}

$developmentDirectories = @(Get-ChildItem -LiteralPath $portableDirectory -Recurse -Directory | Where-Object { $_.Name -in @('bin', 'obj') })
if ($developmentDirectories.Count -ne 0) {
    throw "Portable distribution contains development directories:$([Environment]::NewLine)$($developmentDirectories.FullName -join [Environment]::NewLine)"
}

Compress-Archive -LiteralPath $portableDirectory -DestinationPath $zipPath -CompressionLevel Optimal
if (-not (Test-Path -LiteralPath $zipPath)) {
    throw "ZIP creation failed: $zipPath"
}

Write-Host "Published ReportExtract backend to $publishDirectory"
Write-Host "Created launcher $launcherPath"
Write-Host "Created ZIP archive $zipPath"
