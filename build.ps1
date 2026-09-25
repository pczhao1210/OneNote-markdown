param(
    [ValidateSet("All", "x86", "x64", "arm64")]
    [string]$Architecture = "All",
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    [ValidatePattern("^[0-9A-Za-z][0-9A-Za-z.-]*$")]
    [string]$ReleaseName = "1.2.1-fix.3"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$msbuild = if (Test-Path $vswhere) {
    & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" |
        Select-Object -First 1
} else {
    $null
}
$isccCandidates = @(
    "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
    "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
)
$iscc = $isccCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $msbuild -or -not (Test-Path $msbuild)) {
    throw "MSBuild not found. Install Visual Studio Build Tools with the MSBuild component."
}
if (-not $iscc) {
    throw "Inno Setup compiler not found."
}

function Assert-PeMachine {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [int]$ExpectedMachine
    )

    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $reader = New-Object System.IO.BinaryReader($stream)
        if ($reader.ReadUInt16() -ne 0x5A4D) {
            throw "Not a PE file: $Path"
        }

        $stream.Position = 0x3C
        $peOffset = $reader.ReadInt32()
        $stream.Position = $peOffset
        if ($reader.ReadUInt32() -ne 0x00004550) {
            throw "Invalid PE signature: $Path"
        }

        $actualMachine = $reader.ReadUInt16()
        if ($actualMachine -ne $ExpectedMachine) {
            throw ("Unexpected PE machine for {0}: expected 0x{1:X4}, actual 0x{2:X4}" -f
                $Path, $ExpectedMachine, $actualMachine)
        }
    }
    finally {
        $stream.Dispose()
    }
}

& $msbuild "$root\tests\OneNoteMarkdown.Tests\OneNoteMarkdown.Tests.csproj" /t:Rebuild "/p:Configuration=$Configuration" /v:minimal
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

& "$root\tests\OneNoteMarkdown.Tests\bin\$Configuration\OneNoteMarkdown.Tests.exe"
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$architectures = if ($Architecture -eq "All") { @("x86", "x64", "arm64") } else { @($Architecture) }
$platformTargets = @{
    x86 = "x86"
    x64 = "x64"
    arm64 = "ARM64"
}
$peMachines = @{
    x86 = 0x014c
    x64 = 0x8664
    arm64 = 0xaa64
}

foreach ($arch in $architectures) {
    $frameworkVersion = if ($arch -eq "arm64") { "v4.8.1" } else { "v4.8" }
    $outputPath = "bin\$Configuration\$arch\"
    & $msbuild "$root\src\OneNoteMarkdown.AddIn\OneNoteMarkdown.AddIn.csproj" `
        /t:Rebuild `
        "/p:Configuration=$Configuration" `
        "/p:PlatformTarget=$($platformTargets[$arch])" `
        "/p:TargetFrameworkVersion=$frameworkVersion" `
        "/p:OutputPath=$outputPath" `
        /v:minimal
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    $addInPath = Join-Path $root "src\OneNoteMarkdown.AddIn\$outputPath\OneNoteMarkdown.AddIn.dll"
    Assert-PeMachine -Path $addInPath -ExpectedMachine $peMachines[$arch]

    & $iscc `
        "/DInstallerArch=$arch" `
        "/DBuildConfiguration=$Configuration" `
        "/DReleaseName=$ReleaseName" `
        "$root\src\OneNoteMarkdown.Installer\setup.iss"
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

$artifacts = foreach ($arch in $architectures) {
    Get-Item "$root\src\OneNoteMarkdown.Installer\Output\OneNoteMarkdownSetup-$ReleaseName-$arch.exe"
}
$artifacts |
    Select-Object Name, Length, LastWriteTime
