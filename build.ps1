param(
    [ValidateSet("All", "x86", "x64", "arm64")]
    [string]$Architecture = "All",
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    [ValidatePattern("^[0-9A-Za-z][0-9A-Za-z.-]*$")]
    [string]$ReleaseName = "1.2.1-fix.4",
    [switch]$ArmCompatibilityValidation,
    [ValidatePattern("^[0-9]+\.[0-9]+\.[0-9]+-armcompat\.[0-9]+$")]
    [string]$ValidationReleaseName = "1.2.1-armcompat.1"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
if ($ArmCompatibilityValidation) {
    if ($Architecture -notin @("All", "arm64") -or $PSBoundParameters.ContainsKey("ReleaseName")) {
        throw "Validation mode only builds the ARM AnyCPU probe. Use -ValidationReleaseName, not -ReleaseName."
    }
    . "$root\scripts\ManagedPayload.ps1"
}
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

if ($ArmCompatibilityValidation) {
    & "$root\tests\Test-ManagedPayload.ps1"
    $outputPath = "bin\$Configuration\arm-anycpu-probe\"
    & $msbuild "$root\src\OneNoteMarkdown.AddIn\OneNoteMarkdown.AddIn.csproj" `
        /t:Rebuild `
        "/p:Configuration=$Configuration" `
        /p:Platform=AnyCPU /p:PlatformTarget=AnyCPU /p:Prefer32Bit=false `
        /p:TargetFrameworkVersion=v4.8.1 `
        "/p:OutputPath=$outputPath" `
        "/p:BaseIntermediateOutputPath=obj\arm-anycpu-probe\" `
        "/p:IntermediateOutputPath=obj\$Configuration\arm-anycpu-probe\" `
        /v:minimal
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    $payload = Join-Path $root "src\OneNoteMarkdown.AddIn\$outputPath"
    Assert-ArmAnyCpuPayload -Directory $payload
    & $iscc /DInstallerArch=arm64 /DArmCompatibilityValidation=1 `
        "/DBuildConfiguration=$Configuration" "/DReleaseName=$ValidationReleaseName" `
        "$root\src\OneNoteMarkdown.Installer\setup.iss"
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

    $bundle = Join-Path $root "src\OneNoteMarkdown.Installer\Output\arm-anycpu-probe"
    if (Test-Path $bundle) { Remove-Item $bundle -Recurse -Force }
    New-Item -ItemType Directory -Path $bundle -Force | Out-Null
    Copy-Item "$root\src\OneNoteMarkdown.Installer\Output\OneNoteMarkdownSetup-$ValidationReleaseName-arm64.exe" $bundle
    Copy-Item "$root\scripts\ManagedPayload.ps1", "$root\scripts\Test-ArmCompatibility.ps1", "$root\PLUGIN_TEST_GUIDE.md" $bundle
    # Also allow the first-stage CLR test without installing or registering the add-in.
    New-Item -ItemType Directory -Path "$bundle\payload" -Force | Out-Null
    Copy-Item "$payload\*.dll" "$bundle\payload"
    Get-ChildItem $bundle -File -Recurse |
        Where-Object Name -ne "SHA256SUMS.txt" |
        Sort-Object FullName |
        ForEach-Object {
            $relative = $_.FullName.Substring($bundle.Length + 1)
            "{0}  {1}" -f (Get-FileHash $_.FullName -Algorithm SHA256).Hash, $relative
        } | Set-Content "$bundle\SHA256SUMS.txt" -Encoding UTF8
    Get-ChildItem $bundle
    return
}

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
