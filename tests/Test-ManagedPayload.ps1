# Standalone regression tests; run with powershell.exe or pwsh -NoProfile -File.
# Fixtures are unsigned, disposable assemblies. Production/ThirdParty files are
# inspected and hashed only, never edited (including when testing bad flags).
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
. (Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts/ManagedPayload.ps1')
$work = Join-Path ([IO.Path]::GetTempPath()) ('OneNoteMarkdown-managed-payload-tests-' + [Guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $work
$script:passed = 0
$script:failed = 0
$script:skipped = 0
$script:mutationRevision = 45678

function Test-Case {
    param([string]$Name, [scriptblock]$Body, [switch]$RequiresTableOffsets)
    if ($RequiresTableOffsets -and $script:valid.AssemblyFlagsOffset -lt 0) {
        $script:skipped++
        Write-Host "SKIP $Name (standard table offsets exposed only by CoreCLR metadata API)."
        return
    }
    try {
        & $Body
        $script:passed++
        Write-Host "PASS $Name"
    }
    catch {
        $script:failed++
        Write-Host "FAIL $Name : $($_.Exception.Message)"
    }
}

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Assert-Throws {
    param([scriptblock]$Body, [string]$Pattern)
    $caught = $null
    try { $null = & $Body } catch { $caught = $_.Exception.Message }
    if ($null -eq $caught) { throw "Expected failure matching '$Pattern', but command succeeded." }
    if ($caught -notmatch $Pattern) { throw "Expected '$Pattern'; actual: $caught" }
}

function Set-U16 {
    param([byte[]]$Bytes, [long]$Offset, [uint16]$Value)
    [BitConverter]::GetBytes($Value).CopyTo($Bytes, [int]$Offset)
}

function Set-U32 {
    param([byte[]]$Bytes, [long]$Offset, [uint32]$Value)
    [BitConverter]::GetBytes($Value).CopyTo($Bytes, [int]$Offset)
}

function Find-Bytes {
    param([byte[]]$Bytes, [byte[]]$Needle)
    $found = -1
    for ($i = 0; $i -le $Bytes.Length - $Needle.Length; $i++) {
        $equal = $true
        for ($j = 0; $j -lt $Needle.Length; $j++) {
            if ($Bytes[$i + $j] -ne $Needle[$j]) { $equal = $false; break }
        }
        if ($equal) {
            if ($found -ge 0) { throw 'Test fixture byte pattern is not unique.' }
            $found = $i
        }
    }
    if ($found -lt 0) { throw 'Test fixture byte pattern was not found.' }
    return $found
}

function Get-TestStreamOffset {
    param([byte[]]$Bytes, [string]$Name)
    $root = $script:valid.MetadataOffset
    $length = [BitConverter]::ToUInt32($Bytes, ($root + 12))
    $p = $root + 16 + (($length + 3) -band -4)
    $count = [BitConverter]::ToUInt16($Bytes, ($p + 2))
    $p += 4
    for ($i = 0; $i -lt $count; $i++) {
        $offset = [BitConverter]::ToUInt32($Bytes, $p)
        $start = $p + 8
        $end = $start
        while ($Bytes[$end] -ne 0) { $end++ }
        if ([Text.Encoding]::ASCII.GetString($Bytes, $start, ($end - $start)) -eq $Name) {
            return $root + $offset
        }
        $p = $start + (($end - $start + 4) -band -4)
    }
    throw "No test metadata stream '$Name'."
}

function New-FixtureAssembly {
    param([string]$Path, [string]$Framework = '.NETFramework,Version=v4.8.1', [switch]$OmitTargetFramework)
    $class = 'Fixture_' + [Guid]::NewGuid().ToString('N')
    $frameworkAttribute = if ($OmitTargetFramework) { '' } else {
        '[assembly: TargetFramework("{0}", FrameworkDisplayName = "Regression fixture")]' -f $Framework
    }
    $source = @"
using System.Reflection;
using System.Runtime.Versioning;
[assembly: AssemblyVersion("12345.23456.34567.45678")]
$frameworkAttribute
[assembly: ${class}Trap]
public sealed class ${class}TrapAttribute : System.Attribute {
    public ${class}TrapAttribute() { throw new System.Exception("Attribute constructors must never run during inspection."); }
}
public sealed class $class { public static int Answer() { return 42; } }
"@
    if ($PSVersionTable.PSEdition -eq 'Core') {
        # Add-Type on CoreCLR chooses a random assembly identity even with
        # -OutputAssembly. Use its already-loaded Roslyn compiler for fixed names.
        $tree = [Microsoft.CodeAnalysis.CSharp.CSharpSyntaxTree]::ParseText($source)
        $reference = [Microsoft.CodeAnalysis.MetadataReference]::CreateFromFile((Join-Path $PSHOME 'ref/System.Runtime.dll'))
        $options = New-Object Microsoft.CodeAnalysis.CSharp.CSharpCompilationOptions ([Microsoft.CodeAnalysis.OutputKind]::DynamicallyLinkedLibrary)
        $compilation = [Microsoft.CodeAnalysis.CSharp.CSharpCompilation]::Create(
            [IO.Path]::GetFileNameWithoutExtension($Path),
            [Microsoft.CodeAnalysis.SyntaxTree[]]@($tree),
            [Microsoft.CodeAnalysis.MetadataReference[]]@($reference), $options)
        $stream = [IO.File]::Create($Path)
        try {
            $emitted = $compilation.Emit($stream)
            if (-not $emitted.Success) { throw ($emitted.Diagnostics -join "`n") }
        }
        finally { $stream.Dispose() }
    }
    else {
        Add-Type -TypeDefinition $source -OutputAssembly $Path -OutputType Library -CompilerOptions '/platform:anycpu'
    }
    if ($OmitTargetFramework) {
        # Assign a fresh identity before reflection-only inspection can cache it.
        $bytes = [IO.File]::ReadAllBytes($Path)
        Set-FixtureMutationIdentity $bytes
        [IO.File]::WriteAllBytes($Path, $bytes)
    }
    $info = Get-PeImageInfo $Path
    # On CoreCLR, Roslyn's trivial fixture references System.Runtime rather than
    # mscorlib. Retarget this UNSIGNED TEST FIXTURE's single AssemblyRef to the
    # equivalent Framework contract; its IL uses only System.Object/attributes.
    # This permits the same real PE + metadata + IL positive fixture on Linux.
    if (@($info.AssemblyReferences | Where-Object { $_.Name -eq 'System.Runtime' }).Count -eq 1) {
        $bytes = [IO.File]::ReadAllBytes($Path)
        $reference = @($info.AssemblyReferences | Where-Object { $_.Name -eq 'System.Runtime' })[0]
        Assert-True ($info.AssemblyReferences.Count -eq 1) 'Unexpected additional compiler references in fixture.'
        $name = Find-Bytes $bytes ([Text.Encoding]::UTF8.GetBytes("System.Runtime`0"))
        $replacement = New-Object byte[] 15
        [Text.Encoding]::UTF8.GetBytes('mscorlib').CopyTo($replacement, 0)
        $replacement.CopyTo($bytes, $name)
        $token = Find-Bytes $bytes ([byte[]]@(0xb0, 0x3f, 0x5f, 0x7f, 0x11, 0xd5, 0x0a, 0x3a))
        ([byte[]]@(0xb7, 0x7a, 0x5c, 0x56, 0x19, 0x34, 0xe0, 0x89)).CopyTo($bytes, $token)
        Set-U16 $bytes $reference.MetadataOffset 4
        Set-U16 $bytes ($reference.MetadataOffset + 2) 0
        Set-U16 $bytes ($reference.MetadataOffset + 4) 0
        Set-U16 $bytes ($reference.MetadataOffset + 6) 0
        [IO.File]::WriteAllBytes($Path, $bytes)
    }
    $info = Get-PeImageInfo $Path
    Assert-True ($info.AssemblyReferences[0].Name -eq 'mscorlib') 'Fixture must have a real Framework AssemblyRef.'
    if ($OmitTargetFramework) {
        Assert-True ($null -eq $info.TargetFramework) 'Fixture must have no TargetFrameworkAttribute.'
    }
    else {
        Assert-True ($info.TargetFramework -eq $Framework) 'Fixture TargetFramework did not round-trip.'
    }
    Assert-True ($info.AssemblyIdentity -eq [Reflection.AssemblyName]::GetAssemblyName($Path).FullName) 'Fixture identity must be CLR-readable.'
}

function Set-FixtureMutationIdentity {
    param([byte[]]$Bytes)
    # Each unsigned negative fixture gets a unique assembly identity, even when
    # its file name is unchanged. Reflection-only identity caches therefore
    # cannot hide mutations within one test process on Windows Framework.
    [byte[]]$marker = @(
        [BitConverter]::GetBytes([uint16]12345)
        [BitConverter]::GetBytes([uint16]23456)
        [BitConverter]::GetBytes([uint16]34567)
        [BitConverter]::GetBytes([uint16]45678)
    )
    $version = Find-Bytes $Bytes $marker
    $script:mutationRevision++
    Set-U16 $Bytes ($version + 6) ([uint16]$script:mutationRevision)
}

function New-Mutation {
    param([string]$Name, [scriptblock]$Change)
    $path = Join-Path $work ($Name + '.dll')
    $bytes = [IO.File]::ReadAllBytes($script:main)
    Set-FixtureMutationIdentity $bytes
    & $Change $bytes
    [IO.File]::WriteAllBytes($path, $bytes)
    return $path
}

function Test-PayloadMutation {
    param([string]$Name, [scriptblock]$Change, [string]$Pattern)
    $bytes = [IO.File]::ReadAllBytes($script:main)
    try {
        $changed = $bytes.Clone()
        Set-FixtureMutationIdentity $changed
        & $Change $changed
        [IO.File]::WriteAllBytes($script:main, $changed)
        Assert-Throws { Assert-ArmAnyCpuPayload $script:payload } $Pattern
    }
    finally { [IO.File]::WriteAllBytes($script:main, $bytes) }
}

try {
    $script:payload = Join-Path $work 'payload'
    $null = New-Item -ItemType Directory -Path $payload
    $names = @('OneNoteMarkdown.AddIn', 'WpfMath', 'XamlMath.Shared', 'Extensibility',
        'office', 'Microsoft.Office.Interop.OneNote')
    foreach ($name in $names) { New-FixtureAssembly (Join-Path $payload ($name + '.dll')) }
    $script:main = Join-Path $payload 'OneNoteMarkdown.AddIn.dll'
    $script:valid = Get-PeImageInfo $main
    $script:pe = [BitConverter]::ToUInt32([IO.File]::ReadAllBytes($main), 0x3c)
    $script:optional = $pe + 24
    $script:section = $optional + [BitConverter]::ToUInt16([IO.File]::ReadAllBytes($main), ($pe + 20))
    $originalHash = (Get-FileHash -LiteralPath $main -Algorithm SHA256).Hash

    Test-Case 'real compiled managed AnyCPU assembly and SHA256' {
        Assert-True ($valid.Machine -eq 0x014c -and $valid.OptionalHeaderMagic -eq 0x10b) 'Expected I386 PE32.'
        Assert-True ($valid.CorFlags -eq 1 -and $valid.IsManaged) 'Expected ILONLY.'
        Assert-True ($valid.SHA256 -eq $originalHash) 'Hash differs from Get-FileHash.'
        Assert-True ($valid.AssemblyName -eq 'OneNoteMarkdown.AddIn') 'Missing assembly manifest identity.'
        Assert-True ($valid.AssemblyReferences.Count -eq 1) 'Missing assembly reference metadata.'
    }
    Test-Case 'complete payload succeeds and reports every DLL' {
        $reports = @(Assert-ArmAnyCpuPayload $payload)
        Assert-True ($reports.Count -eq 6) 'Expected one report per shipped DLL.'
        foreach ($report in $reports) {
            Assert-True ($report.Path -and $report.SHA256.Length -eq 64 -and $report.AssemblyIdentity) 'Incomplete report.'
        }
    }
    Test-Case 'dot-sourcing twice is safe' {
        . (Join-Path (Split-Path -Parent $PSScriptRoot) 'scripts/ManagedPayload.ps1')
        Assert-True ((Get-PeImageInfo $main).SHA256 -eq $originalHash) 'Dot sourcing changed input.'
    }
    Test-Case 'negative fixtures have distinct identities without changing file simple names' {
        $first = Get-PeImageInfo (New-Mutation 'identity-first' { param($b) })
        $second = Get-PeImageInfo (New-Mutation 'identity-second' { param($b) })
        Assert-True ($first.AssemblyIdentity -ne $second.AssemblyIdentity -and
            $first.AssemblyIdentity -ne $valid.AssemblyIdentity -and
            $first.AssemblyName -eq $second.AssemblyName) 'Fixture identities could collide in reflection-only caches.'
    }
    Test-Case '32BITREQUIRED rejected' {
        Test-PayloadMutation 'required32' { param($b) Set-U32 $b ($valid.CliHeaderOffset + 16) 3 } '32BITREQUIRED'
    }
    Test-Case '32BITPREFERRED rejected' {
        Test-PayloadMutation 'preferred32' { param($b) Set-U32 $b ($valid.CliHeaderOffset + 16) 0x20001 } '32BITPREFERRED'
    }
    Test-Case 'ILONLY clear rejected' {
        Test-PayloadMutation 'mixedmode' { param($b) Set-U32 $b ($valid.CliHeaderOffset + 16) 0 } 'ILONLY'
    }
    Test-Case 'native entry point rejected' {
        Test-PayloadMutation 'nativeentry' { param($b) Set-U32 $b ($valid.CliHeaderOffset + 16) 0x11 } 'Native entry'
    }
    Test-Case 'AMD64 machine rejected even with ILONLY' {
        Test-PayloadMutation 'amd64' { param($b) Set-U16 $b ($pe + 4) 0x8664 } 'I386|Cannot inspect'
    }
    Test-Case 'ARM64 machine rejected even with ILONLY' {
        Test-PayloadMutation 'arm64' { param($b) Set-U16 $b ($pe + 4) 0xaa64 } 'I386|Cannot inspect'
    }
    Test-Case 'AssemblyName architecture contradicting PE is rejected' {
        Test-PayloadMutation 'architecture' { param($b) Set-U32 $b $valid.AssemblyFlagsOffset 0x40 } 'processor architecture|Cannot inspect'
    } -RequiresTableOffsets
    Test-Case 'executable renamed to DLL rejected' {
        Test-PayloadMutation 'executable' {
            param($b) Set-U16 $b ($pe + 22) ([BitConverter]::ToUInt16($b, ($pe + 22)) -band 0xdfff)
        } 'managed DLL'
    }
    Test-Case 'wrong main framework rejected' {
        Test-PayloadMutation 'framework' {
            param($b)
            $p = Find-Bytes $b ([Text.Encoding]::UTF8.GetBytes('.NETFramework,Version=v4.8.1'))
            [Text.Encoding]::UTF8.GetBytes('.NETFramework,Version=v4.8.0').CopyTo($b, $p)
        } 'TargetFrameworkAttribute'
    }
    Test-Case 'missing main TargetFrameworkAttribute rejected' {
        $directory = Join-Path $work 'missingframework'
        $null = New-Item -ItemType Directory -Path $directory
        $path = Join-Path $directory 'OneNoteMarkdown.AddIn.dll'
        New-FixtureAssembly $path -OmitTargetFramework
        $saved = [IO.File]::ReadAllBytes($main)
        try {
            Copy-Item -LiteralPath $path -Destination $main -Force
            Assert-Throws { Assert-ArmAnyCpuPayload $payload } 'TargetFrameworkAttribute'
        }
        finally { [IO.File]::WriteAllBytes($main, $saved) }
    }
    Test-Case 'CLR reserved flags rejected' {
        Test-PayloadMutation 'reservedflags' { param($b) Set-U32 $b ($valid.CliHeaderOffset + 16) 0x41 } 'Unknown/reserved'
    }
    Test-Case 'wrong mscorlib reference version rejected' {
        Test-PayloadMutation 'coreversion' {
            param($b) Set-U16 $b $valid.AssemblyReferences[0].MetadataOffset 8
        } 'mscorlib'
    } -RequiresTableOffsets
    Test-Case 'architecture-specific AssemblyRef rejected' {
        Test-PayloadMutation 'refarchitecture' {
            param($b) Set-U32 $b ($valid.AssemblyReferences[0].MetadataOffset + 8) 0x40
        } 'incompatible AssemblyRef'
    } -RequiresTableOffsets
    Test-Case 'missing required DLL rejected' {
        $path = Join-Path $payload 'WpfMath.dll'
        $saved = [IO.File]::ReadAllBytes($path)
        try {
            Remove-Item -LiteralPath $path
            Assert-Throws { Assert-ArmAnyCpuPayload $payload } 'Missing required payload.*WpfMath'
        }
        finally { [IO.File]::WriteAllBytes($path, $saved) }
    }
    Test-Case 'missing main DLL rejected' {
        $saved = [IO.File]::ReadAllBytes($main)
        try {
            Remove-Item -LiteralPath $main
            Assert-Throws { Assert-ArmAnyCpuPayload $payload } 'Missing required payload.*OneNoteMarkdown'
        }
        finally { [IO.File]::WriteAllBytes($main, $saved) }
    }
    Test-Case 'optional stdole absent is accepted when unreferenced' {
        Assert-True (@(Assert-ArmAnyCpuPayload $payload).Count -eq 6) 'Unreferenced optional stdole required.'
    }
    Test-Case 'ALL additional DLLs are inspected, including nested and uppercase extension' {
        $nested = Join-Path $payload 'nested'
        $null = New-Item -ItemType Directory -Path $nested
        try {
            [IO.File]::WriteAllBytes((Join-Path $nested 'Unexpected.DLL'), [byte[]]@(1, 2, 3))
            Assert-Throws { Assert-ArmAnyCpuPayload $payload } 'Unexpected.DLL'
        }
        finally { Remove-Item -LiteralPath $nested -Recurse -Force }
    }
    Test-Case 'hidden DLLs are not skipped' {
        $path = Join-Path $payload '.hidden.dll'
        try {
            [IO.File]::WriteAllBytes($path, [byte[]]@(1, 2, 3))
            if ($PSVersionTable.PSEdition -ne 'Core' -or $env:OS -eq 'Windows_NT') {
                [IO.File]::SetAttributes($path, [IO.FileAttributes]::Hidden)
            }
            Assert-Throws { Assert-ArmAnyCpuPayload $payload } '\.hidden.dll'
        }
        finally { Remove-Item -LiteralPath $path -Force }
    }
    Test-Case 'duplicate assembly simple names rejected' {
        $nested = Join-Path $payload 'duplicate'
        $null = New-Item -ItemType Directory -Path $nested
        try {
            Copy-Item -LiteralPath $main -Destination $nested
            Assert-Throws { Assert-ArmAnyCpuPayload $payload } 'Duplicate assembly simple name'
        }
        finally { Remove-Item -LiteralPath $nested -Recurse -Force }
    }
    Test-Case 'wrong optional stdole is inspected too' {
        $path = Join-Path $payload 'stdole.dll'
        try {
            [IO.File]::WriteAllBytes($path, [byte[]]@(1, 2, 3))
            Assert-Throws { Assert-ArmAnyCpuPayload $payload } 'stdole.dll'
        }
        finally { Remove-Item -LiteralPath $path -Force }
    }
    Test-Case 'nonmanaged PE is reported with null CorFlags and rejected by policy' {
        $path = New-Mutation 'native' {
            param($b) Set-U32 $b ($optional + 96 + 14 * 8) 0; Set-U32 $b ($optional + 96 + 14 * 8 + 4) 0
        }
        $info = Get-PeImageInfo $path
        Assert-True (-not $info.IsManaged -and $null -eq $info.CorFlags) 'Native PE must not claim CLI flags.'
        $destination = Join-Path $payload 'native.dll'
        try {
            Copy-Item -LiteralPath $path -Destination $destination
            Assert-Throws { Assert-ArmAnyCpuPayload $payload } 'Native DLLs'
        }
        finally { Remove-Item -LiteralPath $destination -Force }
    }
    Test-Case 'PE32+ native PE is readable (x64 clr.dll layout)' {
        $path = New-Mutation 'native64' {
            param($b)
            Set-U16 $b ($pe + 4) 0x8664
            Set-U16 $b $optional 0x20b
            # Native fixture needs no directories; section table remains fixed.
            Set-U32 $b ($optional + 108) 0
        }
        $info = Get-PeImageInfo $path
        Assert-True ($info.Machine -eq 0x8664 -and $info.OptionalHeaderMagic -eq 0x20b -and $null -eq $info.CorFlags) 'Native PE32+ inspection failed.'
    }
    Test-Case 'truncated DOS/PE/CLI/metadata data rejected' {
        $bytes = [IO.File]::ReadAllBytes($main)
        foreach ($length in @(0, 2, 63, ($pe + 10), ($optional + 10), ($valid.CliHeaderOffset + 20),
            ($valid.MetadataOffset + 20), ($bytes.Length - 1))) {
            $path = Join-Path $work ('truncated-' + $length + '.dll')
            $part = New-Object byte[] $length
            [Array]::Copy($bytes, $part, $length)
            [IO.File]::WriteAllBytes($path, $part)
            Assert-Throws { Get-PeImageInfo $path } 'Cannot inspect'
        }
    }
    Test-Case 'bad MZ signature rejected' {
        $path = New-Mutation 'badmz' { param($b) Set-U16 $b 0 0 }
        Assert-Throws { Get-PeImageInfo $path } 'MZ signature'
    }
    Test-Case 'bad PE signature rejected' {
        $path = New-Mutation 'badpe' { param($b) Set-U32 $b $pe 0 }
        Assert-Throws { Get-PeImageInfo $path } 'PE signature'
    }
    Test-Case 'overflowing PE offset rejected' {
        $path = New-Mutation 'overflowpe' { param($b) Set-U32 $b 0x3c ([uint32]::MaxValue) }
        Assert-Throws { Get-PeImageInfo $path } 'out-of-range'
    }
    Test-Case 'invalid optional-header magic rejected' {
        $path = New-Mutation 'badmagic' { param($b) Set-U16 $b $optional 0x999 }
        Assert-Throws { Get-PeImageInfo $path } 'optional header magic'
    }
    Test-Case 'directory count beyond optional header rejected' {
        $path = New-Mutation 'directorycount' { param($b) Set-U32 $b ($optional + 92) 1000 }
        Assert-Throws { Get-PeImageInfo $path } 'directory count'
    }
    Test-Case 'zero section count rejected' {
        $path = New-Mutation 'nosections' { param($b) Set-U16 $b ($pe + 6) 0 }
        Assert-Throws { Get-PeImageInfo $path } 'section count'
    }
    Test-Case 'raw section pointer outside file rejected' {
        $path = New-Mutation 'badraw' { param($b) Set-U32 $b ($section + 20) ([uint32]::MaxValue) }
        Assert-Throws { Get-PeImageInfo $path } 'section raw data'
    }
    Test-Case 'CLI RVA outside all sections rejected' {
        $path = New-Mutation 'unmapped' { param($b) Set-U32 $b ($optional + 96 + 14 * 8) 4294901760 }
        Assert-Throws { Get-PeImageInfo $path } 'Unmapped.*CLI'
    }
    Test-Case 'CLI RVA in unbacked virtual section tail rejected' {
        $path = New-Mutation 'virtualtail' {
            param($b)
            $rva = [BitConverter]::ToUInt32($b, ($section + 12))
            $rawSize = [BitConverter]::ToUInt32($b, ($section + 16))
            Set-U32 $b ($section + 8) ($rawSize + 0x200)
            Set-U32 $b ($optional + 96 + 14 * 8) ($rva + $rawSize + 8)
        }
        Assert-Throws { Get-PeImageInfo $path } 'not backed|Overlapping'
    }
    Test-Case 'undersized CLI directory rejected' {
        $path = New-Mutation 'shortcli' { param($b) Set-U32 $b ($optional + 96 + 14 * 8 + 4) 16 }
        Assert-Throws { Get-PeImageInfo $path } 'Truncated CLI'
    }
    Test-Case 'I386 PE32 ILONLY without valid metadata rejected' {
        $path = New-Mutation 'badmetadata' { param($b) Set-U32 $b $valid.MetadataOffset 0 }
        Assert-Throws { Get-PeImageInfo $path } 'metadata signature'
    }
    Test-Case 'overflowing metadata length rejected' {
        $path = New-Mutation 'metadataoverflow' { param($b) Set-U32 $b ($valid.CliHeaderOffset + 12) ([uint32]::MaxValue) }
        Assert-Throws { Get-PeImageInfo $path } 'Overflowing|Unmapped|out-of-range'
    }
    Test-Case 'bad metadata stream bounds rejected' {
        $path = New-Mutation 'streamoverflow' {
            param($b)
            $versionLength = [BitConverter]::ToUInt32($b, ($valid.MetadataOffset + 12))
            $stream = $valid.MetadataOffset + 16 + (($versionLength + 3) -band -4) + 4
            Set-U32 $b $stream ([uint32]::MaxValue)
        }
        Assert-Throws { Get-PeImageInfo $path } 'Cannot inspect'
    }
    Test-Case 'metadata stream overlapping its directory rejected' {
        $path = New-Mutation 'streamoverlap' {
            param($b)
            $length = [BitConverter]::ToUInt32($b, ($valid.MetadataOffset + 12))
            $stream = $valid.MetadataOffset + 16 + (($length + 3) -band -4) + 4
            Set-U32 $b $stream 0
        }
        Assert-Throws { Get-PeImageInfo $path } 'Cannot inspect'
    }
    Test-Case 'overflowing metadata table row count rejected' {
        $path = New-Mutation 'tableoverflow' {
            param($b)
            $table = Get-TestStreamOffset $b '#~'
            Set-U32 $b ($table + 24) ([uint32]::MaxValue)
        }
        Assert-Throws { Get-PeImageInfo $path } 'Cannot inspect'
    }
    Test-Case 'assembly manifest is required, not a netmodule' {
        $path = New-Mutation 'netmodule' {
            param($b)
            $table = Get-TestStreamOffset $b '#~'
            $mask = [BitConverter]::ToUInt64($b, ($table + 8))
            $row = $table + 24
            for ($i = 0; $i -lt 32; $i++) {
                if (($mask -band ([uint64]1 -shl $i)) -ne 0) { $row += 4 }
            }
            Set-U32 $b $row 0
        }
        Assert-Throws { Get-PeImageInfo $path } 'Cannot inspect'
    }
    Test-Case 'invalid metadata string heap index rejected' {
        $path = New-Mutation 'stringindex' {
            param($b) Set-U16 $b ($valid.AssemblyFlagsOffset + 6) 0xffff
        }
        Assert-Throws { Get-PeImageInfo $path } 'Cannot inspect'
    } -RequiresTableOffsets
    Test-Case 'invalid metadata blob heap index rejected' {
        $path = New-Mutation 'blobindex' {
            param($b) Set-U16 $b ($valid.AssemblyFlagsOffset + 4) 0xffff
        }
        Assert-Throws { Get-PeImageInfo $path } 'Cannot inspect'
    } -RequiresTableOffsets
    Test-Case 'malformed target-framework attribute blob rejected' {
        $path = New-Mutation 'attributeblob' {
            param($b)
            $p = Find-Bytes $b ([Text.Encoding]::UTF8.GetBytes('.NETFramework,Version=v4.8.1'))
            Set-U16 $b ($p - 3) 0
        }
        Assert-Throws { Get-PeImageInfo $path } 'Cannot inspect'
    }
    Test-Case 'invalid managed entry point token rejected' {
        $path = New-Mutation 'badentry' { param($b) Set-U32 $b ($valid.CliHeaderOffset + 20) 0x0600ffff }
        Assert-Throws { Get-PeImageInfo $path } 'Cannot inspect'
    }
    Test-Case 'unsupported CLI runtime header version rejected' {
        $path = New-Mutation 'cliversion' { param($b) Set-U16 $b ($valid.CliHeaderOffset + 4) 99 }
        Assert-Throws { Get-PeImageInfo $path } 'CLI header runtime version'
    }
    Test-Case 'missing file and directory-as-file are actionable' {
        Assert-Throws { Get-PeImageInfo (Join-Path $work 'missing.dll') } 'Cannot inspect'
        Assert-Throws { Get-PeImageInfo $work } 'Expected a file'
    }
    Test-Case 'all shipped ThirdParty assemblies have valid read-only metadata' {
        $thirdParty = Join-Path (Split-Path -Parent $PSScriptRoot) 'src/OneNoteMarkdown.AddIn/ThirdParty'
        $files = @(Get-ChildItem -LiteralPath $thirdParty -Recurse -Filter '*.dll')
        Assert-True ($files.Count -eq 6) 'Expected the six vendored dependency DLLs.'
        foreach ($file in $files) {
            $hash = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
            $info = Get-PeImageInfo $file.FullName
            Assert-True ($info.Machine -eq 0x14c -and $info.OptionalHeaderMagic -eq 0x10b -and
                ($info.CorFlags -band 0x20003) -eq 1) "Incompatible shipped flags: $($file.Name)."
            Assert-True ($info.SHA256 -eq $hash -and $info.AssemblyIdentity) 'Missing hash/identity.'
            Assert-True ((Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash -eq $hash) 'Shipped file modified.'
            Write-Host ("  {0} Machine=0x{1:X4} CorFlags=0x{2:X8} SHA256={3} Identity={4} TFM={5}" -f
                $file.Name, $info.Machine, $info.CorFlags, $hash, $info.AssemblyIdentity, $info.TargetFramework)
        }
    }
    Test-Case 'real ThirdParty set passes alongside valid main (including legacy interop refs)' {
        $real = Join-Path $work 'real-dependencies'
        $null = New-Item -ItemType Directory -Path $real
        Copy-Item -LiteralPath $main -Destination $real
        $thirdParty = Join-Path (Split-Path -Parent $PSScriptRoot) 'src/OneNoteMarkdown.AddIn/ThirdParty'
        Get-ChildItem -LiteralPath $thirdParty -Recurse -Filter '*.dll' |
            ForEach-Object { Copy-Item -LiteralPath $_.FullName -Destination $real }
        Assert-True (@(Assert-ArmAnyCpuPayload $real).Count -eq 7) 'Actual dependency set failed policy.'
        $xaml = Join-Path $real 'XamlMath.Shared.dll'
        $xamlSaved = [IO.File]::ReadAllBytes($xaml)
        try {
            # Swap in an unsigned fixture, not an edited signed production DLL.
            Copy-Item -LiteralPath (Join-Path $payload 'XamlMath.Shared.dll') -Destination $xaml -Force
            Assert-Throws { Assert-ArmAnyCpuPayload $real } 'references.*XamlMath.Shared.*but shipped'
        }
        finally { [IO.File]::WriteAllBytes($xaml, $xamlSaved) }
        Remove-Item -LiteralPath (Join-Path $real 'stdole.dll')
        Assert-Throws { Assert-ArmAnyCpuPayload $real } 'unresolved.*stdole'
    }
    Test-Case 'validation leaves unsigned fixture unchanged too' {
        $null = Assert-ArmAnyCpuPayload $payload
        Assert-True ((Get-FileHash -LiteralPath $main -Algorithm SHA256).Hash -eq $originalHash) 'Validation changed fixture.'
    }
    if ($env:OS -eq 'Windows_NT') {
        Test-Case 'installed native .NET Framework clr.dll remains inspectable' {
            $clrFiles = @(
                (Join-Path $env:WINDIR 'Microsoft.NET/Framework/v4.0.30319/clr.dll'),
                (Join-Path $env:WINDIR 'Microsoft.NET/Framework64/v4.0.30319/clr.dll')
            ) | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf }
            Assert-True (@($clrFiles).Count -gt 0) 'No installed .NET Framework clr.dll found.'
            foreach ($clr in $clrFiles) {
                $info = Get-PeImageInfo $clr
                Assert-True (-not $info.IsManaged -and $null -eq $info.CorFlags -and $info.Machine -ne 0) 'Native clr.dll misidentified.'
            }
        }
    }
    else { $script:skipped++; Write-Host 'SKIP installed Windows clr.dll/.NET Framework runtime check (Linux host).' }
}
finally {
    Remove-Item -LiteralPath $work -Recurse -Force
}

Write-Host ("Managed payload tests: {0} passed, {1} failed, {2} skipped." -f $passed, $failed, $skipped)
if ($failed -ne 0) { throw "$failed managed payload regression test(s) failed." }
