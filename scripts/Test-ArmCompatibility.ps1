[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AssemblyPath,
    [switch]$InstantiateConnect,
    [switch]$ComActivation
)

$ErrorActionPreference = "Stop"
. "$PSScriptRoot\ManagedPayload.ps1"

function Write-LoadException {
    param([System.Exception]$Exception, [int]$Depth = 0)
    if ($null -eq $Exception -or $Depth -gt 16) { return }
    $prefix = "  " * $Depth
    Write-Output ("{0}{1}, HRESULT=0x{2}: {3}" -f $prefix,
        $Exception.GetType().FullName, $Exception.HResult.ToString("X8"), $Exception.Message)
    foreach ($property in @("FileName", "FusionLog")) {
        try {
            if ($Exception.PSObject.Properties[$property] -and $Exception.$property) {
                Write-Output ("{0}{1}: {2}" -f $prefix, $property, $Exception.$property)
            }
        } catch {
            Write-Output "$prefix${property}: unavailable"
        }
    }
    if ($Exception -is [System.Reflection.ReflectionTypeLoadException]) {
        foreach ($loaderException in $Exception.LoaderExceptions) {
            Write-LoadException $loaderException ($Depth + 1)
        }
    }
    Write-LoadException $Exception.InnerException ($Depth + 1)
}

try {
    if ($PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSEdition -ne "Desktop") {
        throw "Use a fresh Windows PowerShell 5.1 process, not pwsh/PowerShell 7."
    }
    Write-Output "This host is PowerShell, NOT OneNote. Success does not prove ARM64EC Office compatibility."
    Write-Output ("PowerShell={0}; CLR={1}; pointer size={2}" -f
        $PSVersionTable.PSVersion, [Environment]::Version, [IntPtr]::Size)
    $clrModules = @(Get-Process -Id $PID -Module | Where-Object ModuleName -ieq "clr.dll")
    if ($clrModules.Count -ne 1) { throw "Cannot identify the actual loaded clr.dll module." }
    $clrPath = $clrModules[0].FileName
    $clrImage = Get-PeImageInfo -Path $clrPath
    Write-Output ("Host CLR: {0}; Machine=0x{1:X4}" -f $clrPath, $clrImage.Machine)
    if ($clrImage.Machine -ne 0xAA64 -or [IntPtr]::Size -ne 8) {
        throw "This is not the ARM64 CLR. Launch native ARM64 Windows PowerShell 5.1 (avoid SysWOW64/x64 emulation)."
    }

    $path = (Resolve-Path -LiteralPath $AssemblyPath).ProviderPath
    if ([IO.Path]::GetFileName($path) -ne "OneNoteMarkdown.AddIn.dll") {
        throw "AssemblyPath must point to OneNoteMarkdown.AddIn.dll in the complete trusted payload directory."
    }
    if (@([AppDomain]::CurrentDomain.GetAssemblies() |
        Where-Object { $_.GetName().Name -eq "OneNoteMarkdown.AddIn" }).Count -ne 0) {
        throw "The add-in is already loaded. Start a fresh powershell.exe -NoProfile process."
    }
    foreach ($file in Get-ChildItem -LiteralPath (Split-Path $path) -Filter *.dll -File | Sort-Object Name) {
        Write-Output ("Assembly file: {0}; SHA256={1}" -f $file.FullName,
            (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash)
        $image = Get-PeImageInfo -Path $file.FullName
        Write-Output ("Machine=0x{0:X4}; CorFlags={1}" -f $image.Machine, $image.CorFlags)
    }
    $assembly = [Reflection.Assembly]::LoadFrom($path)
    Write-Output ("Assembly.LoadFrom OK: {0}; {1}" -f $assembly.FullName, $assembly.Location)
    $connect = $assembly.GetType("OneNoteMarkdown.AddIn.Connect", $true, $false)
    Write-Output ("Connect type OK: {0}" -f $connect.FullName)
    if ($InstantiateConnect) {
        $instance = [Activator]::CreateInstance($connect)
        Write-Output "Connect ctor returned. Check the ctor log separately; this does not invoke OnConnection."
    }
    if ($ComActivation) {
        Write-Output "COM probe uses registered CLSID, which may point elsewhere. It can invoke ctor/logging, NOT OneNote OnConnection."
        $comType = [Type]::GetTypeFromCLSID([Guid]"0A92B61B-98B8-4E5D-BE2D-48EDB01ED177", $true)
        $comInstance = [Activator]::CreateInstance($comType)
        try {
            Write-Output ("COM activation returned: {0}" -f $comInstance.GetType().FullName)
        } finally {
            if ([Runtime.InteropServices.Marshal]::IsComObject($comInstance)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comInstance)
            }
        }
    }
} catch {
    Write-LoadException $_.Exception
    exit 1
}
