# Only header/RVA parsing is custom. Assembly metadata is read by the runtime's
# standard reader: reflection-only on Framework, bundled PEReader on CoreCLR.
if (-not ('OneNoteMarkdown.PayloadInspector' -as [type])) {
    $source = @'
using System;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Collections.Generic;
#if CORECLR
using System.Reflection.Metadata;
using System.Reflection.Metadata.Ecma335;
using System.Reflection.PortableExecutable;
#endif
namespace OneNoteMarkdown {
    public sealed class PayloadReference {
        public string Name, Version, Culture, PublicKeyToken, Identity;
        public uint Flags;
        public long MetadataOffset = -1;
        public PayloadReference(AssemblyName n) {
            Name = n.Name; Version = n.Version.ToString(); Culture = n.CultureName;
            byte[] token = n.GetPublicKeyToken();
            PublicKeyToken = token == null || token.Length == 0 ? "null" : BitConverter.ToString(token).Replace("-", "").ToLowerInvariant();
            Identity = n.FullName;
        }
    }
    public sealed class PayloadImage {
        public string Path, SHA256, MetadataVersion, AssemblyName, AssemblyIdentity, ProcessorArchitecture, TargetFramework;
        public ushort Machine, OptionalHeaderMagic, Characteristics;
        public uint? CorFlags;
        public bool IsManaged, HasNativeCode;
        public long CliHeaderOffset, MetadataOffset, AssemblyFlagsOffset = -1;
        public uint MetadataSize;
        public PayloadReference[] AssemblyReferences = new PayloadReference[0];
    }
    public sealed class PayloadInspector {
        private byte[] bytes;
        private uint headers;
        private readonly List<uint[]> sections = new List<uint[]>();
        private static void Require(bool valid, string message) { if (!valid) throw new BadImageFormatException(message); }
        private void Range(long p, long n, string label) {
            Require(p >= 0 && n >= 0 && p <= bytes.LongLength && n <= bytes.LongLength - p, "Truncated or out-of-range " + label + ".");
        }
        private ushort W(long p) { Range(p, 2, "field"); return BitConverter.ToUInt16(bytes, (int)p); }
        private uint D(long p) { Range(p, 4, "field"); return BitConverter.ToUInt32(bytes, (int)p); }
        private long Map(uint rva, uint size, string label) {
            long end = (long)rva + size, found = -1;
            Require(rva != 0 && size != 0 && end <= 0x100000000L, "Missing or overflowing " + label + " RVA/size.");
            if (rva < headers && end <= headers) found = rva;
            foreach (uint[] s in sections) {
                if (rva < s[0] || end > (long)s[0] + Math.Max(s[1], s[3])) continue;
                long delta = (long)rva - s[0];
                Require(delta + size <= s[3], label + " is not backed by raw file data.");
                Require(found < 0, "Ambiguous overlapping sections for " + label + ".");
                found = s[2] + delta;
            }
            Require(found >= 0, "Unmapped " + label + " RVA.");
            Range(found, size, label); return found;
        }
        private PayloadImage Headers(string path) {
            PayloadImage info = new PayloadImage { Path = System.IO.Path.GetFullPath(path) };
            using (SHA256 hash = SHA256.Create()) info.SHA256 = BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", "");
            Range(0, 64, "DOS header"); Require(W(0) == 0x5a4d, "Missing MZ signature.");
            uint pe = D(0x3c); Require(pe >= 64, "Invalid PE header offset."); Range(pe, 24, "PE header");
            Require(D(pe) == 0x4550, "Invalid PE signature.");
            info.Machine = W(pe + 4); info.Characteristics = W(pe + 22);
            ushort count = W(pe + 6), optionalSize = W(pe + 20);
            Require(count > 0 && count <= 96, "Invalid section count.");
            long optional = pe + 24L; Range(optional, optionalSize, "optional header");
            Require(optionalSize >= 2, "Missing optional header.");
            info.OptionalHeaderMagic = W(optional);
            Require(info.OptionalHeaderMagic == 0x10b || info.OptionalHeaderMagic == 0x20b, "Invalid optional header magic.");
            int directories = info.OptionalHeaderMagic == 0x10b ? 96 : 112;
            Require(optionalSize >= directories, "Truncated optional header.");
            uint directoryCount = D(optional + directories - 4);
            Require(directoryCount <= (optionalSize - directories) / 8, "Data directory count exceeds optional header.");
            headers = D(optional + 60);
            long table = optional + optionalSize; Range(table, count * 40L, "section table");
            Require(headers >= table + count * 40L && headers <= bytes.LongLength, "Invalid SizeOfHeaders.");
            for (int i = 0; i < count; i++) {
                long p = table + i * 40;
                uint[] s = { D(p + 12), D(p + 8), D(p + 20), D(p + 16) };
                if (s[3] != 0) { Range(s[2], s[3], "section raw data"); Require(s[2] >= headers, "Section overlaps PE headers."); }
                Require((long)s[0] + Math.Max(s[1], s[3]) <= 0x100000000L, "Overflowing section RVA.");
                foreach (uint[] old in sections) {
                    Require(s[3] == 0 || old[3] == 0 || (long)s[2] + s[3] <= old[2] || (long)old[2] + old[3] <= s[2], "Overlapping raw sections.");
                    Require(Math.Max(s[1], s[3]) == 0 || Math.Max(old[1], old[3]) == 0 ||
                        (long)s[0] + Math.Max(s[1], s[3]) <= old[0] || (long)old[0] + Math.Max(old[1], old[3]) <= s[0], "Overlapping section RVAs.");
                }
                sections.Add(s);
            }
            if (directoryCount <= 14) return info;
            long cliDirectory = optional + directories + 14 * 8;
            uint cliRva = D(cliDirectory), cliSize = D(cliDirectory + 4);
            if (cliRva == 0 && cliSize == 0) return info; // Native clr.dll is valid here.
            Require(cliSize >= 72, "Truncated CLI header.");
            long cli = Map(cliRva, cliSize, "CLI header");
            Require(D(cli) >= 72 && D(cli) <= cliSize, "Invalid CLI header size.");
            Require(W(cli + 4) == 2 && W(cli + 6) <= 5, "Unsupported CLI header runtime version.");
            info.CliHeaderOffset = cli; info.CorFlags = D(cli + 16); info.IsManaged = true;
            info.MetadataSize = D(cli + 12); info.MetadataOffset = Map(D(cli + 8), info.MetadataSize, "CLI metadata");
            Require(info.MetadataSize >= 20 && D(info.MetadataOffset) == 0x424a5342, "Invalid CLI metadata signature.");
            for (int i = 24; i <= 64; i += 8) {
                uint rva = D(cli + i), size = D(cli + i + 4);
                if (rva == 0 && size == 0) continue;
                Map(rva, size, "CLI data directory");
                if (i >= 40) info.HasNativeCode = true;
            }
            return info;
        }
        private void Metadata(PayloadImage info) {
            AssemblyName identity = AssemblyName.GetAssemblyName(info.Path); // Metadata only; never executes code.
            info.AssemblyName = identity.Name; info.AssemblyIdentity = identity.FullName;
            info.ProcessorArchitecture = typeof(AssemblyName).GetProperty("ProcessorArchitecture").GetValue(identity, null).ToString();
            List<PayloadReference> references = new List<PayloadReference>();
#if CORECLR
            using (PEReader pe = new PEReader(new MemoryStream(bytes, false))) {
                MetadataReader md = pe.GetMetadataReader();
                Require(md.IsAssembly, "Metadata has no assembly manifest.");
                AssemblyDefinition assembly = md.GetAssemblyDefinition();
                info.MetadataVersion = md.MetadataVersion;
                info.AssemblyFlagsOffset = info.MetadataOffset + md.GetTableMetadataOffset(TableIndex.Assembly) + 12;
                uint flags = (uint)assembly.Flags;
                int architecture = (int)((flags >> 4) & 7);
                if (architecture != 0) info.ProcessorArchitecture = new [] { "None", "MSIL", "X86", "IA64", "Amd64", "Arm", "Reserved", "Reserved" }[architecture];
                foreach (AssemblyReferenceHandle handle in md.AssemblyReferences) {
                    AssemblyReference reference = md.GetAssemblyReference(handle);
                    AssemblyName name = new AssemblyName { Name = md.GetString(reference.Name), Version = reference.Version, CultureName = md.GetString(reference.Culture) };
                    byte[] key = md.GetBlobBytes(reference.PublicKeyOrToken);
                    if (((uint)reference.Flags & 1) != 0) name.SetPublicKey(key); else name.SetPublicKeyToken(key);
                    PayloadReference item = new PayloadReference(name) { Flags = (uint)reference.Flags };
                    item.MetadataOffset = info.MetadataOffset + md.GetTableMetadataOffset(TableIndex.AssemblyRef) +
                        (MetadataTokens.GetRowNumber(handle) - 1) * md.GetTableRowSize(TableIndex.AssemblyRef);
                    references.Add(item);
                }
                foreach (CustomAttributeHandle handle in assembly.GetCustomAttributes()) {
                    CustomAttribute attribute = md.GetCustomAttribute(handle);
                    if (attribute.Constructor.Kind != HandleKind.MemberReference) continue;
                    MemberReference member = md.GetMemberReference((MemberReferenceHandle)attribute.Constructor);
                    if (member.Parent.Kind != HandleKind.TypeReference || md.GetString(member.Name) != ".ctor") continue;
                    TypeReference type = md.GetTypeReference((TypeReferenceHandle)member.Parent);
                    if (md.GetString(type.Namespace) != "System.Runtime.Versioning" || md.GetString(type.Name) != "TargetFrameworkAttribute") continue;
                    Require(type.ResolutionScope.Kind == HandleKind.AssemblyReference, "Invalid TargetFrameworkAttribute scope.");
                    AssemblyReference scope = md.GetAssemblyReference((AssemblyReferenceHandle)type.ResolutionScope);
                    string scopeName = md.GetString(scope.Name);
                    Require(scopeName == "mscorlib" || scopeName == "System.Runtime" || scopeName == "netstandard", "Invalid TargetFrameworkAttribute assembly.");
                    string token = references[MetadataTokens.GetRowNumber((AssemblyReferenceHandle)type.ResolutionScope) - 1].PublicKeyToken;
                    Require(token == "b77a5c561934e089" || token == "b03f5f7f11d50a3a" || token == "cc7b13ffcd2ddd51", "Invalid TargetFrameworkAttribute framework identity.");
                    byte[] signature = md.GetBlobBytes(member.Signature);
                    Require(signature.Length == 4 && signature[0] == 0x20 && signature[1] == 1 && signature[2] == 1 && signature[3] == 0x0e, "Invalid TargetFrameworkAttribute constructor.");
                    BlobReader value = md.GetBlobReader(attribute.Value);
                    Require(value.ReadUInt16() == 1, "Invalid TargetFrameworkAttribute prolog.");
                    Require(info.TargetFramework == null, "Duplicate TargetFrameworkAttribute.");
                    info.TargetFramework = value.ReadSerializedString();
                    ushort named = value.ReadUInt16();
                    for (int n = 0; n < named; n++) {
                        Require(value.ReadByte() == 0x54 && value.ReadByte() == 0x0e &&
                            value.ReadSerializedString() == "FrameworkDisplayName", "Invalid TargetFrameworkAttribute named property.");
                        value.ReadSerializedString();
                    }
                    Require(value.RemainingBytes == 0, "Trailing TargetFrameworkAttribute bytes.");
                }
                uint entry = D(info.CliHeaderOffset + 20);
                Require((info.CorFlags.Value & 0x10) != 0 || entry == 0 ||
                    ((entry >> 24) == 6 && (entry & 0xffffff) > 0 && (entry & 0xffffff) <= md.MethodDefinitions.Count), "Invalid managed entry point token.");
            }
#else
            // Loading bytes avoids ReflectionOnlyLoadFrom's path/identity cache
            // when a build is rebuilt and validated again in the same process.
            ResolveEventHandler resolve = delegate(object sender, ResolveEventArgs args) {
                string local = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(info.Path), new AssemblyName(args.Name).Name + ".dll");
                return File.Exists(local) ? Assembly.ReflectionOnlyLoadFrom(local) : Assembly.ReflectionOnlyLoad(args.Name);
            };
            AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve += resolve;
            try {
                Assembly assembly = Assembly.ReflectionOnlyLoad(bytes);
                info.MetadataVersion = assembly.ImageRuntimeVersion;
                if ((info.CorFlags.Value & 0x10) == 0 && D(info.CliHeaderOffset + 20) != 0)
                    Require(assembly.EntryPoint != null, "Invalid managed entry point token.");
                foreach (AssemblyName name in assembly.GetReferencedAssemblies()) {
                    PayloadReference item = new PayloadReference(name);
                    string arch = typeof(AssemblyName).GetProperty("ProcessorArchitecture").GetValue(name, null).ToString();
                    item.Flags = arch == "None" ? 0U : arch == "MSIL" ? 0x10U : 0x20U;
                    references.Add(item);
                }
                foreach (CustomAttributeData attribute in CustomAttributeData.GetCustomAttributes(assembly)) {
                    if (attribute.Constructor.DeclaringType != typeof(System.Runtime.Versioning.TargetFrameworkAttribute) &&
                        attribute.Constructor.DeclaringType.AssemblyQualifiedName != typeof(System.Runtime.Versioning.TargetFrameworkAttribute).AssemblyQualifiedName) continue;
                    Require(info.TargetFramework == null && attribute.ConstructorArguments.Count == 1, "Invalid or duplicate TargetFrameworkAttribute.");
                    info.TargetFramework = attribute.ConstructorArguments[0].Value as string;
                }
            } finally { AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve -= resolve; }
#endif
            info.AssemblyReferences = references.ToArray();
        }
        public static PayloadImage Inspect(string path) {
            PayloadInspector reader = new PayloadInspector();
            // Keep the read-sharing guard through metadata inspection so Windows
            // writers cannot change the binary between hashing and metadata APIs.
            using (FileStream guard = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                Require(guard.Length <= Int32.MaxValue, "PE image too large to inspect safely.");
                reader.bytes = new byte[(int)guard.Length];
                int read = 0;
                while (read < reader.bytes.Length) {
                    int count = guard.Read(reader.bytes, read, reader.bytes.Length - read);
                    Require(count > 0, "Truncated file during inspection."); read += count;
                }
                PayloadImage info = reader.Headers(path);
                if (info.IsManaged) reader.Metadata(info);
                return info;
            }
        }
    }
}
'@
    if ($PSVersionTable.PSEdition -eq 'Core') {
        # These reference assemblies ship with pwsh; nothing is downloaded.
        Add-Type -TypeDefinition $source -CompilerOptions '/define:CORECLR' -ReferencedAssemblies @(
            Get-ChildItem -LiteralPath (Join-Path $PSHOME 'ref') -Filter '*.dll' | ForEach-Object { $_.FullName }
        ) -ErrorAction Stop
    }
    else { Add-Type -TypeDefinition $source -ErrorAction Stop }
}

function Get-PeImageInfo {
    <#
    .SYNOPSIS
    Reports PE/CLI headers, SHA256 and assembly metadata without running payload code.
    .DESCRIPTION
    Native PE images (including clr.dll) have IsManaged=false and CorFlags=$null.
    Uses Framework reflection-only APIs or pwsh's bundled metadata reader.
    Structural inspection is not IL verification or cryptographic signature verification.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)
    try {
        $file = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
        if ($file.PSIsContainer) { throw 'Expected a file, not a directory.' }
        [OneNoteMarkdown.PayloadInspector]::Inspect($file.FullName)
    }
    catch { throw "Cannot inspect PE/CLR image '$Path': $($_.Exception.GetBaseException().Message) Rebuild or obtain the correct original binary; never patch signed files." }
}

function Assert-ArmAnyCpuPayload {
    <#
    .SYNOPSIS
    Validates every shipped DLL and returns per-file read-only inspection reports.
    .DESCRIPTION
    Required DLLs must be directly under Directory. Also inspects optional, hidden
    and nested DLLs. Use -Verbose for per-file hashes/flags/identity on failure.
    Legacy Framework interop references are allowed; no binary is rewritten.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Directory)
    $folder = Get-Item -LiteralPath $Directory -Force -ErrorAction Stop
    if (-not $folder.PSIsContainer) { throw "Payload directory is not a directory: '$Directory'." }
    $problems = New-Object 'System.Collections.Generic.List[string]'
    $reports = New-Object 'System.Collections.Generic.List[object]'
    foreach ($name in @('OneNoteMarkdown.AddIn.dll', 'WpfMath.dll', 'XamlMath.Shared.dll', 'Extensibility.dll', 'office.dll', 'Microsoft.Office.Interop.OneNote.dll')) {
        if (-not (Test-Path -LiteralPath (Join-Path $folder.FullName $name) -PathType Leaf)) { $problems.Add("Missing required payload '$name' directly under '$($folder.FullName)'.") }
    }
    $byName = @{}
    foreach ($file in @(Get-ChildItem -LiteralPath $folder.FullName -Recurse -File -Force -ErrorAction Stop | Where-Object { $_.Extension -ieq '.dll' } | Sort-Object FullName)) {
        try {
            $info = Get-PeImageInfo $file.FullName
            $reports.Add($info)
            $flags = if ($null -eq $info.CorFlags) { 'none' } else { '0x{0:X8}' -f $info.CorFlags }
            $summary = "'$($info.Path)' SHA256=$($info.SHA256) Machine=0x{0:X4} PE=0x{1:X3} CorFlags={2} Identity={3}" -f $info.Machine, $info.OptionalHeaderMagic, $flags, $info.AssemblyIdentity
            Write-Verbose $summary
            if (-not $info.IsManaged) { $problems.Add("$summary : Native DLLs cannot be shipped in the ARM AnyCPU payload."); continue }
            if ($info.Machine -ne 0x014c -or $info.OptionalHeaderMagic -ne 0x10b) { $problems.Add("$summary : Require I386 (0x014C), PE32 (0x10B) managed AnyCPU.") }
            if (($info.CorFlags -band 1) -eq 0 -or ($info.CorFlags -band 0x20002) -ne 0) { $problems.Add("$summary : Require ILONLY, with 32BITREQUIRED and 32BITPREFERRED clear.") }
            if (($info.CorFlags -band 0x10) -ne 0 -or $info.HasNativeCode) { $problems.Add("$summary : Native entry points, mixed-mode and precompiled native code are not AnyCPU.") }
            if (($info.CorFlags -band 0xfffcffe0L) -ne 0) { $problems.Add("$summary : Unknown/reserved CLR flags.") }
            if ($info.ProcessorArchitecture -notin @('None', 'MSIL')) { $problems.Add("$summary : AssemblyName processor architecture '$($info.ProcessorArchitecture)' is not MSIL/neutral.") }
            if (($info.Characteristics -band 0x2000) -eq 0) { $problems.Add("$summary : Expected a managed DLL, not a renamed executable.") }
            if ($file.BaseName -ine $info.AssemblyName) { $problems.Add("$summary : File name differs from assembly simple name '$($info.AssemblyName)'.") }
            if ($byName.ContainsKey($info.AssemblyName)) { $problems.Add("$summary : Duplicate assembly simple name.") } else { $byName[$info.AssemblyName] = $info }
            if ($file.Name -ieq 'OneNoteMarkdown.AddIn.dll' -and $info.TargetFramework -cne '.NETFramework,Version=v4.8.1') {
                $problems.Add("$summary : Main TargetFrameworkAttribute must be '.NETFramework,Version=v4.8.1'; actual '$($info.TargetFramework)'.")
            }
            elseif ($info.TargetFramework -and ($info.TargetFramework -notmatch '^\.NETFramework,Version=v([0-9]+(\.[0-9]+){1,3})$' -or [version]$Matches[1] -gt [version]'4.8.1')) {
                $problems.Add("$summary : TargetFrameworkAttribute '$($info.TargetFramework)' is not Framework 4.8.1 compatible.")
            }
        }
        catch { $problems.Add($_.Exception.Message) }
    }
    foreach ($info in $reports) {
        if (-not $info.IsManaged) { continue }
        $core = @($info.AssemblyReferences | Where-Object { $_.Name -ieq 'mscorlib' })
        if ($core.Count -ne 1 -or $core[0].Version -notin @('1.0.3300.0', '1.0.5000.0', '2.0.0.0', '4.0.0.0') -or $core[0].PublicKeyToken -ne 'b77a5c561934e089') {
            $problems.Add("'$($info.Path)' must reference signed .NET Framework mscorlib v1/v2/v4, not CoreCLR.")
        }
        if ($info.AssemblyName -ieq 'OneNoteMarkdown.AddIn' -and ($core.Count -ne 1 -or $core[0].Version -ne '4.0.0.0')) { $problems.Add("'$($info.Path)' main assembly must reference mscorlib, Version=4.0.0.0.") }
        foreach ($reference in $info.AssemblyReferences) {
            if ((($reference.Flags -band 0x70) -shr 4) -notin @(0, 1) -or $reference.Name -in @('System.Private.CoreLib', 'System.Runtime.Loader')) { $problems.Add("'$($info.Path)' has incompatible AssemblyRef '$($reference.Identity)'.") }
            if ($byName.ContainsKey($reference.Name)) {
                $dependency = $byName[$reference.Name]
                if ($reference.Identity -cne $dependency.AssemblyIdentity) { $problems.Add("'$($info.Path)' references '$($reference.Identity)' but shipped '$($dependency.Path)' is '$($dependency.AssemblyIdentity)'.") }
            }
            else {
                $framework = $reference.Name -match '^(mscorlib|netstandard|System(\..+)?|Microsoft\.CSharp|WindowsBase|PresentationCore|PresentationFramework|Accessibility|UIAutomation(Client|Types|Provider))$'
                if (-not $framework -or $reference.PublicKeyToken -notin @('b77a5c561934e089', 'b03f5f7f11d50a3a', '31bf3856ad364e35', 'cc7b13ffcd2ddd51')) {
                    $problems.Add("'$($info.Path)' has unresolved non-framework AssemblyRef '$($reference.Identity)'. Include its compatible AnyCPU DLL.")
                }
                elseif ([version]$reference.Version -gt [version]'4.0.0.0' -or ($reference.Name -ieq 'netstandard' -and [version]$reference.Version -gt [version]'2.0.0.0')) { $problems.Add("'$($info.Path)' has incompatible framework AssemblyRef '$($reference.Identity)'.") }
            }
        }
    }
    if ($problems.Count) { throw ("ARM AnyCPU payload validation failed (read-only; no files modified):`n - " + ($problems -join "`n - ") + "`nRebuild or replace incompatible originals; never patch signed DLL headers/CorFlags.") }
    $reports.ToArray()
}
