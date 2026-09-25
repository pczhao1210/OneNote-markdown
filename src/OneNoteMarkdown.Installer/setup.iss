#ifndef InstallerArch
  #define InstallerArch "x64"
#endif
#ifndef ReleaseName
  #define ReleaseName "1.2.1-fix.2"
#endif
#ifndef BuildConfiguration
  #define BuildConfiguration "Release"
#endif

#if InstallerArch == "x86"
  #define InstallerSuffix "x86"
  #define InstallerDefaultDir "{autopf32}\OneNoteMarkdown"
  #define InstallerArchitectures "x86compatible"
  #define ComRoot "HKCR32"
  #define TypeLibPlatform "win32"
  #define RequiredDotNetRelease 528040
  #define RequiredDotNetName ".NET Framework 4.8"
#elif InstallerArch == "x64"
  #define InstallerSuffix "x64"
  #define InstallerDefaultDir "{autopf64}\OneNoteMarkdown"
  #define InstallerArchitectures "x64compatible"
  #define ComRoot "HKCR64"
  #define TypeLibPlatform "win64"
  #define RequiredDotNetRelease 528040
  #define RequiredDotNetName ".NET Framework 4.8"
#elif InstallerArch == "arm64"
  #define InstallerSuffix "arm64"
  #define InstallerDefaultDir "{autopf64}\OneNoteMarkdown"
  #define InstallerArchitectures "arm64"
  #define ComRoot "HKCR64"
  #define TypeLibPlatform "win64"
  #define RequiredDotNetRelease 533320
  #define RequiredDotNetName ".NET Framework 4.8.1"
#else
  #error Unsupported InstallerArch
#endif

#define AddInOutput "..\OneNoteMarkdown.AddIn\bin\" + BuildConfiguration + "\" + InstallerSuffix

[Setup]
AppId={{7D0C9CB6-5D83-4E17-89D5-1C8A285F52B4}
AppName=OneNote Markdown
AppVersion={#ReleaseName}
AppVerName=OneNote Markdown {#ReleaseName}
AppPublisher=OneNote Markdown
DefaultDirName={#InstallerDefaultDir}
DefaultGroupName=OneNote Markdown
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=OneNoteMarkdownSetup-{#ReleaseName}-{#InstallerSuffix}
ArchitecturesAllowed={#InstallerArchitectures}
#if InstallerArch != "x86"
ArchitecturesInstallIn64BitMode={#InstallerArchitectures}
#endif
#if InstallerArch == "arm64"
MinVersion=10.0.22000
#endif
PrivilegesRequired=admin
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
VersionInfoVersion=1.2.1.2
VersionInfoProductVersion=1.2.1.2

[Languages]
Name: "chinesesimplified"; MessagesFile: "ChineseSimplified.isl"

[Files]
Source: "{#AddInOutput}\OneNoteMarkdown.AddIn.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddInOutput}\WpfMath.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddInOutput}\XamlMath.Shared.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddInOutput}\Extensibility.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddInOutput}\office.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddInOutput}\Microsoft.Office.Interop.OneNote.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#AddInOutput}\stdole.dll"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "{#AddInOutput}\*.pdb"; DestDir: "{app}"; Flags: ignoreversion skipifsourcedoesntexist
Source: "..\..\HELP.md"; DestDir: "{app}"; Flags: ignoreversion

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Office\OneNote\AddIns\OneNoteMarkdown.Connect"; ValueType: string; ValueName: "Description"; ValueData: "OneNote Markdown - Markdown 工具插件"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\OneNote\AddIns\OneNoteMarkdown.Connect"; ValueType: string; ValueName: "FriendlyName"; ValueData: "OneNote Markdown"
Root: HKCU; Subkey: "Software\Microsoft\Office\OneNote\AddIns\OneNoteMarkdown.Connect"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: "3"

Root: {#ComRoot}; Subkey: "CLSID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}"; ValueType: string; ValueData: "OneNoteMarkdown.AddIn.Connect"; Flags: uninsdeletekey
Root: {#ComRoot}; Subkey: "CLSID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}"; ValueType: string; ValueName: "AppID"; ValueData: "{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}"
Root: {#ComRoot}; Subkey: "CLSID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}\Implemented Categories\{{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"; Flags: uninsdeletekey
Root: {#ComRoot}; Subkey: "CLSID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}\InprocServer32"; ValueType: string; ValueData: "mscoree.dll"; Flags: uninsdeletekey
Root: {#ComRoot}; Subkey: "CLSID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}\InprocServer32"; ValueType: string; ValueName: "ThreadingModel"; ValueData: "Both"
Root: {#ComRoot}; Subkey: "CLSID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}\InprocServer32"; ValueType: string; ValueName: "Class"; ValueData: "OneNoteMarkdown.AddIn.Connect"
Root: {#ComRoot}; Subkey: "CLSID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}\InprocServer32"; ValueType: string; ValueName: "Assembly"; ValueData: "OneNoteMarkdown.AddIn, Version=1.2.1.0, Culture=neutral, PublicKeyToken=null"
Root: {#ComRoot}; Subkey: "CLSID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}\InprocServer32"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: {#ComRoot}; Subkey: "CLSID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}\InprocServer32"; ValueType: string; ValueName: "CodeBase"; ValueData: "{app}\OneNoteMarkdown.AddIn.dll"
Root: {#ComRoot}; Subkey: "CLSID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}\InprocServer32\1.2.1.0"; ValueType: string; ValueName: "Class"; ValueData: "OneNoteMarkdown.AddIn.Connect"
Root: {#ComRoot}; Subkey: "CLSID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}\InprocServer32\1.2.1.0"; ValueType: string; ValueName: "Assembly"; ValueData: "OneNoteMarkdown.AddIn, Version=1.2.1.0, Culture=neutral, PublicKeyToken=null"
Root: {#ComRoot}; Subkey: "CLSID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}\InprocServer32\1.2.1.0"; ValueType: string; ValueName: "RuntimeVersion"; ValueData: "v4.0.30319"
Root: {#ComRoot}; Subkey: "CLSID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}\InprocServer32\1.2.1.0"; ValueType: string; ValueName: "CodeBase"; ValueData: "{app}\OneNoteMarkdown.AddIn.dll"
Root: {#ComRoot}; Subkey: "CLSID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}\ProgId"; ValueType: string; ValueData: "OneNoteMarkdown.Connect"
Root: {#ComRoot}; Subkey: "CLSID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}\Programmable"; ValueType: string; ValueData: ""
Root: {#ComRoot}; Subkey: "CLSID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}\TypeLib"; ValueType: string; ValueData: "{{5C3E37E8-7D8A-41E5-9D3D-6B5A1C92B7D1}"
Root: {#ComRoot}; Subkey: "CLSID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}\VersionIndependentProgID"; ValueType: string; ValueData: "OneNoteMarkdown.Connect"

Root: {#ComRoot}; Subkey: "AppID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}"; ValueType: string; ValueData: "OneNoteMarkdown.AddIn"; Flags: uninsdeletekey
Root: {#ComRoot}; Subkey: "AppID\{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}"; ValueType: string; ValueName: "DllSurrogate"; ValueData: ""
Root: {#ComRoot}; Subkey: "AppID\OneNoteMarkdown.AddIn.dll"; ValueType: string; ValueName: "AppID"; ValueData: "{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}"; Flags: uninsdeletekey

Root: {#ComRoot}; Subkey: "TypeLib\{{5C3E37E8-7D8A-41E5-9D3D-6B5A1C92B7D1}\1.0"; ValueType: string; ValueData: "OneNote Markdown 插件"; Flags: uninsdeletekey
Root: {#ComRoot}; Subkey: "TypeLib\{{5C3E37E8-7D8A-41E5-9D3D-6B5A1C92B7D1}\1.0\0\{#TypeLibPlatform}"; ValueType: string; ValueData: "{app}\OneNoteMarkdown.AddIn.tlb"
Root: {#ComRoot}; Subkey: "TypeLib\{{5C3E37E8-7D8A-41E5-9D3D-6B5A1C92B7D1}\1.0\FLAGS"; ValueType: string; ValueData: "0"
Root: {#ComRoot}; Subkey: "TypeLib\{{5C3E37E8-7D8A-41E5-9D3D-6B5A1C92B7D1}\1.0\HELPDIR"; ValueType: string; ValueData: "{app}"

Root: {#ComRoot}; Subkey: "OneNoteMarkdown.Connect"; ValueType: string; ValueData: "OneNoteMarkdown.AddIn.Connect"; Flags: uninsdeletekey
Root: {#ComRoot}; Subkey: "OneNoteMarkdown.Connect\CLSID"; ValueType: string; ValueData: "{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}"
Root: {#ComRoot}; Subkey: "OneNoteMarkdown.Connect\CurVer"; ValueType: string; ValueData: "OneNoteMarkdown.Connect.1"
Root: {#ComRoot}; Subkey: "OneNoteMarkdown.Connect.1"; ValueType: string; ValueData: "OneNoteMarkdown.AddIn.Connect"; Flags: uninsdeletekey
Root: {#ComRoot}; Subkey: "OneNoteMarkdown.Connect.1\CLSID"; ValueType: string; ValueData: "{{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}"

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Code]
const
  RequiredDotNetRelease = {#RequiredDotNetRelease};

function IsRequiredDotNetInstalled: Boolean;
var
  Release: Cardinal;
begin
  Result := RegQueryDWordValue(HKLM32, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release) and
    (Release >= RequiredDotNetRelease);

  if (not Result) and IsWin64 then
    Result := RegQueryDWordValue(HKLM64, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release) and
      (Release >= RequiredDotNetRelease);
end;

function InitializeSetup: Boolean;
begin
  Result := IsRequiredDotNetInstalled;
  if not Result then
  begin
    MsgBox('需要先安装 {#RequiredDotNetName} 才能继续安装 OneNote Markdown。', mbCriticalError, MB_OK);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
  begin
    RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER, 'Software\Microsoft\Office\OneNote\AddInsData\OneNoteMarkdown.Connect');
  end;
end;
