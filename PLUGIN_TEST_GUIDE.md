# OneNote Markdown 插件测试说明

## 1. 构建与安装

1. 执行 `powershell -ExecutionPolicy Bypass -File .\build.ps1`。
2. 按 OneNote 架构安装 `OneNoteMarkdownSetup-1.2.1-fix.1-x86.exe`、`OneNoteMarkdownSetup-1.2.1-fix.1-x64.exe` 或 `OneNoteMarkdownSetup-1.2.1-fix.1-arm64.exe`。
3. 启动 OneNote，确认 `Markdown` 功能区可见且启动无“无法顺利启动”提示。

## 2. 自动回归

`OneNoteMarkdown.Tests` 应全部通过，至少覆盖：

- 分隔线与列表优先级；
- 不同长度代码围栏；
- 单行块级 LaTeX；
- 表格和相对图片路径；
- 原地多元素源码恢复安全；
- 托管预览导出排重。

## 3. 对照预览

准备两个独立正文框，并分别输入 Markdown。

| 用例 | 操作 | 预期 |
|---|---|---|
| 整页预览 | 点击“渲染整页”，连续按三次 `F5` | 只有一个整页预览，内容原位更新，不包含旧预览 |
| 正文框预览 | 光标放入每个正文框并点击“渲染正文框” | 两个来源各自维护独立预览 |
| 选区预览 | 完整选中 Markdown 后点击“渲染选区” | 预览关联选区；无真实选区时仅提示，不处理整页 |
| 布局 | 将位置设为 `right`/`below` 并新建预览 | 预览位于来源右侧/下方，顶部或左侧对齐 |
| 保留调整 | 手动移动预览、调整宽度并刷新 | 位置和宽度保持 |
| 重置布局 | 点击“重置布局” | 整页预览按当前间距和宽度重新放置 |
| 标题 | 修改或删除标题后刷新 | 不强制恢复；隐藏标题时无空标题行 |
| 删除 | 光标置于预览并点击“删除预览” | 只删除托管预览，来源保持不变 |

## 4. 刷新、并发与关闭

1. 开启实时模式，在一行 Markdown 后快速连续回车和切页。
2. 预期：只处理最后一个有效任务，不向新页面写入旧页面结果。
3. 关闭实时模式并立即退出 OneNote。
4. 预期：无卡死，`ONENOTE.EXE` 正常退出。
5. 启用 `preview.autoRefresh=true` 后重复快速输入。
6. 预期：输入被防抖合并，只刷新活动页对照预览；不同时原地渲染同一来源。
7. 手动修改预览正文后触发自动刷新。
8. 预期：不覆盖；手动刷新时显示覆盖确认。只移动、改宽度或改标题时不提示冲突。

## 5. Markdown 内容

验证以下样例：

```markdown
# 标题

* * *

- [x] 任务
  - 嵌套项

| 名称 | 值 |
|---|---:|
| alpha | 1 |

[GitHub](https://github.com)
![本地图片](images/demo.png)

````csharp
```
var x = 1;
````

$$E=mc^2$$

```mermaid
flowchart TD
  A[开始] --> B{通过?}
  B -->|是| C[完成]
```
```

预期：

- 分隔线不变成列表；
- 四反引号围栏内的三反引号不会提前关闭；
- 表格生成 OneNote 表格；
- 链接保留文字和 HTTP/HTTPS 目标；
- 导入文件的相对图片能显示；
- 单行公式成功渲染，错误公式显示原因并保留源码；
- Mermaid 流程图离线生成图片；不支持语句回退为完整围栏源码。

## 6. 导入导出

1. 导入包含标题、空行、缩进、代码、表格、公式、Mermaid 和相对图片的 `.md`。
2. 立即导出，比较 Markdown 结构和必要空行/缩进。
3. 预期：托管预览不重复导出，插件标题不导出，原始 Mermaid/公式源码保留。
4. 手动修改导入预览正文后导出。
5. 预期：提示将使用保存的原始 Markdown；取消时不产生成功结果。
6. 验证普通 OneNote 富文本采用尽力转换，不支持对象有明确提示或保留，不静默宣称完全无损。

## 7. 隐私、资源与兼容性

- 默认 `image.allowRemote=false`；确认不会请求远程图片。
- 启用后仅测试 HTTP/HTTPS，并验证 10 MB 与超时限制。
- 使用超大 Mermaid、无效公式和无效图片，确认单个失败不阻断其他块。
- 检查日志不包含笔记正文或 Markdown 源码。
- 分别验证 x86/x64/Arm64 安装包和旧 `theme.ini`；Arm64 使用 Windows 11 + .NET Framework 4.8.1，缺失新键时使用默认值。
- 标题恰好为 `Markdown Render` 的普通正文框不得被自动认领或删除。

## 8. 结果记录

```text
测试时间：
OneNote 版本与位数：
系统版本：

启动/退出：
整页/正文框/选区预览：
刷新取消与冲突：
表格/链接/图片：
公式/Mermaid：
导入导出：
x86/x64：
日志隐私：

失败现象与日志摘要：
```

## 9. ARM AnyCPU 独立实验 / Experimental compatibility probe

**这是诊断实验，不是已确认的 ARM64EC OneNote 修复。** 正式 x86/x64/arm64
构建默认不变。AnyCPU 是纯 IL 托管载荷，不是 ARM64 原生 DLL；安装器仍仅允许
Windows 11 ARM64，使用 HKCR64，并要求 .NET Framework 4.8.1（Release ≥ 533320）。
保留程序集身份、CLSID、ProgID 及启动逻辑，不引入原生 shim 或 .NET 迁移。

### 构建与获取

Windows 开发机需要 Visual Studio Build Tools（MSBuild、C# 桌面构建工具）、
.NET Framework 4.8 和 4.8.1 Developer Pack、Inno Setup 6（含现有中文语言文件）。
在仓库根目录运行：

```powershell
# 默认正式包：保持原有接口
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1
# 仅实验包；不使用 -ReleaseName，也不同时构建正式包
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -ArmCompatibilityValidation
# 后续实验序号只能使用明确的 armcompat 预发布版本
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -ArmCompatibilityValidation -ValidationReleaseName 1.2.1-armcompat.2
```

主 DLL 使用 `Platform=AnyCPU / PlatformTarget=AnyCPU / Prefer32Bit=false / v4.8.1`，
输出到 `src\OneNoteMarkdown.AddIn\bin\Release\arm-anycpu-probe`，
中间文件隔离到 `obj\arm-anycpu-probe` 和 `obj\Release\arm-anycpu-probe`。
构建会先运行既有回归和载荷检查测试，逐一检查随包 DLL（Office interop、
Extensibility、stdole 如存在、WpfMath、XamlMath.Shared）的 PE/CLR 标志和程序集
元数据，确认纯 IL、无 32 位限定/偏好，并验证主程序集目标框架。
检查失败时停止打包：获取适用的上游纯 IL 依赖或继续调查，**不要改写 CorFlags、
PE 头或已签名第三方文件**。生产包仍执行原来的架构 PE 检查。

也可在 GitHub Actions → Build → Run workflow 勾选
`arm_compatibility_validation`。该手动路径仅上传独立实验 artifact，
**不会附加到正常 v* tag Release**。x64 runner 只证明构建和元数据检查，
不证明 ARM64 硬件、CLR 或 Office 行为。

默认实验包名为 `OneNoteMarkdownSetup-1.2.1-armcompat.1-arm64.exe`。
`src\OneNoteMarkdown.Installer\Output\arm-anycpu-probe` 中包含安装器、
完整未安装 `payload`、两个诊断脚本、本说明和 `SHA256SUMS.txt`；
复制/下载并解压整个目录，不要只复制主 DLL。可运行
`Get-ChildItem -Recurse -File | Get-FileHash -Algorithm SHA256` 与清单比对。
只对可信来源的载荷运行加载/构造测试。

### 安装前：备份、检查，禁止并排注册

实验与正式包共用 AppId、CLSID `{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}`、
ProgID `OneNoteMarkdown.Connect` 和默认安装目录，**不能并排安装**。
先退出 OneNote 并确认无 `ONENOTE.EXE`，保存当前正式安装器及插件配置
（`%APPDATA%\OneNoteMarkdown`，自行保管，不上传），完成下面的只读检查/备份。
第一阶段可在独立目录直接测试，无需安装；第二阶段开始前必须卸载旧插件再安装实验包。

此前手工脚本可能留下 `Office\16.0\OneNote\AddIns` 重复项和
`HKCU\Software\Classes` COM 覆盖；重装机器级包不一定移除它们。
在原用户的 ARM64 PowerShell 中执行以下**仅限本插件**的检查/导出
（备份文件含本机路径，勿直接公开）：

```powershell
$backup = Join-Path $env:USERPROFILE ('OneNoteMarkdown-reg-backup-' + (Get-Date -Format yyyyMMdd-HHmmss))
New-Item -ItemType Directory $backup | Out-Null
$keys = @(
  'HKCU\Software\Microsoft\Office\OneNote\AddIns\OneNoteMarkdown.Connect',
  'HKCU\Software\Microsoft\Office\16.0\OneNote\AddIns\OneNoteMarkdown.Connect',
  'HKCU\Software\Classes\OneNoteMarkdown.Connect',
  'HKCU\Software\Classes\OneNoteMarkdown.Connect.1',
  'HKCU\Software\Classes\CLSID\{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}',
  'HKCU\Software\Classes\AppID\{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}',
  'HKCU\Software\Classes\AppID\OneNoteMarkdown.AddIn.dll',
  'HKCU\Software\Classes\TypeLib\{5C3E37E8-7D8A-41E5-9D3D-6B5A1C92B7D1}'
)
for ($i = 0; $i -lt $keys.Count; $i++) {
  & reg.exe query $keys[$i] /s /reg:64
  if ($LASTEXITCODE -eq 0) { & reg.exe export $keys[$i] (Join-Path $backup "$i.reg") /y }
}
& reg.exe query 'HKCR\CLSID\{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}\InprocServer32' /s /reg:64
& reg.exe query 'HKLM\Software\Classes\CLSID\{0A92B61B-98B8-4E5D-BE2D-48EDB01ED177}\InprocServer32' /s /reg:64
```

键不存在的查询报错正常。HKCR 是合并视图，必须比较 HKCU/HKLM 的 Class、
Assembly、CodeBase 与实际安装文件。**不要批量删除上述键**：如确认某个
用户级覆盖或重复项是此前误建且妨碍实验，先核实导出成功，再在注册表编辑器中
只处理那个插件项并记录操作；不确定就停止并寻求帮助。不要清空 Office
Resiliency，不自动 RegAsm，不启用全局 Fusion 注册表日志。
实际观测的 OneNote 读取非版本路径 `Office\OneNote\AddIns`；
`LoadBehavior=3` 不等于成功加载，`2` 也不能证明策略封禁。

### 第一关：ARM64 CLR（独立于 OneNote）

使用**全新原生 ARM64 Windows PowerShell 5.1**（不是 PowerShell 7，也不是
SysWOW64/x64 仿真宿主）。从原生终端启动下列命令，按解压位置替换路径：

```powershell
# 假设完整 artifact 已解压到 C:\ArmProbe
& "$env:windir\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File C:\ArmProbe\Test-ArmCompatibility.ps1 -AssemblyPath C:\ArmProbe\payload\OneNoteMarkdown.AddIn.dll
# 上一步通过后，重新开进程，显式允许 ctor（会安装 AssemblyResolve 并可能写日志）
& "$env:windir\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass -File C:\ArmProbe\Test-ArmCompatibility.ps1 -AssemblyPath C:\ArmProbe\payload\OneNoteMarkdown.AddIn.dll -InstantiateConnect
```

脚本报告**实际已加载 clr.dll 路径和 PE Machine**（应为 `0xAA64`，通常来自
`FrameworkArm64\v4.0.30319`），而非依赖可能为空的 ProcessArchitecture。
默认仅报告路径/哈希并 `Assembly.LoadFrom`、解析 Connect 类型，不改注册表、
不实例化插件；异常含嵌套类型、HRESULT、FileName、已有 FusionLog 及
ReflectionTypeLoadException.LoaderExceptions。构造测试返回后还需人工确认
`%APPDATA%\OneNoteMarkdown\logs\onenotemarkdown.log` 出现新的
`Connect ctor invoked`（构造函数可能吞掉日志初始化异常）。
脚本不读取笔记内容；分享输出前脱敏用户名、内部路径及已有 FusionLog 中的敏感信息。

安装后可在新进程中另加 `-ComActivation`，显式测试注册 CLSID 的 COM 激活。
此开关会执行构造函数/写日志，可能使用注册到其他目录的 DLL；先检查 CodeBase。
它不调用 OneNote 的 OnConnection，**不代替真实宿主验收**。
如第一关失败，停止安装测试，保留具体异常和依赖名称。

### 第二关：实际 ARM64EC OneNote（待硬件人工验证）

- [ ] 第一关 ARM64 CLR LoadFrom、Connect 类型、显式 ctor 及新日志全部通过。
- [ ] 卸载旧包、安装实验包后，在实际 OneNote 中启用插件；捕获真实 COM/CLR 激活证据。
- [ ] 新日志既有 `Connect ctor invoked`，又有 `OneNoteMarkdown connected`；功能区正常。
- [ ] 渲染、LaTeX、设置、导入、导出、快捷键及上述回归用例通过。
- [ ] 至少 3 次完全退出/启动，无崩溃、残留进程或再次禁用。
- [ ] 原有 x86/x64 正式包的既有测试、安装及功能回归通过。

已知 x64 DLL 在 ARM64 PowerShell 报 `0x8007000B` 只证明该 PowerShell
架构不匹配，**不是 OneNote 的已证实根因**。无插件日志/无 ProcMon 文件事件
也不足以断言没有加载尝试；捕获前启用文件系统与进程/线程类别，以
`Process Name is ONENOTE.EXE` 开始，确认能看到 CreateFile/Load Image。
同字段多个 Include 不是 AND。不要从 LoadBehavior 或缺失事件猜测实际 HRESULT。
若 PowerShell 通过而 OneNote 失败，停止声称 AnyCPU 能修复，转查实际宿主
COM/CLR 激活和可信 DLL-load/HRESULT 证据。现有安装器还注册了未随包提供的
`.tlb` 路径；这是独立待调查事项，本实验不凭推测修改生产 TypeLib 行为。

### 回滚

完全退出 OneNote → 卸载实验包 → 重装先前保存的、适合原环境的正式包。
不要直接覆盖 DLL 或在两个安装包间并排切换。若手动移除过确定的用户覆盖项，
可在退出 OneNote 后用 `reg.exe import <对应备份.reg>` 恢复**仅该项**；
先核对 CodeBase，避免把过期误建项重新引入。配置如需恢复，仅恢复自己的备份。
不做任何全局注册表清理。正式包在该 ARM64EC 宿主上此前也可能无法激活；
回滚只是恢复基线，并不保证 OneNote 兼容。
