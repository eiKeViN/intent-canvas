# 首轮编辑能力验证

状态：2026-10-19 已接受“复用成熟编辑器、保留原生文档、向 AI 提供任务视图”的验证方向。这里是能力探针与实验材料，尚不是意绘应用，也没有任何编辑器兼容性通过结论。

PPT 比较 PowerPoint 原生接口与 ONLYOFFICE，矢量交互先以 Penpot 官方 MCP 为基线。先在一个宿主中验证价值，再决定产品形态。两种正式入口（生成初稿／打开已有 PPT）和深入编辑要求保持不变。

## 已准备及实际验证情况

| 材料 | 用途 | 本次检查 |
| --- | --- | --- |
| `powerpoint-inspect.ps1` | 从已打开的指定文稿读取形状 ID、分组子对象、文字摘要、图表系列和当前选区 | 按微软公开 API 编写；当前 Linux 无 PowerPoint／Windows PowerShell，未做宿主执行与 PowerShell 解析验证 |
| `onlyoffice/` | 只读插件：读取当前页与选区、图表系列数量，显式记录缺失接口 | JavaScript 语法及 JSON 格式检查通过；未安装、未连接编辑器 |
| `pptx_parts.py` | 比较 PPTX 解压后各部件的内容摘要及部分 XML 结构，不修改输入 | Python 3.12 下 6 项合成测试通过；命令行生成报告及拒绝覆盖已有文件的检查通过；这不是实际编辑器往返测试 |
| `three-steps.svg` | 嵌套分组、文字、跨区箭头和干扰对象的统一样例 | XML 可解析、ID 唯一、4 个文字节点；未在 Penpot／Inkscape 实测 |

本次没有安装产品依赖、购买授权或接入模型。当前插件目录检索没有找到相应编辑器连接，因此不能在本会话内完成宿主实验。更多连接可能出现在[插件目录](https://chatgpt.com/plugins)，本地验证不依赖该目录。

## 本地 Codex 从这里继续

在实际使用 PPT／SVG 的电脑上打开本仓库，让 Codex 阅读本文件即可继续。建议给它的任务是：

> 继续意绘的首轮验证。先读 AGENTS.md、README.md 和 validation/README.md，检查本机已有 PowerPoint／ONLYOFFICE 与 Penpot 连接。使用单独的工作副本执行以下实验，修复探针并记录真实结果。保留用户已经打开的其他文档；不把生成的报告或用户文件提交到公开仓库。接口缺失或测试失败应如实记录，不以重画、栅格化或先拆组重组代替通过。先完成一个宿主的对象读取、定向修改、撤销、保存重开闭环，再比较另一候选；不要先搭建完整产品或多宿主框架。需要新账号、商业授权或原生宿主不可用时，记录具体缺口并完成其他可执行项。

本地运行产物放在 `.local/validation/`，已由仓库忽略。探针输出可能包含对象名称；PowerPoint 的 `-IncludeContent` 还会包含原文和图表数值。这里只提交人工审阅后的能力结论，不提交原始报告。

## 1. 选择验证文件

PPT 使用一份在候选内核之外制作的真实 `.pptx`，另存工作副本。至少覆盖：嵌套分组内的文字、原生图表及其嵌入工作簿、表格、连接符、母版／主题。单份文件不具备的类型另补小样例；不能把缺少对象记作通过。首轮工具只检查通常路径布局的未加密 PPTX；旧二进制 `.ppt` 和其他格式仍待决定。

SVG 使用随附 `three-steps.svg` 做交互基线，并补充一份真实文件检查字体、路径、变换和复杂效果。自制样例不能代替外来文件兼容性测试。

## 2. PowerPoint：原生对象读取

在 Windows PowerPoint 中打开工作副本并激活其窗口。使用 Windows PowerShell 5.1（`powershell.exe`），脚本的 `PresentationPath` 必须与当前活动文稿一致：

```powershell
powershell.exe -NoProfile -File .\validation\powerpoint-inspect.ps1 -PresentationPath "C:\work\sample.pptx" -OutputPath ".\.local\validation\powerpoint-before.json"
```

探针只附着已有 PowerPoint 进程，不打开其他文件，不保存、关闭或退出 PowerPoint。输出路径必须是尚不存在的 JSON 文件。若本机执行策略阻止脚本，按本机现有开发规范处理，不修改系统级策略来绕过限制。

选择组内文字、选择图表后分别再运行，换一个输出文件名，检查选区是否落到预期子对象。读取期间暂停手动操作；本探针不是原子快照、实时监听器或完整对象序列化器。逐项检查 `errors`，不能把缺字段当成没有对象。分组几何是原生读数，尚未转换为屏幕坐标。

下一步在工作副本中实现小范围写入实验：以文稿身份、SlideID、形状 ID／父级链定位对象，再核对旧值；只改指定文本片段，检查其他富文本样式。图表实验必须核对嵌入工作簿及公式，不能仅给 `Series.Values` 赋字面数组就宣称数据关系保持。将一次 AI 操作与宿主撤销机制结合，实际按撤销检查边界。当前脚本只实现读取，这些写入实验尚未实现或执行。

官方接口依据：[SlideID](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.slideid)、[GroupItems](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.groupitems)、[ChildShapeRange](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.selection.childshaperange)、[Series.Values](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.series.values)。

## 3. ONLYOFFICE：公共接口探针

在已有合适授权的 Desktop Editors 或 Docs 测试环境中使用。`onlyoffice/` 包含一个常规演示文稿插件；不需要外部 Automation API 连接器。按[官方安装与测试说明](https://api.onlyoffice.com/docs/plugins/development-workflow/installing-and-testing/desktop-editors-installation)安装，记录宿主版本及版本类型；该页侧栏另有 Docs 部署说明。若需要 `.plugin` 包，在仓库根目录执行：

```sh
python -c "from pathlib import Path; import zipfile; out=Path('.local/validation/intentcanvas-probe.plugin'); out.parent.mkdir(parents=True,exist_ok=True); z=zipfile.ZipFile(out,'x'); [z.write(p,p.name) for p in Path('validation/onlyoffice').iterdir() if p.is_file()]; z.close()"
```

插件入口引用官方 CDN 上的 `plugins.js`；本仓库没有复制供应商代码。探针没有发送文稿内容的网络逻辑，下载记录只在本地生成；仍需遵循所用编辑器、运行时和部署环境的条款。独立打开 `index.html` 不会连接编辑器。

打开工作副本，分别选形状、组内对象和图表，点击读取并下载记录。缺少方法、异常返回和读取失败会单独标记；空选区与接口不可用不能混为一谈。探针不读取全部复杂子结构，也不执行写入。某写入方法存在不等于实际修改或往返测试成功。

随后执行与 PowerPoint 相同的修改与保存实验，重点验证组内对象和 SmartArt 是否能通过正式 API 定位，不调用编辑器私有内部变量。插件 API 验证通过，也不能自动证明外部 Automation API 的授权与控制能力相同。

依据：[插件配置](https://api.onlyoffice.com/docs/plugins/configuration)、[调用命令与隔离上下文](https://api.onlyoffice.com/docs/plugins/interacting-with-editors/overview/how-to-call-commands)、[ApiSelection](https://api.onlyoffice.com/docs/office-api/usage-api/presentation-api/ApiSelection)、[ApiChart](https://api.onlyoffice.com/docs/office-api/usage-api/presentation-api/ApiChart)。

## 4. PPTX 保存前后检查

分别在两个编辑器中打开同一个原始副本，先不修改、另存一次，作为各自的基线；再做定向修改，另存新文件。不要让第二个编辑器读取第一个编辑器的输出作为唯一测试输入。

```sh
python validation/pptx_parts.py original.pptx --output .local/validation/original-parts.json
python validation/pptx_parts.py original.pptx powerpoint-noop.pptx --output .local/validation/powerpoint-noop-diff.json
python validation/pptx_parts.py original.pptx onlyoffice-noop.pptx --output .local/validation/onlyoffice-noop-diff.json
```

检查器忽略 ZIP 压缩方式、顺序和时间戳，逐部件比较实际内容。任何 XML 字节变化都会报告，可能只是序列化变化，须人工判断；被删除的图表／嵌入文件需重点检查。嵌入工作簿只计算整体摘要，不解析单元格和公式；关系文件计入比较，但不会解析关系图或判断引用有效性。所有 XML 元素保留计数，额外列出基础形状及组内父级链，不穷举格式对象语义。

返回码 0 表示生成报告成功，不表示保真通过；2 表示检查失败。`inspection_complete=false` 表示存在未完成的 XML 检查。最后必须在 PowerPoint 重开文件，检查外观及内部编辑；图片完全相同也不能替代对象检查。

## 5. Penpot：矢量交互基线

沿用[官方 MCP 接入流程](https://help.penpot.app/mcp/)，在现有账户／部署中连接当前工作文件。先只读确认文件和页面，避免切页后操作了另一份文件。当前仓库没有存放凭据，也没有建立连接。

导入 `three-steps.svg` 后先核查是否仍有独立文字、分组和路径，并记录宿主分配的 ID 映射；不能假定 SVG ID 原样保持。然后执行：

1. 要求列出三个步骤、其标签、两条箭头和右侧干扰对象，先验证理解。
2. 精确选中中间的 `Edit` 文字，仅改为 `Refine`；检查文字仍可编辑，边框和其他对象不变。
3. 在页面坐标中给出粗略区域 `x=70, y=45, width=420, height=560`，请求“三个步骤框等宽，保持中心线”。区域故意包含两条箭头并碰到右侧便签；目标应为三个模块，不能机械修改全部相交对象。
4. 用户手动移动一个模块后继续 AI 修改，再测试 AI 规划期间的手动修改、撤销和导出 SVG 重开。

第 3 步只是“粗区域＋对象结构”的语义实验。若当前 MCP 没有画框交互，明确记录这个 UI 缺口；传坐标不是已经完成意绘的框选产品功能。箭头当前只是路径，没有自动跟随关系；若实验要求移动后跟随，必须明确新增连接语义并验证。

## 结果怎么记录与决定

每项记录宿主版本、输入文件摘要、操作、期望、实际结果及证据位置；状态使用“通过／失败／未运行／接口不可用”。不要只列一张最终截图。独立记录显示、结构保留、手动编辑、API 编辑和撤销五方面。

只有在对象读写、保存重开与手动／AI 接续有实际证据后才确定内核。无操作保存已损坏结构、无法定位关键子对象或不能可靠撤销时，先定位原因，不能靠重画降级满足要求。部分样例通过只支持相应范围的结论。

检查器自测：

```sh
python -m unittest discover -s validation -p "test_*.py" -v
node --check validation/onlyoffice/plugin.js
```

## Unresolved Questions

- 实际目标电脑有哪些已安装／已授权编辑器及版本？当前云端无法代替本机验证。
- 哪份真实 PPT 和 SVG 作为首轮验收输入？使用本地副本，不默认发布用户内容。
- 粗框放在同步预览是否满足体验，还是必须扩展宿主原生画布？
- 宿主实测后，再决定插件／伴随程序／独立网页及生产授权；当前不购买或锁定。
