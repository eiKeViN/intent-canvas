# 意绘：编辑能力复用与产品形态比较

调研日期：2026-10-19。状态：公开文档核查与方案建议，未运行兼容性测试，尚未决定架构、供应商或许可证。

后续进展：用户已接受优先复用成熟编辑器的方向与推荐验证顺序；可执行材料和实际检查状态见 [首轮验证](validation/README.md)。本文保留调研时的依据和待验证边界，不把接受方向视为通过兼容性测试。

## 结论与前提

建议优先验证“成熟编辑器中的 AI 编辑层”。PPT 重点比较 PowerPoint 原生扩展和 ONLYOFFICE；矢量图重点比较 Penpot 与 Inkscape。独立 Web App 保留为候选，是否需要自建画布应由接口限制和交互验证决定。

这不是降低已确认需求：生成初稿与打开已有 PPT 仍是两个正式入口；格式中已有的子对象仍需能够深入编辑。只保留外观、把图表变成图片或普通路径，都不算完成。没有候选在本次调研中被证明覆盖全部对象和属性。

对每个候选应分开判断四件事：编辑器能否手动编辑；公开 API 能否读写同一对象；扩展 UI 能否承载我们的指向交互；保存后原有结构是否仍可编辑。前一项成立，不代表后三项成立。

## 主要候选

| 候选与意绘形态 | 文档可确认的基础 | 关键缺口与代价 | 当前判断 |
| --- | --- | --- | --- |
| PowerPoint 桌面原生扩展／本地伴随程序 | 原生对象模型可访问分组子对象、图表系列；有选择变化事件和撤销入口 | COM／VSTO 路线依赖 Windows 和已安装的 PowerPoint；画布覆盖层、完整属性覆盖与事务行为待测 | 若目标环境是 Windows，最值得先验证既有 PPT 的深入编辑 |
| PowerPoint Office.js 网页插件 | 官方任务窗格与网页技术；可访问形状、文字、表格、分组等，跨平台能力按宿主与 requirement set 检查 | 本次查阅的 Shape API 未见与原生 Chart／Series 对等的入口；不能把原生接口能力直接记到 Office.js 名下 | 交互入口有吸引力，暂不能作为全部 PPT 深度编辑的唯一接口 |
| ONLYOFFICE Docs 嵌入式 Web App／编辑器插件 | 成套演示文稿编辑器；有选择、图表系列、坐标轴等 API；插件和外部 Automation API 两种接入 | PPTX 往返保真、嵌套组与 SmartArt 内部编辑待测；商业授权和外部连接器费用需核清 | 独立网页产品的首要候选 |
| LibreOffice 本地扩展／Collabora Online 网页集成 | 成熟办公编辑器；LibreOffice 提供 UNO，Collabora 提供网页嵌入及消息接口 | 格式转换保真待测；网页消息接口不能直接等同于完整 UNO 对象访问；部署与适配工作需评估 | 开源与自主部署候选，暂不排在首轮前两条 PPT 路线之前 |
| Penpot 插件／基于 Penpot 的工作区 | 矢量对象、选择、历史及插件 API；官方 MCP 已能读写设计结构 | 不是已经验证的 PPT 原生编辑内核；任意 SVG 的语义与效果往返也未验证 | 最适合快速验证矢量场景的 AI 协作交互 |
| Inkscape 扩展／本地伴随程序 | 以 SVG 为文档基础；扩展能处理文档和选择 | 普通效果扩展不等于持续在线的编辑会话；状态同步、撤销与自定义 UI 待测 | 更重视 SVG 文件结构时值得优先验证 |
| Figma 插件 | 官方插件 API 可读写图层、层级、文字、向量等 | 运行于宿主内；经典插件有运行生命周期限制；不是取得可独立部署的 Figma 内核 | 用户已在 Figma 工作时的候选；本轮不把它当 PPT 保真方案 |

以上排序是针对意绘需求的判断，不是通用性能排名。各候选的公开依据见下文。

## PPT：原生宿主与独立网页的取舍

### PowerPoint 原生路线

微软原生对象模型的 GroupItems 可以访问组内形状；Series.Values 可以读写系列值。WindowSelectionChange 可接收选择变化，StartNewUndoEntry 提供建立撤销条目的接口。它们为精确选择、深入修改和手动／AI 接续提供了基础，但没有证明任意修改都能组成可靠的原子事务。[M1–M4]

Office.js 是另一套接口。其形状 API 已覆盖分组、位置和文本等，但本次查阅没有发现与原生图表系列操作等价的公开入口。需要按具体宿主、版本及 requirement set 建立能力表，不能用“PowerPoint 支持”笼统代替插件支持。[M5–M6]

建议原型让用户继续在 PowerPoint 中手动编辑，意绘提供任务窗格或伴随窗口。精确指向使用宿主选区；粗略框选可先在同步的幻灯片预览上完成，再映射到原生对象。直接覆盖原生画布是另一项待验证能力，不能把任务窗格误当作画布扩展权限。

复用的是用户环境中的 PowerPoint。官方扩展机制不代表获得把 PowerPoint 内核嵌入自有网页、再分发或作为普通多租户服务的授权。微软也明确提示无人值守 Office 自动化的支持与稳定性限制，不能把本地可用直接外推为云端可用。[M5, M7]

### ONLYOFFICE 路线

公开 API 中，ApiSelection 可取得选中形状或文字范围，ApiChart 可访问系列并设置数据及轴等属性。编辑器内部插件通过公开 API 操作；从意绘自有网页外部控制编辑器，则另有 Automation API 连接器。[O1–O3]

仍有具体疑点：本次 ApiGroup 文档列出了 Ungroup、文字遍历及 ToJSON，未见通用的直接子形状遍历入口；ApiSmartArt 页面列出的主要是整体绘图操作，尚不足以确认内部节点编辑。不能用“先拆组再重组”默认绕过，因为对象身份、关系和历史可能改变。这些应成为验证题，不宜提前写成不可能或已经解决。[O4–O5]

授权存在正式路径，但版本与接入方式必须对应：
- 社区版适用 AGPL 及相关条件；不能把“开源”理解为无条件嵌入闭源产品。
- Developer 商业授权面向把编辑器集成进自有产品；开发／评估许可与面向终端用户的生产许可不同。
- Automation API 页面称它是 Developer 专属且需额外费用；开发者产品页又称包含于 Developer 许可。两份官方页面表述不一致，选型前应取得适用版本和合同范围的明确说明，不能预设免费或固定费用。[O3, O6–O8]

## 矢量图：可以先复用现成 AI 通道

Penpot 的官方 MCP 已通过插件连接当前设计文件，让代理读取和修改图层、组件、样式等；既有远程也有本地接入方式。这意味着验证“AI 读懂对象并回写”未必需要先自建整套编辑器。意绘可以先补充粗略区域指向、候选高亮、修改反馈等交互。[V1]

Penpot 插件 API 另有 selection、history 和对象操作；源代码采用 MPL-2.0。自托管或修改应检查对应代码及分发义务，托管服务条款与代码许可也应分开看。它的 SVG 导入导出仍须测试，不能由“使用矢量对象”推导出任意 SVG 无损。[V2–V3]

Inkscape 的 EffectExtension 处理 SVG 文档与选区，适合直接围绕 SVG 工作。官方将 Inkscape 以 GPL 发布；复用、修改并分发编辑器与单独开发扩展的具体义务需按方案核查。效果扩展只是接入起点，不能直接承诺实时双向同步。[V4–V5]

Figma 的官方插件 API 也能深入读写图层。经典插件由用户启动，文档明确限制后台运行和同时运行多个插件；这会影响常驻助手设计。选择 Figma 时，意绘更自然地成为插件，不能由插件 API 推导出独立嵌入编辑器的权利。[V6]

## 看似相近、但不能独自解决问题的组件

| 类别 | 可复用价值 | 为什么不足以单独满足需求 |
| --- | --- | --- |
| Polotno 商业编辑器 SDK | 可嵌入画布、文档状态与多格式输出，包括 PPTX | 官方明确目前没有第一方 PPTX 导入；不能承担已有 PPT 深入编辑闭环。授权按产品集成范围另行确认。[B1–B2] |
| Aspose.Slides 商业文档 SDK | 可操作演示文稿和图表工作簿，可作为文件处理后端 | 不是现成的交互编辑器，画布、选择与历史仍要解决；内部使用、对外部署及 API／SDK 再分发对应的授权不同。[B3–B4] |
| Open XML／直接 XML 编辑 | 保留格式原始信息、定向修改文件的候选基础 | 不自带成熟画布；关系、排版、渲染、保存及宿主内状态同步仍需处理。[M7] |
| 当前会话可调用的演示文稿工具 | 可辅助制作样例、验证交互想法 | 本次未核实有可供意绘分发或嵌入的公开 SDK 授权，不能直接列作产品依赖 |

LibreOffice／Collabora 不是上表所说的“只有后端”：它们有完整编辑界面。其难点是具体接入层与 PPT 兼容性。LibreOffice 的 MPL 等代码许可、Collabora 的版本与商业服务范围应分别核查；CODE 官方定位为开发／测试用途，不应直接视作生产部署承诺。[L1–L4]

## 对意绘设计的建议

建议把意绘定义为“理解视觉修改意图、定位对象，并在成熟编辑环境中执行和反馈的协作层”。最终可以表现为插件、本地伴随程序或嵌入式 Web App；不必先为品牌统一而重写两个编辑内核。

这里最值得调整的是中间表示的职责：给 AI 的简化对象描述，不必成为文件的权威格式。宿主保持原生文档状态，意绘按任务读取画面、选区、对象 ID、层级及所需属性；修改再通过对应接口回写。图表数据仍作为图表数据，SVG 路径仍作为路径。HTML 可用于界面，SVG 可用于预览，都不必承担所有 PPT 语义。

建议的验证顺序：
1. 在可用的目标设备上比较 PowerPoint 原生扩展与 ONLYOFFICE 的对象闭环，再决定 PPT 的落脚点。若必须浏览器跨平台运行，优先验证 ONLYOFFICE；若 Windows 原生 PPT 保真优先，先验证 PowerPoint。
2. 用 Penpot 现有 MCP 做矢量交互基线；若 SVG 原始结构是主要工作对象，再比较 Inkscape。
3. 比较现成宿主加 AI 通道已经能做到什么，以及意绘新增交互实际节省多少操作。若现成组合已满足需求，产品可能只需一个小插件或工作流包。
4. 在一个宿主上证明价值后，再决定是否抽出共享适配层；目前不建设多宿主框架。

## 最小验证与淘汰标准

这是一组待执行实验，不是已经通过的测试。至少使用一份在候选编辑器之外制作的 PPTX，以及包含嵌套组、文字和路径的 SVG，避免只验证自己生成的文件。

| 实验 | 需要证明什么 |
| --- | --- |
| 打开后不修改，直接保存并重新打开 | 外观、对象类型、未触碰关系与复杂结构是否保持；不要求 ZIP 字节完全相同 |
| 组内只改一段文字 | 能定位到原子对象；分组、身份、其他文字不被重建或丢失 |
| 修改图表某系列的数据及轴标题 | 数据语义、嵌入工作簿／公式关系及后续原生编辑保持，不只是更新缓存或画面 |
| 粗框三步骤并请求等宽 | 正确排除跨区箭头、背景与邻近干扰对象；预览坐标和实际对象对应 |
| 精确选取文字、手改后继续 AI 修改 | 宿主与 AI 使用同一最新状态；选区、对象身份和版本可追踪 |
| AI 规划期间手动修改，再撤销 AI 批量操作 | 不覆盖新修改；撤销边界清楚且不损失更早操作 |
| SmartArt、公式、母版、主题、连接符及复杂 SVG 特性 | 分别记录显示、保留、手动编辑、API 编辑的覆盖情况与差距 |

无法读写关键子对象、无操作保存就破坏结构、无法可靠同步手动修改，都是不能直接进入完整产品实现的信号。部分样例通过只能支持相应范围的结论，不代表全格式兼容；未覆盖能力继续记作需求缺口。

## 官方资料

以下为本次查阅依据。API 与合同会变化，正式选型需固定版本重新核验。

- M1：[PowerPoint GroupItems](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.groupitems)
- M2：[PowerPoint Series.Values](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.series.values)
- M3：[WindowSelectionChange](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.windowselectionchange)
- M4：[StartNewUndoEntry](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.startnewundoentry)
- M5：[Office Add-ins 平台与 COM／VSTO 区别](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins)
- M6：[Office.js PowerPoint Shape，含 preview](https://learn.microsoft.com/en-us/javascript/api/powerpoint/powerpoint.shape?view=powerpoint-js-preview)
- M7：[微软：无人值守 Office 自动化及替代方案](https://learn.microsoft.com/en-us/office/client-developer/integration/considerations-unattended-automation-office-microsoft-365-for-unattended-rpa)
- O1：[ApiSelection](https://api.onlyoffice.com/docs/office-api/usage-api/presentation-api/ApiSelection)；[ApiChart](https://api.onlyoffice.com/docs/office-api/usage-api/presentation-api/ApiChart)
- O2：[演示文稿插件 API](https://api.onlyoffice.com/docs/plugins/interacting-with-editors/presentation-api)
- O3：[Automation API](https://api.onlyoffice.com/docs/docs-api/usage-api/automation-api)
- O4：[ApiGroup](https://api.onlyoffice.com/docs/office-api/usage-api/presentation-api/ApiGroup)
- O5：[ApiSmartArt](https://api.onlyoffice.com/docs/office-api/usage-api/presentation-api/ApiSmartArt)
- O6：[ONLYOFFICE 许可 FAQ](https://www.onlyoffice.com/license-faq)
- O7：[Developer Edition 许可 FAQ](https://helpcenter.onlyoffice.com/docs/faq/developer.aspx)
- O8：[ONLYOFFICE 面向开发者的产品页](https://www.onlyoffice.com/for-developers)
- V1：[Penpot 官方 MCP](https://help.penpot.app/mcp/)
- V2：[Penpot Plugin API](https://doc.plugins.penpot.app/interfaces/Penpot)
- V3：[Penpot LICENSE](https://github.com/penpot/penpot/blob/develop/LICENSE)
- V4：[Inkscape 扩展 API](https://inkscape.gitlab.io/extensions/documentation/source/inkex.extensions.html)
- V5：[Inkscape 许可证](https://inkscape.org/en/about/license/)（本次正文抓取受限，GPL 信息依据官方搜索摘要）
- V6：[Figma Plugin API 与运行限制](https://developers.figma.com/docs/plugins/)
- L1：[LibreOffice 许可证](https://www.libreoffice.org/licenses/)
- L2：[LibreOffice UNO 绘图页接口](https://api.libreoffice.org/docs/idl/ref/interfacecom_1_1sun_1_1star_1_1drawing_1_1XDrawPagesSupplier.html)
- L3：[Collabora SDK 手册](https://sdk.collaboraonline.com/CO-SDK-manual.pdf)（本次完整抓取受限；仅确认官方搜索结果中的消息接口，不据此主张对象覆盖完整）
- L4：[Collabora CODE](https://www.collaboraonline.com/code/)；[许可说明](https://www.collaboraonline.com/terms/collabora-online-mplv2/)（官方搜索摘要，完整页面抓取受限）
- B1：[Polotno 导入导出范围](https://polotno.com/docs/import-and-export)
- B2：[Polotno 商业授权](https://polotno.com/sdk/pricing)
- B3：[Aspose.Slides 图表工作簿](https://docs.aspose.com/slides/net/chart-workbook/)
- B4：[Aspose 商业许可类型](https://purchase.aspose.com/policies/license-types/)

## Unresolved Questions

- 首轮实际目标设备和 PowerPoint 版本是什么？跨平台与统一画布是否必须从第一版满足？
- 粗略框选放在同步预览中能否满足体验，还是必须直接发生在宿主原生画布上？
- 深入编辑如何分配给 AI 指令与宿主的手动界面？关键操作是否均有稳定、正式的 API？
- 接受哪种部署与授权模式：用户自备宿主、履行开源许可义务，或购买商业嵌入授权？
- 哪些真实 PPT／SVG 构成首轮验收集，如何持续扩大到格式定义的子对象和属性？
