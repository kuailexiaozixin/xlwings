# Ribbon 构建指南（6.5.6 聚合）

> 本文件聚合 Ribbon 侧全部既有指南（入口规划 / 回调签名速查 / 图标原则），并补充官方 `customaddin.md` Ribbon 节导航。Ribbon 侧编码以本文档为准。

## 一、Ribbon XML 版本规范

- **一律用 2010 版起步**：`customUI14.xml`，命名空间 `xmlns="http://schemas.microsoft.com/office/2009/07/customui"`
- **2006 版（`customUI.xml`）**：仅兼容旧文件保留（import 流程原样保真）；新写内容一律 2010 版
- 构建脚本（build_addin.ps1 / apply_vba_to_workbook.ps1）按此规范注入/回注
- quickstart 初始生成的 `customUI.xml` 为 2006 命名空间，构建时按本规范重写为 2010 版 `customUI14.xml`（quickstart 模板已知问题）

## 二、什么时候真的需要 Ribbon

- 需要**常驻入口**（每次打开 Excel 都可见的按钮）
- 需要**分组/图标/动态状态**（getEnabled/getVisible/getLabel）
- 纯表内按钮/工作表入口足够时**不必用 Ribbon**

### 普通用户入口优先级

1. Ribbon 按钮（全局可见）
2. 工作表内按钮（形状/表单控件，随文件走）
3. 自动事件（Workbook_Open 等，慎用——打扰用户）

## 三、回调签名速查

### 先记住这条原则

- 只有被 Ribbon XML 的 `onAction`/`getLabel`/`getVisible`/`getEnabled`/`getContent`/`onLoad`/`loadImage` 等属性**直接绑定**的过程，才需要写成回调签名
- 普通宏/工具方法/业务子过程**不需要** `IRibbonControl` 参数
- **回调与业务过程必须分离命名**：`onDownload_Click`（回调）→ `DownloadAnnouncements`（业务），不得同名

### 常用控件签名对照

| 控件 | 回调属性 | 签名 |
|------|---------|------|
| `button` | `onAction` | `Sub OnAction(control As IRibbonControl)` |
| `button` | `getLabel`/`getEnabled`/`getVisible`/`getImage`/`getScreentip`/`getSupertip` | `Sub GetXxx(control As IRibbonControl, ByRef xxx)` |
| `checkBox`/`toggleButton` | `onAction` | `Sub OnAction(control As IRibbonControl, pressed As Boolean)` |
| `checkBox`/`toggleButton` | `getPressed` | `Sub GetPressed(control As IRibbonControl, ByRef returnValue)` |
| `comboBox` | `onChange` | `Sub OnChange(control As IRibbonControl, text As String)` |
| `comboBox` | `getText`/`getItemCount`/`getItemLabel` | `ByRef` 返回；`getItemLabel` 含 `index As Integer` |
| `dropDown` | `onAction` | `Sub OnAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)` |
| `gallery` | `onAction` | `Sub OnAction(control As IRibbonControl, selectedId As String, selectedIndex As Integer)` |
| `dynamicMenu` | `getContent` | `Sub GetContent(control As IRibbonControl, ByRef ribbonXml)` |
| `editBox` | `onChange` | `Sub OnChange(control As IRibbonControl, text As String)` |
| `customUI` | `onLoad` | `Sub OnLoad(ribbon As IRibbonUI)` |

**关键规则**：
- 回调属性名 = 过程名前缀（`onAction`→`OnAction`），参数顺序必须完全匹配
- `ByRef` 返回型回调必须在过程内给参数赋值
- `onLoad` 需保存 `ribbon` 对象供 `Invalidate`/`InvalidateControl` 刷新

## 四、Ribbon 图标原则

- 优先内置 `imageMso`（Office 内置图标名，如 `HappyFace`），免资源文件
- 自定义图标：`getImage` 返回 `IPictureDisp`；icon 建议 PNG（16×16 / 32×32）
- 用户提供 logo 时：原始文件放技能根目录 `./assets/`，缩小为 16×16 / 32×32 后放项目 `src/ribbon/`，构建时嵌入
- 图标命名与回调命名分离：`imageMso` 属性用 Office 内置图标名（如 `MacroPlay`），`getImage` 回调返回自定义图标

## 五、构建脚本如何引用

- 从零 xlam：quickstart 工程 → `src/ribbon/customUI14.xml` → `build_addin.ps1` 注入
- 现有文件：`src/ribbon/customUI14.xml` → `apply_vba_to_workbook.ps1` zip 级回注
- 删除 `src/ribbon/` 文件 = 保留基线原 Ribbon
- Ribbon XML 编码必须 UTF-8（防乱码）

**AI 必须理解的注入结果**：构建后 .xlam/.xlsm 包内 `[Content_Types].xml` 含 customUI 部件声明、`_rels/.rels` 恰好 1 条 customUI 关系、`customUI/customUI14.xml` 存在——门禁 A 自动核验这三项，缺失任一项 Ribbon 不显示。

## 六、官方文档导航

- `xlwings/docs/customaddin.md`：xlwings 加载项打包与 Ribbon 集成官方语义
- `../VBA-Docs/Language/`：VBA 语言参考
- Ribbon XML 规范完整参考：微软 customUI 文档（2009/07 命名空间）

## 可配套阅读

- `docs/troubleshooting.md`（customUI.xml 2006/2010 坑）
