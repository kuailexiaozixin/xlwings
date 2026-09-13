# 项目设计指南（6.5.3 深度支撑）

> 本文件承接 SKILL.md 6.5.3.2（设计六件套 + 模块化原则 + 验收标准，正文为准），聚焦**列 Schema 设计决策法**——设计阶段③数据设计的核心方法（列定义决策 + 动态列 + 与 ListObject/表格的映射 + 三方案对照）。

## 什么时候优先阅读本文件

- 项目设计阶段的数据设计环节：定义表格/数据区列定义（列名/类型/来源/格式）
- 任何"输出列较多 / 含日期数字格式敏感列 / 用 ListObject 输出"的场景
- 列 schema 完整性自检

## 常见列类型与格式决策

| 列类型 | 写入策略 | 格式决策 |
|--------|---------|---------|
| 文本列（含前导零：股票代码/身份证/编号） | **先设 `NumberFormat = "@"` 再写值** | 文本格式；前导零归一化（如 `Format(code, "000000")`） |
| 日期列 | 写入日期值后补 `NumberFormat`（如 `yyyy-mm-dd`） | 显示格式与存储值分离 |
| 时间列 | 同上，`hh:mm:ss` | 注意 Excel 时间序列值语义 |
| 日期时间列 | 写入后补 `yyyy-mm-dd hh:mm:ss` | 毫秒/时区在 Python 侧先归一 |
| 数值列 | 直接写数值，`Value2` 读取 | 长数字（15 位以上）先设文本再写，防科学计数 |
| 公式列 | 写公式字符串，配合 `.Formula`/`.Formula2` | `Formula` 与 `Formula2` 差异见 `../VBA-Docs/excel/Concepts/Cells-and-Ranges/range-formula-vs-formula2.md` |
| 布尔/标记列 | 写 `True/False` 或 `是/否` 常量 | 与业务侧约定一致 |
| 混合展示列 | 生成展示字符串（如 `代码+名称`） | 明确"展示列"定位，不与数据列混用 |

## 二维数组读取后的 Schema 检查

- `Value2` 读回后重新确认格式语义（文本列可能被 Excel 转数值、日期列可能是序列值）
- 空值处理：明确空单元格 → `None`/空串的约定
- 前导零丢失：读回文本列必须检查（Excel 可能已转数值）

## 推荐的列级处理策略

- **纯数据列**：整块二维数组写入，速度快
- **文本列**：先设整列 `NumberFormat="@"` 再写
- **日期/时间列**：写值后补列格式；读回用 `Value2` 再转
- **公式列**：`ListObject` 中穿插公式列时，**禁止用纯数据数组整块覆盖**（会抹掉公式）——分开写数据区与公式区

## 适合和 ListObject 一起使用的列设计

- 结构化明细表优先 `ListObject`（自动扩展、样式、汇总行）
- ListObject 扩容/缩容前检查周边内容（防覆盖相邻数据）
- 列设计固定后 ListObject 列数 = Schema 列数

## 可配套阅读

- `../VBA-Docs/api/Excel.ListObject*.md`（ListObject 成员官方语义）
- `../VBA-Docs/excel/Concepts/Cells-and-Ranges/range-formula-vs-formula2.md`（Formula/Formula2 差异）
