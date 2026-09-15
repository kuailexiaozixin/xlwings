# xlwings Office.js 引擎深度技术分析与构建工作流

> 本文件由 `references/13-scenario-d-source-code.md` 的 1.2.5 与 1.8 节独立而来：以开发生命周期工作流重排，深度研读 Office.js 引擎源码（`xlwings/pro/_xlofficejs.py` 160 行 + `xlwings/pro/udfs_officejs.py` 1255 行），并结合官方教程（`xlwings-lite/custom-functions.md`、`custom-scripts.md`）与加载项工程实践编排成可执行工作流。

> **三文档分工**：Office.js 引擎同时服务 xlwings Server 与 xlwings Lite/Pyodide。
>
> - 本文件（16）＝ Office.js 引擎**客户端侧整体实现**（值转换层 + UDF/脚本全链路 + socket.io 会话）＋ 构建 Excel Web 加载项的工作流；
> - `references/14-xlwings-server-guidance.md` ＝ 服务端运行时（xlwings_server 包、FastAPI + Socket.IO、认证部署）；
> - `references/15-xlwings-lite-guidance.md` ＝ Lite 特有适配（Pyodide 浏览器端、JsNull 空值归一、streaming_callback 直推、BookAsync `_lazy` 懒加载）与自托管。
>
> **官方教程的层次**（lite.xlwings.org 文档体系，本文件按阶段语义引用，不重复全文）：
>
> 1. 函数教程 `custom-functions.md`：Basic syntax → 类型提示/装饰器 → Enum → 日期时间 → 数组维度 → 错误单元格 → 动态数组 → Volatile → Object handles → 绘图 → 流式函数 → 调用格地址 → 函数后脚本；
> 2. 脚本教程 `custom-scripts.md`：Basic syntax → 脚本参数与 App Mode 表单渲染 → 配置 → 性能；
> 3. 版本演进 `whatsnew.md`：Office.js 通道的能力时间线。

## 目录

- [工作流总览](#工作流总览)
- [阶段〇：认知与选型](#阶段〇认知与选型)
- [阶段一：工程骨架与本地运行环境](#阶段一工程骨架与本地运行环境)
- [阶段二：manifest 清单与侧载](#阶段二manifest-清单与侧载)
- [阶段三：数据协议基础——值转换层研读](#阶段三数据协议基础值转换层研读)
- [阶段四：自定义函数开发](#阶段四自定义函数开发)
- [阶段五：自定义脚本与元数据生成](#阶段五自定义脚本与元数据生成)
- [阶段六：流式任务与 socket.io 会话](#阶段六流式任务与-socketio-会话)
- [阶段七：测试与验证（双层门禁）](#阶段七测试与验证双层门禁)
- [阶段八：打包与分发](#阶段八打包与分发)
- [阶段九：工程借鉴与能力时间线](#阶段九工程借鉴与能力时间线)
- [附录：源码与教程路径](#附录源码与教程路径)

## 工作流总览

| 阶段 | 目的 | 输出物 | 验收要点 |
|------|------|--------|---------|
| 〇 认知与选型 | 明确引擎定位与运行时选型 | 选型结论 | 引擎与 Server/Lite 关系讲清、运行时匹配需求 |
| 一 工程骨架与本地环境 | 后端 HTTPS + WebView2 + 目录结构 | 可启动的 HTTPS 服务 | 证书受信任、任务窗格页可访问 |
| 二 manifest 清单与侧载 | 把加载项装进 Excel | manifest.xml + 侧载成功 | 清单校验通过、任务窗格可打开 |
| 三 数据协议基础 | 研读值转换层（JS ↔ Python） | 值转换对照笔记 | 日期/错误/NaN 双向转换机制讲清 |
| 四 自定义函数开发 | 用 @xlfunc 写第一个 UDF | 可被公式调用的 UDF | 单元格公式调用成功、类型转换正确 |
| 五 自定义脚本与元数据 | 用 @script 写按钮动作 | 可被按钮触发的脚本 | 脚本参数表单渲染正确 |
| 六 流式任务与会话 | 研读长任务生命周期 | 流式函数可推流 | 断连取消、订阅去重生效 |
| 七 测试与验证 | 双层门禁（引擎 + 加载项） | 测试记录与门禁输出 | 引擎 pytest 全绿、加载项级验证通过 |
| 八 打包与分发 | 产出可分发的加载项 | 可分发包 | 目标机可用（EXE/App Mode/商店） |
| 九 工程借鉴与能力时间线 | 提炼可复用模式 | 借鉴清单 | 对照能力时间线评估兼容面 |

> **执行原则**：开工前先产出 `task_plan.md`（逐条列出：工作步骤、该步任务、手段、产出物、可量化的验收标准）；**逐步推进、逐步验收**——只有完成当前阶段全部目标并留证（产物/日志/门禁输出）才能进入下一阶段，任何一步未达标不得声称完成。

## 阶段〇：认知与选型

- **任务**：明确引擎定位、适用边界，选定运行时（Server 自托管 / Lite 商店 / 纯研读）。
- **要求**：先读三文档分工与 1.1-1.4，再按"选型判定"做决策。
- **输出物**：选型结论（运行时 + 通道 + 许可边界）。
- **验收**：能说清引擎在 xlwings 架构中的位置、与 Server/Lite 的关系，所选运行时与实际需求匹配。

### 〇.1 产品定位：Office.js 通道在 xlwings 架构中的位置

- **Office.js 是什么**：微软 Office 加载项（Office Add-ins）的统一 JavaScript API。xlwings 借它让 **Excel Web 加载项**（任务窗格 + 自定义函数）调用 Python——函数在单元格里以公式形式出现，脚本以按钮触发，任务窗格承载 UI。
- **xlwings 的落地方式**：`pro/_xlofficejs.py` 与 `pro/udfs_officejs.py` 实现 Office.js 引擎。引擎**不是独立服务**，而是"协议转换 + 元数据生成"的客户端侧实现：它把 Python 函数包装成 Office.js 自定义函数/脚本清单与 JS 桥，实际执行位置由宿主决定（Server 的服务器或 Lite 的浏览器）。
- **文件头注释即定位**："This engine is only used in connection with Office.js UDFs, not with runPython."——runPython（VBA 桥接）走 COM 通道，本引擎专为 Web 加载项的自定义函数/脚本而生。
- **许可**：PRO 双许可（PolyForm Noncommercial / xlwings PRO License），商业使用需购买许可——研读与自用学习不受限，对外商业化分发注意授权边界。

### 〇.2 引擎事实总览

- `xlwings/pro/_xlofficejs.py`（160 行）：值转换层——Engine 单例（`name="officejs"`、`type="remote"`）+ 读写两侧的数据清洗/编码。
- `xlwings/pro/udfs_officejs.py`（1255 行）：UDF/脚本全链路——装饰器族、签名解析、类型注入、值转换、调用管线、元数据生成、socket.io 会话管理。
- 引擎类型：`remote`（走 socket.io 推流），与 `excel`（COM）、`calamine`（只读）并列（四引擎见 `references/13-scenario-d-source-code.md` 1.2）。
- 配套测试：`tests/test_custom_functions_officejs.py` / `test_custom_scripts_call.py` / `test_jsnull.py` / `test_streaming_*.py`。

### 〇.3 适用场景与边界

- **适用**：
  - 构建 **Excel Web 加载项**：自定义函数（单元格公式）+ 自定义脚本（按钮动作）+ 任务窗格（可选 UI）；
  - 多人在线协作（Excel on the web）、浏览器端使用——无需本机安装 Python/Office；
  - 跨 Web/Excel 端统一的数据类型语义（DataFrame、日期、错误值、对象句柄）。
- **边界**：
  - 不承载 runPython——VBA 桥接仍走 COM（场景 C）；
  - 引擎本身不产生 UI——任务窗格由 Office.js 前端承载（Server 的 `taskpane.py` 或 Lite 的 App Mode）；
  - 流式函数（asyncgen）**拿不到调用格地址**（`stream` 与 `requiresAddress` 在 Office.js 中互斥），注册期即拒绝——见 4.5 与 5.2。

### 〇.4 与 Server / Lite 的分工

| 维度 | 本文件（16） | 14（Server） | 15（Lite） |
|------|------------|-------------|-----------|
| 角色 | 客户端引擎（协议 + 元数据） | 服务端运行时（FastAPI + Socket.IO） | 浏览器端运行时（Pyodide/WASM） |
| Python 执行位置 | 无（由宿主决定） | 自己的服务器 | 用户浏览器 |
| 通信 | socket.io（Server）/ 直推（Lite） | socket.io | Pyodide 进程内 |
| 特有适配 | — | 认证授权、Redis、部署 | JsNull 归一、streaming_callback、`_lazy` 懒加载 |

- 同一套 `@xlfunc`/`@script` 装饰器与值转换层，被 Server 与 Lite 两端的调用管线共用——这是三份文档共享的"语义内核"。

### 〇.5 运行时选型判定

按需求选运行时，决定后续阶段走哪条工程路径：

| 场景 | 选 Server（14） | 选 Lite（15） | 纯研读 |
|------|----------------|---------------|--------|
| 需要自托管、SSO/RBAC、内网 | ✅ | — | — |
| 免费、免安装、商店分发 | — | ✅ | — |
| 只学引擎机制/借鉴实践 | — | — | ✅ |

- 选型确定后：Server → 14 号阶段一（部署）+ 本文件阶段一至八（客户端侧）；Lite → 15 号（Lite 工作流）+ 本文件阶段三至六（引擎机制）。

## 阶段一：工程骨架与本地运行环境

- **任务**：搭起加载项可运行的本地环境——Python 后端（HTTPS）+ WebView2 就绪 + 目录结构。
- **要求**：加载项 `SourceLocation` 必须是 `https://`（Office 硬性要求）；自签证书（SAN 含 localhost+127.0.0.1）、环回豁免、WebView2 运行时按官方 WebView 平台映射核对。
- **输出物**：`src/` 后端骨架 + `certs/` + 可启动的 HTTPS 服务（访问 `https://127.0.0.1:<port>/taskpane.html` 正常）。
- **验收**：证书装入信任库、环回豁免生效、任务窗格页在浏览器/WebView2 可打开。
- **官方支撑**：官方文档 `develop/addressing-same-origin-policy-limitations.md`（SourceLocation 必须 https 的原因）、`concepts/browsers-used-by-office-web-add-ins.md`（各宿主用哪个 WebView/浏览器版本）、`concepts/privacy-and-security.md`（安全模型）。
- **引擎关联**：引擎本身无 UI、无服务，运行容器由宿主决定——本阶段只是把"容器"准备好，引擎参与从阶段三开始。

## 阶段二：manifest 清单与侧载

- **任务**：编写 `manifest.xml`（把加载项声明给 Excel：Id/版本/权限/宿主/函数与脚本的注册位置/任务窗格 SourceLocation），并侧载进 Excel。
- **要求**：清单元素按微软官方清单参考（docs/manifest，105 篇）核对；侧载走官方机制 `WEF\Developer` 注册 + sideload 容器（Ribbon 命令渲染的必要条件）；开发期用 `127.0.0.1` 而非 `localhost`（环回豁免与证书 SAN 对 IP 更稳）。
- **输出物**：`manifest.xml` + 侧载后 Excel 中出现功能区按钮/任务窗格。
- **验收**：清单静态校验通过（结构/必填项/资源 id 长度 ≤32）+ 真实宿主任务窗格可打开。
- **官方支撑**：官方 `testing/sideload-office-add-ins-for-testing.md`（侧载总述）、`testing/clear-cache.md`（改代码不生效先清缓存）、`testing/uninstall-add-in.md`（卸载）。
- **引擎关联**：xlwings 的函数/脚本清单（`custom_functions_meta` / `custom_scripts_meta`，见阶段五）最终要并入 manifest 的 `<FunctionFile>`/函数注册区——引擎生成的元数据是 manifest 的**内容来源之一**。

## 阶段三：数据协议基础——值转换层研读

> 官方教程对应：`custom-functions.md` 的 "Date and time"（写日期格式）、"Error cells"（错误值）、"Writing NaN values"。教程讲"用法"，本节讲"引擎如何把用法变成 Office.js 协议"。**构建任何函数/脚本前先过此关**——值转换是所有 UDF 的数据底座。

- **任务**：理解读侧（JS 值 → Python）与写侧（Python → JS 值）两条数据流及 data types 协议。
- **要求**：按 3.1 → 3.2 → 3.3 顺序读（读侧 → 写侧 → 协议开关），串起一次调用的完整值旅程。
- **输出物**：值转换对照笔记（Python 值 ↔ Office.js 协议映射表）。
- **验收**：能回答"日期 / 错误值 / NaN / JsNull 在 JS↔Python 之间如何转换、受哪个 runtime 开关控制"。

### 3.1 引擎骨架

- `Engine` 单例（文件末尾 `engine = Engine()`）：只读侧 `clean_value_data` + 写侧 `prepare_xl_data_element` 两个静态方法，`name`/`type` 两个属性——最小引擎契约，被远程调用管线按 `engine_name` 分发。

### 3.2 读侧：clean_value_data（JS 值 → Python）

- 逐元素 `_clean_value_data_element`，对传入的二维数组做规整：
  - **JsNull 哨兵**：Pyodide ≥ 0.28 中 JS `null`（Office.js 空单元格）以 `JsNull` 呈现而非 `None`——`_is_jsnull` 探测并归一为 `empty_as`；空字符串同样归 `empty_as`。**Lite 特有含义见 15 号文档**；
  - **data types 协议的 dict**（`Error`/日期包装）：取 `basicValue`；`Error` 时按 `err_to_str` 决定返回错误字符串或 `None`；
  - **float 走 `number_builder`**：配合调用侧的转换器（如 DataFrame 的数值构建）做精度规整。
- 注释明确指出：datetime_builder 在此不受支持——普通日期格式单元格不被识别（该引擎 UDF-only，所有读值都落在自定义函数参数上）。

### 3.3 写侧：prepare_xl_data_element（Python → JS 值）

按类型逐类编码，返回 Office.js 可直接消费的值：

| Python 值 | 编码结果 |
|-----------|---------|
| `None` | `""`（空字符串） |
| `pd.isna` / `np.isnan`（float） | `#NUM!` Error 类型 |
| `np.number` | `float(x)` |
| `np.datetime64` / `pd.Timestamp` / `NaT`（兜底） / `date` / `datetime` | `datetime_to_formatted_number`（见 3.4） |
| `str` 以 `#` 开头 | `errorstr_to_errortype`（错误字符串 → Error 类型） |
| 其余 | 原样返回 |

### 3.4 data types 协议与 runtime 版本开关

- **日期写侧** `datetime_to_formatted_number`：`runtime >= 1.4` 且 `date_format` 非空时包装为 `{"type":"FormattedNumber","basicValue":serial,"numberFormat":...}`；`runtime < 1.4` 降级裸值。
- **错误写侧** `errorstr_to_errortype`：`runtime >= 1.4` 时映射为 `{"type":"Error","errorType":...}`；错误名映射表：`#DIV/0!`→`Div0`、`#N/A`→`NotAvailable`、`#NAME?`→`Name`、`#NULL!`→`Null`、`#NUM!`→`Num`、`#REF!`→`Ref`、`#VALUE!`→`Value`。
- **设计要点**：值与协议分离——协议包装只依赖 `runtime` 一个开关，业务代码不需要感知客户端能力版本。官方教程 `custom-functions.md` 的 "Error cells" / "Writing NaN values" 即对应此处的 Error 类型与 `#NUM!` 写侧。

## 阶段四：自定义函数开发

> 官方教程对应：`custom-functions.md` 全篇（Basic syntax → 类型提示 → Enum → 日期时间 → 数组维度 → 动态数组 → Volatile → Object handles → 异步）。教程定义"开发者怎么写"，本节揭示"引擎如何解释这些写法"。

- **任务**：用 `@xlfunc` 写第一个自定义函数，跑通"单元格公式 → Python 函数 → 返回值进单元格"闭环。
- **要求**：先过签名约束（4.3），再按 4.1-4.2 用装饰器与注入类型；对照官方教程按能力阶梯（类型提示 → Enum → 数组 → volatile → Object handle）逐个验证。
- **输出物**：可被 Excel 公式调用的 UDF（.py + 注册元数据）。
- **验收**：单元格公式调用成功、类型转换正确（枚举下拉、数组维度、错误单元格、动态数组各验一例）；研读 `tests/test_custom_functions_officejs.py` 看契约的可执行规格。

### 4.1 注入类型机制

- **框架值不进 Excel 签名**：`register_injectable_typehint()` 注册注入类型（如 Server 的 `CurrentUser`）→ `func_sig` 从签名剔除注入参数（可放任意位置，含 keyword-only）；`_unwrap_optional` 处理 Optional 包装。
- **Book 不注册**：`Book`/`BookAsync` 不在此机制内——脚本经 `custom_scripts_call` 单独注入（见 5.1），UDF 的 `Caller` 同理走特殊参数路径。
- **使用**：Server 侧在 `xlwings_server/__init__.py` 注册 `CurrentUser` 后，函数签名写 `current_user: CurrentUser` 即自动注入，Excel 公式里不出现该参数——服务端框架扩展点的参考实现。

### 4.2 装饰器族：xlfunc / xlret / xlarg

- `xlfunc`：核心 UDF 装饰器——
  - 类型提示 → 枚举描述符（`Literal[...]` → Office.js `customEnumId`）或转换选项；
  - `Annotated[X, {"doc","convert",...}]` 抽取选项字典；
  - `ObjectHandle` 出现在 Annotated 中时顶层类型归一为 `object`（走对象句柄缓存而非自身转换器；`-> ObjectHandle` 裸别名同理）；
  - 关键词：`volatile` / `name=`（Excel 名校验：正则 `[^\W\d_][\w.]*`，≤128 字符）/ `namespace` / `help_url` / `required_roles`。
- `xlret` / `xlarg`：返回值与按参数名覆盖转换选项（等价于官方教程的 `@ret(...)` / `@arg(...)` 装饰器，引擎侧统一走 conversion 注册表）。
- **官方对应**：`custom-functions.md` 的 "Enum parameters"（Literal → 下拉枚举）、"Custom function name"（name=）、"Namespace"（namespace 与子命名空间）、"Help URL"（help_url）、"Volatile"（volatile）。

### 4.3 签名约束

- UDF **不支持 keyword 参数**（传参即 `XlwingsError`）；
- **vararg 与 optional 参数不能共存**（`*args` 与默认值同时出现即拒绝）；
- 原因：Office.js 单元格公式的调用模型——参数按位置填充，无关键字概念；引擎在注册期就把非法签名挡下，避免运行期才暴露。

### 4.4 值转换与本地化

- `js_to_none`（JsNull→None）、`to_scalar`（单元素 list/tuple 降标量）；
- `date_format_language_map`：**22 个 locale 的日期格式字符映射**（de `j→y`/`t→d`、fr `a→y`/`j→d`、ru `г→y`/`м→m`/`д→d` 等）——Office.js 交付本地化 shortDatePattern 需转成 Excel 格式字符；
- `convert`（async 写侧）：date_format 三级回退（`@ret` 装饰器 → `XLWINGS_DATE_FORMAT` 环境变量 → Excel cultureInfo）+ locale 归一 + `conversion.async_write(engine_name)`；
- **官方对应**：`custom-functions.md` 的 "Date and time"——教程说明"日期以本地化格式写回需要 date_format"，引擎侧的三级回退就是该承诺的实现。

### 4.5 调用管线 custom_functions_call

调用一次自定义函数的完整步骤：

1. `check_user_roles`（`required_roles` 缺失 → Access Denied 报错并日志）；
2. 版本一致性（client version ≠ `__version__` → 提示重启 Excel / reload 任务窗格）；
3. varargs 展平 → 逐参数 `conversion.read` → 可选参数默认值 → `provide_values_for_special_args` 注入；
4. **三形态执行**：
   - **asyncgen → 流式**：`background_tasks` 字典、task_key 去重复用、socket.io `set-result-{task_key}` 推流或 Lite `streaming_callback` 直推、`on_task_done` 清理竞态；
   - **coroutine**：await 执行；
   - **普通函数**：同步调用；
5. **WithScript 解包**：`CustomFunctionResult(value, script=payload)`——脚本随结果返回（官方对应 "Running a script after a custom function"；流式函数拒绝 WithScript，因脚本与持续推流语义冲突）；
6. **对象句柄代际回收**：仅句柄生产函数（`ret` convert ∈ `object_handles.CONVERTER_KEYS`）+ 带 caller_address 时算 `producer_discriminator`，调用后 `evict_superseded` 清理该代句柄。

## 阶段五：自定义脚本与元数据生成

> 官方教程对应：`custom-scripts.md`（Basic syntax → Script arguments 与 App Mode 表单 → 配置）。脚本是"点击按钮运行的 Sub"，引擎侧是其参数校验与元数据支撑。

- **任务**：用 `@script` 写按钮动作，跑通"任务窗格按钮 → Python 脚本 → 文档变更"闭环。
- **要求**：脚本必须恰一个 book 参数（`Book` 或 `BookAsync`）；对照官方参数类型提示表选表单控件。
- **输出物**：可被按钮触发的脚本 + 函数/脚本清单元数据。
- **验收**：按钮触发成功、脚本参数表单渲染正确（str→文本框、Literal→下拉、date→日期选择器）；`custom_scripts_meta` 与 manifest 的函数注册区一致。

### 5.1 脚本管线 custom_scripts_call

- current_user 前置 → 注入书对象（`_inject_value`：**BookAsync 注解 → `value.impl._lazy = True`** 懒加载，见 remote 后端 `Range.raw_value`；Lite 侧含义见 15 号文档 3.7）→ 按参数类型提示强制转换 → 执行 → 返回 book。
- **book 参数必须恰一个**（`_book_param_hint`：多 Book/BookAsync 注解 → 报错）；`BookAsync` 注解与 `lazy=False` 冲突 → 报错（注解优先，仅与显式 False 矛盾才拒绝）。
- `script` 装饰器关键词：`include` / `exclude` / `button` / `show_taskpane`；`lazy=` 已弃用 → **`book: xw.BookAsync` 注解**等价（内部统一发 `"lazy"` wire 键）。
- **官方对应**：`custom-scripts.md` 的 "Script arguments" 类型提示表（`str`→文本框、`int`/`float`→数字框、`bool`→复选框、`Literal`→下拉、`datetime.date`→日期选择器）——引擎的 `_coerce_script_arg` 逐类型强制即该表的实现。

### 5.2 元数据生成与注册期校验

- `custom_functions_code`：从 `custom_functions_code.js` 模板注入版本/路径，为每个函数生成 JS 包装 + `CustomFunctions.associate`——**这就是"Python 函数出现在 Excel 公式里"的桥**。
- `custom_functions_meta`：清单 JSON——`stream` / `requiresAddress` / `volatile` 选项、matrix 维度、枚举注册、**重复 Excel 名检测**。
- **流式函数注册期拒绝**：asyncgen 带 `Caller` 注解 → 注册即报错——Office.js 中 `stream` 与 `requiresAddress` 互斥，流式函数永远拿不到调用格地址，注册期拒绝而非运行期注入 None。
- `custom_scripts_meta`：脚本清单 JSON（与函数清单并列，供任务窗格/App Mode 渲染按钮与表单）。

## 阶段六：流式任务与 socket.io 会话

> 官方教程对应：`custom-functions.md` 的 "Streaming functions"（asyncgen 持续 yield）。教程给三个例子（时钟/随机数/BTC 价格），本节讲引擎如何管理这些长任务。

- **任务**：研读流式函数生命周期，理解长任务在 Server（socket.io）与 Lite（直推）两端的差异。
- **要求**：先读官方三个流式示例（用法），再读引擎双字典管理（机制）。
- **输出物**：一个可运行的流式函数（asyncgen + 订阅）。
- **验收**：多单元格订阅同一任务不重复起任务、断连后任务取消、重连去重复用。

- `sio_connect`（认证 token）/ `sio_disconnect`（订阅计数归零取消任务）/ `sio_custom_function_call`（sid 计数）/ `sio_cancel_task`。
- **双字典管理**：`task_key_to_sid_counts`（一个任务被多个单元格订阅时的计数）+ `task_key_to_task`（任务句柄）——断开时按计数归零取消，重连时任务去重复用，不重复起任务。
- **对构建应用系统的借鉴**：WebSocket/socket.io 服务的可靠模式——订阅计数、任务去重、断连取消三件套是长连接推送的通用骨架。

## 阶段七：测试与验证（双层门禁）

> 质量心智图：**Check 管规约、Test 管行为、Debug 管失败归因**——Check/Test 失败后进入 Debug 闭环（复现→定位→修复→回归→沉淀），每轮改动后跑门禁，全绿才进下一阶段。

- **任务**：引擎级 + 加载项级两层验证。
- **要求**：
  - **引擎级**（改动 `udfs_officejs.py` / `_xlofficejs.py` 后必跑）——pytest 直接测引擎契约；
  - **加载项级**（引擎装配进加载项后必跑）——清单静态校验、HTTP 冒烟、真实宿主侧载断言。
- **输出物**：测试记录 + 门禁输出（pytest 结果 / 清单校验输出 / 冒烟结果）。
- **验收**：引擎 pytest 全绿；清单校验通过、HTTP 冒烟双 200、真实宿主任务窗格可开；任何 failed 进入 Debug 闭环，不得跳过。

- **引擎级测试**（`xlwings-0.37.3/tests/`）：
  - `test_custom_functions_officejs.py`：UDF 全链路（装饰器 → 调用管线 → 返回值编码）；
  - `test_custom_scripts_call.py`：脚本管线（book 注入、参数强制、返回 book）；
  - `test_jsnull.py`：JsNull 哨兵归一（Pyodide ≥ 0.28 兼容分支）；
  - `test_streaming_*.py`：流式函数（asyncgen 推流、task_key 去重、断连取消）。
  - 研读测试 = 看引擎契约的"可执行规格"。
- **加载项级验证**（引擎装配进加载项后执行）：
  - **清单静态校验**：manifest 结构/必填项/资源 id 长度（≤32）——改清单后立即跑；
  - **HTTP 冒烟**：起 HTTPS、`/api/health` 与 `/taskpane.html` 双 200、页面含 office.js 引用；
  - **宿主断言**：真实 Excel 侧载后任务窗格 WebView2 加载、Ribbon 命令渲染（UIA 枚举 + 截图留证）；
  - **调试第一手证据**：Office 运行时日志（注册表启用 RuntimeLogging）——"命令不渲染"时直接看宿主拒绝/解析 manifest 的具体原因，比截图更快更可靠。

## 阶段八：打包与分发

- **任务**：把后端 + 证书逻辑 + manifest + 前端打成分发包，用户端零配置可用。
- **要求**：打包前门禁全绿；按目标形态选分发路径（EXE 自包含 / App Mode 工作簿 / 商店包）。
- **输出物**：可分发包。
- **验收**：目标机双击即用（EXE 自含 HTTPS+证书+清单注册）；App Mode 表单可渲染；商店包通过官方校验。
- **官方支撑**：官方 `publish/publish.md`（发布总览）、`publish/publish-add-in-vs-code.md`（VS Code 发布到 Azure，对照）、`publish/deploy-office-add-in-sso-to-azure.md`（SSO 部署）、`publish/maintain-breaking-changes.md`（升级兼容）。
- **引擎关联**：
  - **App Mode**：`custom_scripts_meta` 的表单定义即 App Mode 渲染数据源（见 5.2）；
  - **EXE 分发**：xlwings Server 侧可自托管（14 号阶段七），Lite 侧走商店或自托管（15 号阶段六）——本阶段面向"用引擎构建自有加载项"的分发，三种路径按选型取一。

## 阶段九：工程借鉴与能力时间线

- **借鉴清单**：
  - 跨 Web/Excel 端 UDF 的**数据类型协议**（FormattedNumber/Error 包装）与 JsNull 归一——多端统一值语义的模板；
  - **locale 日期格式归一**（`date_format_language_map`）——处理多语言客户端格式差异的实例；
  - 流式函数的**订阅计数 + 任务去重 + 断连取消**生命周期管理（sid/task_key 双字典）；
  - 框架注入类型（CurrentUser/Book）的**签名隐藏与位置无关注入**——服务端框架扩展点的设计；
  - **注册期校验优先**（签名约束、重复名、stream+requiresAddress 互斥都在元数据生成期拒绝）——把错误挡在加载前，而非运行中。
- **加载项工程化借鉴**（引擎知识落地为可运行加载项）：
  - **manifest 清单与侧载**——引擎跑在 manifest.xml 定义的任务窗格容器里；清单元素、WEF 注册、sideload 容器的权威结构与做法以微软官方清单参考（docs/manifest，105 篇）为准（见阶段二）；
  - **本地 HTTPS 与 WebView2**——加载项本地开发前提（自签证书、环回豁免、WebView2 运行时），细节见阶段一；
  - **加载项级验证**——清单静态校验、HTTP 冒烟、宿主侧载断言与阶段七的引擎级 pytest 互补，构成"引擎内 + 容器外"双层验证；
  - **要求集兼容评估**——Office.js requirement-sets 矩阵（92 篇）是版本兼容评估的权威依据，与阶段九能力时间线对照使用。
- **能力时间线**（`docs/whatsnew.md`，演进即"哪些能力是哪一版才有"）：
  - Office.js 支持随 PRO 引入，逐步覆盖：值转换 → 自定义函数 → 流式函数 → 对象句柄 → 脚本（WithScript）→ 动态数组/volatile/枚举 → App Mode 表单；
  - 版本差异直接影响引擎分支：`runtime < 1.4` 的裸值降级、Pyodide ≥ 0.28 的 JsNull 哨兵——升级 Excel/Office.js runtime 前先对照此时间线评估兼容面。

## 附录：源码与教程路径

| 资源 | 路径 |
|------|------|
| 值转换层源码 | `xlwings-0.37.3/xlwings/pro/_xlofficejs.py` |
| UDF/脚本全链路源码 | `xlwings-0.37.3/xlwings/pro/udfs_officejs.py` |
| 函数教程 | `xlwings-lite/custom-functions.md`（lite.xlwings.org） |
| 脚本教程 | `xlwings-lite/custom-scripts.md` |
| 版本演进 | `xlwings-0.37.3/docs/whatsnew.md` |
| 配套测试 | `xlwings-0.37.3/tests/test_custom_functions_officejs.py` 等（见阶段七） |
| 官方清单参考 | 微软 office-js-docs-reference 的 docs/manifest（105 篇）与 requirement-sets（92 篇） |
| 官方测试/发布文档 | Microsoft Learn `testing/` 与 `publish/` 目录（见阶段七/八官方支撑） |
| 服务端侧 | `references/14-xlwings-server-guidance.md`（Server 运行时） |
| Lite 适配 | `references/15-xlwings-lite-guidance.md`（浏览器端分支） |
