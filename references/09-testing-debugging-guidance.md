# 测试、调试与检验指南（6.5.9 单元测试 + 6.5.12 集成验证 聚合）

> 本文件聚合测试调试侧全部既有指南（COM 冒烟 / 真机 XLSTART 验证铁律）+ TDD 通用理念 + 门禁体系速查 + 官方调试文档导航。

## 一、TDD 理念（三阶段门禁）

- **Red-Green-Refactor**：新业务函数先写失败测试 → 实现 → 重构；禁止先实现后补测
- **Prove-It**：修 bug 先写复现测试，再修复
- **持续验证**：每次改代码后 `py_compile` + 模块导入 + pytest 全绿
- **测试分层**：单元（业务逻辑）→ 集成（跨模块/转换器）→ 端到端（真实运行）；UDF 用"单元格公式断言 + 纯 Python"双路径

## 二、COM 冒烟测试（原 06）

### 总体原则
冒烟 = "能不能跑起来"，不是业务正确性验证；业务正确性靠单元/集成测试。

### 建议测试流程

- **xlsm**：打开 → 触发入口 → 校验输出区域 → 退出
- **xlam**：新实例（避免污染用户 Excel）→ 确认 AddIns 加载 → 触发回调 → 校验输出 → 退出
- 推荐最小测试数据：1 行正常 + 1 行边界（空值/超长）+ 1 行错误（异常码）

### 测试时要检查什么
入口可见、回调无报错、输出区写入正确、状态位反馈、异常路径不崩溃

### UserForm 场景的测试分层（默认方案）

1. **第一层：业务方法测试**——窗体逻辑解耦为可单测的业务方法
2. **第二层：窗体打开冒烟**——Show 不报错
3. **第三层：UI 自动化（可选增强）**——默认不做（脆弱）；COM UI 自动化仅在有明确收益时用

### 关键风险：模态对话框
MsgBox 会阻塞自动化——测试时用参数化/回调替代弹窗，或确保弹窗可自动关闭（见 `../VBA-Docs/excel/Concepts/Controls-DialogBoxes-Forms/automatically-dismiss-a-message-box.md`）

### UDF 冒烟三步
1. `import_udfs` 注册 2. 单元格写公式 3. 重算断言结果；禁忌：不验证 UDF 依赖服务（xw.serve）就交付

**白标 xlam UDF 测试前提（实测高频坑）**：当开发机同时安装了 xlwings 标准加载项（`xlwings.xlam`，通常在 XLSTART）和白标 xlam 时，两者的 `xlwings` VBA 模块同名冲突，白标 xlam 的 UDF 会返回 `#NAME?`。测试前须在 COM 自动化中禁用标准加载项：
```python
for a in app.api.AddIns:
    if 'xlwings.xlam' in a.Name and 'myproject' not in a.Name:
        a.Installed = False
```
且白标 xlam 须以 `app.books.open()` 方式保持打开（仅 `AddIns.Add + Installed=True` 时 UDF 可能不初始化）。

### UDF 回归测试骨架（pytest，两条路径互补）

`conftest.py`（xw_app fixture：启动 Excel → 禁用标准加载项 → 打开白标 xlam → 清理）：

```python
import pytest
import xlwings as xw
import os

XLAM_PATH = os.path.join(os.path.dirname(__file__), '..', 'myproject.xlam')

@pytest.fixture(scope='module')
def xw_app():
    app = xw.App(visible=False)
    # 禁用标准 xlwings.xlam（与白标 xlam 冲突）
    for a in app.api.AddIns:
        if 'xlwings.xlam' in a.Name and 'myproject' not in a.Name:
            a.Installed = False
    # 白标 xlam 须保持打开，UDF 才可用
    wb_xlam = app.books.open(XLAM_PATH)
    yield app
    wb_xlam.close()
    app.quit()
```

测试用例：

```python
# 路径 A：单元格公式断言（真实用户行为，验转换器）
def test_udf_via_cell(xw_app):
    wb = xw_app.books.add()
    wb.sheets[0].range('A1').value = 5
    wb.sheets[0].range('B1').formula = '=double(A1)'
    xw_app.api.Calculate()  # 注意：wb 无 calculate()，须用 app.api.Calculate()
    assert wb.sheets[0].range('B1').value == 10
    wb.close()

# 路径 B：纯 Python 单测（不依赖 Excel，验业务逻辑）
from myproj import double
def test_double_logic():
    assert double(5) == 10
    assert double(0) == 0
    assert double(-3) == -6
```

### Ribbon XML 编码验证（防乱码）
构建后检查包内 `customUI/customUI14.xml` 首字节不是 `EF BB BF`（UTF-8 无 BOM）；中文按钮名在 Excel 中显示正常（编码规则遵循代码化运行铁律）。

### 推荐测试方式
- 单元/集成：pytest（Python 侧）+ 纯 VBA 业务方法单测（窗体解耦后）
- 冒烟：COM 自动化（PowerShell + Excel COM，见上文销毁规范）
- 端到端：门禁 J（子进程/端口/API）+ 门禁 L（真实 Excel 启动）
- 不推荐：UI 自动化（脆弱，仅 UserForm 第三层可选增强）

### 交付物清单
产物（xlsm/xlam/ZIP）+ 配套测试文件 + 运行证据（输出/截图/日志）

### 示例文件交付逻辑
- **xlsm 示例**：随产物一起交付，用户打开即可看到输入区+按钮+输出区
- **xlam 示例**：单独提供一个"示例工作簿"（含输入区与调用按钮），因为 xlam 本身无工作表——用户在示例工作簿里触发加载项功能
- **生成示例工作簿**：用 COM 自动化新建簿 → 写输入区/按钮/说明 → 保存为 .xlsx（不嵌宏，按钮通过加载项回调触发）

### 测试数据管理

- **数据与代码分离**：测试数据放 `tests/fixtures/`（JSON/CSV/Excel），测试代码放 `tests/test_*.py`，禁止在测试代码中硬编码大量业务数据。
- **准备/清理用 fixture**：pytest 的 `setup`/`teardown` 或 `@pytest.fixture` 管理测试数据生命周期，测试结束后必须清理（删除临时文件、关闭工作簿、恢复环境）。
- **不得污染交付示例**：测试用的工作簿/数据与交付示例文件严格分离，测试不得修改 `dist/` 或示例目录下的文件。
- **外部 API 用 mock**：调用外部 API（如公告查询、网络请求）的测试必须用 `unittest.mock` 或 `responses` 库 mock 响应，不依赖真实网络；真实 API 验证只在手动冒烟时做一次。

### 失败测试处理流程

1. **复现**：单独运行失败用例 `pytest tests/test_xxx.py::TestClass::test_method -v`，确认可稳定复现。
2. **定位**：读 traceback + 日志；COM 测试失败优先检查 Excel 进程残留（`tasklist | findstr EXCEL`），有残留先 `taskkill /F /IM EXCEL.EXE` 再重跑。
3. **修复**：修改业务代码或测试代码，不得注释掉失败用例绕过。
4. **回归**：修复后跑全量 `pytest`，确认无新引入的失败。
5. **跳过须记录**：确因环境限制无法运行的用例，用 `@pytest.mark.skip(reason="环境限制：...")` 标记并在注释中说明，不得静默禁用。

### 测试运行策略（并行化与选择）

- **纯 Python 单元测试可并行**：`pytest -n auto`（需 `pytest-xdist`），加速快速反馈。
- **COM/Excel 集成测试禁止并行**：Excel 实例冲突，用 `@pytest.mark.serial` 标记，pytest 自动串行执行。
- **增量开发**：跑 `pytest -k "not serial"`（纯 Python 快速反馈，通常 < 5 秒）。
- **提交前**：跑全量 `pytest`（含 COM 集成，通常 < 1 分钟）。
- **CI 环境**：无 GUI 环境可能无法启动 Excel，CI 只跑纯 Python 单元测试；COM 测试在本地/专用机器跑。

## 三、COM 对象销毁规范（最高优先级）

```powershell
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    # ... 业务 ...
} finally {
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}
```

**清理铁律**：
- 必须 finally：异常时 Excel 进程不残留
- **禁止无差别杀进程**（`Stop-Process -Name EXCEL` 会误杀用户正在工作的 Excel）
- 只 Quit 不 ReleaseComObject → RCW 引用计数泄漏
- 释放顺序：子对象先失效（Workbook/Worksheet 引用先置空），父对象后 Quit
- WPS 场景补充：WPS 的 COM 清理与 Excel 类似，注意进程名（et/wps）差异

## 四、真实机器验证（原 18，7 条实测结论）

1. **COM 方式启动的 Excel 不加载 XLSTART**——验证 XLSTART 加载项必须用"双击文件/正常启动"路径
2. **Workbooks.Count 不枚举 XLSTART 加载项**——别用 Count 判断加载成功
3. **批量结束 Python 进程会误杀验证脚本自身**——精确按 PID 清理
4. **COM 新起实例的 AddIns2 不含未打开过的 XLSTART 文件**（2026-09-05 补测）
5. **XLSTART 分发不需要任何注册表注册**（2026-09-05 补测）
6. **门禁脚本内部 throw 会被同层容错 catch 吞掉**——门禁失败要检查返回码而非只看输出
7. **对已部署插件的写盘操作必须避开用户 Excel 会话**（2026-09-05 事故实录）
- **操作守则**：任何 COM 保存后必须回注 Ribbon（2026-09-05 实录）

## 五、安装与部署验证（6.5.10/11 门禁）

- 门禁 L：安装到 XLSTART 后、无 Excel 实例时执行——验证真实启动 + 加载项注册 + RunPython 链路
- 门禁 J：端到端运行（子进程/端口/API）
- 门禁 E：交付前最终验证（文件/导入/冒烟）

## 六、调试技巧（官方导航）

- `xlwings-0.37.3/docs/debugging.md`：UDF 断点（xw.serve）、日志
- `xlwings-0.37.3/docs/troubleshooting.md`：官方常见问题
- `docs/troubleshooting.md`：本技能沉淀的坑点（发布流水线/COM/Ribbon）
- `docs/troubleshooting-python.md`：Python/面板专项坑（WebView2 Runtime、端口、编码）



