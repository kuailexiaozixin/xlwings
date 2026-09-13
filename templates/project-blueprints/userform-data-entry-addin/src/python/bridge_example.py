# -*- coding: utf-8 -*-
"""
userform-data-entry-addin 蓝图 —— Python 侧桥接示例
对应 VBA 骨架中的 RunPythonAction 入口：
    RunPython "import __PROJECT_NAME__; __PROJECT_NAME__.python_action()"

本文件写入 quickstart 生成的同名 .py 中（例如 __PROJECT_NAME__.py）。
演示在 xlwings 加载项中，VBA 负责 Excel 对象模型（UserForm/ListObject 结构），
Python 负责复杂计算（pandas 聚合、网络请求、正则清洗）的混合形态。
"""
import xlwings as xw
import pandas as pd


def python_action():
    """示例：读取当前工作表数据区，用 pandas 做按列求和后写回说明单元格。"""
    wb = xw.Book.caller()
    sheet = wb.sheets.active

    # 读取数据（A1 起向下扩展；期望含表头）
    data = sheet.range("A1").expand("down").options(pd.DataFrame).value

    if data is None or data.empty:
        sheet["H1"].value = "未读取到数据"
        return

    # 示例：对数值列求和（业务逻辑请按需替换）
    numeric = data.select_dtypes(include="number")
    total = float(numeric.sum().sum()) if not numeric.empty else 0.0

    sheet["H1"].value = f"Python 处理完成，数值合计: {total:,.2f}"


# 可在单元格调用的 UDF 示例
@xw.func
@xw.arg("df", pd.DataFrame)
def order_status_count(region: str, df) -> int:
    """=order_status_count("华南", 数据区域)：统计指定区域的订单行数。

    注意：pandas DataFrame 转换要求区域第一行为表头。
    """
    if df is None or df.empty:
        return 0
    col = df.iloc[:, 1]  # 第二列视为"区域"
    return int((col.astype(str) == region).sum())


if __name__ == "__main__":
    # 无 Excel 时模拟调用环境（quickstart 标准写法）
    xw.Book("__PROJECT_NAME__.xlsm").set_mock_caller()
    python_action()
