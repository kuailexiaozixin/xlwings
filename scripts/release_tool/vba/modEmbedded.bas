"""
xlwings release 复刻工具 - VBA 嵌入式 RunPython 模块

这是复刻版的无加载项 VBA 模块（简化版 xlwings.bas），
只包含嵌入式代码运行所需的最小逻辑：
- RunEmbedded 子过程：检测 .py sheet → 调用 Python 解出并执行
- 不依赖 xlwings 加载项引用

实际部署说明：
- 完整复刻建议直接使用官方生成的 xlwings.bas（开源），
  但需把其中的嵌入式分支从
      "import xlwings.pro;xlwings.pro.runpython_embedded_code('...')"
  改为
      "import release_tool.runtime_embedded;release_tool.runtime_embedded.run('...')"
- 本文件是裁剪版参考，仅用于演示嵌入式调用链
"""
Option Explicit

' 嵌入式 RunPython：从 .py sheet 解出代码并执行
Public Sub RunEmbedded(PythonCommand As String)
    Dim interpreter As String
    Dim pyCmd As String
    
    interpreter = GetConfig("INTERPRETER_WIN", "")
    If interpreter = "" Then
        interpreter = GetConfig("INTERPRETER", "python")
    End If
    
    ' 调用复刻版运行时（替代 xlwings.pro）
    pyCmd = "import release_tool.runtime_embedded;release_tool.runtime_embedded.run('" & _
            Replace(PythonCommand, "'", "\'") & "')"
    
    ExecuteCommand interpreter, pyCmd
End Sub

' 检测工作簿是否有嵌入式代码（.py sheet）
Public Function HasEmbeddedCode() As Boolean
    Dim sht As Worksheet
    For Each sht In ThisWorkbook.Worksheets
        If Right$(sht.Name, 3) = ".py" Then
            HasEmbeddedCode = True
            Exit Function
        End If
    Next
    HasEmbeddedCode = False
End Function

' 读取配置表
Private Function GetConfig(key As String, Optional default As String) As String
    On Error Resume Next
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Worksheets("xlwings.conf")
    If sheet Is Nothing Then
        GetConfig = default
        Exit Function
    End If
    ' 简化：遍历两列找键
    Dim cell As Range
    For Each cell In sheet.Range("A:A")
        If cell.Value = key Then
            GetConfig = cell.Offset(0, 1).Value
            Exit Function
        End If
    Next
    GetConfig = default
End Function

' 执行命令（简化版 ExecuteWindows）
Private Sub ExecuteCommand(interpreter As String, pyCmd As String)
    Dim runCmd As String
    runCmd = """" & interpreter & """ -B -c """ & _
             "import sys;sys.path.insert(0, r'RELEASE_TOOL_DIR');" & pyCmd & """"
    ' 实际部署时由 WScript.Shell 或 Shell 执行
    ' 此处留给集成脚本（真实运行需 Excel 环境）
    Debug.Print "CMD: " & runCmd
End Sub
