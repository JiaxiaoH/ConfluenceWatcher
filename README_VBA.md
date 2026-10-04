# ConfluenceWatcher Excel VBA 方案

由于当前生成环境不能向新工作簿写入 VBA 工程，本包提供：

- `ConfluenceWatcher.xlsx`：已设计好的配置、当前状态、历史版本、更新记录、运行日志工作表
- `ConfluenceWatcher.bas`：完整宏模块

## 生成 .xlsm

1. 打开 `ConfluenceWatcher.xlsx`。
2. 按 `Alt+F11`。
3. 在 VBA 编辑器选择 `文件 > 导入文件`，导入 `ConfluenceWatcher.bas`。
4. 返回 Excel，另存为 `Excel 启用宏的工作簿 (*.xlsm)`。
5. 在“配置”填写认证信息。
6. 按 `Alt+F8` 运行 `CheckConfluenceTree`。

## 建议的认证方式

在 Windows 用户环境变量中设置 `CONFLUENCE_TOKEN`，宏会优先读取该变量。也可以直接填在“配置”页，但工作簿中会保存明文。

## 功能

- 递归扫描根页面和所有后代页面
- REST API 分页
- 当前状态刷新
- 每次扫描追加完整历史快照
- 更新记录：新增、更新、重命名、移动、删除或无权访问
- 运行成功和错误日志
- 初次运行默认只建立基线

## 可选自动运行

若希望打开工作簿时自动检查，在 VBA 编辑器的 `ThisWorkbook` 中加入：

```vb
Private Sub Workbook_Open()
    If CBool(ThisWorkbook.Worksheets("配置").Range("B12").Value) Then
        CheckConfluenceTree
    End If
End Sub
```

然后把“配置”页的“打开工作簿时自动检查”设为 TRUE，并通过 Windows 任务计划程序定时启动这个 `.xlsm`。
