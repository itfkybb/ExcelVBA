Sub MergeAllExcelFilesInDirectory()
    Dim path As String
    Dim filePattern As String
    Dim currentFile As String
    Dim targetWorkbook As Workbook
    Dim sourceWorkbook As Workbook
    Dim targetSheet As Worksheet
    Dim sourceSheet As Worksheet
    Dim lastRow As Long
    Dim destLastRow As Long
    Dim fileCount As Integer
    
    ' 设置当前工作簿为目标工作簿
    Set targetWorkbook = ThisWorkbook
    
    ' 获取当前文件所在目录路径
    path = ThisWorkbook.path & "\"
    filePattern = "*.xls*"
    
    ' 创建新工作表用于存放合并的数据
    Set targetSheet = targetWorkbook.Sheets.Add(After:=targetWorkbook.Sheets(targetWorkbook.Sheets.Count))
    targetSheet.Name = "合并数据"
    
    ' 初始化文件计数器
    fileCount = 0
    
    ' 获取第一个Excel文件
    currentFile = Dir(path & filePattern)
    
    ' 禁用屏幕更新和自动计算以提高性能
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' 遍历目录中的所有Excel文件
    Do While currentFile <> ""
        ' 跳过当前工作簿自身
        If currentFile <> ThisWorkbook.Name Then
            ' 打开源工作簿
            Set sourceWorkbook = Workbooks.Open(path & currentFile)
            
            ' 遍历源工作簿中的所有工作表
            For Each sourceSheet In sourceWorkbook.Sheets
                ' 查找源工作表的最后一行
                lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 1).End(xlUp).Row
                
                ' 如果工作表有数据（超过1行）
                If lastRow > 1 Then
                    ' 查找目标工作表的最后一行
                    destLastRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row + 1
                    
                    ' 复制数据（从第2行开始，避免重复标题）
                    If destLastRow = 2 Then ' 第一次复制时包括标题
                        sourceSheet.Range("A1").CurrentRegion.Copy targetSheet.Cells(destLastRow - 1, 1)
                    Else ' 后续复制跳过标题
                        sourceSheet.Range("A2:A" & lastRow).EntireRow.Copy targetSheet.Cells(destLastRow, 1)
                    End If
                End If
            Next sourceSheet
            
            ' 关闭源工作簿，不保存更改
            sourceWorkbook.Close SaveChanges:=False
            fileCount = fileCount + 1
        End If
        
        ' 获取下一个文件
        currentFile = Dir
    Loop
    
    ' 恢复屏幕更新和自动计算
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    ' 显示完成消息
    MsgBox "合并完成！共处理了 " & fileCount & " 个文件。", vbInformation, "完成"
End Sub
