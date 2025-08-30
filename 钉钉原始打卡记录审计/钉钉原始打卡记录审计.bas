Sub FindMultiUserDevicesWithCount()
    ' 声明变量
    Dim wsSource As Worksheet, wsResult As Worksheet
    Dim dict As Object ' 用于存储设备号和对应的员工集合
    Dim lastRow As Long, i As Long
    Dim deviceID As String, employeeName As String
    Dim key As Variant, empKey As Variant
    Dim outputRow As Long
    Dim nameList As String
    
    ' 设置源数据工作表 (假设数据在原始记录)
    On Error Resume Next
    Set wsSource = ThisWorkbook.Worksheets("原始记录")
    On Error GoTo 0
    
    If wsSource Is Nothing Then
        MsgBox "找不到工作表 '原始记录'，请修改代码中的工作表名称。", vbExclamation
        Exit Sub
    End If
    
    ' 创建或清空结果工作表
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets("异常设备报告").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set wsResult = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    wsResult.Name = "异常设备报告"
    
    ' 在结果表创建标题
    With wsResult
        .Range("A1").Value = "设备编号"
        .Range("B1").Value = "使用员工数量"
        .Range("C1").Value = "使用员工名单及打卡次数"
        .Range("D1").Value = "设备持有人"
        .Range("E1").Value = "代打卡员工"
        .Range("A1:E1").Font.Bold = True
        .Columns("A:E").AutoFit
    End With
    
    outputRow = 2 ' 从第2行开始输出结果
    
    ' 创建字典对象来存储数据
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' 找出源数据的最后一行 (使用P列确定行数)
    lastRow = wsSource.Cells(wsSource.Rows.Count, "P").End(xlUp).Row
    
    ' 检查是否有数据
    If lastRow < 2 Then
        wsResult.Cells(2, 1).Value = "源数据表中没有找到有效数据。"
        wsResult.Columns("A:E").AutoFit
        wsResult.Activate
        Exit Sub
    End If
    
    ' 遍历所有数据行 (假设第1行是标题，从第2行开始)
    For i = 2 To lastRow
        On Error Resume Next ' 防止类型转换错误
        deviceID = Trim(CStr(wsSource.Cells(i, "P").Value)) ' P列是设备号
        employeeName = Trim(CStr(wsSource.Cells(i, "A").Value)) ' A列是员工姓名
        On Error GoTo 0
        
        ' 跳过空设备号或空姓名的行
        If deviceID = "" Or employeeName = "" Then GoTo NextRow
        
        ' 如果字典中没有这个设备号，则添加一个新字典
        If Not dict.Exists(deviceID) Then
            dict.Add deviceID, CreateObject("Scripting.Dictionary")
        End If
        
        ' 如果该设备号的字典中没有这个员工，则添加并初始化计数为1
        If Not dict(deviceID).Exists(employeeName) Then
            dict(deviceID).Add employeeName, 1
        Else
            ' 如果已存在，则计数加1
            dict(deviceID)(employeeName) = dict(deviceID)(employeeName) + 1
        End If
        
NextRow:
    Next i
    
    ' 遍历字典，找出使用员工数 > 1 的设备
    For Each key In dict.Keys
        If dict(key).Count > 1 Then
            ' 输出到结果表
            wsResult.Cells(outputRow, 1).Value = key ' 设备号
            wsResult.Cells(outputRow, 2).Value = dict(key).Count ' 员工数量
            
            ' 将员工姓名和打卡次数集合连接成一个字符串
            nameList = ""
            For Each empKey In dict(key).Keys
                nameList = nameList & empKey & "(" & dict(key)(empKey) & "次), "
            Next empKey
            
            ' 去掉最后一个逗号和空格
            If Len(nameList) > 0 Then
                nameList = Left(nameList, Len(nameList) - 2)
            End If
            
            wsResult.Cells(outputRow, 3).Value = nameList ' 员工名单及打卡次数
            
            ' 新增功能：识别设备持有人和代打卡员工
            Dim maxCount As Long
            Dim holder As String
            Dim proxyEmployees As String
            
            ' 找出打卡次数最多的员工
            maxCount = 0
            holder = ""
            For Each empKey In dict(key).Keys
                If dict(key)(empKey) > maxCount Then
                    maxCount = dict(key)(empKey)
                    holder = empKey
                End If
            Next
            
            ' 构建代打卡员工名单（包含打卡次数）
            proxyEmployees = ""
            For Each empKey In dict(key).Keys
                If empKey <> holder Then
                    proxyEmployees = proxyEmployees & empKey & "(" & dict(key)(empKey) & "次), "
                End If
            Next
            
            ' 去掉最后一个逗号和空格
            If Len(proxyEmployees) > 0 Then
                proxyEmployees = Left(proxyEmployees, Len(proxyEmployees) - 2)
            End If
            
            ' 写入持有人和代打卡员工信息
            wsResult.Cells(outputRow, 4).Value = holder & "(" & maxCount & "次)"
            wsResult.Cells(outputRow, 5).Value = proxyEmployees
            
            outputRow = outputRow + 1
        End If
    Next key
    
    ' 如果没有找到异常设备，提示用户
    If outputRow = 2 Then
        wsResult.Cells(2, 1).Value = "未发现一个设备对应多个员工的情况。"
    Else
        ' 对结果表进行排序（按员工数量降序）
        With wsResult
            If outputRow > 2 Then
                .Range("A1:E" & outputRow - 1).Sort Key1:=.Range("B2"), Order1:=xlDescending, Header:=xlYes
            End If
        End With
    End If
    
    ' 自动调整列宽
    wsResult.Columns("A:E").AutoFit
    wsResult.Activate ' 切换到结果工作表
    
    MsgBox "分析完成！共找到 " & (outputRow - 2) & " 个异常设备。结果已输出到工作表【异常设备报告】。", vbInformation
End Sub
