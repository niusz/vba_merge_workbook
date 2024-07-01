Option Explicit
Sub 合并工作表1()
    'Application.ScreenUpdating = False
    Dim FileName As String, Sht As Worksheet, Wb As Workbook,FolderName As String
    Dim fd As FileDialog
    ' 创建FileDialog对象，设置为打开文件选择对话框
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "请选择一个目录"
        .Show
        If .SelectedItems.Count > 0 Then
            FolderName = .SelectedItems(1)
        Else
            MsgBox "没有选择任何文件！"
            Exit Sub
        End If
    End With
    FileName = Dir(FolderName & "\*.xls?")
    Do While FileName <> ""
       'Workbooks.Open FileName:="E:\项目\VBA\合并多个工作表\source\" & FileName
       Workbooks.Open FileName:= FolderName & "\" & FileName
       Set Wb = ActiveWorkbook
       Set Sht = ThisWorkbook.Worksheets("Sheet1")
       For Each Sht In Wb.Worksheets
           If (Sht.Name = "个人身份信息") Then
               Sht.Copy after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
           End If
       Next Sht
       Wb.Close savechanges:=False
       FileName = Dir
    Loop
    Application.ScreenUpdating = True
    '调用Sub
    填入数据1
End Sub
Sub 填入数据1()
    Dim Sht As Worksheet, Wb As Workbook, Ws As Worksheet
    Dim EndRow As Long
    Set Sht = ActiveWorkbook.Worksheets("Sheet1")
    Set Wb = ActiveWorkbook
    For Each Ws In Wb.Worksheets
        EndRow = Sht.Range("A1048576").End(xlUp).Row
        If (Ws.Name <> "Sheet1") Then
            Sht.Range("A" & EndRow + 1).Value = Ws.Range("B3").Value '姓名
            Sht.Range("B" & EndRow + 1).Value = Ws.Range("D3").Value '性别
            Sht.Range("C" & EndRow + 1).Value = Ws.Range("F3").Value '出生年月
            Sht.Range("D" & EndRow + 1).Value = Ws.Range("B4").Value '国籍
            Sht.Range("E" & EndRow + 1).Value = Ws.Range("D4").Value '民族
            Sht.Range("F" & EndRow + 1).Value = Ws.Range("F4").Value '籍贯
            Sht.Range("G" & EndRow + 1).Value = Ws.Range("B5").Value '身份证号
            Sht.Range("H" & EndRow + 1).Value = Ws.Range("D5").Value '身份证开始时间
            Sht.Range("I" & EndRow + 1).Value = Ws.Range("F5").Value '身份证到期时间
            Sht.Range("J" & EndRow + 1).Value = Ws.Range("B6").Value '婚育状态
            Sht.Range("K" & EndRow + 1).Value = Ws.Range("D6").Value '户口类型
            Sht.Range("L" & EndRow + 1).Value = Ws.Range("B7").Value '户口地址
            Sht.Range("M" & EndRow + 1).Value = Ws.Range("B8").Value  '现居住地
            Sht.Range("N" & EndRow + 1).Value = Ws.Range("B10").Value  '本人联系电话
            Sht.Range("O" & EndRow + 1).Value = Ws.Range("D10").Value  '邮箱
            Sht.Range("P" & EndRow + 1).Value = Ws.Range("F10").Value '微信号
            Sht.Range("Q" & EndRow + 1).Value = Ws.Range("B11").Value '紧急联系人
            Sht.Range("R" & EndRow + 1).Value = Ws.Range("D11").Value '紧急联系人关系
            Sht.Range("S" & EndRow + 1).Value = Ws.Range("F11").Value  '紧急联系人电话
            '社会关系
            Sht.Range("T" & EndRow + 1).Value = Ws.Range("B17").Value  '政治面貌
            Sht.Range("U" & EndRow + 1).Value = Ws.Range("D17").Value  '组织关系所在地
            Sht.Range("V" & EndRow + 1).Value = Ws.Range("F17").Value  '档案所在地
            Sht.Range("W" & EndRow + 1).Value = Ws.Range("B19").Value  '宗教信仰
        End If
    Next Ws
End Sub