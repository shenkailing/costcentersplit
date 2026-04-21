Sub costcentersplit_Click()
Dim wsSrc As Worksheet, wsRes As Worksheet
    Dim lastRow As Long, lastCol As Long, resRow As Long
    Dim i As Long, j As Long
    Dim brand As String, storeName As String
    Dim startCol As Long
    Dim groupRowsKilian As Collection, groupRowsFM As Collection
    Dim rowDict As Object
    Dim subtotalRow As Long
    Set groupRowsKilian = New Collection
    Set groupRowsFM = New Collection
    Set wsSrc = ThisWorkbook.Sheets(1)
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("CostCenterGroupBySumResult").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsRes = ThisWorkbook.Sheets.Add(After:=wsSrc)
    wsRes.Name = "CostCenterGroupBySumResult"

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    lastCol = wsSrc.Cells(2, wsSrc.Columns.Count).End(xlToLeft).Column

    ' 1. 只复制原表数据和表头，不复制按钮和控件

    wsSrc.UsedRange.Copy
    wsRes.Cells(1, 1).PasteSpecial xlPasteValuesAndNumberFormats
    wsRes.Cells(1, 1).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    ' 删除新SHEET第一行（按钮行）
    wsRes.Rows(1).Delete


    ' 找到新表常规数据的最后一行（C列为空的第一行）
    Dim dataEndRow As Long
    dataEndRow = 2
    Do While wsRes.Cells(dataEndRow, 3).Value <> ""
        dataEndRow = dataEndRow + 1
    Loop
    dataEndRow = dataEndRow - 1

    ' 在常规数据最后一行下插入合计行
    wsRes.Rows(dataEndRow + 1).Insert
    Dim sumRow As Long
    sumRow = dataEndRow + 1
    wsRes.Cells(sumRow, 4).Value = "合计："
    ' 先收集分组，只执行一次
    For i = 2 To dataEndRow
        If wsRes.Cells(i, 1).Value = "Y" Then
            groupRowsKilian.Add i
        End If
        If wsRes.Cells(i, 2).Value = "Y" Then
            groupRowsFM.Add i
        End If
    Next i

    For j = 5 To lastCol
        If j < 19 Then
            wsRes.Cells(sumRow, j).Value = ""
        Else
            Dim sumFormula As String, first As Boolean
            sumFormula = ""
            first = True
            For i = 2 To dataEndRow
                ' 统计合计用
                If wsRes.Cells(i, 1).Value <> "Y" And wsRes.Cells(i, 2).Value <> "Y" And wsRes.Cells(i, 3).Value <> "" Then
                    If first Then
                        sumFormula = wsRes.Cells(i, j).Address
                        first = False
                    Else
                        sumFormula = sumFormula & "," & wsRes.Cells(i, j).Address
                    End If
                End If
            Next i
            If sumFormula <> "" Then
                wsRes.Cells(sumRow, j).Formula = "=SUM(" & sumFormula & ")"
            Else
                wsRes.Cells(sumRow, j).Value = ""
            End If
        End If
    Next j

    resRow = sumRow + 1
    
    wsRes.Rows(resRow).Insert Shift:=xlDown
    resRow = resRow + 1
    ' Kilian分组
    Dim kilianStartRow As Long
    kilianStartRow = resRow
    For i = 1 To groupRowsKilian.Count
        wsRes.Rows(resRow).Insert Shift:=xlDown
        wsRes.Rows(groupRowsKilian(i)).Copy
        wsRes.Rows(resRow).PasteSpecial xlPasteValuesAndNumberFormats
        wsRes.Rows(resRow).PasteSpecial xlPasteFormats
        resRow = resRow + 1
    Next i
    ' Kilian小计
    wsRes.Cells(resRow, 4).Value = "Kilian合计"
    For j = 5 To lastCol
        If j < 19 Then
            wsRes.Cells(resRow, j).Value = ""
        Else
            wsRes.Cells(resRow, j).Formula = "=SUM(" & wsRes.Cells(kilianStartRow, j).Address & ":" & wsRes.Cells(resRow - 1, j).Address & ")"
        End If
    Next j
    resRow = resRow + 1
    ' 合计行下插入空行
    wsRes.Rows(resRow).Insert Shift:=xlDown
    resRow = resRow + 1
    ' FM分组
    Dim fmStartRow As Long
    fmStartRow = resRow
    For i = 1 To groupRowsFM.Count
        wsRes.Rows(resRow).Insert Shift:=xlDown
        wsRes.Rows(groupRowsFM(i)).Copy
        wsRes.Rows(resRow).PasteSpecial xlPasteValuesAndNumberFormats
        wsRes.Rows(resRow).PasteSpecial xlPasteFormats
        resRow = resRow + 1
    Next i
    ' FM小计
    wsRes.Cells(resRow, 4).Value = "FM合计"
    For j = 5 To lastCol
        If j < 19 Then
            wsRes.Cells(resRow, j).Value = ""
        Else
            wsRes.Cells(resRow, j).Formula = "=SUM(" & wsRes.Cells(fmStartRow, j).Address & ":" & wsRes.Cells(resRow - 1, j).Address & ")"
        End If
    Next j
    resRow = resRow + 1
    ' 合计行下插入空行
    wsRes.Rows(resRow).Insert Shift:=xlDown
    resRow = resRow + 1
    
    
    
    dataEndRow = 2
    Do While wsRes.Cells(dataEndRow, 3).Value <> ""
        dataEndRow = dataEndRow + 1
    Loop
    Dim idx As Long
    Dim delRows As Collection
    Set delRows = New Collection
    dataEndRow = dataEndRow - 1
    For i = 2 To dataEndRow
        ' 统计合计用
        If wsRes.Cells(i, 1).Value = "Y" Or wsRes.Cells(i, 2).Value = "Y" Then
            delRows.Add i
        End If
    Next i
    For idx = delRows.Count To 1 Step -1
        wsRes.Rows(delRows(idx)).Delete
    Next idx

    ' ====== 重新设置隔行变色格式 ======
    Dim fmtLastRow As Long, fmtLastCol As Long
    fmtLastRow = wsRes.UsedRange.Rows(wsRes.UsedRange.Rows.Count).Row
    fmtLastCol = wsRes.Cells(1, wsRes.Columns.Count).End(xlToLeft).Column
    Dim rowIdx As Long, srcRow As Long
    For rowIdx = 2 To fmtLastRow
        ' 2、3行分别为第2、3行的格式模板，交替刷
        If (rowIdx Mod 2) = 0 Then
            srcRow = 2
        Else
            srcRow = 3
        End If
        wsRes.Range(wsRes.Cells(srcRow, 1), wsRes.Cells(srcRow, fmtLastCol)).Copy
        wsRes.Range(wsRes.Cells(rowIdx, 1), wsRes.Cells(rowIdx, fmtLastCol)).PasteSpecial Paste:=xlPasteFormats
    Next rowIdx
    Application.CutCopyMode = False
    ' ====== 隔行变色格式结束 ======

    MsgBox "分组数据已追加到新表 'CostCenterGroupBySumResult' 中！", vbInformation

    ' ====== 新增：将新Sheet单独保存为Excel文件 ======
    Dim newWb As Workbook
    Dim savePath As String, fName As String
    Dim overwrite As VbMsgBoxResult

    ' 获取当前工作簿路径和文件名
    fName = ThisWorkbook.Name
    If InStrRev(fName, ".") > 0 Then
        fName = Left(fName, InStrRev(fName, ".") - 1)
    End If
    savePath = ThisWorkbook.Path & "\" & fName & ".xlsx"

    ' 检查是否存在同名文件
    If Dir(savePath) <> "" Then
        overwrite = MsgBox("文件 '" & savePath & "' 已存在，是否覆盖？", vbYesNo + vbQuestion, "文件已存在")
        If overwrite = vbNo Then Exit Sub
    End If

    ' 复制Sheet到新工作簿
    wsRes.Copy
    Set newWb = ActiveWorkbook
    ' 保存为xlsx格式
    Application.DisplayAlerts = False
    newWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    newWb.Close SaveChanges:=False
    MsgBox "Sheet已单独保存为：" & savePath, vbInformation
    ' ====== 新增结束 ======
End Sub


