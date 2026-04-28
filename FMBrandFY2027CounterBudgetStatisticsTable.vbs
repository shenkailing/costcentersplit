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
    lastCol = wsSrc.Cells(3, wsSrc.Columns.Count).End(xlToLeft).Column

    wsSrc.UsedRange.Copy
    wsRes.Cells(1, 1).PasteSpecial xlPasteValuesAndNumberFormats
    wsRes.Cells(1, 1).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    ' 删除新SHEET第一行空白行（如有），内容会自动上移
    If Application.WorksheetFunction.CountA(wsRes.Rows(1)) = 0 Then
        wsRes.Rows(1).Delete
    End If
    ' 此时前两行为表头和副表头，从第3行开始为有效数据
    Dim dataEndRow As Long
    dataEndRow = 3
    Do While wsRes.Cells(dataEndRow, 3).Value <> ""
        dataEndRow = dataEndRow + 1
    Loop
    dataEndRow = dataEndRow - 1

    wsRes.Rows(dataEndRow + 1).Insert
    Dim sumRow As Long
    sumRow = dataEndRow + 1
    wsRes.Cells(sumRow, 4).Value = "合计："
    For i = 3 To dataEndRow
        If wsRes.Cells(i, 1).Value = "Y" Then
            groupRowsKilian.Add i
        End If
        If wsRes.Cells(i, 2).Value = "Y" Then
            groupRowsFM.Add i
        End If
    Next i

    For j = 5 To lastCol
        If j < 17 Then
            wsRes.Cells(sumRow, j).Value = ""
        Else
            Dim sumFormula As String, first As Boolean
            sumFormula = ""
            first = True
            For i = 3 To dataEndRow
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
    Dim kilianStartRow As Long, kilianEndRow As Long
    kilianStartRow = resRow
    For i = 1 To groupRowsKilian.Count
        wsRes.Rows(resRow).Insert Shift:=xlDown
        wsRes.Rows(groupRowsKilian(i)).Copy
        wsRes.Rows(resRow).PasteSpecial xlPasteValuesAndNumberFormats
        wsRes.Rows(resRow).PasteSpecial xlPasteFormats
        resRow = resRow + 1
    Next i
    kilianEndRow = resRow - 1
    wsRes.Cells(resRow, 4).Value = "KL合计："
    For j = 5 To lastCol
        If j < 17 Then
            wsRes.Cells(resRow, j).Value = ""
        ElseIf kilianEndRow >= kilianStartRow Then
            wsRes.Cells(resRow, j).Formula = "=SUM(" & wsRes.Cells(kilianStartRow, j).Address & ":" & wsRes.Cells(kilianEndRow, j).Address & ")"
        Else
            wsRes.Cells(resRow, j).Value = ""
        End If
    Next j
    resRow = resRow + 1
    wsRes.Rows(resRow).Insert Shift:=xlDown
    resRow = resRow + 1
    Dim fmStartRow As Long, fmEndRow As Long
    fmStartRow = resRow
    For i = 1 To groupRowsFM.Count
        wsRes.Rows(resRow).Insert Shift:=xlDown
        wsRes.Rows(groupRowsFM(i)).Copy
        wsRes.Rows(resRow).PasteSpecial xlPasteValuesAndNumberFormats
        wsRes.Rows(resRow).PasteSpecial xlPasteFormats
        resRow = resRow + 1
    Next i
    fmEndRow = resRow - 1
    wsRes.Cells(resRow, 4).Value = "FM合计："
    For j = 5 To lastCol
        If j < 17 Then
            wsRes.Cells(resRow, j).Value = ""
        ElseIf fmEndRow >= fmStartRow Then
            wsRes.Cells(resRow, j).Formula = "=SUM(" & wsRes.Cells(fmStartRow, j).Address & ":" & wsRes.Cells(fmEndRow, j).Address & ")"
        Else
            wsRes.Cells(resRow, j).Value = ""
        End If
    Next j
    resRow = resRow + 1
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
    For i = 3 To dataEndRow
        If wsRes.Cells(i, 1).Value = "Y" Or wsRes.Cells(i, 2).Value = "Y" Then
            delRows.Add i
        End If
    Next i
    For idx = delRows.Count To 1 Step -1
        wsRes.Rows(delRows(idx)).Delete
    Next idx

    ' find fm sum row number
    Dim fmTotalRow As Long
    fmTotalRow = 0
    For i = 1 To wsRes.UsedRange.Rows.Count
        If wsRes.Cells(i, 4).Value = "FM合计：" Then
            fmTotalRow = i
            Exit For
        End If
    Next i

    ' If found, delete all rows below it
    If fmTotalRow > 0 Then
        Dim lastUsedRow As Long
        lastUsedRow = 0
        For j = 1 To lastCol
            Dim colLastRow As Long
            colLastRow = wsRes.Cells(wsRes.Rows.Count, j).End(xlUp).Row
            If colLastRow > lastUsedRow Then
                lastUsedRow = colLastRow
            End If
        Next j
        If fmTotalRow < lastUsedRow Then
            wsRes.Rows((fmTotalRow + 1) & ":" & lastUsedRow).Delete
        End If
    End If



    Dim fmtLastRow As Long, fmtLastCol As Long
    fmtLastRow = wsRes.UsedRange.Rows(wsRes.UsedRange.Rows.Count).Row
    fmtLastCol = wsRes.Cells(2, wsRes.Columns.Count).End(xlToLeft).Column '修正为第2行的最后一列，确保所有数据列都刷样式
    Dim rowIdx As Long
    ' 复制第4、5行样式到第6行及以后所有数据行
    For rowIdx = 6 To fmtLastRow
        If (rowIdx Mod 2) = 0 Then
            wsRes.Range(wsRes.Cells(4, 1), wsRes.Cells(4, fmtLastCol)).Copy
        Else
            wsRes.Range(wsRes.Cells(5, 1), wsRes.Cells(5, fmtLastCol)).Copy
        End If
        wsRes.Range(wsRes.Cells(rowIdx, 1), wsRes.Cells(rowIdx, fmtLastCol)).PasteSpecial Paste:=xlPasteFormats
    Next rowIdx
    Application.CutCopyMode = False



    Dim newWb As Workbook
    Dim savePath As String, fName As String
    Dim overwrite As VbMsgBoxResult

    fName = ThisWorkbook.Name
    If InStrRev(fName, ".") > 0 Then
        fName = Left(fName, InStrRev(fName, ".") - 1)
    End If
    savePath = ThisWorkbook.Path & "\" & fName & ".xlsx"

   
    If Dir(savePath) <> "" Then
        overwrite = MsgBox("The file '" & savePath & "' already exists. Do you want to overwrite it?", vbYesNo + vbQuestion, "The file already exists")
        If overwrite = vbNo Then Exit Sub
    End If


    wsRes.Copy
    Set newWb = ActiveWorkbook

    Application.DisplayAlerts = False
    newWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    Application.DisplayAlerts = True
    newWb.Close SaveChanges:=False
    MsgBox "The report has been generated successfully to " & savePath, vbInformation

End Sub



