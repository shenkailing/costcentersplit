Sub costcentersplit_Click()
Dim wsSrc As Worksheet, wsRes As Worksheet
    Dim lastCol As Long, resRow As Long
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
    wsRes.Cells(sumRow, 4).Value = "Total:"
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
    wsRes.Cells(resRow, 4).Value = "KL Total:"
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
    wsRes.Cells(resRow, 4).Value = "FM Total:"
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
        If wsRes.Cells(i, 4).Value = "FM Total:" Then
            fmTotalRow = i
            Exit For
        End If
    Next i

    Dim fmtLastRow As Long, fmtLastCol As Long
    fmtLastRow = wsRes.UsedRange.Rows(wsRes.UsedRange.Rows.Count).Row
    fmtLastCol = wsRes.Cells(3, wsRes.Columns.Count).End(xlToLeft).Column '修正为第2行的最后一列，确保所有数据列都刷样式
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


    Dim r As Long
    Dim valA As String
    Dim valB As String

    For r = 2 To fmtLastRow
        valA = Trim$(UCase$(wsRes.Cells(r, "A").Value))
        valB = Trim$(UCase$(wsRes.Cells(r, "B").Value))
    
        '1. Total行不赋值
        If valB = "TOTAL:" _
           Or valB = "KL TOTAL:" _
           Or valB = "FM TOTAL:" Then
            '不处理
    
        '2. 特殊空行不赋值：A、B、C都为空
        ElseIf valA = "" And valB = "" And Trim$(wsRes.Cells(r, "C").Value) = "" Then
            '不处理
    
        '3. A、B都为空，但C原来有内容，赋值KL/FM
        ElseIf valA = "" And valB = "" Then
            wsRes.Cells(r, "C").Value = "KL/FM"
    
        '4. A列有Y，赋值KL
        ElseIf valA = "Y" Then
            wsRes.Cells(r, "C").Value = "KL"
    
        '5. B列有Y，赋值FM
        ElseIf valB = "Y" Then
            wsRes.Cells(r, "C").Value = "FM"
    
        Else
            wsRes.Cells(r, "C").Value = ""
        End If
    Next r
    'delete column A and B, after style is applied, to avoid impact on formatting
    wsRes.Columns("A:B").Delete Shift:=xlToLeft

    ' set brand to KL/FM only for detail rows before the first Total section
    Dim firstTotalRow As Long
    firstTotalRow = 0
    For i = 3 To wsRes.UsedRange.Rows.Count
        If Trim$(CStr(wsRes.Cells(i, 2).Value)) = "Total:" Then
            firstTotalRow = i
            Exit For
        End If
    Next i

    If firstTotalRow > 3 Then
        For i = 3 To firstTotalRow - 1
            If Application.WorksheetFunction.CountA(wsRes.Rows(i)) > 0 Then
                wsRes.Cells(i, 1).Value = "KL/FM"
            End If
        Next i
    End If

    ' final cleanup: remove all non-empty rows below "FM Total:" in result sheet
    Dim lastUsedRowCleanup As Long
    Dim lastContentCell As Range
    If fmTotalRow > 0 Then
        Set lastContentCell = wsRes.Cells.Find(What:="*", _
                                               After:=wsRes.Cells(1, 1), _
                                               LookIn:=xlFormulas, _
                                               LookAt:=xlPart, _
                                               SearchOrder:=xlByRows, _
                                               SearchDirection:=xlPrevious, _
                                               MatchCase:=False)
        If Not lastContentCell Is Nothing Then
            lastUsedRowCleanup = lastContentCell.Row
        Else
            lastUsedRowCleanup = fmTotalRow
        End If

        For i = lastUsedRowCleanup To fmTotalRow + 1 Step -1
            If Application.WorksheetFunction.CountA(wsRes.Rows(i)) > 0 Then
                wsRes.Rows(i).Delete
            End If
        Next i
    End If

    Dim newWb As Workbook
    Dim savePath As String, fName As String
    Dim basePath As String, localFolder As String
    Dim saveResult As Variant
    Dim tempPath As String, tempFile As String
    Dim fileExists As Boolean
    Dim overwrite As VbMsgBoxResult
    Dim errMsg As String

    fName = ThisWorkbook.Name
    If InStrRev(fName, ".") > 0 Then
        fName = Left(fName, InStrRev(fName, ".") - 1)
    End If

    basePath = Trim$(ThisWorkbook.Path)

    If IsHttpPath(basePath) Then
        localFolder = MapOneDriveUrlToLocalFolder(basePath)
        If localFolder <> "" Then
            savePath = localFolder & "\" & fName & "_copy.xlsx"
        Else
            saveResult = Application.GetSaveAsFilename( _
                InitialFileName:=fName & "_copy.xlsx", _
                FileFilter:="Excel Workbook (*.xlsx), *.xlsx", _
                Title:="Save report as")
            If VarType(saveResult) = vbBoolean And saveResult = False Then Exit Sub
            savePath = CStr(saveResult)
        End If
    ElseIf basePath = "" Then
        saveResult = Application.GetSaveAsFilename( _
            InitialFileName:=fName & "_copy.xlsx", _
            FileFilter:="Excel Workbook (*.xlsx), *.xlsx", _
            Title:="Save report as")
        If VarType(saveResult) = vbBoolean And saveResult = False Then Exit Sub
        savePath = CStr(saveResult)
    Else
        savePath = basePath & "\" & fName & "_copy.xlsx"
    End If

    If LCase$(Right$(savePath, 5)) <> ".xlsx" Then
        savePath = savePath & ".xlsx"
    End If

    fileExists = FileExistsNoErr(savePath)

    If fileExists Then
        overwrite = MsgBox("The file '" & savePath & "' already exists. Do you want to overwrite it?", vbYesNo + vbQuestion, "The file already exists")
        If overwrite = vbNo Then Exit Sub
    End If


    wsRes.Copy
    Set newWb = ActiveWorkbook

    tempPath = Environ$("TEMP")
    If Right$(tempPath, 1) <> "\" Then
        tempPath = tempPath & "\"
    End If
    tempFile = tempPath & "CostCenterGroupBySum_" & Format(Now, "yyyymmdd_hhnnss") & "_copy.xlsx"

    On Error GoTo SaveErr
    Application.DisplayAlerts = False
    newWb.SaveAs Filename:=tempFile, FileFormat:=xlOpenXMLWorkbook, Local:=True
    Application.DisplayAlerts = True
    On Error GoTo 0
    newWb.Close SaveChanges:=False

    On Error GoTo CopyErr
    If FileExistsNoErr(savePath) Then Kill savePath
    FileCopy tempFile, savePath
    On Error GoTo 0

    On Error Resume Next
    If FileExistsNoErr(tempFile) Then Kill tempFile
    On Error GoTo 0

    MsgBox "The report has been generated successfully to " & savePath, vbInformation
    Exit Sub

SaveErr:
    errMsg = Err.Description
    Application.DisplayAlerts = True
    On Error Resume Next
    Set newWb = Nothing
    On Error GoTo 0
    MsgBox "Failed to generate xlsx file." & vbCrLf & _
           "Details: " & errMsg, vbExclamation
    Exit Sub

CopyErr:
    errMsg = Err.Description
    On Error Resume Next
    Set newWb = Nothing
    If FileExistsNoErr(tempFile) Then Kill tempFile
    On Error GoTo 0
    MsgBox "The report file was created, but copy to target folder failed." & vbCrLf & _
           "Target: " & savePath & vbCrLf & _
           "Details: " & errMsg, vbExclamation

End Sub

Private Function FileExistsNoErr(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExistsNoErr = (Len(Dir$(filePath, vbNormal)) > 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function IsHttpPath(ByVal pathText As String) As Boolean
    Dim p As String
    p = LCase$(Trim$(pathText))
    IsHttpPath = (Left$(p, 7) = "http://" Or Left$(p, 8) = "https://")
End Function

Private Function MapOneDriveUrlToLocalFolder(ByVal urlPath As String) As String
    On Error GoTo FailMap
    Dim normalized As String
    Dim relPath As String
    Dim p As Long
    Dim rootPath As String

    normalized = Replace(urlPath, "/", "\\")
    p = InStr(1, LCase$(normalized), "d.docs.live.net\\", vbTextCompare)
    If p = 0 Then GoTo FailMap

    relPath = Mid$(normalized, p + Len("d.docs.live.net\\"))
    p = InStr(1, relPath, "\\")
    If p = 0 Then GoTo FailMap

    ' remove CID segment, keep sub-folder under OneDrive root
    relPath = Mid$(relPath, p + 1)

    rootPath = Environ$("OneDriveConsumer")
    If rootPath = "" Then rootPath = Environ$("OneDrive")
    If rootPath = "" Then GoTo FailMap

    If Right$(rootPath, 1) = "\" Then
        MapOneDriveUrlToLocalFolder = rootPath & relPath
    Else
        MapOneDriveUrlToLocalFolder = rootPath & "\" & relPath
    End If

    Exit Function

FailMap:
    MapOneDriveUrlToLocalFolder = ""
End Function




