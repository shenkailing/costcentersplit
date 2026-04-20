Sub costcentersplit_Click()
Dim wsSrc As Worksheet, wsRes As Worksheet
    Dim lastRow As Long, lastCol As Long, resRow As Long
    Dim i As Long, j As Long
    Dim brand As String, storeName As String
    Dim startCol As Long
    Dim groupRowsKilian As Collection, groupRowsFM As Collection
    Dim rowDict As Object
    Dim subtotalRow As Long

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
    For j = 5 To lastCol
        If j < 19 Then
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
    MsgBox "分组数据已追加到新表 'CostCenterGroupBySumResult' 中！", vbInformation
End Sub
