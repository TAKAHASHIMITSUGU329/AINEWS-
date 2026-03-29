Attribute VB_Name = "Module_統計集計"
Option Explicit

Sub SelectAndEditFile()

    Dim FileDialog As FileDialog
    Dim strFilePath As Variant
    Dim openedWorkbook As Workbook
    Dim srcWs As Worksheet
    Dim destWs As Worksheet
    Dim cell As Range
    Dim targetRange As Range
    Dim fileCount As Integer
    Dim varPath As Variant

    ' 不要なシートを削除
    DeleteSheets

    ' ファイルダイアログオブジェクトを初期化
    Set FileDialog = Application.FileDialog(msoFileDialogOpen)
    FileDialog.AllowMultiSelect = True
    FileDialog.Title = "ファイルを選択してください"

    If FileDialog.Show = -1 Then
        fileCount = 0

        strFilePath = FileDialog.SelectedItems(1)
        ThisWorkbook.Sheets("Menu").Range("basePath").Value = strFilePath

        For Each varPath In FileDialog.SelectedItems
            fileCount = fileCount + 1
            strFilePath = varPath
            Set openedWorkbook = Workbooks.Open(Filename:=strFilePath)

            If fileCount = 1 Then
                Set srcWs = openedWorkbook.Sheets("清掃確認書")
                srcWs.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                Set destWs = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                destWs.Name = "Copied_清掃確認書"

                Set targetRange = destWs.Range("A1:Z1000")

                For Each cell In targetRange
                    If cell.Column = 7 And cell.Row > 1 And cell.Value <> "" Then
                        cell.Interior.Color = RGB(255, 105, 180)
                    ElseIf cell.Value = "出発" Or cell.Value = "ＲＣ" Then
                        cell.Interior.Color = RGB(255, 255, 0)
                    ElseIf cell.Value = "ダブル解除" Or cell.Value = "ダブルセット" Or _
                           cell.Value = "トリプルセット" Or cell.Value = "トリプル解除" Or _
                           cell.Value = "ツインセット" Or cell.Value = "ツイン解除" Then
                        cell.Interior.Color = RGB(255, 105, 180)
                    ElseIf cell.Value = "STAY" Then
                        cell.Interior.Color = RGB(0, 255, 0)
                    End If
                Next cell

            ElseIf fileCount = 2 Then
                Set srcWs = openedWorkbook.Sheets("Sheet2")
                srcWs.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                Set destWs = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
                destWs.Name = "Copied_清掃指示書"
            End If

            openedWorkbook.Close SaveChanges:=False
        Next varPath

        CreateSummary
        CreateWorkTimeSheet
    End If

    Set FileDialog = Nothing

End Sub

Sub DeleteSheets()

    Dim sheetName As String

    Application.DisplayAlerts = False

    sheetName = "Copied_清掃確認書"
    If SheetExists(sheetName) Then ThisWorkbook.Sheets(sheetName).Delete

    sheetName = "Copied_指示集計表"
    If SheetExists(sheetName) Then ThisWorkbook.Sheets(sheetName).Delete

    sheetName = "指示書一覧"
    If SheetExists(sheetName) Then ThisWorkbook.Sheets(sheetName).Delete

    sheetName = "作業時間合計"
    If SheetExists(sheetName) Then ThisWorkbook.Sheets(sheetName).Delete

    Application.DisplayAlerts = True

End Sub

Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then SheetExists = True
End Function

Sub CreateSummary()

    Dim destWs As Worksheet, summaryWs As Worksheet
    Dim i As Integer, j As Integer, lastRow As Long, k As Long
    Dim floorPrefix As String
    Dim totalS As Long, totalT As Long, totalD As Long, total3 As Long
    Dim totalDD As Long, totalTD As Long, total3D As Long
    Dim totalTS As Long, totalUR As Long, totalFA As Long
    Dim totalZ As Long
    Dim cell As Range
    Dim c As Long

    Sheets("指示集計表").Copy After:=Sheets(Sheets.Count)
    ActiveSheet.Name = "Copied_指示集計表"

    ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).Name = "指示書一覧"
    Set summaryWs = ThisWorkbook.Sheets("指示書一覧")

    summaryWs.Cells(1, 3).Value = "ハートンホテル東品川清掃指示集計表"
    summaryWs.Cells(1, 3).Font.Size = 20

    summaryWs.Cells(2, 11).Value = "フロント"
    summaryWs.Cells(2, 12).Value = "プライム"
    summaryWs.Range("K2:L2").Borders.LineStyle = True
    summaryWs.Range("K3:K5").BorderAround True
    summaryWs.Range("L3:L5").BorderAround True

    Dim dayOfWeek As String
    dayOfWeek = "(" & Format(Now, "aaa") & ")"
    summaryWs.Cells(6, 10).Value = "日付: " & Format(Now, "yyyy年mm月dd日 ") & dayOfWeek

    Dim headers As Variant
    headers = Array("フロア", "S", "T", "UR", "FA", "D解除", "Dセット", "T解除", "Tメイク", "3解除", "3メイク", "済")
    For i = 1 To UBound(headers) + 1
        summaryWs.Cells(7, i).Value = headers(i - 1)
    Next i

    Set destWs = ThisWorkbook.Sheets("Copied_清掃確認書")
    lastRow = destWs.Cells(destWs.Rows.Count, 2).End(xlUp).Row

    j = 8
    For i = 19 To 3 Step -1
        floorPrefix = CStr(i)
        totalS = 0: totalT = 0: totalD = 0: total3 = 0
        totalDD = 0: totalTD = 0: total3D = 0
        totalTS = 0: totalUR = 0: totalFA = 0
        totalZ = 0

        For k = 2 To lastRow
            If Left(destWs.Cells(k, 2).Value, Len(floorPrefix)) = floorPrefix Then
                If destWs.Cells(k, 4).Value = "済" Then
                    totalZ = totalZ + 1
                End If

                If destWs.Cells(k, 4).Value <> "済" And _
                   (InStr(1, destWs.Cells(k, 7).Value, "未掃") = 0 And _
                    Not (destWs.Cells(k, 7).Value = "ノータッチ" Or destWs.Cells(k, 7).Value = "ﾉｰﾀｯﾁ") And _
                    Not (destWs.Cells(k, 7).Value = "入室禁止")) And _
                   (InStr(1, destWs.Cells(k + 1, 7).Value, "未掃") = 0 And _
                    Not (destWs.Cells(k + 1, 7).Value = "ノータッチ" Or destWs.Cells(k + 1, 7).Value = "ﾉｰﾀｯﾁ") And _
                    Not (destWs.Cells(k + 1, 7).Value = "入室禁止")) Then

                    If destWs.Cells(k, 3).Value = "S" Or _
                       (destWs.Cells(k, 3).Value = "T" And InStr(1, destWs.Cells(k, 7).Value, "ツインセット") > 0) Then
                        totalS = totalS + 1
                    ElseIf destWs.Cells(k, 3).Value = "T" Then
                        If destWs.Cells(k, 2).Value = "1615" Then
                            totalUR = totalUR + 1
                        Else
                            totalT = totalT + 1
                        End If
                    ElseIf destWs.Cells(k, 3).Value = "D" Then
                        If destWs.Cells(k, 2).Value = "1513" Then
                            totalFA = totalFA + 1
                        Else
                            totalD = totalD + 1
                        End If
                    ElseIf destWs.Cells(k, 3).Value = "3" Then
                        total3 = total3 + 1
                    End If

                    If InStr(1, destWs.Cells(k, 7).Value, "ダブル解除") > 0 Or _
                       InStr(1, destWs.Cells(k + 1, 7).Value, "ダブル解除") > 0 Then
                        totalDD = totalDD + 1
                    ElseIf InStr(1, destWs.Cells(k, 7).Value, "ツイン解除") > 0 Or _
                           InStr(1, destWs.Cells(k + 1, 7).Value, "ツイン解除") > 0 Then
                        totalTD = totalTD + 1
                    ElseIf InStr(1, destWs.Cells(k, 7).Value, "トリプル解除") > 0 Or _
                           InStr(1, destWs.Cells(k + 1, 7).Value, "トリプル解除") > 0 Then
                        total3D = total3D + 1
                    ElseIf InStr(1, destWs.Cells(k, 7).Value, "ツインセット") > 0 Or _
                           InStr(1, destWs.Cells(k + 1, 7).Value, "ツインセット") > 0 Then
                        totalTS = totalTS + 1
                    End If
                End If
            End If
        Next k

        summaryWs.Cells(j, 1).Value = floorPrefix & "F"
        If totalS + totalD > 0 Then summaryWs.Cells(j, 2).Value = totalS + totalD
        If totalT + total3 > 0 Then summaryWs.Cells(j, 3).Value = totalT + total3
        If totalD > 0 Then summaryWs.Cells(j, 7).Value = totalD
        If totalTD > 0 Then summaryWs.Cells(j, 8).Value = totalTD
        If total3 > 0 Then summaryWs.Cells(j, 11).Value = total3
        If total3D > 0 Then summaryWs.Cells(j, 10).Value = total3D
        If totalTS > 0 Then summaryWs.Cells(j, 9).Value = totalTS
        If totalFA > 0 Then summaryWs.Cells(j, 5).Value = totalFA
        If totalUR > 0 Then summaryWs.Cells(j, 4).Value = totalUR
        If totalZ > 0 Then summaryWs.Cells(j, 12).Value = totalZ

        j = j + 1
    Next i

    summaryWs.Range("A7:L" & j - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
    summaryWs.Range("A7:L" & j - 1).Borders(xlEdgeTop).LineStyle = xlContinuous
    summaryWs.Range("A7:L" & j - 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
    summaryWs.Range("A7:L" & j - 1).Borders(xlEdgeRight).LineStyle = xlContinuous
    summaryWs.Range("A7:L" & j - 1).Borders(xlInsideVertical).LineStyle = xlContinuous
    summaryWs.Range("A7:L" & j - 1).Borders(xlInsideHorizontal).LineStyle = xlContinuous

    For Each cell In summaryWs.Range("D8:F24")
        If Not (cell.Address = "$D$11" Or cell.Address = "$E$12") Then
            cell.Borders(xlDiagonalUp).LineStyle = xlContinuous
        End If
    Next cell
    For Each cell In summaryWs.Range("H8:I21")
        cell.Borders(xlDiagonalUp).LineStyle = xlContinuous
    Next cell

    summaryWs.Cells(25, 1).Value = "小計"
    For c = 2 To 12
        summaryWs.Cells(25, c).Formula = "=SUM(" & Cells(8, c).Address(False, False) & ":" & Cells(24, c).Address(False, False) & ")"
    Next c

    summaryWs.Cells(26, 1).Value = "未掃分"
    summaryWs.Cells(27, 1).Value = "追加分"
    summaryWs.Cells(28, 1).Value = "合計"

    summaryWs.Range("A7:L28").Borders(xlEdgeBottom).LineStyle = xlContinuous
    summaryWs.Range("A7:L28").Borders(xlEdgeTop).LineStyle = xlContinuous
    summaryWs.Range("A7:L28").Borders(xlEdgeLeft).LineStyle = xlContinuous
    summaryWs.Range("A7:L28").Borders(xlEdgeRight).LineStyle = xlContinuous
    summaryWs.Range("A7:L28").Borders(xlInsideVertical).LineStyle = xlContinuous
    summaryWs.Range("A7:L28").Borders(xlInsideHorizontal).LineStyle = xlContinuous

    Set destWs = ThisWorkbook.Sheets("Copied_指示集計表")
    Dim destRow As Integer
    Dim summaryRow As Integer
    summaryRow = 0
    For destRow = 6 To 54 Step 3
        destWs.Cells(destRow, 2).Value = summaryWs.Cells(8 + summaryRow, 2).Value
        destWs.Cells(destRow, 4).Value = summaryWs.Cells(8 + summaryRow, 3).Value
        destWs.Cells(destRow, 7).Value = destWs.Cells(destRow, 2).Value + destWs.Cells(destRow, 4).Value
        summaryRow = summaryRow + 1
    Next destRow
    destWs.Cells(16, 6).Value = summaryWs.Cells(11, 4).Value
    destWs.Cells(19, 6).Value = summaryWs.Cells(12, 5).Value
End Sub


Public Sub CreateWorkTimeSheet()
    Const SRC_SHEET As String = "Copied_清掃確認書"
    Const OUT_SHEET As String = "作業時間合計"

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim src As Worksheet, ws As Worksheet
    Dim lastRow As Long
    Dim floor As Long, r As Long, rowOut As Long
    Dim roomNo As String
    Dim stayCnt As Long, outCnt As Long, doneCnt As Long
    Dim valE As String, valF As String, valD As String

    If Not SheetExistsVB(wb, SRC_SHEET) Then
        MsgBox "集計元シート「" & SRC_SHEET & "」が見つかりません。", vbExclamation
        Exit Sub
    End If
    Set src = wb.Sheets(SRC_SHEET)
    lastRow = src.Cells(src.Rows.Count, 2).End(xlUp).Row

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets(OUT_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = OUT_SHEET

    ws.Cells(1, 1).Value = "フロア別作業時間計算シート（stay 3分 / out 6分 / 済 1分）"
    With ws.Range("A1")
        .Font.Bold = True
        .Font.Size = 14
    End With

    ws.Cells(3, 1).Value = "フロア"
    ws.Cells(3, 2).Value = "stay（部屋数）"
    ws.Cells(3, 3).Value = "out（部屋数）"
    ws.Cells(3, 4).Value = "済（部屋数）"
    ws.Cells(3, 5).Value = "合計時間（分）"
    ws.Range("A3:E3").Font.Bold = True
    ws.Range("A3:E3").Interior.Color = RGB(242, 242, 242)

    rowOut = 4

    For floor = 3 To 19
        stayCnt = 0: outCnt = 0: doneCnt = 0

        For r = 2 To lastRow
            roomNo = CStr(src.Cells(r, 2).Value)
            If Len(roomNo) > 0 Then
                If Left$(roomNo, Len(CStr(floor))) = CStr(floor) Then
                    valE = CStr(src.Cells(r, 5).Value)
                    valF = CStr(src.Cells(r, 6).Value)
                    valD = CStr(src.Cells(r, 4).Value)

                    If StrComp(valE, "STAY", vbTextCompare) = 0 Then
                        stayCnt = stayCnt + 1
                    End If

                    If (InStr(1, valF, "出発", vbTextCompare) > 0) Or _
                       (InStr(1, valF, "ＲＣ", vbTextCompare) > 0) Or _
                       (InStr(1, valF, "延長", vbTextCompare) > 0) Then
                        outCnt = outCnt + 1
                    End If

                    If StrComp(valD, "済", vbTextCompare) = 0 Then
                        doneCnt = doneCnt + 1
                    End If
                End If
            End If
        Next r

        ws.Cells(rowOut, 1).Value = CStr(floor) & "F"
        ws.Cells(rowOut, 2).Value = stayCnt
        ws.Cells(rowOut, 3).Value = outCnt
        ws.Cells(rowOut, 4).Value = doneCnt
        ws.Cells(rowOut, 5).Formula = "=" & _
            ws.Cells(rowOut, 2).Address(False, False) & "*3+" & _
            ws.Cells(rowOut, 3).Address(False, False) & "*6+" & _
            ws.Cells(rowOut, 4).Address(False, False) & "*1"

        rowOut = rowOut + 1
    Next floor

    With ws.Range("A3:E" & rowOut - 1).Borders
        .LineStyle = xlContinuous
    End With

    ws.Cells(rowOut + 2, 1).Value = "終了目標時間"
    ws.Cells(rowOut + 3, 1).Value = "担当者"
    ws.Cells(rowOut + 3, 2).Value = "合計(分)"
    ws.Cells(rowOut + 3, 3).Value = "total時間"
    ws.Cells(rowOut + 3, 4).Value = "9:00スタート終了時間"
    ws.Cells(rowOut + 3, 5).Value = "目標時間"
    ws.Cells(rowOut + 3, 6).Value = "終了時間"
    ws.Range(ws.Cells(rowOut + 3, 1), ws.Cells(rowOut + 3, 6)).Font.Bold = True
    ws.Range(ws.Cells(rowOut + 3, 1), ws.Cells(rowOut + 3, 6)).Interior.Color = RGB(242, 242, 242)

    Dim names As Variant, i As Long, baseRow As Long
    names = Array("イラノ", "アルマ", "エマ", "シャムナ", "ゴマ")
    baseRow = rowOut + 4

    For i = LBound(names) To UBound(names)
        ws.Cells(baseRow + i, 1).Value = names(i)
        ws.Cells(baseRow + i, 2).Value = 0
        ws.Cells(baseRow + i, 3).FormulaR1C1 = "=INT(RC[-1]/60)&""時間""&MOD(RC[-1],60)&""分"""
        ws.Cells(baseRow + i, 4).FormulaR1C1 = "=TIME(9,0,0)+RC[-2]/1440"
        ws.Cells(baseRow + i, 4).NumberFormatLocal = "hh:mm"
        ws.Cells(baseRow + i, 5).ClearContents
        ws.Cells(baseRow + i, 6).ClearContents
    Next i

    ws.Columns("A:A").ColumnWidth = 10
    ws.Columns("B:E").ColumnWidth = 16
    ws.Columns("F:F").ColumnWidth = 16
    ws.Range("A1").EntireColumn.HorizontalAlignment = xlLeft
    ws.Range("B:E").HorizontalAlignment = xlCenter

    With ws.Range(ws.Cells(rowOut + 3, 1), ws.Cells(baseRow + UBound(names), 6)).Borders
        .LineStyle = xlContinuous
    End With
End Sub

Private Function SheetExistsVB(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    On Error Resume Next
    SheetExistsVB = Not wb.Sheets(sheetName) Is Nothing
    On Error GoTo 0
End Function
