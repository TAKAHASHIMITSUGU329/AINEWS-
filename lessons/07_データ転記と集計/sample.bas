Attribute VB_Name = "Lesson07_データ転記と集計"
Option Explicit

'=====================================================
' レッスン7: データ転記と集計 - サンプルコード
'=====================================================

' --- 例1: シート間の転記（基本） ---
Sub 転記の基本()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' まず元データを作成
    ws.Cells(1, 1).Value = "ROOM"
    ws.Cells(1, 2).Value = "C/I"
    ws.Cells(2, 1).Value = 401: ws.Cells(2, 2).Value = "S"
    ws.Cells(3, 1).Value = 402: ws.Cells(3, 2).Value = "D"
    ws.Cells(4, 1).Value = 403: ws.Cells(4, 2).Value = "S"
    ws.Cells(5, 1).Value = 404: ws.Cells(5, 2).Value = "D"
    ws.Cells(6, 1).Value = 405: ws.Cells(6, 2).Value = "S"

    ' D列以降に転記先を作成
    ws.Cells(1, 4).Value = "転記先ROOM"
    ws.Cells(1, 5).Value = "色"

    Dim r As Long, destRow As Long
    destRow = 2

    For r = 2 To 6
        ws.Cells(destRow, 4).Value = ws.Cells(r, 1).Value
        ws.Cells(destRow, 5).Value = ws.Cells(r, 2).Value

        ' 色付け
        If ws.Cells(r, 2).Value = "S" Then
            ws.Cells(destRow, 4).Interior.Color = RGB(0, 176, 240)
        ElseIf ws.Cells(r, 2).Value = "D" Then
            ws.Cells(destRow, 4).Interior.Color = RGB(255, 105, 180)
        End If

        destRow = destRow + 1
    Next r

    MsgBox "転記が完了しました。"
End Sub


' --- 例2: Select Case でフロア別の転記位置を決定 ---
Sub フロア別転記位置()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ws.Cells(1, 1).Value = "フロア"
    ws.Cells(1, 2).Value = "転記先列"
    ws.Cells(1, 3).Value = "開始行"
    ws.Cells(1, 4).Value = "終了行"

    Dim floor As Long
    Dim r As Long
    r = 2

    For floor = 4 To 15
        Dim roomColDest As Long
        Dim firstRowDest As Long, lastRowDest As Long

        Select Case floor
            Case 4:  roomColDest = 1:  firstRowDest = 3:  lastRowDest = 19
            Case 5:  roomColDest = 1:  firstRowDest = 32: lastRowDest = 48
            Case 6:  roomColDest = 8:  firstRowDest = 3:  lastRowDest = 19
            Case 7:  roomColDest = 8:  firstRowDest = 32: lastRowDest = 48
            Case 8:  roomColDest = 15: firstRowDest = 3:  lastRowDest = 19
            Case 9:  roomColDest = 15: firstRowDest = 32: lastRowDest = 48
            Case 10: roomColDest = 22: firstRowDest = 3:  lastRowDest = 19
            Case 11: roomColDest = 22: firstRowDest = 32: lastRowDest = 48
            Case 12: roomColDest = 29: firstRowDest = 3:  lastRowDest = 19
            Case 13: roomColDest = 29: firstRowDest = 32: lastRowDest = 49
            Case 14: roomColDest = 36: firstRowDest = 3:  lastRowDest = 19
            Case 15: roomColDest = 36: firstRowDest = 32: lastRowDest = 48
        End Select

        ws.Cells(r, 1).Value = floor & "F"
        ws.Cells(r, 2).Value = roomColDest & "列目"
        ws.Cells(r, 3).Value = firstRowDest & "行"
        ws.Cells(r, 4).Value = lastRowDest & "行"
        r = r + 1
    Next floor

    MsgBox "フロア別の転記位置一覧を作成しました。"
End Sub


' --- 例3: Application.Match で部屋番号を検索 ---
Sub Match検索()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' 指示書側のデータ（転記先）
    ws.Cells(1, 6).Value = "指示書ROOM"
    Dim rooms As Variant
    rooms = Array(401, 402, 403, 404, 405, 406, 407, 408)
    Dim i As Long
    For i = 0 To UBound(rooms)
        ws.Cells(2 + i, 6).Value = rooms(i)
    Next i

    ' 部屋番号 403 を検索
    Dim searchRoom As Long
    searchRoom = 403

    Dim m As Variant
    m = Application.Match(searchRoom, ws.Range("F2:F9"), 0)

    If IsError(m) Then
        MsgBox "部屋 " & searchRoom & " が見つかりませんでした。"
    Else
        Dim destRow As Long
        destRow = 2 + CLng(m) - 1

        ws.Cells(destRow, 7).Value = "← ここに転記"
        ws.Cells(destRow, 6).Interior.Color = RGB(0, 176, 240)

        MsgBox "部屋 " & searchRoom & " は " & destRow & " 行目（" & m & "番目）"
    End If
End Sub


' --- 例4: カウント集計（D/S集計） ---
Sub DS集計()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' テストデータ作成
    ws.Range("A1:E20").Clear
    ws.Cells(1, 1).Value = "ROOM"
    ws.Cells(1, 2).Value = "C/I"

    Dim ciData As Variant
    ciData = Array("SD", "S", "D", "DS", "S", "D", "S", "D", "SD", "S")
    Dim r As Long
    For r = 0 To UBound(ciData)
        ws.Cells(2 + r, 1).Value = 401 + r
        ws.Cells(2 + r, 2).Value = ciData(r)
    Next r

    ' D/S 集計
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim dCount As Long, sCount As Long
    dCount = 0: sCount = 0

    Dim txt As String
    Dim c1 As String, c2 As String

    For r = 2 To lastRow
        txt = CStr(ws.Cells(r, 2).Value)
        If Len(txt) >= 1 Then c1 = Mid$(txt, 1, 1) Else c1 = ""
        If Len(txt) >= 2 Then c2 = Mid$(txt, 2, 1) Else c2 = ""

        If c1 = "D" Or c2 = "D" Then dCount = dCount + 1
        If c1 = "S" Or c2 = "S" Then sCount = sCount + 1
    Next r

    ' 結果を印字
    ws.Cells(lastRow + 2, 1).Value = "集計結果:"
    ws.Cells(lastRow + 2, 2).Value = "D=" & dCount & ", S=" & sCount
    ws.Cells(lastRow + 2, 2).Font.Bold = True

    MsgBox "D=" & dCount & ", S=" & sCount
End Sub


' --- 例5: 罫線付きの集計表を作成 ---
Sub 集計表作成()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' ヘッダー
    Dim headers As Variant
    headers = Array("フロア", "S", "D", "合計")

    Dim c As Long
    For c = 0 To UBound(headers)
        ws.Cells(1, 1 + c).Value = headers(c)
        ws.Cells(1, 1 + c).Font.Bold = True
    Next c

    ' サンプルデータ
    Dim floor As Long
    For floor = 4 To 15
        Dim row As Long
        row = floor - 2

        ws.Cells(row, 1).Value = floor & "F"
        ws.Cells(row, 2).Value = Int(Rnd * 10)  ' ランダムなS数
        ws.Cells(row, 3).Value = Int(Rnd * 10)  ' ランダムなD数
        ws.Cells(row, 4).Formula = "=B" & row & "+C" & row  ' 合計
    Next floor

    ' 罫線
    Dim tblRange As Range
    Set tblRange = ws.Range("A1:D13")

    tblRange.Borders(xlEdgeTop).LineStyle = xlContinuous
    tblRange.Borders(xlEdgeBottom).LineStyle = xlContinuous
    tblRange.Borders(xlEdgeLeft).LineStyle = xlContinuous
    tblRange.Borders(xlEdgeRight).LineStyle = xlContinuous
    tblRange.Borders(xlInsideVertical).LineStyle = xlContinuous
    tblRange.Borders(xlInsideHorizontal).LineStyle = xlContinuous

    MsgBox "集計表を作成しました。"
End Sub
