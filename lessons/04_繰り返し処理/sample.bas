Attribute VB_Name = "Lesson04_繰り返し処理"
Option Explicit

'=====================================================
' レッスン4: 繰り返し処理 - サンプルコード
'=====================================================

' --- 例1: For...Next の基本 ---
Sub For基本()
    Dim i As Long
    For i = 1 To 10
        Cells(i, 1).Value = "部屋 " & (400 + i)
        Cells(i, 2).Value = "4F"
    Next i
    MsgBox "10部屋分のデータを作成しました。"
End Sub


' --- 例2: Step 2 で2行ずつ処理（実務マクロ風） ---
Sub Step2ループ()
    ' テストデータ作成（2行1セット）
    Dim i As Long
    For i = 1 To 10
        Cells(i * 2 - 1, 1).Value = 400 + i      ' 奇数行: 部屋番号
        Cells(i * 2 - 1, 2).Value = "C/Iデータ"   ' 奇数行: C/I
        Cells(i * 2, 1).Value = ""                  ' 偶数行: 空
        Cells(i * 2, 2).Value = "備考データ"        ' 偶数行: 備考
    Next i

    ' 2行ずつ処理
    Dim r As Long
    Dim count As Long
    count = 0

    For r = 1 To 20 Step 2
        If Cells(r, 1).Value = "" Then Exit For

        count = count + 1
        ' 1行目（部屋番号行）を太字にする
        Cells(r, 1).Font.Bold = True
    Next r

    MsgBox count & "部屋を処理しました。"
End Sub


' --- 例3: For Each でセル範囲を一括処理 ---
Sub ForEach色付け()
    ' テストデータ
    Range("D1").Value = "ステータス"
    Range("D2").Value = "STAY"
    Range("D3").Value = "出発"
    Range("D4").Value = "ＲＣ"
    Range("D5").Value = "STAY"
    Range("D6").Value = "ダブルセット"
    Range("D7").Value = "出発"

    ' For Each で各セルを処理
    Dim cell As Range
    For Each cell In Range("D2:D7")
        If cell.Value = "STAY" Then
            cell.Interior.Color = RGB(0, 255, 0)         ' 緑
        ElseIf cell.Value = "出発" Or cell.Value = "ＲＣ" Then
            cell.Interior.Color = RGB(255, 255, 0)       ' 黄色
        ElseIf cell.Value = "ダブルセット" Then
            cell.Interior.Color = RGB(255, 105, 180)     ' ピンク
        End If
    Next cell

    MsgBox "色付けが完了しました。"
End Sub


' --- 例4: 最終行の取得 ---
Sub 最終行の取得()
    ' テストデータ
    Dim i As Long
    For i = 1 To 15
        Cells(i, 6).Value = 400 + i
    Next i

    ' 最終行を取得
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 6).End(xlUp).Row

    MsgBox "F列の最終行: " & lastRow & "行目"

    ' 最終行まで色付け
    Dim r As Long
    For r = 1 To lastRow
        Cells(r, 6).Interior.Color = RGB(242, 242, 242)
    Next r
End Sub


' --- 例5: 列のループでヘッダーを検索 ---
Sub ヘッダー検索()
    ' テストヘッダーを作成
    Cells(1, 1).Value = "NO"
    Cells(1, 2).Value = "ROOM"
    Cells(1, 3).Value = "TYPE"
    Cells(1, 4).Value = "C/I"
    Cells(1, 5).Value = "備考"

    ' "ROOM" と "C/I" の列番号を探す
    Dim col As Long
    Dim maxCol As Long
    Dim roomCol As Long, ciCol As Long

    maxCol = Cells(1, Columns.Count).End(xlToLeft).Column
    roomCol = 0
    ciCol = 0

    For col = 1 To maxCol
        If Cells(1, col).Value = "ROOM" Then
            roomCol = col
        ElseIf Cells(1, col).Value = "C/I" Then
            ciCol = col
        End If
    Next col

    MsgBox "ROOM列: " & roomCol & "列目" & vbCrLf & _
           "C/I列: " & ciCol & "列目"
End Sub


' --- 例6: 実務マクロ風 — テーブルの一括色付け ---
Sub テーブル一括色付け()
    ' テストデータ作成
    Cells(4, 1).Value = "ROOM"
    Cells(4, 2).Value = "TYPE"
    Cells(4, 3).Value = "C/I"

    Dim rooms As Variant
    rooms = Array(401, 402, 403, 404, 405, 406, 407, 408)
    Dim ciValues As Variant
    ciValues = Array("SD", "S", "D", "DS", "", "S", "D", "SD")

    Dim i As Long
    For i = 0 To UBound(rooms)
        Cells(5 + i * 2, 1).Value = rooms(i)
        Cells(5 + i * 2, 3).Value = ciValues(i)
    Next i

    ' 色付け処理（実務マクロと同じロジック）
    Dim r As Long
    Dim txt As String
    Dim c1 As String, c2 As String

    For r = 5 To 20 Step 2
        If IsEmpty(Cells(r, 1)) Then Exit For

        txt = CStr(Cells(r, 3).Value)
        If Len(txt) >= 1 Then c1 = Mid$(txt, 1, 1) Else c1 = ""
        If Len(txt) >= 2 Then c2 = Mid$(txt, 2, 1) Else c2 = ""

        With Cells(r, 1)
            If c1 = "S" Or c2 = "S" Then
                .Interior.Color = RGB(0, 176, 240)       ' 青
            ElseIf c1 = "D" Or c2 = "D" Then
                .Interior.Color = RGB(255, 105, 180)     ' ピンク
            End If
        End With
    Next r

    MsgBox "テーブルの色付けが完了しました。"
End Sub
