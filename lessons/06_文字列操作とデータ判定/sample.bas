Attribute VB_Name = "Lesson06_文字列操作とデータ判定"
Option Explicit

'=====================================================
' レッスン6: 文字列操作とデータ判定 - サンプルコード
'=====================================================

' --- 例1: Mid$ で1文字ずつ取り出す ---
Sub Mid練習()
    ' テストデータ
    Range("A1").Value = "C/I値"
    Range("B1").Value = "1文字目"
    Range("C1").Value = "2文字目"
    Range("D1").Value = "判定"

    Dim testData As Variant
    testData = Array("SD", "S", "D", "DS", "", "S")

    Dim i As Long
    Dim txt As String
    Dim c1 As String, c2 As String

    For i = 0 To UBound(testData)
        txt = testData(i)
        Range("A" & (i + 2)).Value = txt

        If Len(txt) >= 1 Then c1 = Mid$(txt, 1, 1) Else c1 = ""
        If Len(txt) >= 2 Then c2 = Mid$(txt, 2, 1) Else c2 = ""

        Range("B" & (i + 2)).Value = c1
        Range("C" & (i + 2)).Value = c2

        ' S/D 判定（実務マクロと同じロジック）
        If c1 = "S" Or c2 = "S" Then
            Range("D" & (i + 2)).Value = "S（ステイ）"
        ElseIf c1 = "D" Or c2 = "D" Then
            Range("D" & (i + 2)).Value = "D（出発）"
        Else
            Range("D" & (i + 2)).Value = "-"
        End If
    Next i

    MsgBox "Mid$ による判定が完了しました。"
End Sub


' --- 例2: Left$ でフロア番号を判定 ---
Sub フロア判定()
    Range("A1").Value = "部屋番号"
    Range("B1").Value = "フロア"

    Dim rooms As Variant
    rooms = Array("401", "502", "603", "1201", "1513", "1615")

    Dim i As Long
    Dim roomNo As String
    Dim floorNo As Long

    For i = 0 To UBound(rooms)
        roomNo = rooms(i)
        Range("A" & (i + 2)).Value = roomNo

        ' フロアを判定（先頭1〜2文字）
        If Len(roomNo) = 4 Then
            floorNo = CLng(Left$(roomNo, 2))  ' 4桁 → 先頭2文字
        ElseIf Len(roomNo) = 3 Then
            floorNo = CLng(Left$(roomNo, 1))  ' 3桁 → 先頭1文字
        End If

        Range("B" & (i + 2)).Value = floorNo & "F"
    Next i

    MsgBox "フロア判定が完了しました。"
End Sub


' --- 例3: InStr で文字列を検索 ---
Sub InStr検索()
    Range("A1").Value = "備考"
    Range("B1").Value = "種別"

    Dim remarks As Variant
    remarks = Array("ダブル解除", "ツインセット", "14:30", "ノータッチ", "トリプル解除", "通常清掃")

    Dim i As Long
    Dim s As String

    For i = 0 To UBound(remarks)
        s = remarks(i)
        Range("A" & (i + 2)).Value = s

        If InStr(s, "ダブル解除") > 0 Then
            Range("B" & (i + 2)).Value = "D解除"
        ElseIf InStr(s, "ツインセット") > 0 Then
            Range("B" & (i + 2)).Value = "Tメイク"
        ElseIf InStr(s, "トリプル解除") > 0 Then
            Range("B" & (i + 2)).Value = "3解除"
        ElseIf InStr(s, ":") > 0 Then
            Range("B" & (i + 2)).Value = "時刻"
        ElseIf s = "ノータッチ" Then
            Range("B" & (i + 2)).Value = "除外"
        Else
            Range("B" & (i + 2)).Value = "通常"
        End If
    Next i

    MsgBox "InStr による分類が完了しました。"
End Sub


' --- 例4: CStr で安全に文字列に変換 ---
Sub CStr練習()
    ' 数値、日付、文字列など様々な型のデータ
    Cells(1, 1).Value = 401           ' 数値
    Cells(2, 1).Value = "STAY"        ' 文字列
    Cells(3, 1).Value = True          ' Boolean
    Cells(4, 1).Value = ""            ' 空

    Dim r As Long
    Dim val As String
    For r = 1 To 4
        val = CStr(Cells(r, 1).Value)
        Cells(r, 2).Value = "CStr結果: [" & val & "]"
        Cells(r, 3).Value = "長さ: " & Len(val)
    Next r

    MsgBox "CStr の変換結果を確認してください。"
End Sub


' --- 例5: 実務マクロ風 — C/I値の総合判定 ---
Sub CI値の総合判定()
    ' テストデータ作成
    Cells(1, 1).Value = "ROOM"
    Cells(1, 2).Value = "C/I(1行目)"
    Cells(1, 3).Value = "C/I(2行目)"
    Cells(1, 4).Value = "色"
    Cells(1, 5).Value = "Wset"
    Cells(1, 6).Value = "備考"

    ' 2行1セットのテストデータ
    Cells(2, 1).Value = 401: Cells(2, 2).Value = "SD":   Cells(3, 2).Value = "14:30"
    Cells(4, 1).Value = 402: Cells(4, 2).Value = "S":    Cells(5, 2).Value = ""
    Cells(6, 1).Value = 403: Cells(6, 2).Value = "D":    Cells(7, 2).Value = "15：00"

    ' 判定処理
    Dim r As Long
    Dim ci1 As String, ci2text As String
    Dim c1 As String, c2 As String

    For r = 2 To 6 Step 2
        If Cells(r, 1).Value = "" Then Exit For

        ci1 = CStr(Cells(r, 2).Value)
        ci2text = CStr(Cells(r + 1, 2).Value)

        ' 1文字目・2文字目の取り出し
        If Len(ci1) >= 1 Then c1 = Mid$(ci1, 1, 1) Else c1 = ""
        If Len(ci1) >= 2 Then c2 = Mid$(ci1, 2, 1) Else c2 = ""

        ' 色判定
        If c1 = "S" Or c2 = "S" Then
            Cells(r, 4).Value = "青（S）"
        ElseIf c1 = "D" Or c2 = "D" Then
            Cells(r, 4).Value = "ピンク（D）"
        End If

        ' Wset判定
        If Left$(LTrim$(ci1), 1) = ChrW(&H25CF) Then
            Cells(r, 5).Value = "W"
        End If

        ' 備考（時刻チェック）
        If InStr(ci2text, ":") > 0 Or InStr(ci2text, ChrW(&HFF1A)) > 0 Then
            Cells(r, 6).Value = ci2text
        End If
    Next r

    MsgBox "C/I値の総合判定が完了しました。"
End Sub
