Attribute VB_Name = "Lesson03_条件分岐と色付け"
Option Explicit

'=====================================================
' レッスン3: 条件分岐と色付け - サンプルコード
'=====================================================

' --- 例1: 条件分岐の基本 ---
Sub 条件分岐の基本()
    Dim status As String
    status = InputBox("ステータスを入力してください（S または D）")

    If status = "S" Then
        MsgBox "ステイ（Stay）です。青色で表示します。"
    ElseIf status = "D" Then
        MsgBox "出発（Departure）です。ピンク色で表示します。"
    Else
        MsgBox "不明なステータスです: " & status
    End If
End Sub


' --- 例2: セルの背景色を変更 ---
Sub 色付けの基本()
    ' 実務で使っている4色を試す
    Range("A1").Value = "S（ステイ）"
    Range("A1").Interior.Color = RGB(0, 176, 240)       ' 青

    Range("A2").Value = "D（出発）"
    Range("A2").Interior.Color = RGB(255, 105, 180)     ' ピンク

    Range("A3").Value = "出発・RC"
    Range("A3").Interior.Color = RGB(255, 255, 0)       ' 黄色

    Range("A4").Value = "STAY"
    Range("A4").Interior.Color = RGB(0, 255, 0)         ' 緑

    MsgBox "4色の色付けが完了しました。"
End Sub


' --- 例3: 条件に応じた色付け（実務マクロ風） ---
Sub 条件で色付け()
    ' テストデータ作成
    Range("A1").Value = "ROOM"
    Range("B1").Value = "C/I"
    Range("A2").Value = 401: Range("B2").Value = "S"
    Range("A3").Value = 402: Range("B3").Value = "D"
    Range("A4").Value = 403: Range("B4").Value = "SD"
    Range("A5").Value = 404: Range("B5").Value = ""
    Range("A6").Value = 405: Range("B6").Value = "DS"

    ' C/I値に応じてROOM列に色付け
    Dim r As Long
    Dim ci As String

    For r = 2 To 6
        ci = CStr(Cells(r, 2).Value)  ' B列のC/I値

        If ci = "S" Then
            Cells(r, 1).Interior.Color = RGB(0, 176, 240)       ' 青
        ElseIf ci = "D" Then
            Cells(r, 1).Interior.Color = RGB(255, 105, 180)     ' ピンク
        End If
    Next r

    MsgBox "色付けが完了しました。"
End Sub


' --- 例4: Or演算子を使った判定（実務マクロと同じロジック） ---
Sub Or演算子で色付け()
    ' テストデータ
    Range("A1").Value = "ROOM"
    Range("B1").Value = "C/I"
    Range("A2").Value = 401: Range("B2").Value = "SD"
    Range("A3").Value = 402: Range("B3").Value = "DS"
    Range("A4").Value = 403: Range("B4").Value = "S"
    Range("A5").Value = 404: Range("B5").Value = "D"

    Dim r As Long
    Dim txt As String
    Dim c1 As String, c2 As String

    For r = 2 To 5
        txt = CStr(Cells(r, 2).Value)

        ' 1文字目と2文字目を取り出す
        If Len(txt) >= 1 Then c1 = Mid$(txt, 1, 1) Else c1 = ""
        If Len(txt) >= 2 Then c2 = Mid$(txt, 2, 1) Else c2 = ""

        ' S が含まれていれば青、D が含まれていればピンク
        With Cells(r, 1)
            If c1 = "S" Or c2 = "S" Then
                .Interior.Color = RGB(0, 176, 240)       ' 青
            ElseIf c1 = "D" Or c2 = "D" Then
                .Interior.Color = RGB(255, 105, 180)     ' ピンク
            End If
        End With
    Next r

    MsgBox "色付けが完了しました（Or演算子版）。"
End Sub


' --- 例5: 色をリセット ---
Sub 色リセット()
    Range("A1:B10").Interior.ColorIndex = xlNone
    MsgBox "色をリセットしました。"
End Sub
