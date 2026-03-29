Attribute VB_Name = "Lesson08_実践_統計シート自動作成"
Option Explicit

'=====================================================
' レッスン8: 実践 - 統計シート自動作成 - サンプルコード
' このサンプルでは、実務マクロの「作業時間合計」シート作成を
' 簡略化して体験します。
'=====================================================

' --- メイン処理 ---
Public Sub 作業時間シート作成()
    Dim ws As Worksheet

    ' 既存のシートがあれば削除
    Call シート削除("作業時間_練習")

    ' 新しいシートを作成
    Set ws = ThisWorkbook.Sheets.Add( _
                After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "作業時間_練習"

    ' 各処理を順番に呼び出し
    Call タイトル作成(ws)
    Call テストデータ作成(ws)
    Call 集計処理(ws)
    Call 担当者テーブル作成(ws)

    MsgBox "作業時間シートの作成が完了しました。", vbInformation
End Sub


' --- タイトルと見出しの作成 ---
Private Sub タイトル作成(ByVal ws As Worksheet)
    ' タイトル
    ws.Cells(1, 1).Value = "フロア別作業時間計算シート（stay 3分 / out 6分 / 済 1分）"
    With ws.Cells(1, 1)
        .Font.Bold = True
        .Font.Size = 14
    End With

    ' 見出し
    ws.Cells(3, 1).Value = "フロア"
    ws.Cells(3, 2).Value = "stay（部屋数）"
    ws.Cells(3, 3).Value = "out（部屋数）"
    ws.Cells(3, 4).Value = "済（部屋数）"
    ws.Cells(3, 5).Value = "合計時間（分）"

    ' 見出しの書式
    ws.Range("A3:E3").Font.Bold = True
    ws.Range("A3:E3").Interior.Color = RGB(242, 242, 242)
End Sub


' --- テストデータの作成（実務ではCopied_清掃確認書から集計） ---
Private Sub テストデータ作成(ByVal ws As Worksheet)
    Dim floor As Long
    Dim rowOut As Long
    rowOut = 4

    ' ランダムなテストデータを作成
    Randomize

    For floor = 4 To 15
        ws.Cells(rowOut, 1).Value = CStr(floor) & "F"
        ws.Cells(rowOut, 2).Value = Int(Rnd * 8) + 1    ' stay: 1～8
        ws.Cells(rowOut, 3).Value = Int(Rnd * 6) + 1    ' out: 1～6
        ws.Cells(rowOut, 4).Value = Int(Rnd * 3)         ' 済: 0～2
        rowOut = rowOut + 1
    Next floor
End Sub


' --- 集計処理（数式の書き込み） ---
Private Sub 集計処理(ByVal ws As Worksheet)
    Dim rowOut As Long

    ' 各フロアの合計時間を数式で計算
    For rowOut = 4 To 15
        ' E列 = B列*3 + C列*6 + D列*1
        ws.Cells(rowOut, 5).Formula = "=" & _
            ws.Cells(rowOut, 2).Address(False, False) & "*3+" & _
            ws.Cells(rowOut, 3).Address(False, False) & "*6+" & _
            ws.Cells(rowOut, 4).Address(False, False) & "*1"
    Next rowOut

    ' 罫線
    With ws.Range("A3:E15").Borders
        .LineStyle = xlContinuous
    End With

    ' 列幅
    ws.Columns("A:A").ColumnWidth = 10
    ws.Columns("B:E").ColumnWidth = 16
    ws.Range("B:E").HorizontalAlignment = xlCenter
End Sub


' --- 担当者テーブルの作成 ---
Private Sub 担当者テーブル作成(ByVal ws As Worksheet)
    Dim baseRow As Long
    baseRow = 18

    ' 見出し
    ws.Cells(baseRow - 1, 1).Value = "終了目標時間"
    ws.Cells(baseRow - 1, 1).Font.Bold = True

    ws.Cells(baseRow, 1).Value = "担当者"
    ws.Cells(baseRow, 2).Value = "合計(分)"
    ws.Cells(baseRow, 3).Value = "total時間"
    ws.Cells(baseRow, 4).Value = "9:00スタート終了時間"
    ws.Range(ws.Cells(baseRow, 1), ws.Cells(baseRow, 4)).Font.Bold = True
    ws.Range(ws.Cells(baseRow, 1), ws.Cells(baseRow, 4)).Interior.Color = RGB(242, 242, 242)

    ' 担当者データ
    Dim names As Variant
    names = Array("担当者A", "担当者B", "担当者C", "担当者D", "担当者E")

    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim r As Long
        r = baseRow + 1 + i

        ws.Cells(r, 1).Value = names(i)
        ws.Cells(r, 2).Value = Int(Rnd * 120) + 60   ' 60～180分

        ' total時間: 分を「○時間○分」に変換
        ws.Cells(r, 3).FormulaR1C1 = _
            "=INT(RC[-1]/60)&""時間""&MOD(RC[-1],60)&""分"""

        ' 終了時間: 9:00 + 合計分
        ws.Cells(r, 4).FormulaR1C1 = "=TIME(9,0,0)+RC[-2]/1440"
        ws.Cells(r, 4).NumberFormatLocal = "hh:mm"
    Next i

    ' 罫線
    With ws.Range(ws.Cells(baseRow, 1), ws.Cells(baseRow + UBound(names) + 1, 4)).Borders
        .LineStyle = xlContinuous
    End With
End Sub


' --- シート削除（ユーティリティ） ---
Private Sub シート削除(ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0

    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub
