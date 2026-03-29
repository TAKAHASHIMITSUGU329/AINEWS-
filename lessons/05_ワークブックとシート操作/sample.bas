Attribute VB_Name = "Lesson05_ワークブックとシート操作"
Option Explicit

'=====================================================
' レッスン5: ワークブックとシート操作 - サンプルコード
'=====================================================

' --- 例1: ファイル選択ダイアログ ---
Sub ファイル選択()
    Dim filePath As Variant
    filePath = Application.GetOpenFilename( _
                    Title:="ファイルを選択してください")

    If filePath = False Then
        MsgBox "キャンセルされました。"
        Exit Sub
    End If

    MsgBox "選択されたファイル: " & CStr(filePath)

    ' 選択したパスをセルに記録
    ThisWorkbook.Worksheets(1).Range("A1").Value = CStr(filePath)
End Sub


' --- 例2: ブックを開いて情報を取得 ---
Sub ブック情報取得()
    Dim filePath As Variant
    filePath = Application.GetOpenFilename( _
                    Title:="Excelファイルを選択してください")
    If filePath = False Then Exit Sub

    Dim wb As Workbook
    Set wb = Workbooks.Open(CStr(filePath))

    ' ブックの情報を表示
    Dim info As String
    info = "ブック名: " & wb.Name & vbCrLf
    info = info & "シート数: " & wb.Sheets.Count & vbCrLf

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        info = info & "  - " & ws.Name & vbCrLf
    Next ws

    MsgBox info, vbInformation, "ブック情報"

    ' 閉じる
    wb.Close SaveChanges:=False
End Sub


' --- 例3: シートの存在チェック ---
Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then SheetExists = True
End Function


' --- 例4: シートの追加・削除（安全版） ---
Sub シート管理()
    Dim sheetName As String
    sheetName = "テスト用シート"

    ' 既存のシートがあれば削除
    If SheetExists(sheetName) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If

    ' 新しいシートを追加
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add( _
                After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = sheetName

    ' データを書き込み
    ws.Range("A1").Value = "フロア"
    ws.Range("B1").Value = "部屋数"
    ws.Range("A2").Value = "4F"
    ws.Range("B2").Value = 17

    MsgBox "「" & sheetName & "」シートを作成しました。"
End Sub


' --- 例5: シートをコピーして名前変更 ---
Sub シートコピー()
    Dim srcName As String
    srcName = "Sheet1"

    If Not SheetExists(srcName) Then
        MsgBox srcName & " が見つかりません。"
        Exit Sub
    End If

    ' コピー先の名前
    Dim destName As String
    destName = "Copied_" & srcName

    ' 既存のコピーがあれば削除
    If SheetExists(destName) Then
        Application.DisplayAlerts = False
        ThisWorkbook.Sheets(destName).Delete
        Application.DisplayAlerts = True
    End If

    ' コピー実行
    ThisWorkbook.Sheets(srcName).Copy _
        After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count).Name = destName

    MsgBox srcName & " を " & destName & " にコピーしました。"
End Sub


' --- 例6: エラーハンドリングの実演 ---
Sub エラーハンドリング()
    Dim ws As Worksheet

    ' 存在しないシートを参照してみる
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("存在しないシート名")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "シートが見つかりませんでした。" & vbCrLf & _
               "On Error Resume Next でエラーを回避しました。", _
               vbExclamation
    Else
        MsgBox "シートが見つかりました: " & ws.Name
    End If
End Sub
