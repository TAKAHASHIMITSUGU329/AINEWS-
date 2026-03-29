# レッスン5: ワークブックとシート操作

## このレッスンで学ぶこと
- ファイル選択ダイアログ（GetOpenFilename）
- ワークブックを開く・閉じる
- シートのコピーと操作
- エラーハンドリングの基礎

---

## 1. ファイル選択ダイアログ

ユーザーにファイルを選んでもらうダイアログを表示します。

```vba
Sub ファイル選択()
    Dim filePath As Variant
    filePath = Application.GetOpenFilename( _
                    Title:="ファイルを選択してください")

    If filePath = False Then
        MsgBox "キャンセルされました。"
        Exit Sub
    End If

    MsgBox "選択されたファイル: " & filePath
End Sub
```

### ポイント
- 戻り値は `Variant` 型（キャンセル時は `False` が返る）
- `Title` でダイアログのタイトルを設定

> **実務マクロではここで使われています：**
> ```vba
> filePath = Application.GetOpenFilename( _
>                 Title:="【新】清掃指示書ファイルを選択してください")
> If filePath = False Then Exit Sub
> ```

### 複数ファイル選択（FileDialog）

```vba
Sub 複数ファイル選択()
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogOpen)
    fd.AllowMultiSelect = True
    fd.Title = "ファイルを選択してください"

    If fd.Show = -1 Then
        Dim i As Long
        For i = 1 To fd.SelectedItems.Count
            MsgBox i & "番目: " & fd.SelectedItems(i)
        Next i
    End If
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> Set FileDialog = Application.FileDialog(msoFileDialogOpen)
> FileDialog.AllowMultiSelect = True
> ```
> 清掃確認書と清掃指示書の2ファイルを同時に選択させています。

## 2. ワークブックを開く・閉じる

```vba
Sub ブックを開いて閉じる()
    Dim filePath As Variant
    filePath = Application.GetOpenFilename()
    If filePath = False Then Exit Sub

    ' ブックを開く
    Dim wb As Workbook
    Set wb = Workbooks.Open(CStr(filePath))

    MsgBox "開いたブック: " & wb.Name

    ' ブックを閉じる（保存しない）
    wb.Close SaveChanges:=False
End Sub
```

### ThisWorkbook と ActiveWorkbook の違い
- `ThisWorkbook` — マクロが書かれているブック（常に同じ）
- `ActiveWorkbook` — 今アクティブなブック（変わる可能性あり）

> **実務マクロでは：** `ThisWorkbook` を使って自分自身のシートを参照しています。

## 3. シートの参照

```vba
Sub シートの参照()
    ' シート名で取得
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Sheet1")

    ' シートのセルを操作
    ws.Range("A1").Value = "テスト"
    ws.Cells(2, 1).Value = "Cellsでも可能"
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> Set ws4_9 = wbTarget.Worksheets("清掃指示書 (4-9)")
> Set ws10_15 = wbTarget.Worksheets("清掃指示書 (10-15)")
> Set wsShijisho = ThisWorkbook.Worksheets("指示書")
> ```

## 4. シートのコピー

```vba
Sub シートコピー()
    Dim srcWs As Worksheet
    Dim destWs As Worksheet

    Set srcWs = ThisWorkbook.Worksheets("Sheet1")

    ' 最後尾にコピー
    srcWs.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)

    ' コピーされたシートの名前を変更
    Set destWs = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    destWs.Name = "Sheet1のコピー"

    MsgBox "シートをコピーしました: " & destWs.Name
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> srcWs.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
> Set destWs = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
> destWs.Name = "Copied_清掃確認書"
> ```

## 5. シートの追加と削除

```vba
Sub シート追加()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add( _
                After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "新しいシート"
    MsgBox "シートを追加しました。"
End Sub

Sub シート削除()
    Application.DisplayAlerts = False    ' 確認メッセージを非表示
    ThisWorkbook.Sheets("新しいシート").Delete
    Application.DisplayAlerts = True     ' 元に戻す
    MsgBox "シートを削除しました。"
End Sub
```

## 6. シートの存在チェック

シートが存在するか確認してから操作する関数です。

```vba
Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then SheetExists = True
End Function

Sub シート存在チェック()
    If SheetExists("Sheet1") Then
        MsgBox "Sheet1 は存在します。"
    Else
        MsgBox "Sheet1 は存在しません。"
    End If
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> Function SheetExists(sheetName As String) As Boolean
>     Dim ws As Worksheet
>     On Error Resume Next
>     Set ws = ThisWorkbook.Sheets(sheetName)
>     On Error GoTo 0
>     If Not ws Is Nothing Then SheetExists = True
> End Function
> ```
> シート削除前に存在確認しています。

## 7. On Error Resume Next — エラーハンドリング

存在しないシートを参照するとエラーになります。
`On Error Resume Next` でエラーを無視し、`On Error GoTo 0` で元に戻します。

```vba
Sub エラーハンドリング()
    Dim ws As Worksheet

    On Error Resume Next          ' エラーを無視
    Set ws = ThisWorkbook.Worksheets("存在しないシート")
    On Error GoTo 0               ' エラー無視を解除

    If ws Is Nothing Then
        MsgBox "シートが見つかりませんでした。"
    Else
        MsgBox "シートが見つかりました: " & ws.Name
    End If
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> On Error Resume Next
> Set ws4_9 = wbTarget.Worksheets("清掃指示書 (4-9)")
> Set ws10_15 = wbTarget.Worksheets("清掃指示書 (10-15)")
> On Error GoTo 0
> ```
> シートが存在しない場合でもエラーで止まらないようにしています。

---

## まとめ
- `Application.GetOpenFilename` — ファイル選択ダイアログ
- `Workbooks.Open(filePath)` — ブックを開く
- `wb.Close SaveChanges:=False` — ブックを閉じる
- `srcWs.Copy After:=...` — シートをコピー
- `SheetExists()` — シートの存在チェック
- `On Error Resume Next` — エラーを無視する（限定的に使用）

## 次のレッスン
→ [レッスン6: 文字列操作とデータ判定](../06_文字列操作とデータ判定/lesson.md)
