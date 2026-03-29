# レッスン4: 繰り返し処理

## このレッスンで学ぶこと
- For...Next ループ
- For Each...Next ループ
- Step（ステップ）を使った飛ばし処理
- 最終行の取得方法

---

## 1. For...Next の基本

同じ処理を繰り返すときに使います。

```vba
Sub For基本()
    Dim i As Long
    For i = 1 To 5
        Cells(i, 1).Value = "行 " & i
    Next i
End Sub
```

### 構造
```
For カウンタ変数 = 開始値 To 終了値
    繰り返す処理
Next カウンタ変数
```

## 2. 実務で使うパターン — 行のループ

データが入っている行を1行ずつ処理します。

```vba
Sub 行のループ()
    Dim r As Long
    For r = 2 To 10    ' 2行目～10行目
        ' r行目のA列の値を表示
        If Cells(r, 1).Value <> "" Then
            Cells(r, 2).Value = "処理済み"
        End If
    Next r
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> For r = firstRow To lastRow
>     txt = CStr(ws.Cells(r, ciCol).Value)
>     ' ...色付け処理...
> Next r
> ```
> 清掃指示書の各行を順番に見て色を付けています。

## 3. Step — 2行ずつ飛ばす

清掃指示書では1部屋のデータが2行で1セットです。`Step 2` で2行ずつ処理します。

```vba
Sub Step2のループ()
    Dim r As Long
    For r = 5 To 20 Step 2    ' 5, 7, 9, 11, ... と2行ずつ
        ' 1行目（奇数行）: C/I情報
        Cells(r, 3).Value = "1行目のデータ"
        ' 2行目（偶数行）: 備考
        Cells(r + 1, 3).Value = "2行目のデータ"
    Next r
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> For r = SRC_FIRST_DATA_ROW To lastRowSrc Step 2
>     roomNo = wsSrc.Cells(r, SRC_ROOM_COL).Value
>     ci1 = CStr(wsSrc.Cells(r, SRC_CI_COL).Value)
>     ' r+1 行目は備考
>     s = wsSrc.Cells(r + 1, SRC_CI_COL).Text
> Next r
> ```

## 4. For Each...Next — セル範囲のループ

範囲内のセルを1つずつ処理します。

```vba
Sub ForEach基本()
    Dim cell As Range
    For Each cell In Range("A1:A10")
        If cell.Value = "STAY" Then
            cell.Interior.Color = RGB(0, 255, 0)  ' 緑
        End If
    Next cell
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> For Each cell In targetRange
>     If cell.Value = "STAY" Then
>         cell.Interior.Color = RGB(0, 255, 0)
>     ElseIf cell.Value = "出発" Or cell.Value = "ＲＣ" Then
>         cell.Interior.Color = RGB(255, 255, 0)
>     End If
> Next cell
> ```
> 清掃確認書の全セルを走査して色を付けています。

## 5. 最終行の取得

データが何行あるか分からないとき、最終行を自動で取得します。

```vba
Sub 最終行の取得()
    Dim lastRow As Long
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    MsgBox "A列の最終行: " & lastRow

    ' 2行目から最終行までループ
    Dim r As Long
    For r = 2 To lastRow
        ' 処理...
    Next r
End Sub
```

### 仕組み
- `Cells(Rows.Count, 1)` — A列の一番下のセル（1048576行目）
- `.End(xlUp)` — そこから上に向かって最初のデータがあるセルへ移動
- `.Row` — その行番号を取得

> **実務マクロではここで使われています：**
> ```vba
> lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, SRC_ROOM_COL).End(xlUp).Row
> ```

## 6. ループを途中で抜ける — Exit For

条件を満たしたらループを終了します。

```vba
Sub ループを抜ける()
    Dim r As Long
    For r = 1 To 100
        If IsEmpty(Cells(r, 1)) Then
            MsgBox "空白セルが見つかりました: " & r & "行目"
            Exit For
        End If
    Next r
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> If IsEmpty(ws.Cells(r, roomCol)) And IsEmpty(ws.Cells(r + 1, roomCol)) Then Exit For
> ```
> 空白行が2行続いたらループを終了しています。

## 7. 列のループ

行だけでなく列もループできます。ヘッダーを探すときに使います。

```vba
Sub 列のループ()
    Dim col As Long
    Dim maxCol As Long
    maxCol = Cells(1, Columns.Count).End(xlToLeft).Column

    For col = 1 To maxCol
        If Cells(1, col).Value = "ROOM" Then
            MsgBox "ROOM列が見つかりました: " & col & "列目"
            Exit For
        End If
    Next col
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> For col = 1 To maxCol
>     If ws.Cells(HEADER_ROW, col).Value = "ROOM" Then
>         roomCol = col
>         ' ...
>     End If
> Next col
> ```
> ヘッダー行から「ROOM」列を自動検出しています。

---

## まとめ
- `For i = 1 To 10` — 回数を指定してループ
- `Step 2` — 2行ずつ飛ばしてループ（2行1セットのデータ）
- `For Each cell In Range(...)` — セル範囲を1つずつ処理
- `Cells(Rows.Count, 1).End(xlUp).Row` — 最終行の取得
- `Exit For` — ループの途中終了

## 次のレッスン
→ [レッスン5: ワークブックとシート操作](../05_ワークブックとシート操作/lesson.md)
