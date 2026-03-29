# レッスン7: データ転記と集計

## このレッスンで学ぶこと
- シート間のデータ転記
- Select Case による分岐
- Application.Match で値を検索
- カウント変数による集計処理

---

## 1. シート間のデータ転記

あるシートのデータを別のシートにコピーする処理です。

```vba
Sub 転記の基本()
    Dim wsSrc As Worksheet     ' 転記元
    Dim wsDest As Worksheet    ' 転記先

    Set wsSrc = Worksheets("Sheet1")
    Set wsDest = Worksheets("Sheet2")

    ' Sheet1のA1の値をSheet2のA1にコピー
    wsDest.Cells(1, 1).Value = wsSrc.Cells(1, 1).Value
End Sub
```

### ループで複数行を転記

```vba
Sub 複数行転記()
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Set wsSrc = Worksheets("Sheet1")
    Set wsDest = Worksheets("Sheet2")

    Dim r As Long
    Dim destRow As Long
    destRow = 1

    For r = 2 To 20
        If wsSrc.Cells(r, 1).Value <> "" Then
            wsDest.Cells(destRow, 1).Value = wsSrc.Cells(r, 1).Value
            wsDest.Cells(destRow, 2).Value = wsSrc.Cells(r, 2).Value
            destRow = destRow + 1
        End If
    Next r
End Sub
```

> **実務マクロでの転記パターン：**
> 清掃指示書(4-9) → 指示書シートへ、部屋番号・色・Wセット・備考を転記しています。

## 2. Select Case — 複数パターンの分岐

If文が多くなるときは `Select Case` を使うと見やすくなります。

```vba
Sub SelectCaseの基本()
    Dim floorNo As Long
    floorNo = 6

    Dim roomColDest As Long

    Select Case floorNo
        Case 4, 5
            roomColDest = 1      ' A列
        Case 6, 7
            roomColDest = 8      ' H列
        Case 8, 9
            roomColDest = 15     ' O列
        Case 10, 11
            roomColDest = 22     ' V列
        Case 12, 13
            roomColDest = 29     ' AC列
        Case 14, 15
            roomColDest = 36     ' AJ列
    End Select

    MsgBox floorNo & "F → " & roomColDest & "列目に転記"
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> Private Sub GetDestInfo(ByVal floorNo As Long, ...)
>     Select Case floorNo
>         Case 4:  roomColDest = 1:  firstRowDest = 3:  lastRowDest = 19
>         Case 5:  roomColDest = 1:  firstRowDest = 32: lastRowDest = 48
>         ' ...
>     End Select
> End Sub
> ```
> フロア番号から転記先の列・行を決定しています。

## 3. Application.Match — 値の検索

指定した範囲から値を検索し、その位置（何番目か）を返します。

```vba
Sub Match練習()
    ' テストデータ
    Dim ws As Worksheet
    Set ws = ActiveSheet
    ws.Cells(1, 1).Value = 401
    ws.Cells(2, 1).Value = 402
    ws.Cells(3, 1).Value = 403
    ws.Cells(4, 1).Value = 404
    ws.Cells(5, 1).Value = 405

    ' 403を検索
    Dim m As Variant
    m = Application.Match(403, ws.Range("A1:A5"), 0)

    If IsError(m) Then
        MsgBox "見つかりませんでした"
    Else
        MsgBox "403 は " & m & " 番目にあります（" & m & "行目）"
    End If
End Sub
```

### ポイント
- 戻り値は `Variant`（エラーの可能性があるため）
- `IsError(m)` でエラーチェック
- 第3引数の `0` は完全一致

> **実務マクロではここで使われています：**
> ```vba
> m = Application.Match(roomNo, wsDest.Range(...), 0)
> If IsError(m) Then GoTo NextRoom
> destRow = firstRowDest + CLng(m) - 1
> ```
> 部屋番号を指示書シートから検索し、転記先の行を特定しています。

## 4. カウント変数による集計

条件に合うデータの数を数えます。

```vba
Sub カウント集計()
    Dim r As Long, lastRow As Long
    Dim sCount As Long, dCount As Long

    sCount = 0
    dCount = 0

    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        Dim ci As String
        ci = CStr(Cells(r, 2).Value)

        If ci = "S" Then
            sCount = sCount + 1
        ElseIf ci = "D" Then
            dCount = dCount + 1
        End If
    Next r

    MsgBox "S: " & sCount & "件" & vbCrLf & _
           "D: " & dCount & "件"
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> dCount = 0: sCount = 0
> For r = FIRST_DATA_ROW To lastRow Step 2
>     ' ...
>     If c1 = "D" Or c2 = "D" Then dCount = dCount + 1
>     If c1 = "S" Or c2 = "S" Then sCount = sCount + 1
> Next r
> wsSrc.Range(addr).Value = "D=" & dCount & ", S=" & sCount
> ```
> 各フロアのD/S件数を数えて、清掃指示書に「D=3, S=5」の形式で印字しています。

## 5. 罫線の設定

集計表に罫線を付けます。

```vba
Sub 罫線設定()
    ' 外枠
    Range("A1:D5").Borders(xlEdgeTop).LineStyle = xlContinuous
    Range("A1:D5").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A1:D5").Borders(xlEdgeLeft).LineStyle = xlContinuous
    Range("A1:D5").Borders(xlEdgeRight).LineStyle = xlContinuous

    ' 内側の線
    Range("A1:D5").Borders(xlInsideVertical).LineStyle = xlContinuous
    Range("A1:D5").Borders(xlInsideHorizontal).LineStyle = xlContinuous
End Sub
```

## 6. Const — 定数

変更しない値は `Const` で宣言します。マジックナンバーを避けられます。

```vba
Const HEADER_ROW As Long = 4
Const FIRST_DATA_ROW As Long = 5
Const LAST_ROW As Long = 100
```

> **実務マクロでは：** データの開始行やヘッダー行を定数で管理しています。

---

## まとめ
- シート間の転記: `wsDest.Cells(...) = wsSrc.Cells(...)`
- `Select Case` — 多分岐の処理を見やすく書ける
- `Application.Match` — 範囲内で値を検索
- カウント変数（`count = count + 1`）で集計
- `Borders` で罫線を設定
- `Const` で定数を定義

## 次のレッスン
→ [レッスン8: 実践 - 統計シート自動作成](../08_実践_統計シート自動作成/lesson.md)
