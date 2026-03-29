# レッスン3: 条件分岐と色付け

## このレッスンで学ぶこと
- If...Then...Else による条件分岐
- セルの背景色を変更する方法（Interior.Color, RGB）
- 実務マクロの色付け処理を理解する

---

## 1. If...Then...Else の基本

「もし〇〇なら△△する」という処理を書けます。

```vba
Sub 条件分岐の基本()
    Dim status As String
    status = Range("A1").Value

    If status = "S" Then
        MsgBox "ステイです"
    ElseIf status = "D" Then
        MsgBox "出発です"
    Else
        MsgBox "その他です"
    End If
End Sub
```

### 構造
```
If 条件 Then
    条件が真のときの処理
ElseIf 別の条件 Then
    別の条件が真のときの処理
Else
    どれにも当てはまらないときの処理
End If
```

## 2. 比較演算子

| 演算子 | 意味 | 例 |
|---|---|---|
| `=` | 等しい | `status = "S"` |
| `<>` | 等しくない | `status <> ""` |
| `>` | より大きい | `count > 0` |
| `<` | より小さい | `count < 10` |
| `>=` | 以上 | `Len(txt) >= 1` |

## 3. 論理演算子（Or / And）

複数の条件を組み合わせられます。

```vba
' どちらかが真なら
If c1 = "S" Or c2 = "S" Then
    MsgBox "Sが含まれています"
End If

' 両方とも真なら
If floorNo >= 4 And floorNo <= 9 Then
    MsgBox "4F～9Fです"
End If
```

> **実務マクロではここで使われています：**
> ```vba
> If c1 = "S" Or c2 = "S" Then
>     .Interior.Color = RGB(0, 176, 240)       '青
> ElseIf c1 = "D" Or c2 = "D" Then
>     .Interior.Color = RGB(255, 105, 180)     'ピンク
> End If
> ```
> C/Iの1文字目または2文字目が "S" か "D" かで色を分けています。

## 4. セルの背景色を変更する

`Interior.Color` と `RGB` 関数を使います。

```vba
Sub 色付けの基本()
    Range("A1").Interior.Color = RGB(0, 176, 240)     ' 青
    Range("A2").Interior.Color = RGB(255, 105, 180)   ' ピンク
    Range("A3").Interior.Color = RGB(255, 255, 0)     ' 黄色
    Range("A4").Interior.Color = RGB(0, 255, 0)       ' 緑
End Sub
```

### RGB関数
`RGB(赤, 緑, 青)` — 各色を0～255で指定します。

| 色 | RGB | 実務での用途 |
|---|---|---|
| 青 | `RGB(0, 176, 240)` | S（ステイ）の部屋 |
| ピンク | `RGB(255, 105, 180)` | D（出発）の部屋 |
| 黄色 | `RGB(255, 255, 0)` | 出発・RC |
| 緑 | `RGB(0, 255, 0)` | STAY |

## 5. 条件に応じて色を付ける

セルの値を見て、条件に合う色を付けます。

```vba
Sub 条件で色付け()
    ' テストデータを作成
    Range("A1").Value = "S"
    Range("A2").Value = "D"
    Range("A3").Value = "STAY"
    Range("A4").Value = "出発"

    ' 条件に応じて色付け
    Dim cell As Range
    For Each cell In Range("A1:A4")
        If cell.Value = "S" Then
            cell.Interior.Color = RGB(0, 176, 240)       ' 青
        ElseIf cell.Value = "D" Then
            cell.Interior.Color = RGB(255, 105, 180)     ' ピンク
        ElseIf cell.Value = "STAY" Then
            cell.Interior.Color = RGB(0, 255, 0)         ' 緑
        ElseIf cell.Value = "出発" Then
            cell.Interior.Color = RGB(255, 255, 0)       ' 黄色
        End If
    Next cell
End Sub
```

## 6. 色をリセットする

```vba
Sub 色リセット()
    Range("A1:A10").Interior.ColorIndex = xlNone
End Sub
```

## 7. With 文

同じオブジェクトに対して複数の操作をするとき、`With` を使うと簡潔に書けます。

```vba
' Withなし
Cells(1, 1).Interior.Color = RGB(0, 176, 240)
Cells(1, 1).Font.Bold = True
Cells(1, 1).Font.Size = 14

' Withあり（同じ意味）
With Cells(1, 1)
    .Interior.Color = RGB(0, 176, 240)
    .Font.Bold = True
    .Font.Size = 14
End With
```

> **実務マクロではここで使われています：**
> ```vba
> With ws.Cells(r, roomCol)
>     If c1 = "S" Or c2 = "S" Then
>         .Interior.Color = RGB(0, 176, 240)
>     ElseIf c1 = "D" Or c2 = "D" Then
>         .Interior.Color = RGB(255, 105, 180)
>     End If
> End With
> ```

---

## まとめ
- `If...Then...ElseIf...Else...End If` で条件分岐
- `Or` / `And` で複数条件を組み合わせ
- `Interior.Color = RGB(赤, 緑, 青)` でセルの背景色を変更
- `With` 文で同じオブジェクトへの操作をまとめる

## 次のレッスン
→ [レッスン4: 繰り返し処理](../04_繰り返し処理/lesson.md)
