# レッスン8: 実践 - 統計シート自動作成

## このレッスンで学ぶこと
- 数式（Formula）をVBAから書き込む
- フォント・書式の設定
- Sub（サブルーチン）の分割と呼び出し
- 実務マクロの全体像を理解する

---

## 1. 数式の書き込み — .Formula

VBAからセルに数式を入れることができます。

```vba
Sub 数式の書き込み()
    ' データ
    Cells(1, 1).Value = "stay"
    Cells(1, 2).Value = "out"
    Cells(1, 3).Value = "合計時間(分)"
    Cells(2, 1).Value = 5     ' stay 5部屋
    Cells(2, 2).Value = 3     ' out 3部屋

    ' 数式を書き込み（stay*3 + out*6）
    Cells(2, 3).Formula = "=A2*3+B2*6"

    MsgBox "合計時間: " & Cells(2, 3).Value & "分"
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> ws.Cells(rowOut, 5).Formula = "=" & _
>     ws.Cells(rowOut, 2).Address(False, False) & "*3+" & _
>     ws.Cells(rowOut, 3).Address(False, False) & "*6+" & _
>     ws.Cells(rowOut, 4).Address(False, False) & "*1"
> ```
> stay 3分、out 6分、済 1分の計算式を動的に生成しています。

### .Address で動的にセル参照を作る

```vba
Dim cellAddr As String
cellAddr = Cells(2, 1).Address(False, False)   ' → "A2"
```
- `False, False` — 相対参照（$なし）
- `True, True` — 絶対参照（$A$2）

## 2. R1C1形式の数式 — .FormulaR1C1

行列番号で数式を書く形式です。同じ数式を複数行にコピーしたいときに便利です。

```vba
Sub R1C1数式()
    ' R1C1形式: RC[-1] は「同じ行の1つ左のセル」
    Cells(2, 3).FormulaR1C1 = "=INT(RC[-1]/60)&""時間""&MOD(RC[-1],60)&""分"""
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> ws.Cells(baseRow + i, 3).FormulaR1C1 = _
>     "=INT(RC[-1]/60)&""時間""&MOD(RC[-1],60)&""分"""
> ```
> 分を「○時間○分」の形式に変換する数式です。

## 3. 書式設定

### フォント

```vba
Sub フォント設定()
    With Cells(1, 1)
        .Value = "W"
        .Font.Bold = True       ' 太字
        .Font.Size = 30          ' フォントサイズ
        .Font.Color = RGB(255, 0, 0)  ' 文字色（赤）
    End With
End Sub
```

### 配置

```vba
Sub 配置設定()
    With Cells(1, 1)
        .HorizontalAlignment = xlLeft     ' 左揃え
        ' .HorizontalAlignment = xlCenter ' 中央揃え
        ' .HorizontalAlignment = xlRight  ' 右揃え
    End With
End Sub
```

### 列幅

```vba
Sub 列幅設定()
    Columns("A:A").ColumnWidth = 10
    Columns("B:E").ColumnWidth = 16
End Sub
```

### 数値の表示形式

```vba
Sub 表示形式()
    Cells(1, 1).Value = TimeSerial(9, 0, 0)  ' 9:00
    Cells(1, 1).NumberFormatLocal = "hh:mm"   ' → "09:00" と表示
End Sub
```

> **実務マクロでは：** 終了時間を `hh:mm` 形式で表示しています。

## 4. SUM関数の書き込み

```vba
Sub SUM数式()
    ' データ
    Dim r As Long
    For r = 1 To 10
        Cells(r, 1).Value = r * 2
    Next r

    ' SUM関数を書き込み
    Cells(11, 1).Formula = "=SUM(A1:A10)"
    Cells(11, 1).Font.Bold = True

    MsgBox "合計: " & Cells(11, 1).Value
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> For c = 2 To 12
>     summaryWs.Cells(25, c).Formula = "=SUM(" & _
>         Cells(8, c).Address(False, False) & ":" & _
>         Cells(24, c).Address(False, False) & ")"
> Next c
> ```
> 各列の小計を SUM 関数で自動計算しています。

## 5. Sub の分割 — 処理を分けて管理する

大きな処理は小さな Sub に分けると読みやすくなります。

```vba
' メイン処理
Sub メイン()
    Call データ準備
    Call 集計処理
    Call 書式設定処理
    MsgBox "完了しました。"
End Sub

' データ準備
Private Sub データ準備()
    Cells(1, 1).Value = "フロア"
    Cells(1, 2).Value = "部屋数"
End Sub

' 集計処理
Private Sub 集計処理()
    ' 集計ロジック...
End Sub

' 書式設定
Private Sub 書式設定処理()
    Range("A1:B1").Font.Bold = True
End Sub
```

### Public と Private
- `Public Sub` — どこからでも呼び出せる（マクロ一覧に表示される）
- `Private Sub` — 同じモジュール内からのみ呼び出せる

### Call と引数

```vba
Sub メイン処理()
    Call 色付け処理(ActiveSheet)
End Sub

Private Sub 色付け処理(ByVal ws As Worksheet)
    ' ws に対して色付け
    ws.Cells(1, 1).Interior.Color = RGB(0, 176, 240)
End Sub
```

> **実務マクロの構造：**
> ```
> OpenFileAndColorRooms()        ← メイン（Public）
>   ├─ ColorRooms_AllBlocks()    ← 色付け（Private）
>   │    └─ ColorRooms_OneBlock() ← 1ブロック分の色付け
>   ├─ Copy4F_Block()            ← 4F転記
>   ├─ CopyFloors5To15()         ← 5F～15F転記
>   │    └─ CopyFloorsFromSheet() ← 1シート分の転記
>   │         └─ CopyFloorGeneric() ← 1フロア分の転記
>   └─ CountDSTotalsForAll()     ← D/S集計
> ```

## 6. 実務マクロの全体像

ここまで学んだことを振り返ると、実務マクロで使っている技術はすべてカバーしています。

| レッスン | 実務マクロでの使用箇所 |
|---|---|
| 1. MsgBox | 完了メッセージ、エラー通知 |
| 2. 変数・セル操作 | すべての処理の基盤 |
| 3. 条件分岐・色付け | S/D判定、RGB色付け |
| 4. 繰り返し処理 | 全行走査、Step 2 |
| 5. シート操作 | ファイル選択、シートコピー |
| 6. 文字列操作 | Mid$でS/D判定、InStrで備考判定 |
| 7. 転記・集計 | フロア別転記、D/S集計 |
| 8. 数式・書式 | 作業時間計算、罫線、フォント |

---

## まとめ
- `.Formula = "=SUM(...)"` — セルに数式を書き込む
- `.FormulaR1C1` — R1C1形式の数式（相対参照が楽）
- `.Font.Bold`, `.Font.Size` — フォント設定
- `.NumberFormatLocal` — 表示形式
- `.ColumnWidth` — 列幅
- `Call サブ名(引数)` — 処理を分割して呼び出す
- `Public` / `Private` — アクセス範囲の制御

## おわりに

全8レッスンお疲れ様でした。ここまでの内容を理解していれば、`macros/` フォルダにある実務マクロのコードを読んで理解できるはずです。

次のステップとして：
1. `macros/01_指示書色付け転記.bas` を上から順に読んでみる
2. `macros/02_統計集計.bas` を上から順に読んでみる
3. 小さな修正や改善を自分で試してみる

わからない部分があれば、該当するレッスンに戻って復習してください。
