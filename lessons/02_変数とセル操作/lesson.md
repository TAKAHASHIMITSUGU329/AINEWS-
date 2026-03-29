# レッスン2: 変数とセル操作

## このレッスンで学ぶこと
- 変数の宣言と使い方
- 主なデータ型
- Range と Cells でセルを読み書きする方法

---

## 1. 変数とは？

変数は「データを入れる箱」です。名前を付けて、値を入れたり取り出したりできます。

```vba
Dim roomNo As String    ' 部屋番号を入れる箱
roomNo = "401"          ' 箱に値を入れる
MsgBox roomNo           ' 箱の中身を表示
```

## 2. 変数の宣言（Dim）

変数を使うときは `Dim` で宣言します。

```vba
Dim 変数名 As データ型
```

### よく使うデータ型

| データ型 | 用途 | 例 |
|---|---|---|
| `String` | 文字列 | `"401"`, `"STAY"` |
| `Long` | 整数（大きい数） | `1`, `100`, `15` |
| `Double` | 小数 | `3.14` |
| `Boolean` | 真偽値 | `True`, `False` |
| `Variant` | 何でも入る | ファイルパス等 |

> **実務マクロでの例：**
> ```vba
> Dim filePath As Variant      ' ファイルパス（キャンセル時はFalseになるのでVariant）
> Dim wbTarget As Workbook      ' ワークブック
> Dim ws4_9 As Worksheet        ' ワークシート
> Dim r As Long                 ' 行番号
> Dim txt As String             ' 文字列
> ```

## 3. Option Explicit

ファイルの先頭に `Option Explicit` と書くと、宣言していない変数を使ったときにエラーになります。
タイプミスを防げるので、必ず書きましょう。

```vba
Option Explicit    ' ← これを必ず書く

Sub テスト()
    Dim count As Long
    count = 10       ' OK
    ' cont = 10      ' ← スペルミス → エラーで気づける！
End Sub
```

> **実務マクロでは：** すべてのモジュールの先頭に `Option Explicit` が書かれています。

## 4. セルの読み書き — Range

`Range` はセルのアドレス（A1形式）で指定します。

```vba
Sub Range練習()
    ' セルに値を書き込む
    Range("A1").Value = "部屋番号"
    Range("B1").Value = "ステータス"
    Range("A2").Value = 401
    Range("B2").Value = "S"

    ' セルの値を読み取る
    Dim status As String
    status = Range("B2").Value
    MsgBox "部屋401のステータス: " & status
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> ThisWorkbook.Worksheets("指示書作成").Range("c5").Value = CStr(filePath)
> ```
> ファイルパスをセルC5に書き込んでいます。

## 5. セルの読み書き — Cells

`Cells` は行番号と列番号（数値）で指定します。ループ処理で便利です。

```vba
Sub Cells練習()
    ' Cells(行, 列) で指定
    Cells(1, 1).Value = "部屋番号"   ' A1
    Cells(1, 2).Value = "ステータス" ' B1
    Cells(2, 1).Value = 401          ' A2
    Cells(2, 2).Value = "S"          ' B2

    ' 値を読み取る
    Dim roomNo As Variant
    roomNo = Cells(2, 1).Value
    MsgBox "部屋番号: " & roomNo
End Sub
```

### Range と Cells の使い分け
- **Range**: 特定のセルをアドレスで指定するとき → `Range("C5")`
- **Cells**: 行・列を変数で変えたいとき → `Cells(r, 3)`（r行目の3列目）

> **実務マクロでの使い分け：**
> ```vba
> wsSrc.Range("B2").Value = "D=3, S=5"          ' 固定位置にはRange
> ws.Cells(r, roomCol).Value                     ' ループ内ではCells
> ```

## 6. ワークシートの指定

どのシートのセルかを明示するには、シートを先に指定します。

```vba
Sub シート指定()
    ' アクティブシート（今開いているシート）
    ActiveSheet.Range("A1").Value = "テスト"

    ' シート名で指定
    Worksheets("Sheet1").Range("A1").Value = "テスト"

    ' マクロブック自身のシート
    ThisWorkbook.Worksheets("Menu").Range("A1").Value = "テスト"
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> Set wsShijisho = ThisWorkbook.Worksheets("指示書")
> ```
> マクロブック内の「指示書」シートを変数に格納しています。

---

## まとめ
- `Dim 変数名 As データ型` で変数を宣言する
- `Option Explicit` を必ず書く
- `Range("A1")` — アドレスでセル指定
- `Cells(行, 列)` — 番号でセル指定（ループ向き）
- シートを指定するには `Worksheets("シート名")` を使う

## 次のレッスン
→ [レッスン3: 条件分岐と色付け](../03_条件分岐と色付け/lesson.md)
