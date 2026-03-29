# レッスン6: 文字列操作とデータ判定

## このレッスンで学ぶこと
- Mid$, Left$, Right$ で文字を取り出す
- InStr で文字列を検索する
- CStr, Len, Trim などの文字列関数
- 実務マクロのS/D判定ロジックを理解する

---

## 1. 文字列から1文字を取り出す — Mid$

```vba
Sub Mid関数()
    Dim txt As String
    txt = "SD"

    Dim c1 As String, c2 As String
    c1 = Mid$(txt, 1, 1)    ' 1文字目 → "S"
    c2 = Mid$(txt, 2, 1)    ' 2文字目 → "D"

    MsgBox "1文字目: " & c1 & vbCrLf & "2文字目: " & c2
End Sub
```

### Mid$ の書き方
```
Mid$(文字列, 開始位置, 文字数)
```
- 開始位置は1から始まる（0ではない）

> **実務マクロではここで使われています：**
> ```vba
> If Len(txt) >= 1 Then c1 = Mid$(txt, 1, 1) Else c1 = ""
> If Len(txt) >= 2 Then c2 = Mid$(txt, 2, 1) Else c2 = ""
> ```
> C/I列の値から1文字目と2文字目を取り出して、SかDかを判定しています。

## 2. 先頭・末尾の文字を取り出す — Left$, Right$

```vba
Sub LeftRight()
    Dim roomNo As String
    roomNo = "1201"

    ' 先頭2文字 → フロア番号
    Dim floorStr As String
    floorStr = Left$(roomNo, 2)     ' "12"

    MsgBox "部屋番号: " & roomNo & " → フロア: " & floorStr & "F"
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> If Left(destWs.Cells(k, 2).Value, Len(floorPrefix)) = floorPrefix Then
> ```
> 部屋番号の先頭文字でフロアを判定しています（例: "12xx" → 12F）。

## 3. 文字列の長さ — Len

```vba
Sub Len関数()
    Dim txt As String
    txt = "SD"
    MsgBox "文字数: " & Len(txt)    ' → 2

    ' 空文字チェック
    If Len(txt) = 0 Then
        MsgBox "空です"
    Else
        MsgBox "文字が入っています: " & txt
    End If
End Sub
```

## 4. 文字列の検索 — InStr

文字列の中に特定の文字が含まれるか検索します。

```vba
Sub InStr関数()
    Dim s As String
    s = "14:30"

    ' ":" が含まれるか？
    If InStr(s, ":") > 0 Then
        MsgBox "時刻データです: " & s
    Else
        MsgBox "時刻ではありません"
    End If
End Sub
```

### InStr の戻り値
- 見つかった場合：その位置（1以上）
- 見つからない場合：0

> **実務マクロではここで使われています：**
> ```vba
> If InStr(s, ":") > 0 Or InStr(s, "：") > 0 Then
>     ' 時刻データとして備考に転記
> End If
> ```
> C/I 2行目に ":" が含まれていれば時刻として転記しています。

### 業務用語の検索にも使用

```vba
' 実務マクロの集計処理から
If InStr(1, destWs.Cells(k, 7).Value, "ダブル解除") > 0 Then
    totalDD = totalDD + 1
ElseIf InStr(1, destWs.Cells(k, 7).Value, "ツイン解除") > 0 Then
    totalTD = totalTD + 1
End If
```

## 5. 型変換 — CStr

セルの値を文字列に変換します。数値や日付が入っている可能性があるときに使います。

```vba
Sub CStr関数()
    Cells(1, 1).Value = 401    ' 数値として入力

    Dim roomNo As String
    roomNo = CStr(Cells(1, 1).Value)    ' 文字列 "401" に変換

    MsgBox "部屋番号（文字列）: " & roomNo
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> ci1 = CStr(wsSrc.Cells(r, SRC_CI_COL).Value)
> ```
> セルの値を安全に文字列に変換してから操作しています。

## 6. 空白を除去 — Trim, LTrim

```vba
Sub Trim関数()
    Dim txt As String
    txt = "  ●Wセット  "

    MsgBox "元: [" & txt & "]"
    MsgBox "LTrim: [" & LTrim$(txt) & "]"    ' 左の空白を除去
    MsgBox "Trim: [" & Trim$(txt) & "]"      ' 両側の空白を除去

    ' 先頭文字のチェック
    If Left$(LTrim$(txt), 1) = ChrW(&H25CF) Then
        MsgBox "● で始まっています → Wセット"
    End If
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> ci1Trim = LTrim$(ci1)
> If Left$(ci1Trim, 1) = "●" Then
>     ' Wセットとして処理
> End If
> ```

## 7. 大文字・小文字を無視した比較 — StrComp

```vba
Sub StrComp関数()
    Dim val As String
    val = "stay"    ' 小文字

    ' 大文字小文字を無視して比較
    If StrComp(val, "STAY", vbTextCompare) = 0 Then
        MsgBox "STAY です（大文字小文字を無視）"
    End If
End Sub
```

> **実務マクロではここで使われています：**
> ```vba
> If StrComp(valE, "STAY", vbTextCompare) = 0 Then
>     stayCnt = stayCnt + 1
> End If
> ```

---

## まとめ
- `Mid$(txt, 位置, 文字数)` — 文字列から文字を取り出す
- `Left$(txt, 文字数)` — 先頭から取り出す
- `Len(txt)` — 文字列の長さ
- `InStr(txt, "検索文字")` — 文字列の検索（0なら見つからない）
- `CStr(値)` — 文字列に変換
- `LTrim$()` / `Trim$()` — 空白を除去
- `StrComp(a, b, vbTextCompare)` — 大文字小文字を無視した比較

## 次のレッスン
→ [レッスン7: データ転記と集計](../07_データ転記と集計/lesson.md)
