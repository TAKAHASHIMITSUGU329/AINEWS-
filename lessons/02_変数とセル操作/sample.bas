Attribute VB_Name = "Lesson02_変数とセル操作"
Option Explicit

'=====================================================
' レッスン2: 変数とセル操作 - サンプルコード
'=====================================================

' --- 例1: 変数の基本 ---
Sub 変数の基本()
    Dim roomNo As String
    Dim floorNo As Long
    Dim isStay As Boolean

    roomNo = "401"
    floorNo = 4
    isStay = True

    MsgBox "部屋: " & roomNo & vbCrLf & _
           "フロア: " & floorNo & "F" & vbCrLf & _
           "STAY: " & isStay
End Sub


' --- 例2: Rangeでセルの読み書き ---
Sub Range練習()
    ' 書き込み
    Range("A1").Value = "部屋番号"
    Range("B1").Value = "ステータス"
    Range("C1").Value = "フロア"

    Range("A2").Value = 401
    Range("B2").Value = "S"
    Range("C2").Value = "4F"

    Range("A3").Value = 502
    Range("B3").Value = "D"
    Range("C3").Value = "5F"

    ' 読み取り
    Dim status As String
    status = Range("B2").Value
    MsgBox "部屋401のステータス: " & status
End Sub


' --- 例3: Cellsでセルの読み書き ---
Sub Cells練習()
    ' Cells(行, 列) で書き込み
    Cells(1, 1).Value = "部屋番号"
    Cells(1, 2).Value = "C/I"
    Cells(2, 1).Value = 401
    Cells(2, 2).Value = "S"
    Cells(3, 1).Value = 502
    Cells(3, 2).Value = "D"

    ' 読み取り
    Dim roomNo As Variant
    Dim ci As String
    roomNo = Cells(2, 1).Value
    ci = CStr(Cells(2, 2).Value)
    MsgBox "部屋: " & roomNo & ", C/I: " & ci
End Sub


' --- 例4: シートを指定してセル操作 ---
Sub シート指定の練習()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ws.Range("A1").Value = "テスト"
    ws.Cells(2, 1).Value = "Cellsでも書けます"

    MsgBox "A1の値: " & ws.Range("A1").Value
End Sub


' --- 例5: 実務風 — セルから値を取得して表示 ---
Sub 部屋情報の取得()
    ' まずデータを準備
    Cells(1, 1).Value = "ROOM"
    Cells(1, 2).Value = "TYPE"
    Cells(1, 3).Value = "C/I"
    Cells(2, 1).Value = 401
    Cells(2, 2).Value = "S"
    Cells(2, 3).Value = "SD"
    Cells(3, 1).Value = 502
    Cells(3, 2).Value = "T"
    Cells(3, 3).Value = "D"

    ' 2行目のデータを取得
    Dim roomNo As Variant
    Dim roomType As String
    Dim ci As String

    roomNo = Cells(2, 1).Value
    roomType = CStr(Cells(2, 2).Value)
    ci = CStr(Cells(2, 3).Value)

    MsgBox "部屋: " & roomNo & vbCrLf & _
           "タイプ: " & roomType & vbCrLf & _
           "C/I: " & ci
End Sub
