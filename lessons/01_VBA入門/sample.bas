Attribute VB_Name = "Lesson01_VBA入門"
Option Explicit

'=====================================================
' レッスン1: VBA入門 - サンプルコード
' このファイルのコードをExcel VBEにコピーして実行してください
'=====================================================

' --- 例1: 最初のマクロ ---
Sub はじめてのマクロ()
    MsgBox "こんにちは！VBAの世界へようこそ！"
End Sub


' --- 例2: アイコン付きメッセージ ---
Sub あいさつ()
    MsgBox "本日の清掃作業を開始します！", vbInformation, "確認"
End Sub


' --- 例3: 警告メッセージ ---
Sub 警告テスト()
    MsgBox "マクロブックにシート「指示書」がありません。", vbExclamation, "エラー"
End Sub


' --- 例4: はい/いいえの確認 ---
Sub 確認テスト()
    Dim answer As VbMsgBoxResult
    answer = MsgBox("処理を続けますか？", vbYesNo + vbQuestion, "確認")

    If answer = vbYes Then
        MsgBox "処理を開始します。", vbInformation
    Else
        MsgBox "中止しました。"
        Exit Sub
    End If

    MsgBox "処理が完了しました！", vbInformation
End Sub


' --- 例5: 実務マクロ風の完了メッセージ ---
Sub 完了メッセージ()
    MsgBox "色付けと 4F～15F の転記が完了しました。", vbInformation
End Sub
