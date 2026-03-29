Attribute VB_Name = "Module_指示書色付け転記"
Option Explicit

'=====================================================
' メイン処理
'=====================================================
Public Sub OpenFileAndColorRooms()

    Dim filePath As Variant
    Dim wbTarget As Workbook
    Dim ws4_9 As Worksheet
    Dim ws10_15 As Worksheet
    Dim wsShijisho As Worksheet

    filePath = Application.GetOpenFilename( _
                    Title:="【新】清掃指示書ファイルを選択してください")
    If filePath = False Then Exit Sub

    ThisWorkbook.Worksheets("指示書作成").Range("c5").Value = CStr(filePath)

    Set wbTarget = Workbooks.Open(CStr(filePath))

    On Error Resume Next
    Set ws4_9 = wbTarget.Worksheets("清掃指示書 (4-9)")
    Set ws10_15 = wbTarget.Worksheets("清掃指示書 (10-15)")
    Set wsShijisho = ThisWorkbook.Worksheets("指示書")
    On Error GoTo 0

    If wsShijisho Is Nothing Then
        MsgBox "マクロブックにシート「指示書」がありません。", vbExclamation
        Exit Sub
    End If

    ' 清掃指示書側 色付け
    If Not ws4_9 Is Nothing Then Call ColorRooms_AllBlocks(ws4_9)
    If Not ws10_15 Is Nothing Then Call ColorRooms_AllBlocks(ws10_15)

    ' 4F 転記
    If Not ws4_9 Is Nothing Then
        Call Copy4F_Block(ws4_9, wsShijisho)
    End If

    ' 5F～15F 転記
    Call CopyFloors5To15(ws4_9, ws10_15, wsShijisho)

    ' 5) D/S 集計を指示書に印字
    Call CountDSTotalsForAll(ws4_9, ws10_15)

    MsgBox "色付けと 4F～15F の転記が完了しました。", vbInformation

End Sub


'=====================================================
' 清掃指示書の ROOM/C-I 色付け
'=====================================================
Private Sub ColorRooms_AllBlocks(ByVal ws As Worksheet)

    Const HEADER_ROW As Long = 4
    Const FIRST_DATA_ROW As Long = 5
    Const LAST_ROW As Long = 100

    Dim col As Long, roomCol As Long, ciCol As Long
    Dim maxCol As Long, c As Long

    maxCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column

    For col = 1 To maxCol
        If ws.Cells(HEADER_ROW, col).Value = "ROOM" Then
            roomCol = col
            ciCol = 0

            For c = roomCol + 1 To roomCol + 6
                If ws.Cells(HEADER_ROW, c).Value = "C/I" Then
                    ciCol = c
                    Exit For
                End If
            Next c
            If ciCol = 0 Then ciCol = roomCol + 2

            Call ColorRooms_OneBlock(ws, FIRST_DATA_ROW, LAST_ROW, roomCol, ciCol)
        End If
    Next col

End Sub

Private Sub ColorRooms_OneBlock(ByVal ws As Worksheet, _
                                ByVal firstRow As Long, ByVal lastRow As Long, _
                                ByVal roomCol As Long, ByVal ciCol As Long)

    Dim r As Long
    Dim txt As String
    Dim c1 As String, c2 As String

    For r = firstRow To lastRow
        If IsEmpty(ws.Cells(r, roomCol)) And IsEmpty(ws.Cells(r + 1, roomCol)) Then Exit For

        txt = CStr(ws.Cells(r, ciCol).Value)
        If Len(txt) >= 1 Then c1 = Mid$(txt, 1, 1) Else c1 = ""
        If Len(txt) >= 2 Then c2 = Mid$(txt, 2, 1) Else c2 = ""

        With ws.Cells(r, roomCol)
            If c1 = "S" Or c2 = "S" Then
                .Interior.Color = RGB(0, 176, 240)       '青
            ElseIf c1 = "D" Or c2 = "D" Then
                .Interior.Color = RGB(255, 105, 180)     'ピンク
            End If
        End With
    Next r

End Sub


'=====================================================
' 4F 専用転記（清掃指示書(4-9) → 指示書 4F）
'=====================================================
Private Sub Copy4F_Block(ByVal wsSrc As Worksheet, ByVal wsDest As Worksheet)

    Const SRC_FIRST_DATA_ROW As Long = 5
    Const SRC_ROOM_COL As Long = 1
    Const SRC_CI_COL As Long = 3

    Const DEST_ROOM_COL As Long = 1
    Const DEST_WSET_COL As Long = 3
    Const DEST_CHK_COL As Long = 4
    Const DEST_FIRST_DATA_ROW As Long = 3
    Const DEST_LAST_DATA_ROW As Long = 19

    Dim lastRowSrc As Long
    Dim r As Long, destRow As Long
    Dim roomNo As Variant
    Dim ci1 As String
    Dim c1 As String, c2 As String
    Dim ci1Trim As String
    Dim s As String

    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, SRC_ROOM_COL).End(xlUp).Row
    destRow = DEST_FIRST_DATA_ROW

    For r = SRC_FIRST_DATA_ROW To lastRowSrc Step 2

        roomNo = wsSrc.Cells(r, SRC_ROOM_COL).Value
        If roomNo = "" Or IsError(roomNo) Then Exit For

        ci1 = CStr(wsSrc.Cells(r, SRC_CI_COL).Value)

        c1 = "": c2 = ""
        If Len(ci1) >= 1 Then c1 = Mid$(ci1, 1, 1)
        If Len(ci1) >= 2 Then c2 = Mid$(ci1, 2, 1)

        With wsDest.Cells(destRow, DEST_ROOM_COL)
            If c1 = "S" Or c2 = "S" Then
                .Interior.Color = RGB(0, 176, 240)
            ElseIf c1 = "D" Or c2 = "D" Then
                .Interior.Color = RGB(255, 105, 180)
            End If
        End With

        ci1Trim = LTrim$(ci1)
        If Left$(ci1Trim, 1) = ChrW(&H25CF) Then
            With wsDest.Cells(destRow, DEST_WSET_COL)
                .Value = "W"
                .Font.Bold = True
                .Font.Size = 30
            End With
        End If

        s = wsSrc.Cells(r + 1, SRC_CI_COL).Text
        If InStr(s, ":") > 0 Or InStr(s, ChrW(&HFF1A)) > 0 Then
            With wsDest.Cells(destRow, DEST_CHK_COL)
                .Value = s
                .Font.Size = 26
                .HorizontalAlignment = xlLeft
            End With
        End If

        destRow = destRow + 1
        If destRow > DEST_LAST_DATA_ROW Then Exit For

    Next r

End Sub


'=====================================================
' 5F～15F をまとめて転記
'=====================================================
Private Sub CopyFloors5To15(ByVal ws4_9 As Worksheet, _
                            ByVal ws10_15 As Worksheet, _
                            ByVal wsDest As Worksheet)

    If Not ws4_9 Is Nothing Then
        Call CopyFloorsFromSheet(ws4_9, wsDest, _
                                 Array(5, 6, 7, 8, 9), _
                                 Array(2, 3, 4, 5, 6))
    End If

    If Not ws10_15 Is Nothing Then
        Call CopyFloorsFromSheet(ws10_15, wsDest, _
                                 Array(10, 11, 12, 13, 14, 15), _
                                 Array(1, 2, 3, 4, 5, 6))
    End If

End Sub


'=====================================================
' 1つの清掃指示書シートから複数フロアを転記
'=====================================================
Private Sub CopyFloorsFromSheet(ByVal wsSrc As Worksheet, _
                                ByVal wsDest As Worksheet, _
                                ByVal floorList As Variant, _
                                ByVal blockIndexList As Variant)

    Const HEADER_ROW As Long = 4
    Const FIRST_DATA_ROW As Long = 5

    Dim roomCols(1 To 6) As Long
    Dim ciCols(1 To 6) As Long
    Dim countBlocks As Long
    Dim i As Long
    Dim floorNo As Long
    Dim blockIdx As Long

    Call GetRoomBlocks(wsSrc, HEADER_ROW, roomCols, ciCols, countBlocks)

    For i = LBound(floorList) To UBound(floorList)
        floorNo = floorList(i)
        blockIdx = blockIndexList(i)

        If blockIdx >= 1 And blockIdx <= countBlocks Then
            Call CopyFloorGeneric(wsSrc, _
                                  roomCols(blockIdx), ciCols(blockIdx), _
                                  FIRST_DATA_ROW, wsDest, floorNo)
        End If
    Next i

End Sub


'=====================================================
' 清掃指示書の「ROOM」ブロックを検出
'=====================================================
Private Sub GetRoomBlocks(ByVal ws As Worksheet, ByVal headerRow As Long, _
                          ByRef roomCols() As Long, _
                          ByRef ciCols() As Long, _
                          ByRef countBlocks As Long)

    Dim col As Long, c As Long
    Dim maxCol As Long

    countBlocks = 0
    maxCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column

    For col = 1 To maxCol
        If ws.Cells(headerRow, col).Value = "ROOM" Then
            countBlocks = countBlocks + 1
            roomCols(countBlocks) = col

            ciCols(countBlocks) = 0
            For c = col + 1 To col + 6
                If ws.Cells(headerRow, c).Value = "C/I" Then
                    ciCols(countBlocks) = c
                    Exit For
                End If
            Next c
            If ciCols(countBlocks) = 0 Then ciCols(countBlocks) = col + 2

            If countBlocks = 6 Then Exit For
        End If
    Next col

End Sub


'=====================================================
' フロア番号 → 指示書シート上の転記先位置
'=====================================================
Private Sub GetDestInfo(ByVal floorNo As Long, _
                        ByRef roomColDest As Long, _
                        ByRef firstRowDest As Long, _
                        ByRef lastRowDest As Long)

    Select Case floorNo
        Case 4:  roomColDest = 1:  firstRowDest = 3:  lastRowDest = 19
        Case 5:  roomColDest = 1:  firstRowDest = 32: lastRowDest = 48
        Case 6:  roomColDest = 8:  firstRowDest = 3:  lastRowDest = 19
        Case 7:  roomColDest = 8:  firstRowDest = 32: lastRowDest = 48
        Case 8:  roomColDest = 15: firstRowDest = 3:  lastRowDest = 19
        Case 9:  roomColDest = 15: firstRowDest = 32: lastRowDest = 48
        Case 10: roomColDest = 22: firstRowDest = 3:  lastRowDest = 19
        Case 11: roomColDest = 22: firstRowDest = 32: lastRowDest = 48
        Case 12: roomColDest = 29: firstRowDest = 3:  lastRowDest = 19
        Case 13: roomColDest = 29: firstRowDest = 32: lastRowDest = 49
        Case 14: roomColDest = 36: firstRowDest = 3:  lastRowDest = 19
        Case 15: roomColDest = 36: firstRowDest = 32: lastRowDest = 48
    End Select

End Sub


'=====================================================
' 1フロア分の転記（5F～15F）
'=====================================================
Private Sub CopyFloorGeneric(ByVal wsSrc As Worksheet, _
                             ByVal roomColSrc As Long, _
                             ByVal ciColSrc As Long, _
                             ByVal firstRowSrc As Long, _
                             ByVal wsDest As Worksheet, _
                             ByVal floorNo As Long)

    Dim lastRowSrc As Long
    Dim r As Long
    Dim roomNo As Variant
    Dim ci1 As String
    Dim c1 As String, c2 As String

    Dim roomColDest As Long
    Dim firstRowDest As Long, lastRowDest As Long
    Dim m As Variant, destRow As Long

    Dim ci1Trim As String
    Dim s As String

    Call GetDestInfo(floorNo, roomColDest, firstRowDest, lastRowDest)

    lastRowSrc = wsSrc.Cells(wsSrc.Rows.Count, roomColSrc).End(xlUp).Row

    For r = firstRowSrc To lastRowSrc Step 2

        roomNo = wsSrc.Cells(r, roomColSrc).Value
        If roomNo = "" Or IsError(roomNo) Then Exit For

        ci1 = CStr(wsSrc.Cells(r, ciColSrc).Value)

        m = Application.Match(roomNo, wsDest.Range(wsDest.Cells(firstRowDest, roomColDest), _
                                                   wsDest.Cells(lastRowDest, roomColDest)), 0)
        If IsError(m) Then GoTo NextRoom

        destRow = firstRowDest + CLng(m) - 1

        c1 = "": c2 = ""
        If Len(ci1) >= 1 Then c1 = Mid$(ci1, 1, 1)
        If Len(ci1) >= 2 Then c2 = Mid$(ci1, 2, 1)

        With wsDest.Cells(destRow, roomColDest)
            If c1 = "S" Or c2 = "S" Then
                .Interior.Color = RGB(0, 176, 240)
            ElseIf c1 = "D" Or c2 = "D" Then
                .Interior.Color = RGB(255, 105, 180)
            End If
        End With

        ci1Trim = LTrim$(ci1)
        If Left$(ci1Trim, 1) = ChrW(&H25CF) Then
            With wsDest.Cells(destRow, roomColDest + 2)
                .Value = "W"
                .Font.Bold = True
                .Font.Size = 30
            End With
        End If

        s = wsSrc.Cells(r + 1, ciColSrc).Text
        If InStr(s, ":") > 0 Or InStr(s, ChrW(&HFF1A)) > 0 Then
            With wsDest.Cells(destRow, roomColDest + 3)
                .Value = s
                .Font.Size = 26
                .HorizontalAlignment = xlLeft
            End With
        End If

NextRoom:
    Next r

End Sub


'=====================================================
' D/S 集計
'=====================================================
Private Sub CountDSTotalsForAll(ByVal ws4_9 As Worksheet, _
                                ByVal ws10_15 As Worksheet)

    If Not ws4_9 Is Nothing Then Call CountDSTotalsOnSourceSheet(ws4_9)
    If Not ws10_15 Is Nothing Then Call CountDSTotalsOnSourceSheet(ws10_15)

End Sub

Private Sub CountDSTotalsOnSourceSheet(ByVal wsSrc As Worksheet)

    Const HEADER_ROW As Long = 4
    Const FIRST_DATA_ROW As Long = 5

    Dim roomCols(1 To 6) As Long
    Dim ciCols(1 To 6) As Long
    Dim countBlocks As Long
    Dim i As Long
    Dim lastRow As Long
    Dim r As Long
    Dim ci1 As String
    Dim c1 As String, c2 As String
    Dim dCount As Long, sCount As Long
    Dim addr As String

    Call GetRoomBlocks(wsSrc, HEADER_ROW, roomCols, ciCols, countBlocks)

    For i = 1 To countBlocks

        dCount = 0
        sCount = 0

        lastRow = wsSrc.Cells(wsSrc.Rows.Count, roomCols(i)).End(xlUp).Row

        For r = FIRST_DATA_ROW To lastRow Step 2
            ci1 = CStr(wsSrc.Cells(r, ciCols(i)).Value)
            If Len(ci1) >= 1 Then c1 = Mid$(ci1, 1, 1) Else c1 = ""
            If Len(ci1) >= 2 Then c2 = Mid$(ci1, 2, 1) Else c2 = ""
            If c1 = "D" Or c2 = "D" Then dCount = dCount + 1
            If c1 = "S" Or c2 = "S" Then sCount = sCount + 1
        Next r

        Select Case i
            Case 1: addr = "B2"
            Case 2: addr = "F2"
            Case 3: addr = "J2"
            Case 4: addr = "N2"
            Case 5: addr = "R2"
            Case 6: addr = "V2"
            Case Else: addr = ""
        End Select

        If addr <> "" Then
            wsSrc.Range(addr).Value = "D=" & dCount & ", S=" & sCount
        End If

    Next i

End Sub
