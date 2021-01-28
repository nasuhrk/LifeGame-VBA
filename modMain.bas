Option Explicit

Const Sign As String = "■" '使用文字
Const START_X = 2 '開始X座標
Const START_Y = 2 '開始Y座標
Const MAX_X = 20 '最大X座標
Const MAX_Y = 20 '最大Y座標

Private board(MAX_X, MAX_Y) As Boolean
Private board_next(MAX_X, MAX_Y) As Boolean

Dim x As Byte
Dim y As Byte
Dim count As Byte

' ============================================================
'  [readSheet]
' ============================================================
Private Sub readSheet()
    For x = START_X To START_X + MAX_X - 1
        For y = START_Y To START_Y + MAX_Y - 1
            If Cells(x, y) <> "" Then
                board(x - 1, y - 1) = 1
            Else
                board(x - 1, y - 1) = 0
            End If
        Next y
    Next x
End Sub

' ============================================================
'  [writeSheet]
' ============================================================
Private Sub writeSheet()
    Dim flag As Boolean
    For x = START_X To START_X + MAX_X - 1
        For y = START_Y To START_Y + MAX_Y - 1
            flag = board_next(x - 1, y - 1)
            Call writeSign(x, y, flag)
        Next y
    Next x
End Sub

' ============================================================
'  [getSheetValue]
' ============================================================
Private Sub searchBoard()
    For x = 0 To MAX_X - 3
        For y = 0 To MAX_Y - 3
          Call createNewBoard(x, y)
        Next y
    Next x
End Sub

' ============================================================
'  [createNewBoard]
' ============================================================
Private Function createNewBoard(x As Byte, y As Byte)
   
    '初期化
    count = 0

    Call setCount(x, y)         '上段(左)
    Call setCount(x + 1, y)     '上段(中央)
    Call setCount(x + 2, y)     '上段(右)
    Call setCount(x, y + 1)     '中段(左)
    Call setCount(x + 2, y + 1) '中段(右)
    Call setCount(x, y + 2)     '下段(左)
    Call setCount(x + 1, y + 2) '下段(中央)
    Call setCount(x + 2, y + 2) '下段(右)

    '次の板に書込み
    board_next(x + 1, y + 1) = getRule(x, y) '中段(中央)

End Function

' ============================================================
'  [setCount]
' ============================================================
Private Function setCount(x As Byte, y As Byte)
    If board(x, y) = True Then
        count = count + 1
    End If
End Function

' ============================================================
'  [setCount]
' ============================================================
Private Function getRule(x As Byte, y As Byte) As Boolean
    Dim flag As Boolean
    
    If board(x + 1, y + 1) = False Then
        If count = 3 Then
            '(1)誕生
            flag = True
        End If
    Else
        If count = 2 Or count = 3 Then
            '(2)生存
            flag = True
        End If
        If count <= 1 Then
            '(3)過疎→死滅
            flag = False
        End If
        If 4 <= count Then
            '(4)過密→死滅
            flag = False
        End If
    End If
    
    '戻り値
    getRule = flag

End Function

' ============================================================
'  [LifeGame_START]
' ============================================================
Public Sub LifeGame_START()
   
    Dim i As Byte
    For i = 1 To 10
        
        'Sheetから読込み
        Call readSheet
        
        '板を検索
        Call searchBoard
    
        'Sheetへ書込み
        Call writeSheet
    
        'X秒待機
        Call wait
    
        '終了判定
'        If chackGameStatus Then
'            MsgBox ("end game")
'        End If
    Next i

End Sub

' ============================================================
'  [wait]
' ============================================================
Private Function wait()
    Application.wait [Now() + "0:00:00.1"]
End Function

' ============================================================
'  [setRandomWarm]
' ============================================================
Private Sub setRandomWarm()
    Randomize '乱数系列初期化
   
    For x = START_X To START_X + MAX_X - 1
        For y = START_Y To START_Y + MAX_Y - 1
            '1～2で乱数発生
            If Int(2 * Rnd) = 1 Then
                Call writeSign(x, y, True)
            Else
                Call writeSign(x, y, False)
            End If
        Next y
    Next x
End Sub

' ============================================================
'  [chackGameStatus]
' ============================================================
'Private Function chackGameStatus()
'    Dim flag As Boolean
'    Dim cnt As Integer
'    cnt = 0
'    For x = 0 To MAX_X
'        For y = 0 To MAX_Y
'            If board(x, y) <> board_next(x, y) Then
'                cnt = cnt + 1
'                chackGameStatus = False
'            End If
'        Next y
'    Next x
'
''    MsgBox cnt
'    '終了
'    chackGameStatus = True
'End Function

' ============================================================
'  [writeSign]
' ============================================================
Private Function writeSign(x As Byte, y As Byte, flag As Boolean)
    Cells(x, y) = ""
    If flag Then
        Cells(x, y) = Sign
    End If
End Function
