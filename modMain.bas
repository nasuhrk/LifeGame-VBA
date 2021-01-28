Option Explicit

Public aaaa1(20, 20) As Boolean
Public aaaa2(20, 20) As Boolean

Sub TestStart()
    Call read
'    MsgBox (aaaa1(0, 0))
'    aaaa1(1, 2) = True
'    aaaa1(0, 1) = True
    Call mover
    Call outprint
End Sub

Function mover()
   Dim x As Byte
   Dim y As Byte
    
    For x = 0 To 17
        For y = 0 To 17
          Call Macro3(x, y)
        Next y
    Next x

End Function

Function Macro2()
   Dim x As Byte
   Dim y As Byte
    
    For x = 1 To 20
        For y = 1 To 20
          Call Macro(x, y, False)
        Next y
    Next x

End Function

Function read()
   Dim x As Byte
   Dim y As Byte
    
    For x = 1 To 20
        For y = 1 To 20
            If Cells(x, y) <> "" Then
                aaaa1(x - 1, y - 1) = 1
            Else
                aaaa1(x - 1, y - 1) = 0
            End If
        Next y
    Next x

End Function

Function outprint()
   Dim x As Byte
   Dim y As Byte
    
    For x = 1 To 20
        For y = 1 To 20
            If aaaa2(x - 1, y - 1) = True Then
                Cells(x, y) = "■"
            Else
                Cells(x, y) = ""
            End If
        Next y
    Next x
End Function

Function reset2()
   Dim x As Byte
   Dim y As Byte
    
    For x = 1 To 20
        For y = 1 To 20
          aaaa2(x, y) = False
        Next y
    Next x

End Function

Function Macro(x As Byte, y As Byte, flag As Boolean) As Boolean
    
    If (flag) Then
        Cells(x, y).Value = "■"
    Else
        Cells(x, y).Value = ""
    End If
    
End Function

Function Macro3(x As Byte, y As Byte)
        
    Dim count As Byte
    
    '上段(左)
    If aaaa1(x, y) = True Then
        count = count + 1
    End If
    '上段(中央)
    If aaaa1(x + 1, y) = True Then
        count = count + 1
    End If
    '上段(右)
    If aaaa1(x + 2, y) = True Then
        count = count + 1
    End If
    '中段(左)
    If aaaa1(x, y + 1) = True Then
        count = count + 1
    End If
    '中段(右)
    If aaaa1(x + 2, y + 1) = True Then
        count = count + 1
    End If
    '下段(左)
    If aaaa1(x, y + 2) = True Then
        count = count + 1
    End If
    '下段(中央)
    If aaaa1(x + 1, y + 2) = True Then
        count = count + 1
    End If
    '下段(右)
    If aaaa1(x + 2, y + 2) = True Then
        count = count + 1
    End If
    
'    MsgBox count
'
    Dim flag As Boolean
    '中段(中央)
    If aaaa1(x + 1, y + 1) = False Then
        If count = 3 Then
            '(1)誕生
            flag = True
'        Else
'            '(5)その他
'            flag = False
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
    
'    MsgBox flag
    aaaa2(x + 1, y + 1) = flag

End Function

