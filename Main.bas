Attribute VB_Name = "Main"
'
' Various Utils
'

Public Sub SwitchForm(tCurrent As Form, tNew As Form)
    tCurrent.Hide
    tNew.Show
End Sub

' powers
Public Function Power(tValue As Variant, ByVal tPower As Integer) As Variant
    Power = tValue
    For tPos = 1 To tPower
        Power = Power * tValue
    Next tPos
End Function

' bound integer
Public Function BoundInt(tMin As Integer, tCur As Variant, tMax As Integer) As Integer
    BoundInt = Fix(tCur)
    If (tCur < tMin) Then BoundInt = tMin
    If (tCur > tMax) Then BoundInt = tMax
End Function

' string to int conversion
Public Function StrToSingle(tStr As String) As Single
    Dim tVar As Variant
    
    On Error GoTo Error
     
    If (IsNull(tStr) Or tStr = "") Then
        StrToSingle = 0
        Exit Function
    End If
    
    ' replace . to , in string
    For tPos = 1 To Len(tStr)
        If (Mid$(tStr, tPos, 1) = ".") Then
            tVar = tVar & ","
        Else
            tVar = tVar & Mid$(tStr, tPos, 1)
        End If
    Next tPos
    
    StrToSingle = tVar
Quit:
    Exit Function
Error:
    StrToSingle = 0
    Resume Quit
End Function

Public Function StrToInteger(tStr As String) As Integer
    Dim tVar As Variant
    
    On Error GoTo Error
    
    If (IsNull(tStr) Or tStr = "") Then
        StrToInteger = 0
        Exit Function
    End If
    
    tVar = tStr
    StrToInteger = tVar
    StrToInteger = Fix(StrToInteger)
Quit:
    Exit Function
Error:
    StrToInteger = 0
    Resume Quit
End Function

