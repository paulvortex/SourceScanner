VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSSparser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' SS targ parser by Pavel P. [VorteX] Timofeyev
'

'
' Private properties
'

Private SSpos As Long
Private SSeqpos As Long
Private SSfix As Long
Private SSlen As Long
Private SSformat As String
Private SScont As String
Private SSline As String
Private SSkey As String
Private SSval As String
Private SSnl As String
Private SStab As String
Private SSTokens(256) As String
Private SSNumTokens As Integer
Public KeepTrims As Boolean

'
' Public properties
'


'
' functions & Properties
'

Public Property Get Format() As String
 Format = SSformat
End Property

Public Property Get Content() As String
 Content = SScont
End Property

Public Property Get CurLine() As String
 CurLine = SSline
End Property

Public Property Get CurKey() As String
 CurKey = SSkey
End Property

Public Property Get CurVal() As String
 CurVal = SSval
End Property

Public Sub Load(tContent As String, tType As String)
    SScont = Trim$(tContent)
    SSformat = tType
    
    ' flush temps
    SSline = ""
    SSkey = ""
    SSval = ""
    SSpos = 1
    SSfix = 1
    SSeqpos = 0
    SSlen = Len(SScont)
    SSnl = Chr(13) & Chr(10)
    SStab = Chr(9)
    
    ' hack: skip first line if found "/*SS"
    If (Left$(SScont, 4) = "/*SS") Then
        Do While (SSpos < SSlen)
            If (Mid$(SScont, SSpos, 2) = SSnl) Then
                SSpos = SSpos + 2
                SSfix = SSpos
                Exit Do
            End If
            SSpos = SSpos + 1
        Loop
    End If
End Sub

Private Sub SkipTrailingTrims()
    SSpos = SSfix
    Do While (SSfix < SSlen)
        If (Mid$(SScont, SSpos, 2) = SSnl) Then SSpos = SSpos + 2
        If (Mid$(SScont, SSpos, 1) = " ") Then SSpos = SSpos + 1
        If (Mid$(SScont, SSpos, 1) = SStab) Then SSpos = SSpos + 1
        If (SSpos = SSfix) Then Exit Do
        SSfix = SSpos
    Loop
End Sub

' true Trim function with support for TAB's
Private Function StrTrim(tStr As String) As String
    Dim tS As String
    
    ' left trim
    tS = Left$(tStr, 1)
    Do While (tS = " " Or tS = SStab)
        tStr = Mid$(tStr, 2, 10000)
        tS = Left$(tStr, 1)
    Loop
    
    ' right trim
    tS = Right$(tStr, 1)
    Do While (tS = " " Or tS = SStab)
        tStr = Left$(tStr, Len(tStr) - 1)
        tS = Right$(tStr, 1)
    Loop
    
    StrTrim = tStr
End Function

Public Function ReadLine(Optional tKeepTrims As Boolean) As Boolean
    Dim tPos, tLen As Long
    Dim tEle As String
    
    SSkey = ""
    SSval = ""
    SSline = ""
    
    ' early exit
    If (SSpos >= SSlen) Then
        ReadLine = False
        Exit Function
    End If
    
    ' get line
    ReadLine = True
    Do While (SSpos < SSlen)
        tEle = Mid$(SScont, SSpos, 1)
        If ((SSpos + 1) >= SSlen) Then
            tEle = ";"
            SSpos = SSpos + 1
        End If
        If (Mid$(SScont, SSpos, 2) = SSnl) Then
            If (KeepTrims = False And tKeepNewLines = False) Then
                SSline = SSline & StrTrim(Mid$(SScont, SSfix, SSpos - SSfix)) & " "
                SSfix = SSpos
                SkipTrailingTrims
                ' check for #list# in the beginning if line
                If (Mid$(SScont, SSpos, 6) = "#list#") Then
                    SSline = SSline & SSnl & SStab
                    SSfix = SSfix + 6
                    SSpos = SSpos + 6
                End If
            End If
        ElseIf (tEle = ";") Then
            SSline = SSline & StrTrim(Mid$(SScont, SSfix, SSpos - SSfix))
            ' check if it is a key=parm string
            tLen = Len(SSline)
            SSeqpos = 1
            Do While (SSeqpos <= tLen)
                If (Mid$(SSline, SSeqpos, 1) = "=") Then Exit Do
                SSeqpos = SSeqpos + 1
            Loop
            If (SSeqpos <= tLen) Then
                SSkey = StrTrim(LCase$(Left$(SSline, SSeqpos - 1)))
                SSval = StrTrim(Right$(SSline, tLen - SSeqpos))
                'MsgBox "SSkey=" & SSkey & SSnl & "SSval=" & SSval
            End If
            SSfix = SSpos + 1
            ' skip newlines, spacers and tabs following ;
            SkipTrailingTrims
            Exit Do
        End If
        If (tEle = "=") Then SSeqpos = SSpos + 1
        SSpos = SSpos + 1
    Loop
End Function


  ' check/skip format key
  '              If (SSkey = "form") Then
  '                  If (SSval <> SSformat) Then
  '                      ReadLine = False ' wrong format
  '                      Exit Function
  '                  End If
  '                  SSfix = SSpos + 1
  '                  SkipTrailingTrims
  '                  ReadLine = ReadLine()
  '                  Exit Function
  '              End If
                
' tokenize current line into parts, returns how many
Public Function Tokenize(tLine As String) As Integer
    Dim tPos, tLen, tFix As Long
    Dim tEle As String
    
    tFix = 1
    tPos = 1
    SSNumTokens = 0
    tLen = Len(tLine)
    Do While (tPos <= tLen)
        tEle = Mid$(tLine, tPos, 1)
        If ((tPos + 1) > tLen) Then
            tPos = tPos + 1
            tEle = "\"
        End If
        If (tEle = "\" Or tPos + 1 > tLen) Then
            SSTokens(SSNumTokens) = StrTrim(Mid$(tLine, tFix, tPos - tFix))
            SSNumTokens = SSNumTokens + 1
            tFix = tPos + 1
        End If
        tPos = tPos + 1
    Loop
    
    Tokenize = SSNumTokens
End Function

' tokenize current line into parts with given separator, returns how many
Public Function TokenizeWithSeparator(tLine As String, tSep As String) As Integer
    Dim tPos, tLen, tFix As Long
    Dim tSepLen As Integer
    Dim tEle As String
    
    tFix = 1
    tPos = 1
    tSepLen = Len(tSep)
    SSNumTokens = 0
    tLen = Len(tLine)
      Do While (tPos <= tLen)
        tEle = Mid$(tLine, tPos, tSepLen)
        If ((tPos + 1) > tLen) Then
            tPos = tPos + 1
            tEle = tSep
        End If
        If (tEle = tSep) Then
            SSTokens(SSNumTokens) = StrTrim(Mid$(tLine, tFix, tPos - tFix))
            SSNumTokens = SSNumTokens + 1
            tFix = tPos + tSepLen
        End If
        tPos = tPos + 1
    Loop
    TokenizeWithSeparator = SSNumTokens
End Function

' returns numbered arguments
Property Get Argv(ByVal tIndex As Integer) As String
    Argv = ""
    If (tIndex >= 0 And tIndex < SSNumTokens) Then Argv = SSTokens(tIndex)
End Property
