VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTokenizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' Tokenizer - tokenize strings into parts
'

'
' Private options
'
Private SSpos As Long
Private SSfix As Long
Private SSlen As Long
Private SSnl As String
Private SStab As String
Private SSTokens(256) As String
Private SSNumTokens As Integer

' true Trim function with support for TAB's
Public Function StrTrim(tStr As String) As String
    Dim tS As String
    
    SSnl = Chr(13) & Chr(10)
    SStab = Chr(9)
    
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
          
' tokenize current line into parts, returns how many
Public Function Tokenize(tLine As String) As Integer
    Dim tPos, tLen, tFix As Long
    Dim tEle As String
    
    SSnl = Chr(13) & Chr(10)
    SStab = Chr(9)
    
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
        If (tEle = "\") Then
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

