VERSION 5.00
Begin VB.Form frmScan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Source Scanner"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9720
   Icon            =   "frmScan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExit 
      Caption         =   "&Exit"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   11
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton btnRescan 
      Caption         =   "&Rescan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton btnOpenDestDir 
      Caption         =   "&Open destination path"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton btnStop 
      Caption         =   "&Stop"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   7
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Timer RunTimer 
      Enabled         =   0   'False
      Left            =   120
      Top             =   5280
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "frmScan.frx":08CA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   10680
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtOutput 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmScan.frx":1194
      Top             =   840
      Width           =   9450
   End
   Begin VB.CommandButton btnPrev 
      Caption         =   "<< &Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame3"
      Height          =   855
      Left            =   -720
      TabIndex        =   4
      Top             =   -120
      Width           =   12855
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Scanning progress"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   7815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   975
      Left            =   -960
      TabIndex        =   6
      Top             =   5760
      Width           =   12735
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Extraction
'

Const exstateFindFiles = 0
Const exstateExtract = 1

Const outputSingle = 0
Const outputByFileName = 1
Const outputBySSGroup = 2

' extract formats
Const exformatEntityDef = "quaked"

' generic thing
Private INI As clsINI
Private Extract_StartTime As Double
Private Extract_LastFile As String ' last extracted file, needed by GtkRad15Ent exporter in separate file mode to write </classes> at the end of file
Private Extract_InfoMessages As Boolean
Private Extract_ExtMessages As Boolean
Private Extract_DevMessages As Boolean
Private Extract_Ext As String
Private Extract_SrcPath As String
Private Extract_DstPath As String
Private Extract_CurFile As Integer
Private Extract_Format As String
Private Extract_Language As String
Private Extract_Q3RadDef As Boolean
Private Extract_GtkRad15Ent As Boolean
Private Extract_WorldCraft33Fgd As Boolean
Private Extract_WriteSeparateFiles As Integer
Private Extract_UseSSTags As Boolean
Private Extract_UseSSTags_Counter As Integer
Private Extract_UseSSTags_SkipFormCounter As Integer
Private Extract_UseSSTags_SkipLangCounter As Integer
Private Extract_UseSSTags_NumErrors As Integer
Private Extract_SSTemplatesLoaded As Boolean
Private Extract_Q3RadDef_Counter As Integer
Private Extract_GtkRad15Ent_Counter As Integer
Private Extract_WorldCraft33Fgd_Counter As Integer
Private Extract_SSFixTargetFields As Boolean
Private Extract_SSFixTargetFieldsMethod As Integer
Private Extract_State As Integer ' 0 - find files, 1 - extract
Private Extract_CustomTemplates As String

' used to know where template files can be
Private Extract_TemplateFile1 As String
Private Extract_TemplateFile2 As String

' Known EntityDefKey.Type

' EntityDef struct
Private tDef As EntityDef
Private Type EntityDefFlag
    Key As String
    Name As String
    Description As String
End Type
Private Type EntityDefKey
    Name As String          ' key name
    Caption As String       ' key caption
    Description As String   ' key description scrings
    DefValue As String      ' key default value
    Type As String          ' key data type, string
    ListItems As String     ' tokenstring of list items
    IsDuplicate As Boolean  ' (used to duplicated keys with types that Radiant 1.5 does not shown in Entity Inspector)
End Type
Private Type EntityDef
    ClassName As String     ' classname, eg func_button
    Format As String        ' format
    Lang As String          ' lang
    Name As String
    Description As String   ' description
    Notes As String
    MinSize As String   ' bound box
    MaxSize As String   ' bound box
    Color As String
    Group As String     ' group
    EditorModel As String ' editor model
    SkillFlags As Integer   ' skill flags
    flags(7) As EntityDefFlag
    NumFlags As Integer
    Keys(64) As EntityDefKey
    NumKeys As Integer
End Type

' Template structs
Private Type Define
    Name As String
    Value As String
End Type
Private Type Group
    Name As String
    Color As String
    MinSize As String
    MaxSize As String
    Notes As String
    EditorModel As String
    SFlag As Integer
    NumKeys As Integer
    Keys(64) As String
    flags(7) As String
    NumFlags As Integer
End Type
Private Type Templates
    NumDefs As Integer
    NumGroups As Integer
    EditorModelsPath As String
    Defs(256) As Define
    Groups(64) As Group
End Type
Private SSTemplates As Templates
Private SSTokenizer As New clsTokenizer

'''''''''''''''''''''''''
' Utils
'''''''''''''''''''''''''

Function ClearOutput()
    Me!txtOutput.tExt = ""
End Function

Sub AddOutput(aStr As String, Optional tForce As Boolean)
    If (IsMissing(tForce) = False And tForce = True) Then
        GoTo Force
    End If
    If (Extract_InfoMessages = False) Then
        If (Left$(aStr, 5) = "     ") Then
            Exit Sub
        End If
    End If
Force:
    Me!txtOutput.tExt = Me!txtOutput.tExt & aStr & Chr$(13) & Chr$(10)
    Me!txtOutput.SelStart = Len(Me!txtOutput.tExt) - 5
    Me!txtOutput.SelLength = 1
    Me!txtOutput.Refresh
End Sub

Sub CheckUseOfTemplates()
    If (Extract_SSTemplatesLoaded = False) Then
        AddOutput "Error: define or template is used but template file not found", True
        AddOutput "with current path settings template file can have this paths: ", True
        AddOutput "    " & Extract_TemplateFile1, True
        AddOutput "    " & Extract_TemplateFile2, True
        AddOutput "Review your path settings and/or check template files", True
        btnStop_Click
    End If
End Sub

' scan line for possible defines and replace them


'''''''''''''''''''''''''
' Run_Extract
'''''''''''''''''''''''''

Function ExtractGetOutputFileName(tExt As String, Optional tFileName As String, Optional tSSGroupName As String) As String
    If (tSSGroupName = "") Then tSSGroupName = "nogroup"
    If (tFileName = "") Then tFileName = "entities"
    
    If (Extract_WriteSeparateFiles = outputByFileName) Then
        ExtractGetOutputFileName = Extract_DstPath & "\" & FileName_StripPath(FileName_StripExt(FileName_StripExt(tFileName))) & tExt
    ElseIf (Extract_WriteSeparateFiles = outputBySSGroup) Then
        ExtractGetOutputFileName = Extract_DstPath & "\" & tSSGroupName & tExt
    Else
        ExtractGetOutputFileName = Extract_DstPath & "\" & "entities" & tExt
    End If
End Function

Sub Run_Extract()
    Dim tStr As String
    
    ' fix start time
    Extract_StartTime = Timer
    
    ' freeze buttons
    Me.btnStop.Enabled = True
    Me.btnCancel.Enabled = True
    Me.btnOpenDestDir.Enabled = False
    Me.btnPrev.Enabled = False
    Me.btnExit.Enabled = False
    Me.btnRescan.Enabled = False
    
    ' read main options
    Extract_Ext = INI.Read("Settings", "TargetExtensions")
    Extract_SrcPath = INI.Read("Settings", "TargetPath")
    Extract_DstPath = INI.Read("Settings", "DestinationPath")
    Extract_CustomTemplates = INI.Read("Settings", "CustomConfig")
    Extract_Language = INI.Read("Settings", "CustomLanguage")
    
    ' know if we need to display info messages
    If (StrToInteger(INI.Read("Settings", "InfoOutput")) <> 0) Then
        Extract_InfoMessages = True
    Else
        Extract_InfoMessages = False
    End If

    Call ClearOutput
    AddOutput (" Initializing...")
    AddOutput ("     target: " & Extract_SrcPath)
    AddOutput ("     destination: " & Extract_DstPath)
    AddOutput ("     mask: " & Extract_Ext & "")
    If (Extract_Ext = "") Then
        AddOutput (" Error: empty extension field")
        Exit Sub
    End If
    If (Dir(Extract_SrcPath, vbDirectory) = "") Then
        AddOutput (" Error: target path not found")
        Exit Sub
    End If
    If (Dir(Extract_DstPath, vbDirectory) = "") Then
        AddOutput (" Error: destination path not found")
        Exit Sub
    End If
    
    ' set messages level
    tStr = "     logs: 'minimal'"
    If (Extract_InfoMessages = True) Then
        tStr = tStr & "'info' "
    End If
    If (StrToInteger(INI.Read("Settings", "FullOutput")) <> 0) Then
        tStr = tStr & "'extended' "
        Extract_ExtMessages = True
    Else
        Extract_ExtMessages = False
    End If
    If (StrToInteger(INI.Read("Settings", "DevOutput")) <> 0) Then
        tStr = tStr & "'developer'"
        Extract_DevMessages = True
    Else
        Extract_DevMessages = False
    End If
    AddOutput (tStr)
    
    ' make extraction formats
    tStr = "     extraction formats: "
    If (StrToInteger(INI.Read("Extract", "Q3RadDef")) <> 0) Then
        tStr = tStr & "'def' "
        Extract_Format = exformatEntityDef
        Extract_Q3RadDef = True
    Else
        Extract_Q3RadDef = False
    End If
    If (StrToInteger(INI.Read("Extract", "GtkRad150Ent")) <> 0) Then
        tStr = tStr & "'ent'"
        Extract_Format = exformatEntityDef
        Extract_GtkRad15Ent = True
    Else
        Extract_GtkRad15Ent = False
    End If
    If (StrToInteger(INI.Read("Extract", "WorldCraft33Fgd")) <> 0) Then
        tStr = tStr & "'fgd'"
        Extract_Format = exformatEntityDef
        Extract_WorldCraft33Fgd = True
    Else
        Extract_WorldCraft33Fgd = False
    End If
    AddOutput (tStr)
    
    ' make extraction options
    tStr = "     extraction options: "
    If (StrToInteger(INI.Read("Extract", "UseSSTags")) <> 0) Then
        tStr = tStr & "'ss tags' "
        Extract_UseSSTags = True
    Else
        Extract_UseSSTags = False
    End If
    Extract_WriteSeparateFiles = StrToInteger(INI.Read("Extract", "WriteSeparateFiles"))
    If (Extract_WriteSeparateFiles = outputBySSGroup And Extract_UseSSTags = False) Then Extract_WriteSeparateFiles = outputSingle
    If (Extract_WriteSeparateFiles = outputByFileName) Then
        tStr = tStr & "'separate files output (scanned file name)' "
    ElseIf (Extract_WriteSeparateFiles = outputBySSGroup) Then
        tStr = tStr & "'separate files output (SStag groups)' "
    Else
        tStr = tStr & "'single output' "
    End If
    AddOutput (tStr)
    
    ' show language
    If (Extract_Language <> "") Then
        AddOutput "     custom language specified: " & Extract_Language
    End If
    
    ' execute Timer
    Me.btnStop.Enabled = True
    AddOutput (" Finding source files...")
    Extract_State = exstateFindFiles
    Me.RunTimer.Enabled = True
    Me.RunTimer.Interval = 1
End Sub

Private Sub ExtractEnd()
    AddOutput (" Extraction complete in " & Int(Timer - Extract_StartTime) & " seconds")
    ' deinit
    If (Extract_UseSSTags = True) Then Run_Extract_ProcessSSTagsEnd
    If (Extract_Q3RadDef = True) Then Run_Extract_RadiantDefEnd
    If (Extract_GtkRad15Ent = True) Then Run_Extract_Radiant15XmlEnd
    If (Extract_WorldCraft33Fgd = True) Then Run_Extract_WorldCraft33FgdEnd

    ' shutdown
    Me.RunTimer.Enabled = False
    AddOutput (" Shutdown.")
    ' unfreeze buttons
    Me.btnStop.Enabled = False
    Me.btnCancel.Enabled = False
    Me.btnOpenDestDir.Enabled = True
    Me.btnPrev.Enabled = True
    Me.btnExit.Enabled = True
    Me.btnRescan.Enabled = True
End Sub

Private Sub ExtractionStop()
    AddOutput (" Stopped.")
    ExtractEnd
End Sub

Private Sub ExtractFrame()
    Dim NumFiles As Integer
    Dim NumDirs As Integer
    Dim tFile As String
    
    If (Extract_State = exstateFindFiles) Then
        List1.Clear
        Call FindFiles(List1, Extract_SrcPath, Extract_Ext, NumFiles, NumDirs)
        AddOutput ("     " & NumFiles & " file(s)")
        ' run init function for all extract passers
        If (Extract_UseSSTags = True) Then Run_Extract_ProcessSSTagsInit
        If (Extract_Q3RadDef = True) Then Run_Extract_RadiantDefInit
        If (Extract_GtkRad15Ent = True) Then Run_Extract_Radiant15XmlInit
        If (Extract_WorldCraft33Fgd = True) Then Run_Extract_WorldCraft33FgdInit
        AddOutput (" Extracting in progress...")
        Extract_CurFile = 0
        Extract_State = exstateExtract
        Exit Sub
    End If
    
    If (Extract_State = exstateExtract) Then
        If (Extract_CurFile >= List1.ListCount) Then
            ExtractEnd
            Exit Sub
        End If
        
        tFile = List1.List(Extract_CurFile)
        If (Extract_DevMessages = True) Then AddOutput (" scanning " & Mid(tFile, Len(Extract_SrcPath) + 1, 10000)) & "..."
        If (Extract_UseSSTags = True) Then
            Run_Extract_ProcessSSTags tFile
        Else
            If (Extract_Q3RadDef = True) Then Run_Extract_RadiantDefScanfile tFile
            If (Extract_GtkRad15Ent = True) Then Run_Extract_Radiant15XmlScanfile tFile
        End If
        
        Extract_CurFile = Extract_CurFile + 1
    End If
End Sub

'
''
''''
''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Universal SS tags parser/converter /*SS
'
''''''''''''''''''''''''''''''''''''''''''''''''''
''''
''
'

' replaces all ## values components with defines value
' currently it only scans for whole string to be def
Function SSReplaceDef(tInput As Variant)
    Dim tPos As Integer
    
    tInput = Trim$(tInput)
    If (Left$(tInput, 1) = "#" And Right$(tInput, 1) = "#") Then
        tInput = Mid$(tInput, 2, Len(tInput) - 2)
        CheckUseOfTemplates
        For tPos = 0 To SSTemplates.NumDefs - 1
            If (SSTemplates.Defs(tPos).Name = tInput) Then
                tInput = SSTemplates.Defs(tPos).Value
                Exit For
            End If
        Next tPos
        SSReplaceDef = SSReplaceDef(tInput)
        Exit Function
    End If
    SSReplaceDef = tInput
End Function

Sub Run_Extract_ProcessSSTagsInit()
    Dim tString As String
    Dim tGroupName As String
    Dim tEmptyTemplates As Templates
    Dim tIni As New clsINI
    Dim tPos As Integer
    Dim t As New clsTokenizer
    Dim tCfgFile As String
    Dim tDefsCount As Integer
    Dim tDGroupsCount As Integer

    AddOutput (" SS Tags processor init...")
    
    ' extract form
    Select Case Extract_Format
        Case exformatEntityDef
            AddOutput ("     target form: 'entity definitions'")
        Case Else
            AddOutput ("     target form: " & exformatEntityDef)
            AddOutput (" Error: unknown target form, there can be no output")
    End Select

    ' extract
    tString = "     target formats: "
    If (Extract_Q3RadDef = True) Then tString = tString & "'def' "
    If (Extract_GtkRad15Ent = True) Then tString = tString & "'xml'"
    If (Extract_WorldCraft33Fgd = True) Then tString = tString & "'fgd'"
    AddOutput (tString)
    Extract_UseSSTags_Counter = 0
    Extract_UseSSTags_SkipFormCounter = 0
    Extract_UseSSTags_SkipLangCounter = 0
    
    ' fix targets option
    Extract_SSFixTargetFields = False
    If (StrToInteger(INI.Read("Extract", "SSFixTargetStrings")) <> 0) Then
        Extract_SSFixTargetFields = True
        Extract_SSFixTargetFieldsMethod = StrToInteger(INI.Read("Extract", "SSFixTargetStringsMethod"))
        If (Extract_SSFixTargetFieldsMethod = 0) Then
            AddOutput ("     fix target/targetname keys with method: duplicate as strings")
        ElseIf (Extract_SSFixTargetFieldsMethod = 1) Then
            AddOutput ("     fix target/targetname keys with method: become strings")
        Else
            AddOutput ("     fix target/targetname keys with method: become strings only pure fields (with 'target' or 'targetname names')")
        End If
    End If
    
    ' load templates
    AddOutput (" Loading SS templates...")
    SSTemplates = tEmptyTemplates

    ' check for config file
    Extract_SSTemplatesLoaded = False
    If (Extract_CustomTemplates <> "") Then
        tCfgFile = App.path & "\" & Extract_CustomTemplates
        AddOutput "     custom templates file = '" & tCfgFile & "'"
        If (Dir(tCfgFile) = "") Then
            AddOutput "     failed to access custom templates file, using default paths"
        Else
            AddOutput "     custom templates file found"
            Extract_SSTemplatesLoaded = True
        End If
    End If
    If (Extract_SSTemplatesLoaded = False) Then
        tCfgFile = Extract_SrcPath & "config.ss.ini"
        Extract_TemplateFile1 = tCfgFile
        If (Dir(tCfgFile) = "") Then
            AddOutput "     '" & tCfgFile & "' not found"
            tCfgFile = Left$(Extract_SrcPath, Len(Extract_SrcPath) - 1)
            tCfgFile = FileName_GetPath(tCfgFile) & LCase$(FileName_StripPath(tCfgFile) & ".ss.ini")
            Extract_TemplateFile2 = tCfgFile
            If (Dir(tCfgFile) = "") Then
                AddOutput "     '" & tCfgFile & "' not found"
                tCfgFile = ""
            Else
                Extract_SSTemplatesLoaded = True
                AddOutput "     found '" & tCfgFile & "'"
            End If
        Else
            Extract_SSTemplatesLoaded = True
            AddOutput "     found '" & tCfgFile & "'"
        End If
    End If
    
    ' load templates
    If (tCfgFile <> "") Then
        tIni.File = tCfgFile
        SSTemplates.EditorModelsPath = tIni.Read("Index", "EditorModelsPath")
        ' read defs
        tDefsCount = 0
        SSTemplates.NumDefs = StrToInteger(tIni.Read("Index", "NumDefs"))
        For tPos = 0 To SSTemplates.NumDefs - 1
            t.TokenizeWithSeparator tIni.Read("Defs", "Def" & (tPos + 1)), "|"
            SSTemplates.Defs(tPos).Name = t.Argv(0)
            If (SSTemplates.Defs(tPos).Name <> "") Then
                tDefsCount = tDefsCount + 1
            End If
            SSTemplates.Defs(tPos).Value = SSReplaceDef(t.Argv(1))
        Next tPos
        AddOutput ("     define slots: " & SSTemplates.NumDefs & " (" & tDefsCount & " used)")
        ' read groups
        tGroupsCount = 0
        SSTemplates.NumGroups = StrToInteger(tIni.Read("Index", "NumGroups"))
        For tPos = 0 To SSTemplates.NumGroups - 1
            tGroupName = "Group" & (tPos + 1)
            ' load generic options
            SSTemplates.Groups(tPos).Name = SSReplaceDef(tIni.Read(tGroupName, "Name"))
            If (SSTemplates.Groups(tPos).Name <> "") Then
                tGroupsCount = tGroupsCount + 1
            End If
            SSTemplates.Groups(tPos).Color = SSReplaceDef(tIni.Read(tGroupName, "Color"))
            SSTemplates.Groups(tPos).Notes = SSReplaceDef(tIni.Read(tGroupName, "Notes"))
            SSTemplates.Groups(tPos).EditorModel = SSReplaceDef(tIni.Read(tGroupName, "Model"))
            SSTemplates.Groups(tPos).SFlag = StrToInteger(SSReplaceDef(tIni.Read(tGroupName, "SFlag")))
            SSTemplates.Groups(tPos).MinSize = SSReplaceDef(tIni.Read(tGroupName, "Min"))
            SSTemplates.Groups(tPos).MaxSize = SSReplaceDef(tIni.Read(tGroupName, "Max"))
            ' load flags
            SSTemplates.Groups(tPos).NumFlags = StrToInteger(SSReplaceDef(tIni.Read(tGroupName, "NumFlags")))
            For tPos2 = 0 To SSTemplates.Groups(tPos).NumFlags - 1
                SSTemplates.Groups(tPos).flags(tPos2) = tIni.Read(tGroupName, "Flag" & (tPos2 + 1))
            Next tPos2
            ' load keys
            SSTemplates.Groups(tPos).NumKeys = StrToInteger(SSReplaceDef(tIni.Read(tGroupName, "NumKeys")))
            For tPos2 = 0 To SSTemplates.Groups(tPos).NumKeys - 1
                SSTemplates.Groups(tPos).Keys(tPos2) = tIni.Read(tGroupName, "Key" & (tPos2 + 1))
            Next tPos2
        Next tPos
        AddOutput ("     template slots: " & SSTemplates.NumGroups & " (" & tGroupsCount & " used)")
    End If
End Sub

Sub Run_Extract_ProcessSSTags(tFilePath As String)
  Dim tFile As String
    Dim i, p, CutLen As Integer
    Dim tEntitiesProcessed As Integer
    Dim tFileContent As String
    Dim tInputState As Boolean
    Dim tInputBegin As Long
    Dim tInputEnd As Long
    Dim tOutputFileName As String
    Dim tOutPut As String
    
    If (Extract_WriteSeparateFiles = True) Then
        tOutputFileName = Extract_DstPath & "\" & FileName_StripPath(FileName_StripExt(FileName_StripExt(tFilePath))) & ".def"
    Else
        tOutputFileName = Extract_DstPath & "\" & "entities.def"
    End If
    
    ' scan file
    tFile = FreeFile
    Open tFilePath For Input As #tFile
        tInputLine = 0
        Do While Not EOF(tFile)
            tInputBegin = 0
            tInputEnd = 0
            Line Input #tFile, tFileContent
            ' scan each line for /*SS and begin input if found
            For p = 1 To Len(tFileContent)
                If (Mid(tFileContent, p, 4) = "/*SS") Then
                    ' change input state, check for nested /*QUAKED
                    If (tInputState = True) Then
                        AddOutput ("    warning: nested /*SS tag in file  " & Mid(List1.List(i), CutLen + 1, 10000) & " on line " & tInputLine)
                    End If
                    tInputState = True
                    tInputBegin = p
                    If (Extract_ExtMessages = True) Then
                        AddOutput ("     found " & Mid(tFileContent, p, 60)) & " on line " & tInputLine
                    End If
                    ' Extract_Q3RadDef_Counter = Extract_Q3RadDef_Counter + 1
                Else
                    If (tInputState = True) Then
                        If (Mid(tFileContent, p, 2) = "*/") Then
                            tEntitiesProcessed = tEntitiesProcessed + 1
                            tInputState = False
                            tInputEnd = p + 2
                        End If
                    End If
                End If
            Next p
            ' now make output
            If (tInputState = True) Then
                If (tInputBegin = 0) Then tInputBegin = 1
                If (tInputEnd = 0) Then tInputEnd = Len(tFileContent)
                If (tInputEnd = 0) Then tInputEnd = 1
            Else
                If (tInputBegin = 0) Then tInputBegin = 1
            End If
                
            If ((tInputBegin + tInputEnd) > 1) Then
                ' write output
                tOutPut = tOutPut & Mid(tFileContent, tInputBegin, tInputEnd - tInputBegin + 1) & Chr(13) & Chr(10)
                If (tInputState = False) Then
                    ' extract tag
                    Run_Extract_ExtractSSTag tFilePath, tOutPut
                    tOutPut = ""
                End If
            End If
            tInputLine = tInputLine + 1
        Loop
    Close #tFile
    
'    If (tOutPut = "") Then Exit Sub
'
 '   ' write founded output
'    tFile = FreeFile
'    Open tOutputFileName For Append As #tFile
'       Print #tFile, tOutPut & "// ------------------------------------------------------------"
'        Print #tFile, ""
'    Close #tFile
End Sub

Sub RunExtract_SetGroupProperty(tProp As String, tVal As Variant)
    If (IsNull(tVal) = True Or tVal = "") Then Exit Sub
    Run_Extract_SetProperty tProp, tVal
End Sub

Sub Run_Extract_SetProperty(tProp As String, tVal As Variant)
    Dim tLen As Integer
    Dim tStr As String
    Dim tPos, tPos2 As Integer
    Dim tProperty As String
    Dim tTk As New clsTokenizer
    Dim tKey As EntityDefKey
    
    Select Case tProp
        Case "form"
            tDef.Format = SSReplaceDef(tVal)
        Case "lang"
            tDef.Lang = SSReplaceDef(tVal)
        Case "class"
            tDef.ClassName = SSReplaceDef(tVal)
        Case "name"
            tDef.Name = SSReplaceDef(tVal)
        Case "desc"
            tDef.Description = SSReplaceDef(tVal)
        Case "notes"
            If (tDef.Notes <> "") Then
                tDef.Notes = tDef.Notes & Chr(13) & Chr(10) & " " & Chr(13) & Chr(10) & SSReplaceDef(tVal)
            Else
                tDef.Notes = SSReplaceDef(tVal)
            End If
        Case "color"
            tDef.Color = SSReplaceDef(tVal)
        Case "min"
            tDef.MinSize = SSReplaceDef(tVal)
        Case "max"
            tDef.MaxSize = SSReplaceDef(tVal)
        Case "temp"
            tDef.Description = tDef.Description & " " & Chr(13) & Chr(10) & " " & Chr(13) & Chr(10) & " DON'T USE. IT WILL BE REMOVED"
        Case "support"
            tDef.Description = tDef.Description & " " & Chr(13) & Chr(10) & " " & Chr(13) & Chr(10) & " DON'T USE. IT WAS ADDED ONLY FOR COMPATIBILITY WITH " & UCase$(SSReplaceDef(tVal))
        Case "notest"
            Run_Extract_SetProperty "notes", "NOT TESTED YET"
        Case "notdone"
            Run_Extract_SetProperty "notes", "NOT FINISHED YET"
        Case "unfinished"
            Run_Extract_SetProperty "notes", "NOT FINISHED YET"
        Case "deprecated"
            Run_Extract_SetProperty "notes", "This item exists only for compatibility reasons, using of it is deprecated."
        Case "model"
            tDef.EditorModel = SSReplaceDef(tVal)
        Case "sflag"
            tDef.SkillFlags = StrToInteger(SSReplaceDef(tVal))
        Case "base"
            CheckUseOfTemplates
            tLen = tTk.TokenizeWithSeparator(SSReplaceDef(tVal), ",")
            If (tTk.Argv(0) <> "") Then
                For i = 0 To tLen - 1
                    tStr = tTk.Argv(i)
                    ' avoid null groups
                    If (tStr <> "") Then
                        ' parse group
                        ' find group
                        For tPos = 0 To SSTemplates.NumGroups - 1
                            If (SSTemplates.Groups(tPos).Name = tStr) Then Exit For
                        Next tPos
                        ' set group
                        If (tPos < SSTemplates.NumGroups) Then
                            ' set generic options (only if entity does not have group yet!)
                            If (tDef.Group = "") Then
                                RunExtract_SetGroupProperty "color", SSTemplates.Groups(tPos).Color
                                RunExtract_SetGroupProperty "sflag", SSTemplates.Groups(tPos).SFlag
                                RunExtract_SetGroupProperty "min", SSTemplates.Groups(tPos).MinSize
                                RunExtract_SetGroupProperty "max", SSTemplates.Groups(tPos).MaxSize
                                If (SSTemplates.Groups(tPos).EditorModel <> "") Then
                                    If (SSTemplates.EditorModelsPath <> "") Then
                                        RunExtract_SetGroupProperty "model", SSTemplates.EditorModelsPath & "/" & SSTemplates.Groups(tPos).EditorModel
                                    Else
                                        RunExtract_SetGroupProperty "model", SSTemplates.Groups(tPos).EditorModel
                                    End If
                                End If
                            End If
                            ' set notes
                            RunExtract_SetGroupProperty "notes", SSTemplates.Groups(tPos).Notes
                            ' set flags
                            For tPos2 = 0 To SSTemplates.Groups(tPos).NumFlags - 1
                                RunExtract_SetGroupProperty "flag", SSTemplates.Groups(tPos).flags(tPos2)
                            Next tPos2
                            ' set keys
                            For tPos2 = 0 To SSTemplates.Groups(tPos).NumKeys - 1
                                RunExtract_SetGroupProperty "key", SSTemplates.Groups(tPos).Keys(tPos2)
                            Next tPos2
                        End If
                    End If
                Next i
            End If
        Case "group"
            CheckUseOfTemplates
            tLen = tTk.TokenizeWithSeparator(SSReplaceDef(tVal), ",")
            tDef.Group = ""
            Run_Extract_SetProperty "base", tVal
            tDef.Group = tTk.Argv(0)
            Exit Sub
        Case "flag"
            tLen = SSTokenizer.Tokenize(SSReplaceDef(tVal))
            tPos = BoundInt(1, StrToInteger(SSReplaceDef(SSTokenizer.Argv(0))), 8) - 1 ' first is flag bit
            tDef.flags(tPos).Key = SSReplaceDef(SSTokenizer.Argv(1))
            tDef.flags(tPos).Name = SSReplaceDef(SSTokenizer.Argv(2))
            tDef.flags(tPos).Description = SSReplaceDef(SSTokenizer.Argv(3))
            tDef.NumFlags = tDef.NumFlags + 1
        Case "key"
            tLen = SSTokenizer.Tokenize(SSReplaceDef(tVal))
            tProperty = SSReplaceDef(SSTokenizer.Argv(0))
            ' check if field is defined already, in this case
            For tPos = 0 To tDef.NumKeys - 1
                If (tDef.Keys(tPos).Name = tProperty) Then Exit For
            Next tPos
            ' if found - we must override it, else add new key
            If (tPos >= tDef.NumKeys) Then
                tPos = tDef.NumKeys
                tDef.NumKeys = tDef.NumKeys + 1
            End If
            
            tDef.Keys(tPos).Name = tProperty
            tStr = SSReplaceDef(SSTokenizer.Argv(1))
            If tStr = "target" Or tStr = "targ" Then
                tDef.Keys(tPos).Type = "target"
            ElseIf tStr = "targetname" Or tStr = "targname" Then
                tDef.Keys(tPos).Type = "targetname"
            ElseIf tStr = "string" Or tStr = "str" Then
                tDef.Keys(tPos).Type = "string"
            ElseIf tStr = "integer" Or tStr = "int" Then
                tDef.Keys(tPos).Type = "integer"
            ElseIf tStr = "boolean" Or tStr = "bool" Then
                tDef.Keys(tPos).Type = "boolean"
            ElseIf tStr = "direction" Or tStr = "dir" Then
                tDef.Keys(tPos).Type = "direction"
            ElseIf tStr = "angles" Then
                tDef.Keys(tPos).Type = "angles"
            ElseIf tStr = "angle" Or tStr = "ang" Then
                tDef.Keys(tPos).Type = "angle"
            ElseIf tStr = "list" Or tStr = "choices" Then
                tDef.Keys(tPos).Type = "choices"
                ' fill list items
                For tPos2 = 5 To tLen - 1
                    If (tDef.Keys(tPos).ListItems <> "") Then
                        tDef.Keys(tPos).ListItems = tDef.Keys(tPos).ListItems & "\"
                    End If
                    tDef.Keys(tPos).ListItems = tDef.Keys(tPos).ListItems & SSReplaceDef(SSTokenizer.Argv(tPos2))
                Next tPos2
            ElseIf tStr = "real3" Or tStr = "float3" Then
                tDef.Keys(tPos).Type = "real3"
            ElseIf tStr = "color" Then
                tDef.Keys(tPos).Type = "color"
            ElseIf tStr = "sound" Or tStr = "snd" Then
                tDef.Keys(tPos).Type = "sound"
            ElseIf tStr = "model" Or tStr = "mdl" Then
                tDef.Keys(tPos).Type = "model"
            ElseIf tStr = "sprite" Or tStr = "spr" Then
                tDef.Keys(tPos).Type = "sprite"
            ElseIf tStr = "texture" Or tStr = "tex" Then
                tDef.Keys(tPos).Type = "texture"
            Else ' "" Or "real" Or "float"
                tDef.Keys(tPos).Type = "real"
            End If
            tDef.Keys(tPos).DefValue = SSReplaceDef(SSTokenizer.Argv(2))
            tDef.Keys(tPos).Caption = SSReplaceDef(SSTokenizer.Argv(3))
            tDef.Keys(tPos).Description = SSReplaceDef(SSTokenizer.Argv(4))
           
            ' duplicate target/targetname keys as strings (optionaly)
            If (Extract_SSFixTargetFields = True) Then
                tStr = tDef.Keys(tPos).Type
                If (tStr = "target" Or tStr = "targetname") Then
                    If (Extract_SSFixTargetFieldsMethod = 0) Then
                        ' duplicate strings
                        tKey = tDef.Keys(tPos)
                        tKey.Type = "string"
                        tKey.Name = tKey.Name
                        tDef.Keys(tDef.NumKeys) = tKey
                        tDef.NumKeys = tDef.NumKeys + 1
                    ElseIf (Extract_SSFixTargetFieldsMethod = 1) Then
                        ' force strings
                        tDef.Keys(tPos).Type = "string"
                    Else
                        tDef.Keys(tPos).Type = "string"
                        ' force string only pure fields
                        If (tDef.Keys(tPos).Name = "target" Or tDef.Keys(tPos).Name = "targetname" Or tDef.Keys(tPos).Name = "killtarget") Then
                            tDef.Keys(tPos).Type = "string"
                        End If
                    End If
                End If
            End If
    End Select
End Sub

Sub Run_Extract_ExtractSSTag(tFilePath As String, tContent As String)
    Dim tParser As New clsSSparser
    Dim tEmptyDef As EntityDef
    Dim tExtract As Boolean
    
    ' parse tag properties
    tDef = tEmptyDef
    tParser.Load tContent, Extract_Format
    Do While tParser.ReadLine = True
        Run_Extract_SetProperty tParser.CurKey, tParser.CurVal
    Loop
    
    ' check if form and language is matching
    tExtract = True
    If (tDef.Format <> Extract_Format) Then
        tExtract = False
        Extract_UseSSTags_SkipFormCounter = Extract_UseSSTags_SkipFormCounter + 1
    End If
    If (tDef.Lang <> Extract_Language) Then
        tExtract = False
        Extract_UseSSTags_SkipLangCounter = Extract_UseSSTags_SkipLangCounter + 1
    End If
    
    ' final extract SS tag to needed formats
    If (tExtract = True) Then
        Extract_UseSSTags_Counter = Extract_UseSSTags_Counter + 1
        If (Extract_Q3RadDef = True) Then
            Run_Extract_SSTagToQ3RadiantDef tFilePath
        End If
        If (Extract_GtkRad15Ent = True) Then
            Run_Extract_SSTagToGtkRadiant15Ent tFilePath
        End If
        If (Extract_WorldCraft33Fgd = True) Then
            Run_Extract_SSTagToWorldCraft33Fgd tFilePath
        End If
    End If
End Sub
 
Sub Run_Extract_ProcessSSTagsEnd()
    AddOutput ("    " & Extract_UseSSTags_Counter & " tags extracted")
    AddOutput ("    " & (Extract_UseSSTags_Counter + Extract_UseSSTags_SkipFormCounter + Extract_UseSSTags_SkipLangCounter) & " tags processed")
    AddOutput ("    " & Extract_UseSSTags_SkipFormCounter & " tags with unmatched form")
    AddOutput ("    " & Extract_UseSSTags_SkipLangCounter & " tags with unmatched language")
End Sub

'
''
''''
''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Extract a Q3Radiant .DEF from /*SS TAG
'
''''''''''''''''''''''''''''''''''''''''''''''''''
''''
''
'

Sub Run_Extract_SSTagToQ3RadiantDef(tFilePath As String)
    Dim tOutputFileName, tFile As String
    Dim tString As String
    Dim tPos, tPos2, tLen2 As Integer
    Dim tTokenizer As New clsTokenizer
    
    tOutputFileName = ExtractGetOutputFileName(".def", tFilePath, tDef.Group)
    
    ' write file
    tFile = FreeFile
    
     ' if file not exists - create it and write headers
    tFile = FreeFile
    If (Dir(tOutputFileName) = "") Then
        Open tOutputFileName For Output As #tFile
        Print #tFile, "//"
        Print #tFile, "// Q3Radiant entity definition file (.def)"
        Print #tFile, "// Generated by RazorWind SourceScanner"
        Print #tFile, "// Converted from SS tag"
        Print #tFile, "//"
        Print #tFile, ""
    Else
        Open tOutputFileName For Append As #tFile
        Print #tFile, "// ----------------------------------------"
        Print #tFile, ""
    End If
    ' write file
        ' header string
        tString = "/*QUAKED " & tDef.ClassName & " (" & tDef.Color & ") "
        If (tDef.MinSize <> tDef.MaxSize) Then tString = tString & "(" & tDef.MinSize & ") (" & tDef.MaxSize & ") "
        tPos = 0
        Do While tPos < 8
            If (tDef.flags(tPos).Name <> "") Then
                tString = tString & UCase$(tDef.flags(tPos).Key) & " "
            Else
                tString = tString & "unused" & " "
            End If
            tPos = tPos + 1
        Loop
        If (tDef.SkillFlags = 1) Then tString = tString & "NOT_EASY NOT_MEDIUM NOT_HARD NOT_MULTI "
        Print #tFile, tString
        ' Description
        If (tDef.Description <> "" Or tDef.Notes <> "") Then
            Print #tFile, " " & tDef.Description & Chr(13) & Chr(10)
            If (tDef.Notes <> "") Then
                Print #tFile, " " & tDef.Notes & Chr(13) & Chr(10)
            End If
        End If
        ' Keys
        If (tDef.NumKeys > 0) Then
            tString = " --- Keys --- " & Chr(13) & Chr(10)
            tPos = 0
            Do While tPos < tDef.NumKeys
                If (tDef.Keys(tPos).Name <> "") Then
                    If (tDef.Keys(tPos).DefValue <> "") Then
                        tString = tString & " " & Chr(34) & tDef.Keys(tPos).Name & Chr(34) & " " & tDef.Keys(tPos).Description & ". Default is " & tDef.Keys(tPos).DefValue
                    Else
                        tString = tString & " " & Chr(34) & tDef.Keys(tPos).Name & Chr(34) & " " & tDef.Keys(tPos).Description
                    End If
                    If (tDef.Keys(tPos).ListItems <> "") Then
                        tString = tString & ": " & Chr(13) & Chr(10)
                        tLen2 = SSTokenizer.Tokenize(tDef.Keys(tPos).ListItems)
                        For tPos2 = 0 To tLen2 - 1 Step 2
                            tString = tString & "     " & SSTokenizer.Argv(tPos2) & ") " & SSTokenizer.Argv(tPos2 + 1) & Chr(13) & Chr(10)
                        Next tPos2
                    Else
                        tString = tString & "." & Chr(13) & Chr(10)
                    End If
                End If
                tPos = tPos + 1
            Loop
            Print #tFile, tString
        End If
        ' Flags
        If (tDef.NumFlags > 0) Then
            Print #tFile, " --- Flags --- "
            For tPos = 0 To 7
                If (tDef.flags(tPos).Name <> "") Then
                    Print #tFile, " " & UCase$(tDef.flags(tPos).Key) & ": " & tDef.flags(tPos).Description
                End If
            Next tPos
            Print #tFile, ""
        End If
        ' End
        If (tDef.EditorModel <> "") Then
            Print #tFile, "model = " & Chr(34) & tDef.EditorModel & Chr(34) & " */"
        Else
            Print #tFile, "*/"
        End If
        Print #tFile, ""
    Close #tFile
    
    Extract_Q3RadDef_Counter = Extract_Q3RadDef_Counter + 1
End Sub



'
''
''''
''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Extract a GTK Radiant 1.5 .ent from /*SS TAG
'
''''''''''''''''''''''''''''''''''''''''''''''''''
''''
''
'

' for xml file generation
Function Quotes(tInput As Variant)
    If (tInput = "") Then
        tInput = "sstag_data_missed"
    End If
    Quotes = Chr(34) & XMLString(tInput) & Chr(34)
End Function

Function XMLString(ByVal tCont As String)
    Dim tPos, tLen As Integer
    Dim tPrint, tChar As String
    
    ' replace all & with &amp;
    tLen = Len(tCont)
    For tPos = 1 To tLen
        tChar = Mid$(tCont, tPos, 1)
        If (tChar = "<") Then
            tPrint = tPrint & "{"
        ElseIf (tChar = ">") Then
            tPrint = tPrint & "}"
        Else
            tPrint = tPrint & tChar
        End If
    Next tPos
    
    XMLString = tPrint
End Function

Sub WriteXml(ByVal tFile As String, ByVal tCont As String)
    Dim tPos, tLen As Integer
    Dim tPrint, tChar As String
    
    ' replace all & with &amp;
    tLen = Len(tCont)
    For tPos = 1 To tLen
        tChar = Mid$(tCont, tPos, 1)
        If (tChar = "&") Then
            tPrint = tPrint & "&amp;"
        Else
            tPrint = tPrint & tChar
        End If
    Next tPos
    
    Print #tFile, tPrint
End Sub

Sub Run_Extract_SSTagToGtkRadiant15Ent(tFilePath As String)
    Dim tOutputFileName, tFile As String
    Dim tString, tStr As String
    Dim tPos, tPos2, tLen2 As Integer
    Dim tQ As String
    
    ' double quotes
    tQ = Chr(34)
    
    tOutputFileName = ExtractGetOutputFileName(".ent", tFilePath, tDef.Group)
    
    ' if file not exists - create it and write headers
    tFile = FreeFile
    If (Dir(tOutputFileName) = "") Then
        Open tOutputFileName For Output As #tFile
        WriteXml tFile, "<?xml version=" & Chr(34) & "1.0" & Chr(34) & "?>"
        WriteXml tFile, "<!--"
        WriteXml tFile, "GtkRadiant 1.5.0 entity definition file (.ent)"
        WriteXml tFile, "Generated by RazorWind SourceScanner"
        WriteXml tFile, "Converted from SS tag"
        WriteXml tFile, "-->"
        WriteXml tFile, "<classes>"
        WriteXml tFile, ""
    Else
        Open tOutputFileName For Append As #tFile
        WriteXml tFile, "<!--"
        WriteXml tFile, "============================================"
        WriteXml tFile, "-->"
        WriteXml tFile, ""
    End If
    
    ' write all lists
    If (tDef.NumKeys > 0) Then
        For tPos = 0 To tDef.NumKeys - 1
            If (tDef.Keys(tPos).ListItems <> "") Then
                WriteXml tFile, "<list name=" & Quotes(tDef.ClassName & "_" & tDef.Keys(tPos).Name) & ">"
                tLen2 = SSTokenizer.Tokenize(tDef.Keys(tPos).ListItems)
                For tPos2 = 0 To tLen2 - 1 Step 2
                    WriteXml tFile, " <item name=" & Quotes(SSTokenizer.Argv(tPos2 + 1)) & " value=" & Quotes(SSTokenizer.Argv(tPos2)) & "/>"
                Next tPos2
                WriteXml tFile, "</list>"
                WriteXml tFile, ""
            End If
        Next tPos
    End If
        
    ' write file
        ' write header and description
        If (tDef.MinSize <> tDef.MaxSize) Then
            If (tDef.EditorModel <> "") Then
                tString = "<point name=" & Quotes(tDef.ClassName) & " color=" & Quotes(tDef.Color) & " box=" & Quotes(tDef.MinSize & " " & tDef.MaxSize) & " model=" & Quotes(tDef.EditorModel) & ">"
            Else
                tString = "<point name=" & Quotes(tDef.ClassName) & " color=" & Quotes(tDef.Color) & " box=" & Quotes(tDef.MinSize & " " & tDef.MaxSize) & ">"
            End If
        Else
            If (tDef.EditorModel <> "") Then
                tString = "<group name=" & Quotes(tDef.ClassName) & " color=" & Quotes(tDef.Color) & " model=" & Quotes(tDef.EditorModel) & ">"
            Else
                tString = "<group name=" & Quotes(tDef.ClassName) & " color=" & Quotes(tDef.Color) & ">"
            End If
        End If
        If (tDef.Description <> "" Or tDef.Notes <> "") Then
            tString = tString & " " & XMLString(tDef.Description) & Chr(13) & Chr(10)
            If (tDef.Notes <> "") Then
                tString = tString & " " & XMLString(tDef.Notes) & Chr(13) & Chr(10)
            End If
        End If
        WriteXml tFile, tString
        
        ' Keys
        If (tDef.NumKeys > 0) Then
            WriteXml tFile, " --- Keys --- "
            For tPos = 0 To tDef.NumKeys - 1
                If (tDef.Keys(tPos).Name <> "") Then
                    ' head
                    tStr = tDef.Keys(tPos).Type
                    If (tStr = "choices") Then tStr = tDef.ClassName & "_" & tDef.Keys(tPos).Name ' override type for list
                    ' sprite fields currently not supported by GtkRad1.5 XMLDEF
                    If (tStr = "sprite") Then tStr = "string"
                    If (tDef.Keys(tPos).DefValue <> "") Then
                        tString = "<" & tStr & " key=" & Quotes(tDef.Keys(tPos).Name) & " name=" & Quotes(tDef.Keys(tPos).Caption) & ">" & XMLString(tDef.Keys(tPos).Description) & ". Default is " & XMLString(tDef.Keys(tPos).DefValue)
                    Else
                        tString = "<" & tStr & " key=" & Quotes(tDef.Keys(tPos).Name) & " name=" & Quotes(tDef.Keys(tPos).Caption) & ">" & XMLString(tDef.Keys(tPos).Description)
                    End If
                    ' list items
                    If (tDef.Keys(tPos).ListItems <> "") Then
                        tString = tString & ": " & Chr(13) & Chr(10)
                        tLen2 = SSTokenizer.Tokenize(tDef.Keys(tPos).ListItems)
                        For tPos2 = 0 To tLen2 - 1 Step 2
                            tString = tString & "     " & XMLString(SSTokenizer.Argv(tPos2)) & ") " & XMLString(SSTokenizer.Argv(tPos2 + 1)) & Chr(13) & Chr(10)
                        Next tPos2
                    Else
                        tString = tString & "."
                    End If
                    ' ent
                    tString = tString & "</" & tStr & ">"
                    WriteXml tFile, tString
                End If
            Next tPos
            WriteXml tFile, ""
        End If
        ' Flags
        If (tDef.NumFlags > 0 Or tDef.SkillFlags <> 0) Then Print #tFile, " --- Flags --- "
        If (tDef.NumFlags > 0) Then
            For tPos = 0 To 7
                If (tDef.flags(tPos).Name <> "") Then
                    WriteXml tFile, "<flag key=" & Quotes(UCase$(tDef.flags(tPos).Key)) & " name=" & Quotes(tDef.flags(tPos).Name) & " bit=" & Quotes(tPos) & ">" & XMLString(tDef.flags(tPos).Description) & "</flag>"
                End If
            Next tPos
            If (tDef.SkillFlags = 0) Then Print #tFile, ""
        End If
        If (tDef.SkillFlags <> 0) Then
            WriteXml tFile, "<flag key=" & Quotes("NOT_EASY") & " name=" & Quotes("Not easy") & " bit=" & Quotes(8) & ">not spawned in easy skill</flag>"
            WriteXml tFile, "<flag key=" & Quotes("NOT_MEDIUM") & " name=" & Quotes("Not medium") & " bit=" & Quotes(9) & ">not spawned in medium skill</flag>"
            WriteXml tFile, "<flag key=" & Quotes("NOT_HARD") & " name=" & Quotes("Not hard") & " bit=" & Quotes(10) & ">not spawned in hard skill</flag>"
            WriteXml tFile, "<flag key=" & Quotes("NOT_MULTI") & " name=" & Quotes("Not multiplayer") & " bit=" & Quotes(11) & ">not spawned in multiplayer</flag>"
            WriteXml tFile, ""
        End If
        ' End
        If (tDef.MinSize <> tDef.MaxSize) Then
            WriteXml tFile, "</point>"
        Else
            WriteXml tFile, "</group>"
        End If
        WriteXml tFile, ""
    Close #tFile
    Extract_GtkRad15Ent_Counter = Extract_GtkRad15Ent_Counter + 1
End Sub

'
''
''''
''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Extract WorldCraft 3.3 .fgd /*SS TAG
'
''''''''''''''''''''''''''''''''''''''''''''''''''
''''
''
'

Sub Run_Extract_SSTagToWorldCraft33Fgd(tFilePath As String)
    Dim tOutputFileName, tFile As String
    Dim tString, tStr As String
    Dim tPos, tPos2, tLen2 As Integer
    Dim tQ, tQ2 As String

    ' separate files output not supported
    tOutputFileName = ExtractGetOutputFileName(".fgd", tFilePath, tDef.Group)

    ' tab
    tQ = Chr(9)
    tQ2 = Chr(9) & Chr(9)
    
    ' if file not exists - create it and write headers
    tFile = FreeFile
    If (Dir(tOutputFileName) = "") Then
        Open tOutputFileName For Output As #tFile
        Print #tFile, "//"
        Print #tFile, "// Worldcraft 3.3 game definition file (.fgd)"
        Print #tFile, "// Generated by RazorWind SourceScanner"
        Print #tFile, "// Converted from SS tag"
        Print #tFile, "// "
        Print #tFile, ""
    Else
        Open tOutputFileName For Append As #tFile
        Print #tFile, "// ----------------------------------------"
        Print #tFile, ""
    End If
    
    ' write file
    ' write header and description
        ' make color string
        Call SSTokenizer.TokenizeWithSeparator(tDef.Color, " ")
        tStr = BoundInt(0, StrToSingle(SSTokenizer.Argv(0)) * 255, 255) & " " & BoundInt(0, StrToSingle(SSTokenizer.Argv(1)) * 255, 255) & " " & BoundInt(0, StrToSingle(SSTokenizer.Argv(2)) * 255, 255)
        ' write header
        If (tDef.MinSize <> tDef.MaxSize) Then
            tString = "@PointClass size(" & tDef.MinSize & ", " & tDef.MaxSize & ") color(" & tStr & ") = " & tDef.ClassName & " : " & Quotes(tDef.Name)
        Else
            tString = "@SolidClass color(" & tStr & ") = " & tDef.ClassName & " : " & Chr(34) & tDef.Name & Chr(34)
        End If
        Print #tFile, tString
        Print #tFile, "["
        ' write keys
        If (tDef.NumKeys > 0) Then
            For tPos = 0 To tDef.NumKeys - 1
                tStr = tDef.Keys(tPos).Type
                ' convert key type to known thing
                ' supported: integer, string, flags, choices,
                '            color255, studio, sound, sprite
                '            target_destination, target_source)
                If (tStr = "targetname") Then
                    tStr = "target_source"
                ElseIf (tStr = "target") Then
                    tStr = "target_destination"
                ElseIf (tStr = "string" Or tStr = "direction" Or tStr = "angles" Or tStr = "real3" Or tStr = "color" Or tStr = "texture") Then
                    tStr = "string"
                ElseIf (tStr = "sound") Then
                    tStr = "sound"
                ElseIf (tStr = "model") Then
                    tStr = "studio"
                ElseIf (tStr = "choices") Then
                    tStr = "choices"
                Else ' (tStr = "integer" Or Str = "boolean" Or Str = "angle") Then
                    tStr = "integer"
                End If
                ' write key
                If (tDef.Keys(tPos).DefValue <> "") Then
                    If (tStr = "integer") Then
                        tString = tQ & tDef.Keys(tPos).Name & "(" & tStr & ") : " & Quotes(tDef.Keys(tPos).Caption) & " : " & tDef.Keys(tPos).DefValue
                    Else
                        tString = tQ & tDef.Keys(tPos).Name & "(" & tStr & ") : " & Quotes(tDef.Keys(tPos).Caption) & " : " & Quotes(tDef.Keys(tPos).DefValue)
                    End If
                Else
                    tString = tQ & tDef.Keys(tPos).Name & "(" & tStr & ") : " & Quotes(tDef.Keys(tPos).Caption)
                End If
                
                ' for choises
                If (tStr = "choices") Then
                    tString = tString & " = " & Chr(13) & Chr(10) & tQ & "[" & Chr(13) & Chr(10)
                    tLen2 = SSTokenizer.Tokenize(tDef.Keys(tPos).ListItems)
                    For tPos2 = 0 To tLen2 - 1 Step 2
                        tString = tString & tQ2 & SSTokenizer.Argv(tPos2) & " : " & Quotes(SSTokenizer.Argv(tPos2 + 1)) & Chr(13) & Chr(10)
                    Next tPos2
                    tString = tString & tQ & "]"
                End If
                Print #tFile, tString
            Next tPos
        End If
        
        ' write flags
        If (tDef.NumFlags > 0 Or tDef.SkillFlags <> 0) Then
            Print #tFile, tQ & "spawnflags(Flags) ="
            Print #tFile, tQ & "["
            If (tDef.NumFlags > 0) Then
                For tPos = 0 To 7
                    If (tDef.flags(tPos).Name <> "") Then
                        Print #tFile, tQ2 & Power(2, tPos) & " : " & Quotes(tDef.flags(tPos).Name) & " : 0"
                    End If
                Next tPos
                If (tDef.SkillFlags = 0) Then Print #tFile, ""
            End If
            If (tDef.SkillFlags <> 0) Then
                Print #tFile, tQ2 & "256: " & Quotes("Not Easy") & " : 0"
                Print #tFile, tQ2 & "512: " & Quotes("Not Medium") & " : 0"
                Print #tFile, tQ2 & "1024: " & Quotes("Not Hard") & " : 0"
                Print #tFile, tQ2 & "2048: " & Quotes("Not Deathmatch") & " : 0"
            End If
            Print #tFile, tQ & "]"
        End If
        
        Print #tFile, "]"
        Print #tFile, ""
    Close #tFile
    
    Extract_WorldCraft33Fgd_Counter = Extract_WorldCraft33Fgd_Counter + 1
End Sub

'
''
''''
''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Extract a /*QUAKED FROM SOURCE FILE
'
''''''''''''''''''''''''''''''''''''''''''''''''''
''''
''
'

Sub Run_Extract_RadiantDefInit()
    Dim tFile As String
    Dim tNumRemovedFiles As Integer
    
    AddOutput (" Q3Radiant .def extracting init...")
    If (Extract_WriteSeparateFiles = True) Then
         AddOutput ("     output file: separate files")
    Else
         AddOutput ("     output file: " & Extract_DstPath & "\" & "entities.def")
    End If
    ' delete output files if they exists
    tFile = Dir(Extract_DstPath & "\*.def")
    If (tFile <> "") Then
        Do While tFile <> ""
            tNumRemovedFiles = tNumRemovedFiles + 1
            Kill Extract_DstPath & "\" & tFile
            tFile = Dir
        Loop
        AddOutput ("     notice: " & tNumRemovedFiles & " .def file(s) already exist in output directory, removing...")
    End If
    
    ' init stuff
    Extract_Q3RadDef_Counter = 0
End Sub

Private Sub Run_Extract_RadiantDefEnd()
    AddOutput ("     Q3Radiant .def: " & Extract_Q3RadDef_Counter & " definitions")
End Sub

Private Sub Run_Extract_RadiantDefScanfile(tFilePath As String)
    Dim tFile As String
    Dim i, p, CutLen As Integer
    Dim tEntitiesProcessed As Integer
    Dim tFileContent As String
    Dim tInputState As Boolean
    Dim tInputBegin As Long
    Dim tInputEnd As Long
    Dim tOutputFileName As String
    Dim tOutPut As String
    
    tOutPut = ""
    If (Extract_WriteSeparateFiles = True) Then
        tOutputFileName = Extract_DstPath & "\" & FileName_StripPath(FileName_StripExt(FileName_StripExt(tFilePath))) & ".def"
    Else
        tOutputFileName = Extract_DstPath & "\" & "entities.def"
    End If
    
    ' scan file
    tFile = FreeFile
    Open tFilePath For Input As #tFile
        tInputLine = 0
        Do While Not EOF(tFile)
            tInputBegin = 0
            tInputEnd = 0
            Line Input #tFile, tFileContent
            ' scan each line for /*QUAKED and begin input if found
            For p = 1 To Len(tFileContent)
                If (Mid(tFileContent, p, 8) = "/*QUAKED") Then
                    ' change input state, check for nested /*QUAKED
                    If (tInputState = True) Then
                        AddOutput ("    warning: nested /*QUAKED tag in file  " & Mid(List1.List(i), CutLen + 1, 10000) & " on line " & tInputLine)
                    End If
                    tInputState = True
                    tInputBegin = p
                    If (Extract_ExtMessages = True) Then
                        AddOutput ("     found : " & Mid(tFileContent, p + 9, 60))
                    End If
                    Extract_Q3RadDef_Counter = Extract_Q3RadDef_Counter + 1
                Else
                    If (tInputState = True) Then
                        If (Mid(tFileContent, p, 2) = "*/") Then
                            tEntitiesProcessed = tEntitiesProcessed + 1
                            tInputState = False
                            tInputEnd = p + 2
                        End If
                    End If
                End If
            Next p
            ' now make output
            If (tInputState = True) Then
                If (tInputBegin = 0) Then tInputBegin = 1
                If (tInputEnd = 0) Then tInputEnd = Len(tFileContent)
                If (tInputEnd = 0) Then tInputEnd = 1
            Else
                If (tInputBegin = 0) Then tInputBegin = 1
            End If
                
            If ((tInputBegin + tInputEnd) > 1) Then
                tOutPut = tOutPut & Mid(tFileContent, tInputBegin, tInputEnd - tInputBegin + 1) & Chr(13) & Chr(10)
                If (tInputState = False) Then tOutPut = tOutPut & Chr(13) & Chr(10)
            End If
            tInputLine = tInputLine + 1
        Loop
    Close #tFile
    
    If (tOutPut = "") Then Exit Sub
    
    
    ' add headers to scan file
    tFile = FreeFile
    Open tOutputFileName For Append As #tFile
        Print #tFile, "// Q3Radiant entity definition file"
        Print #tFile, "// Generated by RazorWind SourceScanner"
        Print #tFile, ""
    Close #tFile
    
    ' if file not exists - create it and write headers
    tFile = FreeFile
    If (Dir(tOutputFileName) = "") Then
        Open tOutputFileName For Output As #tFile
        Print #tFile, "// Q3Radiant entity definition file"
        Print #tFile, "// Generated by RazorWind SourceScanner"
        Print #tFile, ""
    Else
        Open tOutputFileName For Append As #tFile
    End If
    ' write output
        Print #tFile, tOutPut & "// ------------------------------------------------------------"
        Print #tFile, ""
    Close #tFile
End Sub

'
''
''''
''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Extract a /*RAD15ED FROM SOURCE FILE
'
''''''''''''''''''''''''''''''''''''''''''''''''''
''''
''
'

Private Sub Run_Extract_Radiant15XmlInit()
     Dim tFile As String
    Dim tNumRemovedFiles As Integer
    
    AddOutput (" GtkRadiant 1.5 .ent extracting init...")
    If (Extract_WriteSeparateFiles = True) Then
         AddOutput ("     output file: separate files")
    Else
         AddOutput ("     output file: " & Extract_DstPath & "\" & "entities.ent")
    End If
    ' delete output files if they exists
    tFile = Dir(Extract_DstPath & "\*.ent")
    If (tFile <> "") Then
        Do While tFile <> ""
            tNumRemovedFiles = tNumRemovedFiles + 1
            Kill Extract_DstPath & "\" & tFile
            tFile = Dir
        Loop
        AddOutput ("     notice: " & tNumRemovedFiles & " .ent file(s) already exist in output directory, removing...")
    End If
    
    ' init stuff
    Extract_GtkRad15Ent_Counter = 0
End Sub

Private Sub Run_Extract_Radiant15XmlScanfile(tFilePath As String)
    Dim tFile As String
    Dim i, p, CutLen As Integer
    Dim tEntitiesProcessed As Integer
    Dim tFileContent As String
    Dim tInputState As Boolean
    Dim tInputBegin As Long
    Dim tInputEnd As Long
    Dim tOutputFileName As String
    Dim tOutPut As String
    
    tOutPut = ""
    tOutputFileName = Extract_DstPath & "\" & "entities.ent"
    
    ' scan file
    tFile = FreeFile
    Open tFilePath For Input As #tFile
        tInputLine = 0
        Do While Not EOF(tFile)
            tInputBegin = 0
            tInputEnd = 0
            Line Input #tFile, tFileContent
            ' scan each line for /*QUAKED and begin input if found
            For p = 1 To Len(tFileContent)
                If (Mid(tFileContent, p, 10) = "/*GTKRAD15") Then
                    ' change input state, check for nested /*QUAKED
                    If (tInputState = True) Then
                        AddOutput ("    warning: nested /*GTKRAD15 tag in file  " & Mid(List1.List(i), CutLen + 1, 10000) & " on line " & tInputLine)
                    End If
                    tInputState = True
                    tInputBegin = p
                    If (Extract_ExtMessages = True) Then
                        AddOutput ("     found : " & Mid(tFileContent, p + 11, 60))
                    End If
                    Extract_GtkRad15Ent_Counter = Extract_GtkRad15Ent_Counter + 1
                Else
                    If (tInputState = True) Then
                        If (Mid(tFileContent, p, 2) = "*/") Then
                            tEntitiesProcessed = tEntitiesProcessed + 1
                            tInputState = False
                            tInputEnd = p + 2
                        End If
                    End If
                End If
            Next p
            ' now make output
            If (tInputState = True) Then
                If (tInputBegin = 0) Then tInputBegin = 1
                If (tInputEnd = 0) Then tInputEnd = Len(tFileContent)
                If (tInputEnd = 0) Then tInputEnd = 1
            Else
                If (tInputBegin = 0) Then tInputBegin = 1
            End If
                
            If ((tInputBegin + tInputEnd) > 1) Then
                tOutPut = tOutPut & Mid(tFileContent, tInputBegin, tInputEnd - tInputBegin + 1) & Chr(13) & Chr(10)
                If (tInputState = False) Then tOutPut = tOutPut & Chr(13) & Chr(10)
            End If
            tInputLine = tInputLine + 1
        Loop
    Close #tFile
    
    If (tOutPut = "") Then Exit Sub
    
    ' write founded output
    tFile = FreeFile
    Open tOutputFileName For Append As #tFile
        Print #tFile, tOutPut & "<!-- ================================================================================ --!>"
        Print #tFile, ""
    Close #tFile
End Sub

Private Sub Run_Extract_Radiant15XmlEnd()
    Dim tFile, tEntFile As String
    
    AddOutput ("     GtkRadiant 1.5.0 .ent: " & Extract_GtkRad15Ent_Counter & " definitions")
       
    ' autocomplete each file with </classes>
    tEntFile = Dir(Extract_DstPath & "\*.ent")
    Do While (tEntFile <> "")
        tFile = FreeFile
        Open Extract_DstPath & "\" & tEntFile For Append As #tFile
            Print #tFile, "</classes>"
        Close #tFile
        tEntFile = Dir
    Loop
End Sub

'
''
''''
''''''''''''''''''''''''''''''''''''''''''''''''''
'
' Extract a .FGD FROM SOURCE FILES
'
''''''''''''''''''''''''''''''''''''''''''''''''''
''''
''
'

Private Sub Run_Extract_WorldCraft33FgdInit()
    Dim tFile As String
    Dim tNumRemovedFiles As Integer
    
    AddOutput (" WorldCraft 3.3 .fgd extracting init...")
    AddOutput ("     output file: " & Extract_DstPath & "\" & "entities.fgd")
    If (Extract_WriteSeparateFiles = True) Then
         AddOutput ("     separate files output not supported")
    End If
        
    ' delete output files if they exists
    tFile = Dir(Extract_DstPath & "\*.fgd")
    If (tFile <> "") Then
        Do While tFile <> ""
            tNumRemovedFiles = tNumRemovedFiles + 1
            Kill Extract_DstPath & "\" & tFile
            tFile = Dir
        Loop
        AddOutput ("     notice: " & tNumRemovedFiles & " .fgd file(s) already exist in output directory, removing...")
    End If
    
    ' init stuff
    Extract_WorldCraft33Fgd_Counter = 0
End Sub

Private Sub Run_Extract_WorldCraft33FgdEnd()
    AddOutput ("     WorldCraft 3.3 .fgd: " & Extract_WorldCraft33Fgd_Counter & " definitions")
End Sub




'
''
''''
''''''''''''''''''''''''''''''''''''''''''''''''''
'
' CONTROLS
'
''''''''''''''''''''''''''''''''''''''''''''''''''
''''
''
'

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub btnOpenDestDir_Click()
    Shell "explorer.exe " & Extract_DstPath & "\", vbNormalFocus
End Sub

Private Sub btnPrev_Click()
    SwitchForm Me, frmMain
End Sub

Private Sub btnRescan_Click()
    Call Run_Extract
End Sub

Private Sub btnStop_Click()
    If (Me.RunTimer.Enabled = True) Then
        AddOutput (" Paused.")
        Me.RunTimer.Enabled = False
        Me.btnStop.Caption = "Resume"
    Else
        AddOutput (" Unpaused.")
        Me.RunTimer.Enabled = True
        Me.btnStop.Caption = "Stop"
    End If
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
    Set INI = New clsINI
    INI.File = App.path & "/sourcescanner.ini"

    Call Run_Extract
End Sub

Private Sub RunTimer_Timer()
    Call ExtractFrame
End Sub
