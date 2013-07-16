VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Source Scanner"
   ClientHeight    =   7320
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   9750
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCustomLanguage 
      Height          =   285
      Left            =   4440
      TabIndex        =   36
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtCustomConfig 
      Height          =   285
      Left            =   4440
      TabIndex        =   34
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CheckBox chkFixTargetFields 
      Caption         =   "Fix target/targetname fields"
      Height          =   255
      Left            =   3240
      TabIndex        =   28
      Top             =   3720
      Width           =   2295
   End
   Begin VB.CheckBox chkQ3RadSS 
      Caption         =   "Extract Q3Radiant  .def"
      Height          =   255
      Left            =   3240
      TabIndex        =   27
      Top             =   3000
      Width           =   3615
   End
   Begin VB.CheckBox chkSepFiles 
      Caption         =   "Write separate output files"
      Height          =   255
      Left            =   3000
      TabIndex        =   20
      Top             =   4935
      Width           =   2295
   End
   Begin VB.CheckBox chkInfoOutput 
      Caption         =   "Info logs"
      Height          =   255
      Left            =   3000
      TabIndex        =   26
      Top             =   5760
      Width           =   1095
   End
   Begin VB.CheckBox chkFgd 
      Caption         =   "Extract WorldCraft .fgd"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   3480
      Value           =   1  'Checked
      WhatsThisHelpID =   1
      Width           =   3615
   End
   Begin VB.CheckBox chkUseSSCodes 
      Caption         =   "Use /*SS tags processor (recommended)"
      Height          =   255
      Left            =   3000
      TabIndex        =   21
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton btnAbout 
      Caption         =   "&About"
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
      Left            =   1800
      TabIndex        =   19
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CheckBox chkDevOutput 
      Caption         =   "Developer logs"
      Height          =   255
      Left            =   5520
      TabIndex        =   18
      Top             =   5760
      Width           =   1455
   End
   Begin VB.DirListBox lstDestPath 
      Height          =   4590
      Left            =   7080
      TabIndex        =   16
      Top             =   1800
      Width           =   2415
   End
   Begin VB.DriveListBox lstDestDrive 
      Height          =   315
      Left            =   7080
      TabIndex        =   15
      Top             =   1440
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "Main.frx":08CA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   120
      Width           =   495
   End
   Begin VB.DirListBox lstSrcDir 
      Height          =   4590
      Left            =   360
      TabIndex        =   8
      Top             =   1800
      Width           =   2415
   End
   Begin VB.DriveListBox lstSrcDrive 
      Height          =   315
      Left            =   360
      TabIndex        =   7
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtPattern 
      Height          =   285
      Left            =   4440
      TabIndex        =   6
      Text            =   "*"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton btnNext 
      Caption         =   "&Scan and Extract >>"
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
      Left            =   7440
      TabIndex        =   5
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CheckBox chkFulloutput 
      Caption         =   "Detailed logs"
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CheckBox chkQ3Rad 
      Caption         =   "Direct extract /*QUAKED defs "
      Height          =   255
      Left            =   3000
      TabIndex        =   1
      Top             =   2520
      Value           =   1  'Checked
      WhatsThisHelpID =   1
      Width           =   3855
   End
   Begin VB.CommandButton btnQuit 
      Caption         =   "&Exit"
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
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame3"
      Height          =   855
      Left            =   -720
      TabIndex        =   11
      Top             =   -120
      Width           =   12135
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   " RazorWind Source Scanner v1.15"
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
         TabIndex        =   12
         Top             =   360
         Width           =   7815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   975
      Left            =   -1440
      TabIndex        =   13
      Top             =   6600
      Width           =   11415
   End
   Begin VB.CheckBox chkGtkRad15 
      Caption         =   "Extract GtkRadiant 1.5 .ent "
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Frame frmOutPut 
      Height          =   735
      Left            =   2880
      TabIndex        =   23
      Top             =   4920
      Width           =   4095
      Begin VB.OptionButton opSStagGroups 
         Caption         =   "For /*SStag groups"
         Height          =   240
         Left            =   2040
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton opScannedFiles 
         Caption         =   "For scanned files"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   2880
      TabIndex        =   29
      Top             =   2760
      Width           =   4095
      Begin VB.Frame frmFixTargetNameFields 
         Height          =   975
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   3855
         Begin VB.OptionButton opDuplicateTargetFields3 
            Caption         =   "Become strings only not pure fields"
            Height          =   255
            Left            =   480
            TabIndex        =   33
            Top             =   600
            Width           =   3255
         End
         Begin VB.OptionButton opDuplicateTargetFields2 
            Caption         =   "Become strings"
            Height          =   255
            Left            =   2400
            TabIndex        =   32
            Top             =   360
            Value           =   -1  'True
            Width           =   1415
         End
         Begin VB.OptionButton opDuplicateTargetFields1 
            Caption         =   "Duplicate as strings"
            Height          =   195
            Left            =   480
            TabIndex        =   31
            Top             =   360
            Width           =   1815
         End
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Custom language"
      Height          =   255
      Left            =   3000
      TabIndex        =   37
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Custom config file"
      Height          =   255
      Left            =   3000
      TabIndex        =   35
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "II. Scan and extract options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "III. Destination path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   7080
      TabIndex        =   14
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Search pattern"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "I. Search path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''
' DarkMaster Toolkit v0.5
' Main window
' by Pavel P. VorteX Timofeyev
' Property of RazorWind
''''''''''''''''''''''''''''''''''''''''''''''''''

Public INI As New clsINI    ' exe configuraion file
Public Messages_Level As Byte

Private Sub btnAbout_Click()
    SwitchForm Me, frmAbout
End Sub

'''''''''''''''''''''''''
' Interface
'''''''''''''''''''''''''

Private Sub btnDest_Click()
    Dim tPath As String
    
    If (txtPath.tExt = "") Then
        tPath = "c:\"
    Else
        tPath = txtPath.tExt
    End If
    
    tPath = SelectFolder(Me, "Select destination directory which should contain generated files", tPath)
    MsgBox tPath
    If (tPath <> "") Then
        txtDest.tExt = tPath
    End If
End Sub

Private Sub chkFixTargetFields_Click()
    If (Me.chkFixTargetFields.Value <> 0 And Me.chkFixTargetFields.Enabled = True) Then
        Me.opDuplicateTargetFields1.Enabled = True
        Me.opDuplicateTargetFields2.Enabled = True
        Me.opDuplicateTargetFields3.Enabled = True
    Else
        Me.opDuplicateTargetFields1.Enabled = False
        Me.opDuplicateTargetFields2.Enabled = False
        Me.opDuplicateTargetFields3.Enabled = False
    End If
End Sub

Private Sub chkSepFiles_Click()
    If (Me.chkSepFiles.Value = 1) Then
        Me.opScannedFiles.Enabled = True
        If (Me.chkUseSSCodes.Value = 1) Then
            Me.opSStagGroups.Enabled = True
        Else
            Me.opSStagGroups.Enabled = False
        End If
    Else
        Me.opScannedFiles.Enabled = False
        Me.opSStagGroups.Enabled = False
    End If
End Sub

Private Sub chkUseSSCodes_Click()
    If (Me.chkUseSSCodes.Value = 1) Then
        Me.chkQ3RadSS.Enabled = True
        Me.chkGtkRad15.Enabled = True
        Me.chkFgd.Enabled = True
        Me.chkFixTargetFields.Enabled = True
        If (Me.chkSepFiles.Value = 1) Then
            Me.opSStagGroups.Enabled = True
        Else
            Me.opSStagGroups.Enabled = False
        End If
    Else
        Me.chkQ3RadSS.Enabled = False
        Me.chkGtkRad15.Enabled = False
        Me.chkFgd.Enabled = False
        Me.opSStagGroups.Enabled = False
        Me.chkFixTargetFields.Enabled = False
    End If
    chkFixTargetFields_Click
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub btnPath_Click()
    Dim tPath As String
    
    If (txtPath.tExt = "") Then
        tPath = "c:\"
    Else
        tPath = txtPath.tExt & "\"
    End If
    
    tPath = SelectFolder(Me, "Select target directory which will be scanned", tPath)
    If (tPath <> "") Then
        txtPath.tExt = tPath
        txtPath.Refresh
    End If
End Sub

Private Sub btnQuit_Click()
    Unload Me
End Sub

Private Sub LoadDefaults()
    INI.Add "General", "Version", "1.0"
    INI.Add "Settings", "TargetPath", App.path
    INI.Add "Settings", "TargetExtensions", "*.c"
    INI.Add "Settings", "InfoOutput", "1"
    INI.Add "Settings", "DestinationPath", "[Select destination path]"
    INI.Add "Settings", "FullOutput", "0"
    INI.Add "Settings", "DevOutput", "0"
    INI.Add "Extract", "Q3RadDef", "0"
    INI.Add "Extract", "GtkRad150Ent", "1"
    INI.Add "Extract", "WorldCraft33Fgd", "0"
    INI.Add "Extract", "WriteSeparateFiles", "0"
    INI.Add "Extract", "UseSSTags", "1"
    INI.Add "Extract", "SSFixTargetStrings", "1"
    INI.Add "Extract", "SSFixTargetStringsMethod", "2"
End Sub

Private Sub Form_Activate()
    Dim tStr As String
    Dim tLen As Integer
    
    ' load up generic ini
    INI.File = App.path & "/sourcescanner.ini"
    ' make default ini if not exists
    If (INI.Read("General", "Version") <> "1.0") Then
        LoadDefaults
    End If
    
    ' target path
    tStr = INI.Read("Settings", "TargetPath")
    On Error GoTo Dir_Error1
    If (Dir(tStr, vbDirectory) = "") Then
Dir_Not_Found1:
        INI.Add "Settings", "TargetPath", App.path
        tStr = App.path
    End If
    Me.lstSrcDrive = Left(tStr, 2)
    Me.lstSrcDrive.Refresh
    Me.lstSrcDir.path = tStr
    Me.lstSrcDir.Refresh

    ' source path
    tStr = INI.Read("Settings", "DestinationPath")
    On Error GoTo Dir_Error2
    If (Dir(tStr, vbDirectory) = "") Then
Dir_Not_Found2:
        INI.Add "Settings", "DestinationPath", App.path
        tStr = App.path
    End If
    Me.lstDestDrive = Left(tStr, 2)
    Me.lstDestDrive.Refresh
    Me.lstDestPath.path = tStr
    Me.lstDestPath.Refresh
    
    ' other options
    Me.txtCustomLanguage.tExt = INI.Read("Settings", "CustomLanguage")
    Me.txtCustomConfig.tExt = INI.Read("Settings", "CustomConfig")
    Me.txtPattern.tExt = INI.Read("Settings", "TargetExtensions")
    Me.chkInfoOutput.Value = StrToInteger(INI.Read("Settings", "InfoOutput"))
    Me.chkFulloutput.Value = StrToInteger(INI.Read("Settings", "FullOutput"))
    Me.chkDevOutput.Value = StrToInteger(INI.Read("Settings", "DevOutput"))
    Me.chkQ3Rad.Value = StrToInteger(INI.Read("Extract", "Q3RadDef"))
    Me.chkQ3RadSS.Value = StrToInteger(INI.Read("Extract", "Q3RadDef"))
    Me.chkGtkRad15.Value = StrToInteger(INI.Read("Extract", "GtkRad150Ent"))
    Me.chkFgd.Value = StrToInteger(INI.Read("Extract", "WorldCraft33Fgd"))
    Me.chkUseSSCodes.Value = StrToInteger(INI.Read("Extract", "UseSSTags"))
    Me.chkFixTargetFields = StrToInteger(INI.Read("Extract", "SSFixTargetStrings"))
    tLen = StrToInteger(INI.Read("Extract", "SSFixTargetStringsMethod"))
    If (tLen = 0) Then
        Me.opDuplicateTargetFields1.Value = 1
        Me.opDuplicateTargetFields2.Value = 0
        Me.opDuplicateTargetFields3.Value = 0
    ElseIf (tLen = 1) Then
        Me.opDuplicateTargetFields1.Value = 0
        Me.opDuplicateTargetFields2.Value = 1
        Me.opDuplicateTargetFields3.Value = 0
    Else
        Me.opDuplicateTargetFields1.Value = 0
        Me.opDuplicateTargetFields2.Value = 0
        Me.opDuplicateTargetFields3.Value = 1
    End If
    tLen = StrToInteger(INI.Read("Extract", "WriteSeparateFiles"))
    If (tLen <> 0) Then
        Me.chkSepFiles.Value = 1
        If (tLen <> 1) Then
            Me.opScannedFiles.Value = 0
            Me.opSStagGroups.Value = 1
            Me.chkUseSSCodes.Value = 1
        Else
            Me.opScannedFiles.Value = 1
            Me.opSStagGroups.Value = 0
        End If
    Else
        Me.chkSepFiles.Value = 0
        Me.opScannedFiles.Value = 1
        Me.opSStagGroups.Value = 0
    End If
    chkUseSSCodes_Click
    chkSepFiles_Click
Quit:
    Exit Sub
Dir_Error1:
    Resume Dir_Not_Found1
Dir_Error2:
    Resume Dir_Not_Found2
End Sub

Private Sub btnNext_Click()
    ' set options
    INI.Add "Settings", "TargetPath", Me.lstSrcDir.path
    INI.Add "Settings", "TargetExtensions", Me.txtPattern.tExt
    INI.Add "Settings", "DestinationPath", Me.lstDestPath.path
    INI.Add "Settings", "InfoOutput", Me.chkInfoOutput.Value
    INI.Add "Settings", "FullOutput", Me.chkFulloutput.Value
    INI.Add "Settings", "DevOutput", Me.chkDevOutput.Value
    INI.Add "Extract", "Q3RadDef", Me.chkQ3Rad.Value
    INI.Add "Extract", "GtkRad150Ent", Me.chkGtkRad15.Value
    INI.Add "Extract", "WorldCraft33Fgd", Me.chkFgd.Value
    INI.Add "Extract", "UseSSTags", Me.chkUseSSCodes.Value
    INI.Add "Extract", "SSFixTargetStrings", Me.chkFixTargetFields.Value
    If (Me.opDuplicateTargetFields1.Value <> 0) Then
        INI.Add "Extract", "SSFixTargetStringsMethod", "0"
    ElseIf (Me.opDuplicateTargetFields2.Value <> 0) Then
        INI.Add "Extract", "SSFixTargetStringsMethod", "1"
    Else
        INI.Add "Extract", "SSFixTargetStringsMethod", "2"
    End If
    If (Me.chkUseSSCodes.Value <> 0) Then
        INI.Add "Extract", "UseSSTags", "1"
        INI.Add "Extract", "Q3RadDef", Me.chkQ3RadSS.Value
    Else
        INI.Add "Extract", "UseSSTags", "0"
    End If
    If (Me.chkSepFiles.Value = 1) Then
        If (Me.opScannedFiles = True) Then
            INI.Add "Extract", "WriteSeparateFiles", "1"
        Else
            INI.Add "Extract", "WriteSeparateFiles", "2"
        End If
    Else
        INI.Add "Extract", "WriteSeparateFiles", "0"
    End If
    INI.Add "Settings", "CustomConfig", Me.txtCustomConfig.tExt
    INI.Add "Settings", "CustomLanguage", Me.txtCustomLanguage.tExt
    SwitchForm Me, frmScan
End Sub

Private Sub lstDestDrive_Change()
    Me.lstDestPath.path = Me.lstDestDrive.Drive
End Sub

Private Sub lstSrcDrive_Change()
    Me.lstSrcDir.path = Me.lstSrcDrive.Drive
End Sub
