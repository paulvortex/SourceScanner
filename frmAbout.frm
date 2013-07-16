VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SourceScanner"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Height          =   2535
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1200
      Width           =   6375
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Oki-doki..."
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   4680
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   975
      Left            =   -720
      TabIndex        =   4
      Top             =   4440
      Width           =   10335
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame3"
      Height          =   855
      Left            =   -720
      TabIndex        =   1
      Top             =   -120
      Width           =   10335
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "RazorWind SourceScanner v1.1"
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
         TabIndex        =   2
         Top             =   360
         Width           =   7815
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "By Pavel P. [VorteX] Timofeyev, (C) RazorWind Games"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "If you like this program, or want to get new features/source code of it drop me e-mail to paul.vortex@gmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   6375
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClose_Click()
    SwitchForm Me, frmMain
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim tReadmeFile As String
    Dim tFile As String
    Dim tFileContent As String
    Dim tFileLen As Long
    Dim tNextLine As String
    Dim tText As String
    
    ' load readme
    Me.Text1.tExt = ""
    tReadmeFile = App.path & "\" & "sourcescanner.txt"
    If (Dir(tReadmeFile) <> "") Then
        tFile = FreeFile
        Open tReadmeFile For Input As #tFile
            tFileLen = FileLen(tReadmeFile)
            Do While (EOF(tFile) = False)
                Line Input #tFile, tFileContent
                Me.Text1.tExt = Me.Text1.tExt & " " & tFileContent & Chr(13) & Chr(10)
            Loop
        Close #tFile
        Me.Text1.Refresh
    Else
        ' embedded readme
        tNextLine = Chr(13) & Chr(10)
        
        tText = " This little utility will help programmers to extract /*QUAKED" & tNextLine & _
                "stuff directly from sourcecodes and make .DEF/.ENT/.FGD for level editor" & tNextLine & _
                "================================================================" & tNextLine & _
                " FEATURES " & _
                "================================================================" & tNextLine & _
                " - Extract /*QUAKED entity definitions from source files into entities.def" & tNextLine & _
                "   which can be used by Q3Radiant/GtkRadiant and others" & tNextLine & _
                " - Universal entity definition language named 'SSTag' which allow" & tNextLine & _
                "   exporting to many entity definition formats. Currently supported" & tNextLine & _
                "   export formats are: .Def (Q3Radiant), .Ent (GtkRadiant 1.5) and" & tNextLine & _
                "   .FGD (WorldCraft and Hammer editor)" & tNextLine & _
                " - Defines and templates that readed from custom files (default paths are" & tNextLine & _
                "   [destination folder]\ssconfig.ini and [destination folder]\..\[destination folder].ini" & tNextLine & _
                "   Defines can be used to share parms between multiple entities" & tNextLine & _
                "   Templates is handy for sharing groups of parms/notes and other options " & tNextLine & _
                "   between entities."
        Me.Text1.tExt = " readme file not found"
    End If
End Sub

