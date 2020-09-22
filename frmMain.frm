VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "rion Belt : Commenting Made Easy!"
   ClientHeight    =   4800
   ClientLeft      =   2205
   ClientTop       =   4395
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   552
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timTargetScroll 
      Interval        =   80
      Left            =   7800
      Top             =   480
   End
   Begin VB.Timer timColorScroll 
      Interval        =   450
      Left            =   7800
      Top             =   0
   End
   Begin VB.Frame fra 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Target Project"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4200
      TabIndex        =   35
      Top             =   3600
      Width           =   3855
      Begin VB.CheckBox chkBackup 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Always backup project before"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Value           =   1  'Checked
         Width           =   2520
      End
      Begin VB.TextBox txtTarget 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   255
         Width           =   2520
      End
      Begin VB.Shape shpCommentBorder 
         BorderColor     =   &H00808080&
         Height          =   225
         Left            =   2640
         Top             =   600
         Width           =   990
      End
      Begin VB.Label cmdComment 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Comment!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2670
         TabIndex        =   21
         Top             =   600
         Width           =   945
      End
      Begin VB.Shape shpCommentShadow 
         BorderColor     =   &H00808080&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   2730
         Top             =   690
         Width           =   975
      End
      Begin VB.Label cmdBrowse 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Browse..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2790
         TabIndex        =   19
         Top             =   255
         Width           =   825
      End
      Begin VB.Shape shpBrowseBorder 
         BorderColor     =   &H00C0C0C0&
         Height          =   225
         Left            =   2760
         Top             =   240
         Width           =   870
      End
      Begin VB.Shape shpTargetBorder 
         BorderColor     =   &H00C0C0C0&
         Height          =   225
         Left            =   120
         Top             =   240
         Width           =   2550
      End
      Begin VB.Shape shpTargetShadow 
         BorderColor     =   &H00808080&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   210
         Top             =   330
         Width           =   2535
      End
      Begin VB.Shape shpBrowseShadow 
         BorderColor     =   &H00808080&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   2850
         Top             =   330
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog cdlBrowse 
      Left            =   7320
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".vbp"
      DialogTitle     =   "Browse For Project"
      FileName        =   "*.vbp"
      Filter          =   "Visual Basic Project (*.vbp)|*.vbp"
      Flags           =   4
      InitDir         =   "C:\My Documents\"
   End
   Begin VB.Frame fraProject 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Commenting Choices"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   4200
      TabIndex        =   34
      Top             =   120
      Width           =   3855
      Begin VB.CheckBox chkComment 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Leave space for filling in procedure desc."
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   17
         Top             =   2760
         Width           =   3615
      End
      Begin VB.CheckBox chkComment 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Add basic information to modules (name, etc.)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox chkComment 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Leave space for each module's description"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox chkComment 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Add copyright to modules"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox chkComment 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Add application descriptions to modules"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   3615
      End
      Begin VB.CheckBox chkComment 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Add name, object and event in procedures"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   16
         Top             =   2520
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox chkComment 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Add developer name to modules"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox chkComment 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Add application name to modules"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox chkComment 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Add related info && credit document"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.Shape shpCommentSeper 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   120
         Top             =   2400
         Width           =   3615
      End
      Begin VB.Shape shpCommentSeper 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   120
         Top             =   720
         Width           =   3615
      End
      Begin VB.Shape shpCommentSeper 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   120
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame fraInfo 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Project Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.TextBox txtProjInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   900
         Index           =   5
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   2295
         Width           =   3465
      End
      Begin VB.OptionButton optDecor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "*** ..."
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   2880
         TabIndex        =   7
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton optDecor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   """ "" ' '"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   6
         Top             =   1680
         Width           =   615
      End
      Begin VB.OptionButton optDecor 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "=== ---"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Top             =   1680
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.TextBox txtProjInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   1335
         TabIndex        =   4
         Text            =   "© Copyright 2002"
         Top             =   1335
         Width           =   2280
      End
      Begin VB.TextBox txtProjInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   1
         Top             =   255
         Width           =   2175
      End
      Begin VB.TextBox txtProjInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   2
         Top             =   630
         Width           =   1935
      End
      Begin VB.TextBox txtProjInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   195
         Index           =   3
         Left            =   1920
         TabIndex        =   3
         Top             =   990
         Width           =   1695
      End
      Begin VB.Shape shpProjInfoBorder 
         BorderColor     =   &H00C0C0C0&
         Height          =   225
         Index           =   0
         Left            =   1425
         Top             =   240
         Width           =   2205
      End
      Begin VB.Shape shpProjInfoBorder 
         BorderColor     =   &H00C0C0C0&
         Height          =   225
         Index           =   1
         Left            =   1665
         Top             =   615
         Width           =   1965
      End
      Begin VB.Shape shpProjInfoBorder 
         BorderColor     =   &H00C0C0C0&
         Height          =   225
         Index           =   2
         Left            =   1905
         Top             =   975
         Width           =   1725
      End
      Begin VB.Shape shpProjInfoBorder 
         BorderColor     =   &H00C0C0C0&
         Height          =   225
         Index           =   3
         Left            =   1320
         Top             =   1320
         Width           =   2310
      End
      Begin VB.Shape shpProjInfoBorder 
         BorderColor     =   &H00C0C0C0&
         Height          =   930
         Index           =   4
         Left            =   120
         Top             =   2280
         Width           =   3495
      End
      Begin VB.Shape shpProjInfoShadow 
         BorderColor     =   &H00808080&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   900
         Index           =   4
         Left            =   210
         Top             =   2400
         Width           =   3480
      End
      Begin VB.Shape shpProjInfoShadow 
         BorderColor     =   &H00808080&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   3
         Left            =   1410
         Top             =   1410
         Width           =   2280
      End
      Begin VB.Label lblProjInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application Name:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   255
         Width           =   1290
      End
      Begin VB.Label lblProjInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "One-line description:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   615
         Width           =   1485
      End
      Begin VB.Label lblProjInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Developer/Programmer:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   975
         Width           =   1725
      End
      Begin VB.Label lblProjInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright Text:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   1335
         Width           =   1140
      End
      Begin VB.Label lblProjInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Decoration:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   1695
         Width           =   840
      End
      Begin VB.Label lblProjInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Description:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   22
         Top             =   2055
         Width           =   1140
      End
      Begin VB.Shape shpProjInfoShadow 
         BorderColor     =   &H00808080&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   0
         Left            =   1515
         Top             =   330
         Width           =   2175
      End
      Begin VB.Shape shpProjInfoShadow 
         BorderColor     =   &H00808080&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   1
         Left            =   1755
         Top             =   705
         Width           =   1935
      End
      Begin VB.Shape shpProjInfoShadow 
         BorderColor     =   &H00808080&
         FillColor       =   &H00404040&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   1995
         Top             =   1065
         Width           =   1695
      End
      Begin VB.Label lblProjInfoShadow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full Description:"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   33
         Top             =   2070
         Width           =   1140
      End
      Begin VB.Label lblProjInfoShadow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Decoration:"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   32
         Top             =   1710
         Width           =   840
      End
      Begin VB.Label lblProjInfoShadow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright Text:"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   31
         Top             =   1350
         Width           =   1140
      End
      Begin VB.Label lblProjInfoShadow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Developer/Programmer:"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   30
         Top             =   990
         Width           =   1725
      End
      Begin VB.Label lblProjInfoShadow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "One-line description"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   29
         Top             =   630
         Width           =   1425
      End
      Begin VB.Label lblProjInfoShadow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application Name:"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   28
         Top             =   270
         Width           =   1290
      End
   End
   Begin VB.Label lblOrionBelt 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "    "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   975
      Index           =   0
      Left            =   120
      TabIndex        =   37
      Top             =   3720
      Width           =   3975
   End
   Begin VB.Label lblOrionBelt 
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      Caption         =   "    rion Belt"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   36
      Top             =   3720
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   120
      Picture         =   "frmMain.frx":324A
      Top             =   3720
      Width           =   960
   End
   Begin VB.Shape shpFrameBorder 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   855
      Index           =   2
      Left            =   4320
      Top             =   3840
      Width           =   3855
   End
   Begin VB.Shape shpFrameBorder 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   3015
      Index           =   1
      Left            =   4320
      Top             =   360
      Width           =   3855
   End
   Begin VB.Shape shpFrameBorder 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   3255
      Index           =   0
      Left            =   240
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label lblOBBackground 
      BackColor       =   &H00A56E3A&
      Height          =   975
      Left            =   120
      TabIndex        =   38
      Top             =   3720
      Width           =   3975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'   ============================================================
'    ----------------------------------------------------------
'     Application Name: Orion Belt
'     Developer/Programmer: Alpha Binary
'    ----------------------------------------------------------
'     Module Name: frmMain
'     Module File: frmMain.frm
'     Module Type: Form
'    ----------------------------------------------------------
'     © Copyright 2002
'    ----------------------------------------------------------
'   ============================================================

Option Explicit


'----------------------------------------
'Name: cmdBrowse_MouseDown
'Object: cmdBrowse
'Event: MouseDown
'----------------------------------------
Private Sub cmdBrowse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdBrowse.Left = 2865
    cmdBrowse.Top = 330
    shpBrowseBorder.Left = 2835
    shpBrowseBorder.Top = 315
End Sub


'----------------------------------------
'Name: cmdBrowse_MouseUp
'Object: cmdBrowse
'Event: MouseUp
'----------------------------------------
Private Sub cmdBrowse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo CancelError
    cmdBrowse.Left = 2790
    cmdBrowse.Top = 255
    shpBrowseBorder.Left = 2760
    shpBrowseBorder.Top = 240
    cdlBrowse.ShowOpen
    sTargetProject = cdlBrowse.FileName
    txtTarget = sTargetProject + Space$(50)
    If txtProjInfo(1) <> "" Then If MsgBox("Project located. Begin commenting immediately?", vbYesNo + vbQuestion, "Proceeding with project file...") = vbYes Then cmdComment_MouseUp Button, Shift, X, Y
    Exit Sub
    
CancelError:
    If Err.Number = cdlCancel Then
        Err.Clear
        Exit Sub
    End If
End Sub


'----------------------------------------
'Name: cmdComment_MouseDown
'Object: cmdComment
'Event: MouseDown
'----------------------------------------
Private Sub cmdComment_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdComment.Left = 2745
    cmdComment.Top = 690
    shpCommentBorder.Left = 2715
    shpCommentBorder.Top = 675
End Sub


'----------------------------------------
'Name: cmdComment_MouseUp
'Object: cmdComment
'Event: MouseUp
'----------------------------------------
Private Sub cmdComment_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdComment.Left = 2670
    cmdComment.Top = 615
    shpCommentBorder.Left = 2655
    shpCommentBorder.Top = 600
    
    If Dir$(sTargetProject) <> "" Then 'If the file exist then...
        If LCase$(Right(sTargetProject, 3)) = "vbp" Then 'If the file match the vb project extension then...
            Load frmProgress 'Start commenting it!
            Me.Enabled = False
        Else
            MsgBox "Incorrect file extension!", vbOKOnly + vbCritical, "Commenting Error!"
        End If
    Else
        MsgBox "The specified target file does not exist!", vbOKOnly + vbCritical, "Commenting Error!"
    End If
End Sub


'----------------------------------------
'Name: Form_Load
'Object: Form
'Event: Load
'----------------------------------------
Private Sub Form_Load()
    Dim Count As Byte
    Dim SettingStr As String
    
    With frmMain
        If Dir$(App.Path & "\Settings.ini") = "" Then Exit Sub
        Open App.Path & "\Settings.ini" For Input As #1
            Line Input #1, SettingStr
            For Count = 1 To 9
                .chkComment(Count).Value = Mid$(SettingStr, Count, 1)
            Next Count
        Close
    End With
End Sub


'----------------------------------------
'Name: Form_Unload
'Object: Form
'Event: Unload
'----------------------------------------
Private Sub Form_Unload(Cancel As Integer)
    With frmMain
        Open App.Path & "\Settings.ini" For Output As #1
            Print #1, .chkComment(1).Value & _
                      .chkComment(2).Value & _
                      .chkComment(3).Value & _
                      .chkComment(4).Value & _
                      .chkComment(5).Value & _
                      .chkComment(6).Value & _
                      .chkComment(7).Value & _
                      .chkComment(8).Value & _
                      .chkComment(9).Value
        Close
    End With
End Sub


'----------------------------------------
'Name: timColorScroll_Timer
'Object: timColorScroll
'Event: Timer
'----------------------------------------
Private Sub timColorScroll_Timer()
    Static bCurColor As Byte 'Declare it as be as the only numbers involve is 0 and 1
    
    'Check if it's time to switch the foreground layer
    If Len(lblOrionBelt(bCurColor)) = 13 Then
        bCurColor = 1 - bCurColor
        lblOrionBelt(bCurColor).ZOrder
        lblOrionBelt(bCurColor) = Space$(4)
    End If
    
    'And now we add 1 more character to current layer
    lblOrionBelt(bCurColor) = lblOrionBelt(bCurColor) + Mid$("    rion Belt", Len(lblOrionBelt(bCurColor)) + 1, 1)
    
    'That's all!
    'At first I thought it's gonna be hard to write this part =O]
End Sub


'----------------------------------------
'Name: timTargetScroll_Timer
'Object: timTargetScroll
'Event: Timer
'----------------------------------------
Private Sub timTargetScroll_Timer()
    txtTarget = Mid$(txtTarget + Left$(txtTarget, 1), 2)
End Sub

'Fully commented by Orion Belt®
'Copyright © 2001-2002 Alpha Binary - All Right Reserved
