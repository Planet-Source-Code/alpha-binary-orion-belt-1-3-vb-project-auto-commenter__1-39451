VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProgress 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "rion Belt - Proceeding"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProcess.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   257
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar prgOverall 
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   9
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar prgCurrent 
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   735
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      MousePointer    =   9
      Scrolling       =   1
   End
   Begin VB.Label cmdNo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "&NO!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   2415
      TabIndex        =   12
      Top             =   3495
      Width           =   225
   End
   Begin VB.Shape shpButtonBorder 
      BorderColor     =   &H00808080&
      Height          =   165
      Index           =   1
      Left            =   2400
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label cmdYes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "&YES!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1230
      TabIndex        =   11
      Top             =   3375
      Width           =   1080
   End
   Begin VB.Shape shpButtonBorder 
      BorderColor     =   &H00808080&
      Height          =   345
      Index           =   0
      Left            =   1200
      Top             =   3360
      Width           =   1125
   End
   Begin VB.Label lblHeading 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Are You Sure You Want To Do This?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label lblConfirm 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmProcess.frx":324A
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label lblCurrent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Idle - Waiting for comfirmation..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Label lblCurrent 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Idle - Waiting for comfirmation..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Index           =   1
      Left            =   30
      TabIndex        =   8
      Top             =   1350
      Width           =   4095
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   272
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   272
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Label lblPercentage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Operation: 0%"
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   975
      Width           =   1695
   End
   Begin VB.Label lblPercentage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Overall Operation Completed: 0%"
      Height          =   195
      Index           =   0
      Left            =   810
      TabIndex        =   2
      Top             =   360
      Width           =   2475
   End
   Begin VB.Shape shpPogressBorder 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   1
      Left            =   105
      Top             =   720
      Width           =   3840
   End
   Begin VB.Shape shpPogressBorder 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   240
      Index           =   0
      Left            =   105
      Top             =   105
      Width           =   3840
   End
   Begin VB.Shape shpProgressShadow 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   0
      Left            =   135
      Top             =   195
      Width           =   3855
   End
   Begin VB.Shape shpProgressShadow 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   195
      Index           =   1
      Left            =   135
      Top             =   810
      Width           =   3855
   End
   Begin VB.Label lblShadow 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Overall Operation Completed: 0%"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   0
      Left            =   825
      TabIndex        =   4
      Top             =   375
      Width           =   2475
   End
   Begin VB.Label lblShadow 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Operation: 0%"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   1
      Left            =   1215
      TabIndex        =   5
      Top             =   990
      Width           =   1695
   End
   Begin VB.Label lblCurrentBackground 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Shape shpButtonShadow 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   315
      Index           =   0
      Left            =   1305
      Top             =   3450
      Width           =   1080
   End
   Begin VB.Shape shpButtonShadow 
      BorderColor     =   &H00808080&
      FillColor       =   &H00404040&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   2460
      Top             =   3540
      Width           =   225
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'   ============================================================
'    ----------------------------------------------------------
'     Application Name: Orion Belt
'     Developer/Programmer: Alpha Binary
'    ----------------------------------------------------------
'     Module Name: frmProgress
'     Module File: frmProcess.frm
'     Module Type: Form
'    ----------------------------------------------------------
'     © Copyright 2002
'    ----------------------------------------------------------
'   ============================================================

Option Explicit
Dim modpModules() As ModuleProperties
Dim bModuleSize As Byte


'----------------------------------------
'Name: cmdNo_MouseDown
'Object: cmdNo
'Event: MouseDown
'----------------------------------------
Private Sub cmdNo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdNo.Left = 164
    cmdNo.Top = 236
    shpButtonBorder(1).Left = 163
    shpButtonBorder(1).Top = 235
    cmdNo.ZOrder
    shpButtonBorder(1).ZOrder
End Sub


'----------------------------------------
'Name: cmdNo_MouseUp
'Object: cmdNo
'Event: MouseUp
'----------------------------------------
Private Sub cmdNo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    frmMain.Enabled = True
    Unload Me
End Sub


'----------------------------------------
'Name: cmdYes_MouseDown
'Object: cmdYes
'Event: MouseDown
'----------------------------------------
Private Sub cmdYes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdYes.Left = 87
    cmdYes.Top = 350
    shpButtonBorder(0).Left = 86
    shpButtonBorder(0).Top = 350
    shpButtonBorder(0).ZOrder
    cmdYes.ZOrder
End Sub


'----------------------------------------
'Name: cmdYes_MouseUp
'Object: cmdYes
'Event: MouseUp
'----------------------------------------
Private Sub cmdYes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdYes.Left = 82
    cmdYes.Top = 345
    shpButtonBorder(0).Left = 81
    shpButtonBorder(0).Top = 344
    shpButtonBorder(0).ZOrder
    cmdYes.ZOrder
    DoEvents
    CommentMain
End Sub


'----------------------------------------
'Name: CommentMain
'----------------------------------------
Private Sub CommentMain()
    'Begin Commenting...
    '-------------------
    '  Ok, let's make it clear. There's 3 step in commenting the whole project.
    'The first step is to scan the .vbp project file for the list of all modules
    'in the project, and store it in an array. Next step, we'll scroll through
    'the list of the modules, and determine its type. Modules will also be
    'commented in this step. And then step 3, procedures will be comment,
    'according to the options user have selected. The program will repeat step 2
    'and 3 While Not all modules, according to the project file, is fully commented.
    '  This procedure will be the main loop, which call other step-procedures and
    'manage variables. Actually it can be put in the cmdYes_MouseUp sub but I
    'would prefer putting it in another sub like this, so it would be easier to
    'change/call in newer versions.
    
    'Pre-Step: Set & clear variables
    bModuleSize = 1
    
    'Step 1 -----------------
    ReadProjectFile
    
    'Step 2 -----------------
    ScanModInfo
    
    'Step 3 -----------------
    CommentProject
    
    'Last Step: Notify User
    prgOverall.Value = 100
    prgCurrent.Value = 100
    DoEvents
    If MsgBox("Your project (" & IIf(modpModules(1).sModName = "Project1", frmMain.txtProjInfo(1), modpModules(1).sModName) & ") was fully commented!" & _
    IIf(frmMain.chkBackup.Value = 1, vbNewLine & "The backup project is stored in " & _
    Left$(sTargetProject, LastStr(1, sTargetProject, "\")) & "Backup", "") & _
    vbNewLine & "Do you want to terminate Orion Belt?", vbYesNo + vbQuestion, "Completed!") _
    = vbYes Then End 'Dunno, I've tried to unload every forms manually but it still hangs in my (computer's) memory
    Unload Me
    frmMain.Enabled = True
End Sub


'----------------------------------------
'Name: ReadProjectFile
'----------------------------------------
Private Sub ReadProjectFile()
    Dim sInput As String
    Dim sKeyword As String
    Dim bSeperator As Byte
    Dim bGotNothing As Boolean
    On Error Resume Next
    
    Open sTargetProject For Input As #1
        
        bModuleSize = 1
        Do While Not EOF(1)
            Line Input #1, sInput
            If sInput <> "" Then If InStr(1, sInput, "=") = 0 Then sKeyword = "" Else sKeyword = Left$(sInput, InStr(1, sInput, "=") - 1)
            
            bModuleSize = bModuleSize + 1
            ReDim Preserve modpModules(1 To bModuleSize) As ModuleProperties
            bGotNothing = False
            
            With modpModules(bModuleSize)
                Select Case sKeyword
                Case "Name" 'Project
                    With modpModules(1)
                        .sFileName = sTargetProject
                        .sModName = Mid$(sInput, Len(sKeyword) + 2)
                        .mModType = MProject
                    End With
                Case "Form" 'Form
                    .sFileName = Mid$(sInput, Len(sKeyword) + 2)
                    .mModType = MForm
                Case "Module" 'Module
                    bSeperator = InStr(Len(sKeyword) + 2, sInput, ";")
                    .sFileName = Mid$(sInput, bSeperator + 2)
                    .sModName = Mid$(sInput, Len(sKeyword) + 2, bSeperator - Len(sKeyword) - 1)
                    .mModType = MModule
                Case "Class" 'Class Module
                    bSeperator = InStr(Len(sKeyword) + 2, sInput, ";")
                    .sFileName = Mid$(sInput, bSeperator + 2)
                    .sModName = Mid$(sInput, Len(sKeyword) + 2, bSeperator - Len(sKeyword) - 1)
                    .mModType = MClass
                Case "UserControl" 'ActiveX Control
                    .sFileName = Mid$(sInput, Len(sKeyword) + 2)
                    .mModType = MUserControl
                Case "PropertyPage" 'Property Page
                    .sFileName = Mid$(sInput, Len(sKeyword) + 2)
                    .mModType = MPropPage
                Case Else
                    bGotNothing = True
                End Select
            End With
            If bGotNothing Then
                bModuleSize = bModuleSize - 1
                ReDim Preserve modpModules(1 To bModuleSize) As ModuleProperties
            End If
        Loop
        bModuleSize = bModuleSize - 1
    Close
End Sub


'----------------------------------------
'Name: ScanModInfo
'----------------------------------------
Private Sub ScanModInfo()
    Dim bCount As Byte
    Dim sCurLine As String
    Dim bLoop As Boolean
    
    For bCount = 2 To bModuleSize
        Do
            If Dir$(modpModules(bCount).sFileName) <> "" Then
                If Dir$(modpModules(bCount).sFileName, vbReadOnly) <> "" Then
                    Open modpModules(bCount).sFileName For Input As #1
                        Do While Not EOF(1)
                            Line Input #1, sCurLine
                            If Left$(sCurLine, 20) = "Attribute VB_Name = " Then
                                modpModules(bCount).sModName = Left(Mid$(sCurLine, 22), Len(Mid$(sCurLine, 22)) - 1)
                                Exit Do
                            End If
                        Loop
                    Close
                Else
                    Select Case MsgBox("Orion Error 405: File is Write-Protected" & vbNewLine & _
                        modpModules(bCount).sFileName & " is write-protected." & vbNewLine & vbNewLine & _
                        "Click Abort if you want to cancel the whole operation to re-check your project." & vbNewLine & _
                        "Click Retry if you just manually removed the read-only attribute and want Orion Belt to check the file again." & vbNewLine & _
                        "Click Ignore if you want to skip this file and continue with the next file." _
                        , vbAbortRetryIgnore + vbCritical, "Error!")
                    Case vbAbort
                        Unload Me
                    Case vbRetry
                        bLoop = True
                    Case vbIgnore
                        bLoop = False
                        modpModules(bCount).sFileName = ""
                    End Select
                End If
            Else
                Select Case MsgBox("Orion Error 404: Path Not Found" & vbNewLine & _
                    modpModules(bCount).sFileName & " cannot be found." & vbNewLine & vbNewLine & _
                    "Click Abort if you want to cancel the whole operation to re-check your project." & vbNewLine & _
                    "Click Retry if you want Orion Belt to check for the file's existance again." & vbNewLine & _
                    "Click Ignore if you want to skip this file and continue with the next file." _
                    , vbAbortRetryIgnore + vbCritical, "Error!")
                Case vbAbort
                    Unload Me
                Case vbRetry
                    bLoop = True
                Case vbIgnore
                    bLoop = False
                    modpModules(bCount).sFileName = ""
                End Select
            End If
        Loop While bLoop
    Next bCount
End Sub


'----------------------------------------
'Name: CommentProject
'----------------------------------------
Private Sub CommentProject()
    Dim bCommentOne As Boolean
    Dim bChar As Byte
    Dim bCheck As Byte
    Dim bCount As Byte
    Dim bSeperator As Byte
    Dim sCode() As String
    Dim sCurLine As String
    Dim sReadLine As String
    Dim sBackupPath As String
    Dim sMajorDecor As String
    Dim sMinorDecor As String
    Dim sModuleType As String
    Dim iLine As Integer
    Dim iCurLine As Integer
    Dim iLineToIns As Integer
    On Error Resume Next
    
    'Initialize Backup Folder
    If frmMain.chkBackup.Value = 1 Then
        sBackupPath = Left$(sTargetProject, LastStr(1, sTargetProject, "\")) & "Backup"
        If Dir$(sBackupPath) <> "" Then
            If MsgBox("The backup folder is already exist. All the existing file in the folder must be deleted before Orion Belt could continue. Proceed?", vbYesNo + vbQuestion, "Duplicate Folder Name") Then
                Kill sBackupPath & "\*.*"
                RmDir sBackupPath
                MkDir sBackupPath
            End If
        Else
            MkDir sBackupPath
        End If
    End If
    On Error GoTo 0
    
    'Initialize Decoration
    With frmMain
        If .optDecor(1).Value = True Then
            sMajorDecor = "="
            sMinorDecor = "-"
        ElseIf .optDecor(2).Value = True Then
            sMajorDecor = """"
            sMinorDecor = "'"
        ElseIf .optDecor(3).Value = True Then
            sMajorDecor = "*"
            sMinorDecor = "."
        End If
    End With
    
    For bCount = 1 To bModuleSize
        With modpModules(bCount)
            If .sFileName <> "" Then
                'Pre-Step - Initialize Variables & Backup
                ReDim sCode(1 To 1) As String
                iLine = 0
                CopyFile .sFileName, sBackupPath & "\" & Right$(.sFileName, Len(.sFileName) - LastStr(1, .sFileName, "\")), &O0
                
                'Step 3.1 - Load code io memory
                Open .sFileName For Input As #1
                    Do While Not EOF(1)
                        iLine = iLine + 1
                        ReDim Preserve sCode(1 To iLine) As String
                        Line Input #1, sCode(iLine)
                    Loop
                Close
                
                'Step 3.2 - Find the code line & comment basic info
                Open .sFileName For Output As #1
                    'What to do when we've got the file?
                    '3.2.1 Search for the 'Keyword' that tell us the next line is the code
                    '3.2.2 Check for the specified options
                    '3.2.3 Comment the module
                    
                    '3.2.1
                    Select Case .mModType
                    Case MForm
                        'Keyword = "Attribute VB_Exposed" (20 Char)
                        For iLine = 1 To UBound(sCode)
                            If Left$(sCode(iLine), 20) = "Attribute VB_Exposed" Then
                                iLineToIns = iLine + 1
                                Exit For
                            End If
                        Next iLine
                        sModuleType = "Form"
                    Case MModule
                        'No Keyword, There's only 1 line before the code
                        iLineToIns = 2
                        sModuleType = "Module"
                    Case MClass
                        'Keyword = "Attribute VB_Exposed" (20 Char)
                        For iLine = 1 To UBound(sCode)
                            If Left$(sCode(iLine), 20) = "Attribute VB_Exposed" Then
                                iLineToIns = iLine + 1
                                Exit For
                            End If
                        Next iLine
                        sModuleType = "Class Module"
                    Case MUserControl
                        'Keyword = "Attribute VB_Exposed" (20 Char)
                        For iLine = 1 To UBound(sCode)
                            If Left$(sCode(iLine), 20) = "Attribute VB_Exposed" Then
                                iLineToIns = iLine + 1
                                Exit For
                            End If
                        Next iLine
                        sModuleType = "UserControl"
                    Case MPropPage
                        'Keyword = "Attribute VB_Exposed" (20 Char)
                        For iLine = 1 To UBound(sCode)
                            If Left$(sCode(iLine), 20) = "Attribute VB_Exposed" Then
                                iLineToIns = iLine + 1
                                Exit For
                            End If
                        Next iLine
                        sModuleType = "Property Page"
                    Case MProject
                        iLineToIns = 5
                    End Select
                    
                    '3.2.2 & 3.2.3
                    For iLine = 1 To (iLineToIns - 1)
                        Print #1, sCode(iLine)
                    Next iLine
                    
                    If .mModType <> MProject Then
                        With frmMain
                            bCommentOne = False
                            For bCheck = 2 To 7
                                If .chkComment(bCheck).Value = 1 Then bCommentOne = True
                            Next bCheck
                            If bCommentOne Then
                                'Decoration 1st
                                Print #1,
                                Print #1, "'   " & Multiple(sMajorDecor, 60) '====================
                                Print #1, "'    " & Multiple(sMinorDecor, 58) '------------------
                                If .chkComment(5).Value = 1 Then Print #1, "'     Application Name: " & .txtProjInfo(1)
                                If .chkComment(6).Value = 1 Then Print #1, "'                       " & .txtProjInfo(2)
                                If .chkComment(3).Value = 1 Then Print #1, "'     Developer/Programmer: " & .txtProjInfo(3)
                                Print #1, "'    " & Multiple(sMinorDecor, 58) '------------------
                                If .chkComment(2).Value = 1 Then
                                    Print #1, "'     Module Name: " & modpModules(bCount).sModName
                                    Print #1, "'     Module File: " & modpModules(bCount).sFileName
                                    Print #1, "'     Module Type: " & sModuleType
                                End If
                                If .chkComment(7).Value = 1 Then Print #1, "'     Module Description:"
                                Print #1, "'    " & Multiple(sMinorDecor, 58) '------------------
                                If .chkComment(4).Value = 1 Then Print #1, "'     " & .txtProjInfo(4)
                                Print #1, "'    " & Multiple(sMinorDecor, 58) '------------------
                                Print #1, "'   " & Multiple(sMajorDecor, 60) '====================
                                Print #1,
                            End If
                        End With
                    Else
                        With frmMain
                            If .chkComment(1).Value = 1 Then
                                '3.2.3.1 Create the info file first
                                Open Left$(sTargetProject, LastStr(1, sTargetProject, "\")) & frmMain.txtProjInfo(1) & " Info.txt" For Output As #2
                                    Print #2, "Project info by Orion Belt " & App.Major & "." & App.Minor
                                    Print #2, "Best viewed in Notepad with a fixed-width font, such as Courier New and Terminal"
                                    Print #2,
                                    Print #2, Multiple(sMajorDecor, 70)
                                    Print #2, Multiple(sMinorDecor, 70)
                                    Print #2, .txtProjInfo(1)
                                    Print #2, .txtProjInfo(2)
                                    Print #2, "By " & .txtProjInfo(3)
                                    If .txtProjInfo(5) <> "" Then
                                        Print #2, Multiple(sMinorDecor, 70)
                                        Print #2, .txtProjInfo(5)
                                    End If
                                    If .txtProjInfo(4) <> "" Then
                                        Print #2, Multiple(sMinorDecor, 70)
                                        Print #2, .txtProjInfo(4)
                                    End If
                                    Print #2, Multiple(sMinorDecor, 70)
                                    Print #2, Multiple(sMajorDecor, 70)
                                    Print #2,
                                    Print #2, "Commented by Orion Belt " & App.Major & "." & App.Minor
                                Close #2
                                '3.2.3.2 Then add a link to the file
                                Print #1, "RelatedDoc=" & frmMain.txtProjInfo(1) & " Info.txt"
                            End If
                        End With
                    End If
                    
                    'Step 3.3 - Comment each procedures according to the options chosen
                    '3.3.1 Loop and search for the starting of each procedure
                    '3.3.2 Check for the specified options
                    '3.3.3 Comment procedures according to the option & procedure types
                    If frmMain.chkComment(8).Value = 1 Or frmMain.chkComment(9).Value = 1 Then
                        For iLine = iLineToIns To UBound(sCode)
                            If Right$(sCode(iLine - 1), 2) <> " _" Then
                                sReadLine = sCode(iLine)
                                iCurLine = iLine
                                Do
                                    If Right$(sReadLine, 2) = " _" Then 'Make it block-if for easier debugging
                                        iCurLine = iCurLine + 1
                                        sReadLine = Left$(sReadLine, Len(sReadLine) - 2) & " " & Trim$(sCode(iCurLine))
                                    Else
                                        Exit Do
                                    End If
                                Loop
                                bChar = 0
                                sCurLine = LTrim$(sReadLine) 'In case there're indents before the actual code
                                If Left$(sCurLine, 4) = "Sub " Then bChar = 5
                                If Left$(sCurLine, 9) = "Function " Then bChar = 10
                                If Left$(sCurLine, 9) = "Property " Then bChar = 10
                                If Left$(sCurLine, 11) = "Public Sub " Then bChar = 12
                                If Left$(sCurLine, 12) = "Private Sub " Then bChar = 13
                                If Left$(sCurLine, 13) = "Public Static " Then bChar = 14
                                If Left$(sCurLine, 16) = "Public Function " Then bChar = 17
                                If Left$(sCurLine, 16) = "Public Property " Then bChar = 17
                                If Left$(sCurLine, 17) = "Private Function " Then bChar = 18
                                If Left$(sCurLine, 17) = "Private Property " Then bChar = 18
                                If bChar > 0 Then
                                    If Trim$(sCode(iLine - 1)) = "" Then Print #1,
                                    Print #1, "'" & Multiple(sMinorDecor, 40)
                                    If frmMain.chkComment(8).Value = 1 Then
                                        Print #1, "'Name: " & Mid$(sCurLine, bChar, InStr(bChar, sCurLine, "(") - bChar)
                                        bSeperator = InStr(1, sCurLine, "_")
                                        If bSeperator <> 0 And bSeperator < InStr(1, sCurLine, "(") Then 'Object-Based Procedures
                                            Print #1, "'Object: " & Mid$(sCurLine, bChar, bSeperator - bChar)
                                            Print #1, "'Event: " & Mid$(sCurLine, bSeperator + 1, InStr(bChar, sCurLine, "(") - (bSeperator + 1))
                                        End If
                                    End If
                                    If frmMain.chkComment(9).Value = 1 Then Print #1, "'Description: "
                                    Print #1, "'" & Multiple(sMinorDecor, 40)
                                End If
                            End If
                            Print #1, sCode(iLine)
                        Next iLine
                    Else
                        For iLine = iLineToIns To UBound(sCode)
                            Print #1, sCode(iLine)
                        Next iLine
                    End If
                    'Ending Credits
                    If .mModType <> MProject Then
                        Print #1,
                        Print #1, "'Fully commented by Orion Belt®"
                        Print #1, "'Copyright © 2001-2002 Alpha Binary - All Right Reserved"
                    End If
                Close
            End If
        End With
    Next bCount
End Sub


'----------------------------------------
'Name: Multiple
'----------------------------------------
Private Function Multiple(sInput As String, bAmount As Byte) As String
    Dim bCount As Byte
    For bCount = 1 To bAmount
        Multiple = Multiple + sInput
    Next bCount
End Function


'----------------------------------------
'Name: LastStr
'----------------------------------------
Private Function LastStr(Start As String, String1 As String, String2 As String) As Byte
    Dim bChar As Integer
    Dim bStringPos As Byte
    For bChar = Start To Len(String1)
        bStringPos = InStr(bChar, String1, String2)
        If bStringPos <> 0 Then LastStr = bStringPos
    Next bChar
End Function
'
'Private Function SetSpace(InputLine As String) As String
'    SetSpace = InputLine
'    Do
'        If Replace$(SetSpace, "  ", " ") = SetSpace Then Exit Do Else SetSpace = Replace$(SetSpace, "  ", " ")
'    Loop
'End Function


'----------------------------------------
'Name: Form_Load
'Object: Form
'Event: Load
'----------------------------------------
Private Sub Form_Load()
    Me.Show
End Sub

'Fully commented by Orion Belt®
'Copyright © 2001-2002 Alpha Binary - All Right Reserved
