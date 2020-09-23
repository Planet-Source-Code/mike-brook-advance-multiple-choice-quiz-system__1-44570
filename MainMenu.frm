VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form MainMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8805
   ClientLeft      =   15
   ClientTop       =   -90
   ClientWidth     =   10485
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCPass 
      Caption         =   "CHANGE PASSWORD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton cmdORes 
      Caption         =   "VIEW &RESULTS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   21
      Top             =   5040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdRunQuiz 
      Caption         =   "RUN QUIZ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   5040
      Width           =   1815
   End
   Begin MSComctlLib.StatusBar bar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   8430
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame frmInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   3960
      TabIndex        =   14
      Top             =   0
      Width           =   6375
      Begin VB.Label lblAge 
         BackStyle       =   0  'Transparent
         Caption         =   "Age:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   5895
      End
      Begin VB.Label lblFullName 
         BackStyle       =   0  'Transparent
         Caption         =   "Student: Alex Brook"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   6015
      End
   End
   Begin VB.Frame frmMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   2055
      Begin VB.CommandButton cmdQuizes 
         Caption         =   "&QUIZZES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdResults 
         Caption         =   "RE&SULTS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton cmdTeachers 
         Caption         =   "&ADMINISTRATION"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton cmdLogOut 
         Caption         =   "LOG &OUT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   1815
      End
   End
   Begin VB.Frame frmOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   8
      Top             =   4320
      Width           =   2055
      Begin VB.CommandButton cmdFindStudent 
         Caption         =   "&FIND STUDENT"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton cmdChangePass 
         Caption         =   "CHANGE &PASSWORD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   18
         Top             =   3120
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.Timer timez 
      Interval        =   1
      Left            =   8520
      Top             =   5040
   End
   Begin VB.Data quizList 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   735
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Subjects"
      Top             =   9960
      Width           =   6375
   End
   Begin VB.Frame container 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quizzes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   6975
      Left            =   2400
      TabIndex        =   0
      Top             =   1320
      Width           =   7935
      Begin VB.ListBox moduleList 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5700
         Left            =   2760
         TabIndex        =   3
         Top             =   960
         Width           =   4815
      End
      Begin VB.ListBox subjectList 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5700
         ItemData        =   "MainMenu.frx":0000
         Left            =   360
         List            =   "MainMenu.frx":0002
         TabIndex        =   1
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Modules:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   5
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Subjects:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   2415
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Bindings        =   "MainMenu.frx":0004
      Height          =   2295
      Left            =   3480
      TabIndex        =   2
      Top             =   9960
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4048
      _Version        =   393216
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "12:00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   8760
      TabIndex        =   7
      Top             =   3960
      Width           =   4335
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "12:00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   8760
      TabIndex        =   6
      Top             =   3480
      Width           =   4335
   End
   Begin VB.Line Line2 
      X1              =   3840
      X2              =   3840
      Y1              =   1080
      Y2              =   0
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10680
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image Image1 
      Height          =   990
      Left            =   0
      Picture         =   "MainMenu.frx":001B
      Top             =   0
      Width           =   3780
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim emode As String 'the current mode (quiz/results)



Private Sub cmdCPass_Click()
MainMenu.Enabled = False 'disable main menu
frmCpass.Show 'load change password form
End Sub

Private Sub cmdFindStudent_Click()
frmAdmin.Show 'load admin form
End Sub

Private Sub cmdLogOut_Click()
Unload Me 'unload form
VdataStudent = "" 'clear student info id
LoginFrm.Show 'load login form
End Sub

Private Sub cmdORes_Click()

On Error GoTo resumeCode 'In case no results found

CurrentModuleID = Me.moduleList.ItemData(moduleList.ListIndex) 'Set Module ID Selected
CurrentSubjectID = Me.subjectList.ItemData(subjectList.ListIndex) 'Set Subject ID Selected
CurrentSubject = Me.subjectList.List(subjectList.ListIndex) 'Set Subject name
CurrentModule = Me.moduleList.List(moduleList.ListIndex) 'Set Module Name
frmResultView.Show 'Load results viewer
resumeCode:


End Sub

Private Sub cmdQuizes_Click()
emode = "QUIZ" 'set mode to quiz
container.Caption = "Quizzes" 'set caption
cmdRunQuiz.Visible = True 'enable Run Quiz button
cmdORes.Visible = False 'Disable view results button
End Sub

Private Sub cmdResults_Click()
emode = "" 'set mode to results only
If UserLevel = "TEACHER" Then GoTo skip 'If a Teacher skip this, as teacher can only view results
    container.Caption = "Results"
    cmdRunQuiz.Visible = False 'Disable Run Quiz
    cmdORes.Visible = True 'Enable View Results
skip:
End Sub

Private Sub cmdTeachers_Click()
frmAdminChoice.Show 'Load Admin Menu
End Sub

Private Sub cmdRunQuiz_Click()

On Error GoTo resumeCode 'In case no quiz found

CurrentModuleID = Me.moduleList.ItemData(moduleList.ListIndex) 'Set Module ID Selected
CurrentSubjectID = Me.subjectList.ItemData(subjectList.ListIndex) 'Set Subject ID Selected
CurrentSubject = Me.subjectList.List(subjectList.ListIndex) 'Set Subject name
CurrentModule = Me.moduleList.List(moduleList.ListIndex) 'Set Module Name

quizEngine.Show 'load the quiz engine

resumeCode:
End Sub


Private Sub Form_Load()
emode = "QUIZ" ' sets current mode (Quiz or Results)
bar.SimpleText = appinfo 'set status bar as copyright info
InitLibrary 'Builds Quiz Library
End Sub



Private Sub moduleList_DblClick()

If emode = "QUIZ" Then 'If quiz mode ...

On Error GoTo skp 'In case of no quiz

CurrentModuleID = Me.moduleList.ItemData(moduleList.ListIndex) 'Set Module ID Selected
CurrentSubjectID = Me.subjectList.ItemData(subjectList.ListIndex) 'Set Subject ID Selected
CurrentSubject = Me.subjectList.List(subjectList.ListIndex) 'Set Subject name
CurrentModule = Me.moduleList.List(moduleList.ListIndex) 'Set Module Name
quizEngine.Show 'load quiz engine

skp:
Else
On Error GoTo skp2 'If no results
If cmdORes.Enabled = True Then
CurrentModuleID = Me.moduleList.ItemData(moduleList.ListIndex) 'Set Module ID Selected
CurrentSubjectID = Me.subjectList.ItemData(subjectList.ListIndex) 'Set Subject ID Selected
CurrentSubject = Me.subjectList.List(subjectList.ListIndex) 'Set Subject name
CurrentModule = Me.moduleList.List(moduleList.ListIndex) 'Set Module Name
frmResultView.Show 'load results viewer
End If
skp2:


End If
End Sub

Private Sub subjectList_Click()
Dim SubjectIndex As Long
'lists modules from clicking on the subject
SubjectIndex = (subjectList.ItemData(subjectList.ListIndex)) 'Subject ID
LoadDB "SELECT ModuleDescription,ModuleID from Modules where [SubjectID] =" & SubjectIndex & ";", MainMenu.quizList 'Lists modules in that subject
moduleList.Clear 'Clear Module List
'populate the modules within the subject
    For i = 0 To quizList.Recordset.RecordCount 'Loop until all modules processed and added to list box
        If i + 1 > quizList.Recordset.RecordCount Then GoTo skp 'done
            grid.Col = 1 'set column 1 from flexgrid
            grid.Row = i + 1 'set row i from flexgrid
            moduleList.List(i) = grid.Text 'Set Module name in list to grid text
            grid.Col = 2 'set column 2 from flexgrid
            moduleList.ItemData(i) = grid.Text ' 'Set the newly added Module name with its Module ID.
skp:
    Next i



End Sub

Sub InitLibrary()

'Userlevel check
If UserLevel = "TEACHER" Then 'If a Teacher...
    'setup settings for TEACHER login
    lblFullName.Caption = "TEACHER:" & StudentFname & " " & StudentSName 'Set name (TEACHER)
    'lblAge.Caption = "Age:" & StudentAge
    cmdTeachers.Enabled = True 'Enable Admin
    cmdFindStudent.Enabled = True 'Enable Find a Student
    cmdQuizes.Enabled = False 'Disable Quiz Button
    cmdResults.Enabled = False 'Disable Results Button, as Teacher is locked into this mode
    container.Caption = "Results [No Student Selected.]" 'Set subject/module container to default.
    emode = "" 'Clear Mode
    cmdRunQuiz.Enabled = False 'Disable Quiz System
    cmdORes.Enabled = True 'Enable Results System
Else 'If student...
    'setup settings for STUDENT login
    lblFullName.Caption = "STUDENT:" & StudentFname & " " & StudentSName 'Set name (STUDENT)
    lblAge.Caption = "Age:" & StudentAge 'Set Age
    cmdTeachers.Enabled = False 'Disable Admin
End If


Me.Caption = "advance - Welcome " & StudentFname & " " & StudentSName & ":[" & StudentUserName & "]" 'set main menu caption
LoadDB "Subjects", MainMenu.quizList 'loads subjects table into data control

subjectList.Clear 'clear listbox

'populate subjects
    For i = 0 To quizList.Recordset.RecordCount 'loop until every record has been processed and added to list box.
        If i + 1 > quizList.Recordset.RecordCount Then 'done
        Else
            grid.Col = 2 'set column to 2 on flex grid
            grid.Row = i + 1 'set row to i on flex grid
            subjectList.List(i) = grid.Text 'Subject Name added to listbox
            grid.Col = 1 'Set column to 1 on flex grid
            subjectList.ItemData(i) = grid.Text 'Set the newly added Subject name with its Subject ID.
        End If
    Next i
End Sub

Private Sub timez_Timer()
'update the time and date
lblName(0).Caption = Time
lblName(1).Caption = Date
End Sub




