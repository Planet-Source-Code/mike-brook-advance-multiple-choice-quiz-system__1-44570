VERSION 5.00
Begin VB.Form frmTopScore 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Awards and Certificates"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   3990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdWorst 
      Caption         =   "&Worst Student"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5760
      Top             =   480
   End
   Begin VB.TextBox results 
      DataField       =   "StudentID"
      DataSource      =   "dbTopScore"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   6000
      Width           =   3375
   End
   Begin VB.CommandButton cmdTopStudent 
      Caption         =   "&Top Student"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3495
   End
   Begin VB.Data dbTopScore 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Width           =   3975
   End
   Begin VB.Label Label5 
      Caption         =   "  Find Student in a specific area :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   9255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Awards and Development"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "in currently selected module."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   2760
      Y1              =   480
      Y2              =   480
   End
End
Attribute VB_Name = "frmTopScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdWorst_Click()
LoadDB "SELECT * from StudentScores WHERE [SubjectModuleID] =" & MainMenu.moduleList.ItemData(MainMenu.moduleList.ListIndex) & " ORDER BY SubjectScore", frmTopScore.dbTopScore 'Load Student Results in Ascending Order
If results.Text = "" Then 'If no student results then...
    MsgBox "No results for students in this module.", vbInformation, "advance" 'Inform user
Else 'If results then...
    CurrentSubject = MainMenu.subjectList.List(MainMenu.subjectList.ListIndex) 'Current Subject
    CurrentModule = MainMenu.moduleList.List(MainMenu.moduleList.ListIndex) 'Current Module
    LoadDB "SELECT * FROM Students WHERE [StudentID] =" & results.Text & ";", frmAdmin.userView 'Load Student Table for the Worst Student
    frmAdmin.users.Caption = "User Administration [Worst Student in " & MainMenu.moduleList.List(MainMenu.moduleList.ListIndex) & "]" 'Update caption on Admin form to show Worst Student
    Unload Me 'Unload Form
End If
End Sub

Private Sub CmdTopStudent_Click()

LoadDB "SELECT * from StudentScores WHERE [SubjectModuleID] =" & MainMenu.moduleList.ItemData(MainMenu.moduleList.ListIndex) & " ORDER BY SubjectScore DESC;", frmTopScore.dbTopScore 'Load Student Results in Descending Order

If results.Text = "" Then 'If no student results then...
    MsgBox "No results for students in this module.", vbInformation, "advance" 'Inform user
Else
    CurrentSubject = MainMenu.subjectList.List(MainMenu.subjectList.ListIndex) 'Current Subject
    CurrentModule = MainMenu.moduleList.List(MainMenu.moduleList.ListIndex) 'Current Module
    LoadDB "SELECT * FROM Students WHERE [StudentID] =" & results.Text & ";", frmAdmin.userView 'Load Student Table for the Best Student
    frmAdmin.users.Caption = "User Administration [Top Student in " & MainMenu.moduleList.List(MainMenu.moduleList.ListIndex) & "]" 'Update caption on Admin form to show Best Student
    Unload Me 'Unload Form
End If
End Sub


Private Sub Form_Load()
With MainMenu
    .cmdTeachers.Enabled = False 'Disable Admin
    .cmdQuizes.Enabled = False 'Disable Quiz
    .cmdResults.Enabled = False 'Disable Results
    .cmdRunQuiz.Enabled = False 'Disable Quiz System
    .cmdORes.Enabled = False 'Disable Resuls System
    .cmdFindStudent.Enabled = False 'Disable Find Student
    .cmdCPass.Enabled = False 'Disable Change Password
    .cmdLogOut.Enabled = False 'Disable Log Out
End With
frmAdmin.Visible = False 'Disable Admin Form
Me.top = 0 'Set Form to top of screen
Me.Left = Screen.Width - Me.Width 'Set form to right hand side.
MainMenu.Enabled = True 'Enable Main Menu
If MainMenu.moduleList.ListIndex = -1 Then 'If no module selected then...
    cmdTopStudent.Enabled = False 'Disable Top Student Button
Else
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmAdmin.Visible = True 'Enable Admin Form
MainMenu.Enabled = False 'Disable Main Menu
With MainMenu
    .cmdTeachers.Enabled = True 'Enable Admin
    .cmdQuizes.Enabled = True 'Enable Quiz
    .cmdResults.Enabled = True 'Enable Results
    .cmdRunQuiz.Enabled = True 'Enable Quiz System
    .cmdORes.Enabled = True 'Enable Resuls System
    .cmdFindStudent.Enabled = True 'Enable Find Student
    .cmdCPass.Enabled = True 'Enable Change Password
    .cmdLogOut.Enabled = True 'Enable Log Out
End With
frmAdmin.Show 'Load Admin Form
End Sub

Private Sub Timer1_Timer()
If MainMenu.moduleList.ListIndex = -1 Then 'If no module selected then...
    cmdTopStudent.Enabled = False 'Disable Top Student Button
    cmdWorst.Enabled = False 'Disable Worst Student Button
Else
    cmdTopStudent.Enabled = True 'Enable Top Student Button
    cmdWorst.Enabled = True 'Enable Worst Student Button
End If
End Sub
