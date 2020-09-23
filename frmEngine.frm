VERSION 5.00
Begin VB.Form frmEdit 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "advance - Quiz Library Editor"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   7560
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close Editor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   6480
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuestions 
      Caption         =   "&Edit Questions"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   11
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Subjects"
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
      Height          =   2055
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Width           =   7095
      Begin VB.CommandButton cmdSav 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5400
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete Subject"
         Height          =   495
         Left            =   5400
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add Subject"
         Height          =   495
         Left            =   5400
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Data dbSubjects 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         EOFAction       =   1  'EOF
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtSId 
         DataField       =   "SubjectID"
         DataSource      =   "dbSubjects"
         Height          =   285
         Left            =   1680
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox txtSubjectName 
         DataField       =   "SubjectName"
         DataSource      =   "dbSubjects"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   2895
      End
   End
   Begin VB.Frame frmModules 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Modules"
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
      Height          =   2535
      Left            =   240
      TabIndex        =   3
      Top             =   3840
      Width           =   7095
      Begin VB.TextBox txtCurModId 
         DataField       =   "ModuleID"
         DataSource      =   "dbModules"
         Height          =   285
         Left            =   1680
         TabIndex        =   14
         Text            =   "Text3"
         Top             =   240
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CommandButton cmdDel2 
         Caption         =   "&Delete Module"
         Height          =   495
         Left            =   5400
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddModule 
         Caption         =   "&Add Module"
         Height          =   495
         Left            =   5400
         TabIndex        =   9
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         DataField       =   "SubjectID"
         DataSource      =   "dbModules"
         Height          =   285
         Left            =   1680
         TabIndex        =   7
         Text            =   "Text3"
         Top             =   600
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox txtModDescription 
         DataField       =   "ModuleDescription"
         DataSource      =   "dbModules"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   960
         Width           =   3135
      End
      Begin VB.Data dbModules 
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         EOFAction       =   1  'EOF
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Module Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
   End
   Begin VB.Label Label5 
      Caption         =   "   Q u i z    L i b r a r y    E d i t o r : :  Subject/Module"
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
      TabIndex        =   18
      Top             =   480
      Width           =   7575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Teacher Administration"
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
      TabIndex        =   17
      Top             =   120
      Width           =   4215
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4920
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Select or Edit Subjects and Modules:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1080
      Width           =   4935
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAdd_Click()
On Error Resume Next
dbSubjects.Recordset.AddNew 'Add New Record
refreshdata 'Refresh Tables
cmdSav.Enabled = True 'Enable Update Button
End Sub

Private Sub cmdAddModule_Click()
CurrentSubjectID = txtSId.Text 'Set Current Subject ID
frmAddModule.Show 'Load Add Module Form
End Sub

Private Sub cmdClose_Click()
Unload Me 'Unload Form
frmAdminChoice.Show 'Load Admin Menu
End Sub

Private Sub cmdDel_Click()
response = MsgBox("Are you sure you want to delete this subject?" & vbNewLine & "All Modules and questions within modules will be lost!", vbYesNo, "advance") 'Confirm that modules/questions will be deleted.
If response = vbYes Then 'If Yes then...
    dbSubjects.Recordset.Delete 'Delete Subject
    LoadDB "Subjects", Me.dbSubjects 'Reload Database
    refreshdata 'Refresh Tables
    MsgBox "Deleted Subject", vbInformation, "advance" 'Inform user of success
Else
End If
End Sub

Private Sub cmdDel2_Click()
On Error Resume Next
response = MsgBox("Are you sure you want to delete this module?" & vbNewLine & "All questions within module will also be lost!", vbYesNo, "advance") 'Confirm that questions will be deleted.
If response = vbYes Then 'If Yes then...
    dbModules.Recordset.Delete 'Delete Module
    refreshdata 'Refresh Tables
    MsgBox "Deleted Module", vbInformation, "advance" 'Inform user of success
End If
End Sub

Private Sub cmdQuestions_Click()
On Error GoTo skip
CurrentSubjectID = txtSId.Text 'Current Subject ID
CurrentModuleID = txtCurModId.Text 'Current Module ID
frmQuestions.Show 'Load Question Editor
EditQuestions CurrentSubjectID, CurrentModuleID 'Load Questions
frmQuestions.dbQuestions.Recordset.MoveLast 'Start at end of quiz (for ResultsCount to work)
frmQuestions.lblPos.Caption = "[" & frmQuestions.dbQuestions.Recordset.AbsolutePosition + 1 & "/" & frmQuestions.dbQuestions.Recordset.RecordCount & "]" 'Current Question out of total questions caption
frmQuestions.dbQuestions.Recordset.MoveFirst 'Move to start of questions
skip:

End Sub

Private Sub cmdSav_Click()
On Error Resume Next
cmdSav.Enabled = False 'Disable Update Button
dbSubjects.Recordset.Update 'Update Table
LoadDB "Subjects", frmEdit.dbSubjects 'Reload Database
dbSubjects.Recordset.MoveLast
End Sub

Private Sub dbModules_Reposition()
If txtModDescription <> "" Then 'If Mod Description does not equal "" then...
    cmdQuestions.Enabled = True 'Enable Edit Questions
Else
    cmdQuestions.Enabled = False 'Disable Edit Questions
End If
End Sub

Private Sub dbSubjects_Reposition()
refreshdata 'Refresh Tables
frmModules.Caption = "Modules (within " & txtSubjectName.Text & ")" 'Set Caption telling user which subject the module is in
End Sub

Private Sub Form_Load()
frmAdminChoice.Enabled = False 'Disable Admin Menu
LoadDB "Subjects", frmEdit.dbSubjects 'Load Subjects Table
End Sub

Private Sub Form_Unload(Cancel As Integer)
MainMenu.InitLibrary 'Refresh Library, so that newly added subjects etc is shown
frmAdminChoice.Enabled = True 'Enable Admin Menu
End Sub

Sub refreshdata()
On Error Resume Next
Dim SID As Long 'Subject ID
SID = txtSId.Text 'Subject ID
EditModules (SID) 'Reload Modules
skp:
End Sub


Private Sub txtModDescription_Change()
If txtModDescription <> "" Then 'If Mod Description does not equal "" then...
    cmdQuestions.Enabled = True 'Enable Edit Questions
Else
    cmdQuestions.Enabled = False 'Disable Edit Questions
End If
End Sub

