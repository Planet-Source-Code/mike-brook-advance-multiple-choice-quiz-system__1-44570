VERSION 5.00
Begin VB.Form frmQuestions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Questions"
   ClientHeight    =   9660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSub 
      DataField       =   "SubjectID"
      DataSource      =   "dbQuestions"
      Height          =   285
      Left            =   3600
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   7560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtMod 
      DataField       =   "ModuleID"
      DataSource      =   "dbQuestions"
      Height          =   285
      Left            =   5520
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   7560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox res5 
      CausesValidation=   0   'False
      Height          =   315
      ItemData        =   "frmQuestions.frx":0000
      Left            =   360
      List            =   "frmQuestions.frx":0002
      TabIndex        =   19
      Text            =   "res5"
      Top             =   8040
      Width           =   6735
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5640
      MaskColor       =   &H00800000&
      TabIndex        =   18
      Top             =   8760
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update Questions"
      Height          =   375
      Left            =   3720
      MaskColor       =   &H00800000&
      TabIndex        =   13
      Top             =   8760
      Width           =   1815
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Question"
      Height          =   375
      Left            =   2040
      MaskColor       =   &H00800000&
      TabIndex        =   6
      Top             =   8760
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddQuestion 
      Caption         =   "&Add Question"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MaskColor       =   &H00800000&
      TabIndex        =   5
      Top             =   8760
      Width           =   1815
   End
   Begin VB.TextBox res 
      DataField       =   "Response4"
      DataSource      =   "dbQuestions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   3
      Left            =   3720
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   5640
      Width           =   3375
   End
   Begin VB.TextBox res 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Response3"
      DataSource      =   "dbQuestions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   5640
      Width           =   3135
   End
   Begin VB.TextBox res 
      DataField       =   "Response2"
      DataSource      =   "dbQuestions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   1
      Left            =   3720
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox res 
      BackColor       =   &H00FFFFFF&
      DataField       =   "Response1"
      DataSource      =   "dbQuestions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1815
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3480
      Width           =   3135
   End
   Begin VB.TextBox txtQuestion 
      DataField       =   "Question"
      DataSource      =   "dbQuestions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   2040
      Width           =   7215
   End
   Begin VB.Data dbQuestions 
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   9240
      Width           =   7215
   End
   Begin VB.Frame frmEditInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      TabIndex        =   16
      Top             =   600
      Width           =   7455
      Begin VB.Label lblPos 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   20
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblCurrentEdit 
         BackStyle       =   0  'Transparent
         Caption         =   "Maths : Basic Maths"
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
         TabIndex        =   17
         Top             =   240
         Width           =   7095
      End
   End
   Begin VB.Label lblError 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   8400
      Width           =   6735
   End
   Begin VB.Label lblCorrectAnswer 
      Alignment       =   1  'Right Justify
      DataField       =   "ResponseAnswer"
      DataSource      =   "dbQuestions"
      Height          =   255
      Left            =   1800
      TabIndex        =   23
      Top             =   7680
      Visible         =   0   'False
      Width           =   1575
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
      TabIndex        =   15
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label5 
      Caption         =   "   Q u i z    L i b r a r y    E d i t o r : :  Question Editor"
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
      TabIndex        =   14
      Top             =   360
      Width           =   7575
   End
   Begin VB.Label lblRes 
      BackStyle       =   0  'Transparent
      Caption         =   "Correct Answer:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   7680
      Width           =   2775
   End
   Begin VB.Label lblRes 
      BackStyle       =   0  'Transparent
      Caption         =   "Answer 4:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   11
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label lblRes 
      BackStyle       =   0  'Transparent
      Caption         =   "Answer 3:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label lblRes 
      BackStyle       =   0  'Transparent
      Caption         =   "Answer 2:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3720
      TabIndex        =   9
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblRes 
      BackStyle       =   0  'Transparent
      Caption         =   "Answer 1:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label lblQ 
      BackStyle       =   0  'Transparent
      Caption         =   "Question:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   2895
   End
End
Attribute VB_Name = "frmQuestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddQuestion_Click()
On Error Resume Next
dbQuestions.Recordset.AddNew 'Add New Question
End Sub

Private Sub cmdClose_Click()
On Error Resume Next
If checkclear = True Then 'If clear then...
    CmdDelete_Click 'Delete current record
Else
    dbQuestions.Recordset.MoveLast 'Save Records
End If

Unload Me 'Unload Form
frmEdit.Show 'Load Quiz Library Editor Form

End Sub

Private Sub cmdUpdate_Click()
dbQuestions.Recordset.MoveLast 'Save Results
End Sub

Private Sub CmdDelete_Click()
On Error Resume Next
absdata = dbQuestions.Recordset.AbsolutePosition - 1 'Currently selected question
dbQuestions.Recordset.Delete 'Delete Question
dbQuestions.Recordset.AbsolutePosition = absdata 'Return to question before one deleted.
EditQuestions CurrentSubjectID, CurrentModuleID 'Reload Questions
End Sub

Private Sub dbQuestions_Reposition()
lblPos.Caption = "[" & dbQuestions.Recordset.AbsolutePosition + 1 & "/" & dbQuestions.Recordset.RecordCount & "]" 'Current Question out of total questions
res5.Text = lblCorrectAnswer.Caption 'Set response Combo to database response answer.
checkanswer 'Check Answer
End Sub

Private Sub dbQuestions_Validate(Action As Integer, Save As Integer)
If Action = 4 Then
    txtMod.Text = CurrentModuleID 'Module ID
    txtSub.Text = CurrentSubjectID 'Subject ID
End If
End Sub

Private Sub Form_Load()
lblCurrentEdit.Caption = frmEdit.txtSubjectName.Text & ":" & frmEdit.txtModDescription.Text
frmEdit.Enabled = False 'Enable Quiz Library Editor
For i = 1 To 4 'Populate responses
    res5.AddItem (i)
Next i
res5.ListIndex = 0 'Set Response to (1)


End Sub

Private Sub Form_Unload(Cancel As Integer)
frmEdit.Enabled = True 'Enable Quiz Library Editor
End Sub

Private Sub res_Change(Index As Integer)
checkanswer 'Check Answer
End Sub

Private Sub res_Click(Index As Integer)
checkanswer 'Check Answer
End Sub

Private Sub res_DblClick(Index As Integer)
res5.ListIndex = Index
End Sub

Private Sub res5_Change()
On Error Resume Next
For i = 0 To res.UBound 'Make all boxes white
    res(i).BackColor = vbWhite
Next i

res(res5.Text - 1).BackColor = &HC0C0C0 'Set Selected answer box to grey
If res5.Text = 0 Then res5.Text = 1 'If response is 0 set to default.

End Sub

Private Sub res5_Click()
On Error Resume Next
For i = 0 To res.UBound 'Make all boxes white
    res(i).BackColor = vbWhite 'set white
Next i

res(res5.Text - 1).BackColor = &HC0C0C0 'Set Selected answer box to grey
lblCorrectAnswer.Caption = res5.Text 'update database
checkanswer 'Check Answer
End Sub

Sub checkanswer()
On Error Resume Next
If res(res5.Text - 1).Text = "" Then 'If blank then...
    lblError.Caption = "The correct answer is currently blank!" 'Set caption to inform user that response for correct answer is blank
Else
    lblError.Caption = "" 'Clear caption.
End If
End Sub

Function checkclear()
    checkclear = False 'set boolean to false
    If txtQuestion.Text = "" Then 'If quetion clear then...
        For i = 0 To res.UBound 'Check answers
            check = res(i)
        Next i
        If check = "" Then checkclear = True 'If answers clear then checkclear = True
    End If
End Function

