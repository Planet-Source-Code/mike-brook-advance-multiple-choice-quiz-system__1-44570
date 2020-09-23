VERSION 5.00
Begin VB.Form FrmAddQ 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Question"
   ClientHeight    =   9720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtQuestion 
      CausesValidation=   0   'False
      DataField       =   "Question"
      DataSource      =   "dbAddQ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   120
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2160
      Width           =   7095
   End
   Begin VB.TextBox res 
      CausesValidation=   0   'False
      DataField       =   "Response1"
      DataSource      =   "dbAddQ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3480
      Width           =   3135
   End
   Begin VB.TextBox res 
      CausesValidation=   0   'False
      DataField       =   "Response2"
      DataSource      =   "dbAddQ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   1
      Left            =   3840
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox res 
      CausesValidation=   0   'False
      DataField       =   "Response3"
      DataSource      =   "dbAddQ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   6000
      Width           =   3135
   End
   Begin VB.TextBox res 
      CausesValidation=   0   'False
      DataField       =   "Response4"
      DataSource      =   "dbAddQ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   3
      Left            =   3840
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   6000
      Width           =   3375
   End
   Begin VB.ComboBox res5 
      CausesValidation=   0   'False
      Height          =   315
      ItemData        =   "FrmAddQ.frx":0000
      Left            =   240
      List            =   "FrmAddQ.frx":0002
      TabIndex        =   4
      Text            =   "res5"
      Top             =   8280
      Width           =   6975
   End
   Begin VB.TextBox txtMod 
      CausesValidation=   0   'False
      DataField       =   "ModuleID"
      DataSource      =   "dbAddQ"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   9240
      Width           =   1575
   End
   Begin VB.TextBox txtSub 
      CausesValidation=   0   'False
      DataField       =   "SubjectID"
      DataSource      =   "dbAddQ"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   9240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   8760
      Width           =   1575
   End
   Begin VB.Data dbAddQ 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8760
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
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
      Left            =   360
      TabIndex        =   0
      Top             =   8760
      Width           =   1575
   End
   Begin VB.Frame frmEditInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      TabIndex        =   15
      Top             =   600
      Width           =   7455
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
         TabIndex        =   16
         Top             =   240
         Width           =   7095
      End
   End
   Begin VB.Label lblCorrectAnswer 
      DataField       =   "ResponseAnswer"
      DataSource      =   "dbAddQ"
      Height          =   255
      Left            =   2760
      TabIndex        =   21
      Top             =   8040
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1320
      TabIndex        =   20
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label lblRes 
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
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "   Q u i z    L i b r a r y    E d i t o r : :  Add Question"
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
      Top             =   360
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
      Top             =   0
      Width           =   4215
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
      TabIndex        =   14
      Top             =   3240
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
      Left            =   3840
      TabIndex        =   13
      Top             =   3240
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
      TabIndex        =   12
      Top             =   5760
      Width           =   1095
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
      Left            =   3840
      TabIndex        =   11
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label lblRes 
      BackStyle       =   0  'Transparent
      Caption         =   "Correct Answer Number:"
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
      Left            =   240
      TabIndex        =   10
      Top             =   8040
      Width           =   2775
   End
End
Attribute VB_Name = "FrmAddQ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
MsgBox "Question Added", vbInformation, "advance"
frmQuestions.Enabled = True
Unload frmQuestions

frmEdit.Enabled = True
Unload Me

End Sub

Private Sub Command1_Click()
On Error Resume Next
dbAddQ.Recordset.Delete
Unload Me
frmQuestions.Show
frmQuestions.Enabled = True
End Sub

Private Sub dbAddQ_Reposition()
res5.text = lblCorrectAnswer.Caption
End Sub

Private Sub Form_Load()
For i = 1 To 4
res5.AddItem (i)
Next i
lblCurrentEdit.Caption = frmEdit.txtSubjectName.text & ":" & frmEdit.txtModDescription.text
frmQuestions.Enabled = False
LoadDB "Questions", FrmAddQ.dbAddQ
dbAddQ.Recordset.AddNew

End Sub

Private Sub res5_Click()
lblCorrectAnswer.Caption = res5.text
End Sub
