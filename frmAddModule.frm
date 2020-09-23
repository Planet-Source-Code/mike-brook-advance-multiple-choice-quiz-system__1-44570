VERSION 5.00
Begin VB.Form frmAddModule 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Module"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "C&ancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox txtCurrentSubject 
      DataField       =   "SubjectID"
      DataSource      =   "dbAddMod"
      Height          =   285
      Left            =   3360
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Data dbAddMod 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      EOFAction       =   2  'Add New
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.TextBox txtModName 
      DataField       =   "ModuleDescription"
      DataSource      =   "dbAddMod"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1200
      Width           =   4215
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3240
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Module Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Add Module to Subject:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmAddModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
frmEdit.Enabled = True 'Enable Quiz Library Editor
Unload Me 'Unload Form
End Sub

Private Sub cmdOk_Click()
txtCurrentSubject.Text = CurrentSubjectID 'Set Current Subject ID to database
dbAddMod.Recordset.Update 'Update Database

frmEdit.refreshdata 'Refresh Tables
frmEdit.dbModules.Recordset.MoveLast 'Update
frmEdit.Enabled = True 'Enable Quiz Library Editor
Unload Me 'Unload Form
End Sub



Private Sub Form_Load()
frmEdit.Enabled = False 'Disable Quiz Library Editor
LoadDB "Modules", frmAddModule.dbAddMod 'Load Modules Table
dbAddMod.Recordset.AddNew 'Add New Module Record
End Sub

