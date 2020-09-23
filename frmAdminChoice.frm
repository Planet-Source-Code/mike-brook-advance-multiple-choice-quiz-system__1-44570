VERSION 5.00
Begin VB.Form frmAdminChoice 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Teacher Administration, What do you want to do?"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   7560
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmPrune 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   7575
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back to Teacher Administration"
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
         Left            =   3360
         TabIndex        =   14
         Top             =   3840
         Width           =   2895
      End
      Begin VB.CommandButton cmdPruneRec 
         Caption         =   "Prune Database"
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
         Left            =   720
         TabIndex        =   13
         Top             =   3840
         Width           =   2535
      End
      Begin VB.OptionButton opt1year 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Older than 1 year"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   2640
         Width           =   2895
      End
      Begin VB.OptionButton opt3months 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Older than 3 months"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   2280
         Width           =   2895
      End
      Begin VB.OptionButton opt30days 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Older than 30 days."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   1920
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Prune records that are:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   7095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAdminChoice.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   7095
      End
      Begin VB.Label Label5 
         Caption         =   "   Prune Database : Prune Options"
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
         TabIndex        =   7
         Top             =   0
         Width           =   7575
      End
   End
   Begin VB.CommandButton cmdPrune 
      Caption         =   "&Prune Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
      Width           =   5415
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close Teacher Administration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1080
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6120
      Width           =   5415
   End
   Begin VB.CommandButton cmdQuizLib 
      Caption         =   "Edit &Quiz Library"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   5415
   End
   Begin VB.CommandButton cmdusers 
      Caption         =   "&Manage Users and View Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1080
      MaskColor       =   &H00800000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   5415
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   -120
      X2              =   7560
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAdminChoice.frx":00C1
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   7335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   0
      X2              =   7560
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label lblAdmin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Teacher Administration"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmAdminChoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBack_Click()
frmPrune.Visible = False 'Back to Admin Menu
End Sub

Private Sub cmdPrune_Click()
frmPrune.Visible = True 'Enable Prune menu
End Sub

Private Sub cmdPruneRec_Click()
Dim prunex As Long 'Number of Days to Prune

If opt30days.Value = True Then 'If 30 days selected.
    prunex = 30 'Set Prune Length
    optString = "older than 30 days" 'Set caption
ElseIf opt3months.Value = True Then 'If 1 month selected.
    prunex = 84 'Set Prune Length
    optString = "older than 3 months" 'Set Caption
ElseIf opt1year.Value = True Then 'If 1 year selected.
    prunex = 365 'Set prune length
    optString = "older than 1 year" 'Set Caption.
End If

response = MsgBox("Are you sure you want to prune the database? All student results that are " & optString & ", will be lost!", vbYesNo, "advance : Prune Database") 'Confirm to user that Pruning will remove results
If response = vbYes Then 'If Yes then...
    LoadDB "PruneT", MainMenu.quizList, True, prunex 'Prune Database
    MsgBox "The database has been pruned!!", vbInformation, "Advance" 'Inform user of success
End If
End Sub

Private Sub cmdQuizLib_Click()
frmEdit.Show 'Load Quiz Editor
End Sub

Private Sub cmdusers_Click()
frmAdmin.Show 'Load Admin Form
End Sub


Private Sub CmdClose_Click()
Unload Me 'Unload Form
MainMenu.Enabled = True 'Enable Main Menu
MainMenu.Show 'Bring Main Menu to front
End Sub

Private Sub Form_Load()
MainMenu.Enabled = False 'Disable Main Menu
End Sub

