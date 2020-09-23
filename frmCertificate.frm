VERSION 5.00
Begin VB.Form frmCertificate 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "advance : Certificate"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   7800
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmOptions 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Certificate Options:"
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   0
      TabIndex        =   8
      Top             =   8280
      Width           =   7815
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print Certificate"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Image Image1 
      Height          =   990
      Left            =   1680
      Picture         =   "frmCertificate.frx":0000
      Top             =   480
      Width           =   3780
   End
   Begin VB.Line Line4 
      X1              =   1440
      X2              =   4680
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "Signed:"
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
      Left            =   600
      TabIndex        =   7
      Top             =   7080
      Width           =   2655
   End
   Begin VB.Line Line3 
      BorderWidth     =   5
      X1              =   720
      X2              =   6960
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "08/07/02"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Well Done!!!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   600
      TabIndex        =   5
      Top             =   6240
      Width           =   6495
   End
   Begin VB.Label lblSubject 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "subject$"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   4920
      Width           =   6615
   End
   Begin VB.Label lblComment 
      BackStyle       =   0  'Transparent
      Caption         =   "Has obtained the best results in:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   3
      ToolTipText     =   "Double Click to Edit."
      Top             =   3960
      Width           =   6495
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "name$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "This is to certify that"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   3360
      Width           =   3615
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   480
      X2              =   7200
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   480
      X2              =   7200
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Certificate of Achievement"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   7455
   End
End
Attribute VB_Name = "frmCertificate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me 'Unload Form
End Sub

Private Sub cmdPrint_Click()
frmOptions.Visible = False 'Make Options not visible
Me.PrintForm 'Print the Form
frmOptions.Visible = True 'Make options visible
End Sub
Private Sub Form_Load()
frmAdmin.Enabled = False 'Enable Admin Form
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblComment.BorderStyle = 0 'Clear Selection
lblSubject.BorderStyle = 0 'Clear Selection
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmAdmin.Enabled = True 'Enable Admin Form
End Sub


Private Sub lblComment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblComment.BorderStyle = 1 'Selection set to Comment box
End Sub

Private Sub lblComment_DblClick()
lblComment.Caption = InputBox("Enter award information:", "advance : Certficate", lblComment.Caption) 'Set Comment
End Sub

Private Sub lblSubject_DblClick()
lblSubject.Caption = InputBox("Enter subject/module for award:", "advance : Certficate", lblSubject.Caption) 'Set Subject
End Sub

Private Sub lblSubject_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblSubject.BorderStyle = 1 'Selection set to Subject box
End Sub
