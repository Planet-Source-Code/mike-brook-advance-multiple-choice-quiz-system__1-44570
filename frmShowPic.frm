VERSION 5.00
Begin VB.Form frmShowPic 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Welcome to Advance"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   8775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrWelcome 
      Interval        =   2000
      Left            =   7920
      Top             =   2280
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "enjoy your session..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   480
      TabIndex        =   2
      Top             =   2520
      Width           =   7335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000B&
      BorderWidth     =   3
      X1              =   360
      X2              =   8160
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "a d v a n c e"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   8175
   End
   Begin VB.Label lblWelcome 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7335
   End
End
Attribute VB_Name = "frmShowPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
lblWelcome.Caption = "Welcome, " & StudentFname & " to" 'welcomes user to the software
End Sub

Private Sub tmrWelcome_Timer()
Unload Me 'unload form
MainMenu.Show 'load menu
End Sub
