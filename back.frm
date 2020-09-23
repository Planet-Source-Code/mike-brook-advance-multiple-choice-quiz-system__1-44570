VERSION 5.00
Begin VB.Form back 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6540
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3000
      Top             =   3840
   End
   Begin VB.Image kyro 
      Height          =   1530
      Left            =   0
      Picture         =   "back.frx":0000
      Top             =   7680
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   7500
      Left            =   0
      Picture         =   "back.frx":0BC2
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "back"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Image1.Left = Me.Width - Image1.Width 'Position Logo
Image1.top = Me.Height - Image1.Height - 100 'Position Logo
kyro.top = Me.Height - kyro.Height 'Position Company Name
kyro.Left = 0 'Position Company Name
End Sub



