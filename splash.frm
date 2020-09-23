VERSION 5.00
Begin VB.Form splash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   7305
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSplash 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5400
      Top             =   1800
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1920
      Width           =   7335
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7320
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Image Image1 
      Height          =   1890
      Left            =   0
      Picture         =   "splash.frx":0000
      Top             =   0
      Width           =   7275
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
appinfo = "advance v2.30 :: kyro (c) 2003" 'copyright and software information
lblInfo.Caption = appinfo 'sets a label to appinfo
checkdb 'checks database
End Sub

Private Sub tmrSplash_Timer()
Unload Me 'unload form
LoginFrm.Show 'load login form
End Sub

Sub checkdb()

'### Check Database Function ###
'This function ensures that the main database that the application uses is present.

top:
dircheck = Dir(App.Path & "\main.mdb") 'returns name of file if found
If UCase(dircheck) = "MAIN.MDB" Then 'if found then
Else
    response = MsgBox("Advance Cannot Find/Access the 'main.mdb' database. Please locate file, and copy in application directory and click RETRY.", vbRetryCancel, "advance") 'error message
    If response = vbRetry Then GoTo top 'allow user to retry
    MsgBox "Database Error: Advance will now close.", vbCritical, "advance" 'error message
    End
End If
tmrSplash.Enabled = True 'load splash timer

End Sub
