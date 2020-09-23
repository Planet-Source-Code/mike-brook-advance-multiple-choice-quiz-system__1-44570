VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form LoginFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "advance - Login "
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5145
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   2895
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
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
      Left            =   3360
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
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
      Left            =   1680
      MaskColor       =   &H00800000&
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1680
      TabIndex        =   1
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Data loginDB 
      Caption         =   "Authenicating Login..."
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   780
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Width           =   4980
   End
   Begin VB.Image Image1 
      Height          =   990
      Left            =   0
      Picture         =   "LoginFrm.frx":0000
      Top             =   0
      Width           =   3780
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7320
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblID 
      DataField       =   "StudentID"
      DataSource      =   "loginDB"
      Height          =   255
      Left            =   7080
      TabIndex        =   12
      Top             =   6240
      Width           =   855
   End
   Begin VB.Label lblSname 
      DataField       =   "StudentSname"
      DataSource      =   "loginDB"
      Height          =   495
      Left            =   10560
      TabIndex        =   11
      Top             =   5880
      Width           =   3375
   End
   Begin VB.Label lblFname 
      DataField       =   "StudentFname"
      DataSource      =   "loginDB"
      Height          =   495
      Left            =   10560
      TabIndex        =   10
      Top             =   5160
      Width           =   3375
   End
   Begin VB.Label lblYear 
      DataField       =   "StudentYear"
      DataSource      =   "loginDB"
      Height          =   495
      Left            =   6840
      TabIndex        =   9
      Top             =   5040
      Width           =   3375
   End
   Begin VB.Label lblAge 
      DataField       =   "StudentAge"
      DataSource      =   "loginDB"
      Height          =   495
      Left            =   6840
      TabIndex        =   8
      Top             =   5520
      Width           =   3375
   End
   Begin VB.Label lblUserLevel 
      DataField       =   "UserLevel"
      DataSource      =   "loginDB"
      Height          =   495
      Left            =   6840
      TabIndex        =   7
      Top             =   6000
      Width           =   3375
   End
   Begin VB.Label varPassword 
      DataField       =   "StudentPassword"
      DataSource      =   "loginDB"
      Height          =   495
      Left            =   6960
      TabIndex        =   6
      Top             =   6600
      Width           =   3375
   End
   Begin VB.Label lblPassword 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "LoginFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdLogin_Click()
LoginUser 'login process
End Sub

Sub LoginUser()

'### LoginUser Function ###
' This function checks that the user exists within the database, and that the credentials used are correct





status.SimpleText = "Logging in..." 'Notify user of status
cmdLogin.Enabled = False 'Disable Login button to allow LoginUser function to complete.
LoadDB "SELECT StudentPassword, UserLevel, StudentID, StudentAge, StudentYear, StudentFname, StudentSname from Students where [StudentUserName] = '" & LoginFrm.txtUser.text & "';", LoginFrm.loginDB 'Query Students table to check if login info correct
cmdLogin.Enabled = True 'Enable Login Button
        
On Error GoTo ErrorHnd 'If the user is not found in the database, handle error



    
    If varPassword.Caption = txtPass.text Then 'Check for correct password

        LoginProcess 'Call to function which processes login
    Else
        status.SimpleText = "Incorrect Password" 'Notify user of status
        MsgBox "Incorrect Password", vbExclamation, "advance" 'Notify user of status
        txtPass.text = "" 'Clear password
    End If
ErrorHnd:




If Err.Number = 13 Then
    status.SimpleText = "Incorrect Login Details" 'set status bar as last error message
    MsgBox "Incorrect Login Details", vbExclamation, "advance" 'error message
    txtPass.text = "" 'clear password box
End If


End Sub


Private Sub cmdQuit_Click()
Unload Me 'unload form
End 'end application
End Sub


Sub LoginProcess()
'set variables for session
    StudentUserName = txtUser.text 'Users Username
    StudentAge = lblAge.Caption 'Users age
    StudentYear = lblYear.Caption 'Users Year
    StudentFname = lblFName.Caption 'Users Forename
    StudentSName = lblSName.Caption 'Users surname
    StudentID = lblID.Caption 'Users ID
    UserLevel = lblUserLevel.Caption 'Users Level
    status.SimpleText = "OK" 'Set status bar as OK
    Unload Me 'unload form
    frmShowPic.Show 'load welcome form
End Sub

Private Sub Form_Load()
back.Show 'load background form
status.SimpleText = appinfo 'set status bar as copyright info
End Sub

Private Sub txtPass_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then 'if user presses enter/return
    LoginUser 'login process
End If
End Sub
