VERSION 5.00
Begin VB.Form frmCpass 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "advance : Change Password"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   7110
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox stdPass 
      DataField       =   "StudentPassword"
      DataSource      =   "dbCPass"
      Height          =   285
      Left            =   6360
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data dbCPass 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   3480
      TabIndex        =   7
      Top             =   2400
      Width           =   1815
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   2400
      Width           =   1815
   End
   Begin VB.TextBox txtCNewPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox txtNewPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtOldPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   0
      X2              =   7560
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Form1.frx":0000
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblAdmin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Change Password"
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
      TabIndex        =   9
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm New Password:"
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
      Left            =   840
      TabIndex        =   3
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password:"
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
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label lblOldPass 
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password:"
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
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   2895
   End
End
Attribute VB_Name = "frmCpass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
Unload Me 'Unload Form
End Sub

Private Sub cmdChange_Click()
LoadDB "SELECT * FROM STUDENTS WHERE StudentID =" & StudentID, frmCpass.dbCPass 'match the currently logged in user to their corresponding record.

If txtOldPass.text = stdPass.text Then 'check to see that the old password matches with the database

        If txtNewPass.text = txtCNewPass.text Then 'check to see that the new password and the confirmation are the same.
        
            stdPass.text = txtNewPass.text 'change password
            MsgBox "Password Changed!", vbInformation, "Advance" 'inform user of success
            Unload Me
        Else
            MsgBox "Please reenter your new password.", vbExclamation, "Advance" 'advise user to reenter the new password
            txtNewPass.text = "" 'Clear New Password
            txtCNewPass.text = "" 'Clear Confirm New Password
        End If
        
Else
    MsgBox "Your old password is incorrect, please reenter.", vbExclamation, "Advance" 'advise user that old password is incorrect
    txtNewPass.text = "" 'Clear New Password
    txtCNewPass.text = "" 'Clear Confirm New Password
    txtOldPass.text = "" 'Clear Old Password
End If
    



End Sub

Private Sub Form_Unload(Cancel As Integer)
MainMenu.Enabled = True 'Enable Main Menu
End Sub
