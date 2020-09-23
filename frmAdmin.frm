VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmAdmin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "advance - Administration"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11640
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11640
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmDev 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Awards and Development"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   1440
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   3135
      Begin VB.CommandButton cmdTopScore 
         Caption         =   "Fi&nd Top Student/Worst Student"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   2895
      End
      Begin VB.CommandButton cmdCert 
         Caption         =   "&Generate Certificate for Current User"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   2895
      End
   End
   Begin VB.Frame frmUsers 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Users"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   4695
      Begin VB.CheckBox ckAwards 
         Caption         =   "Awa&rds"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox ckFilter 
         Caption         =   "&Filter"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   5880
         Width           =   1455
      End
      Begin VB.Frame filter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Filter Options:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1215
         Left            =   1440
         TabIndex        =   28
         Top             =   4560
         Visible         =   0   'False
         Width           =   3015
         Begin VB.CommandButton cmdNosearch 
            Caption         =   "&Filter Users"
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
            Left            =   480
            TabIndex        =   30
            Top             =   720
            Width           =   2055
         End
         Begin VB.ComboBox cboFilter 
            Height          =   315
            ItemData        =   "frmAdmin.frx":0000
            Left            =   480
            List            =   "frmAdmin.frx":0019
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   360
            Width           =   2055
         End
      End
      Begin ComctlLib.ListView UserList 
         Height          =   5415
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   9551
         View            =   1
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         SmallIcons      =   "imageList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.CommandButton cmdDel 
         Caption         =   "&Delete User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   26
         Top             =   5880
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add New User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   5880
         Width           =   1455
      End
   End
   Begin VB.Data userView 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   465
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Width           =   5775
   End
   Begin VB.Frame users 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "User Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   5040
      TabIndex        =   0
      Top             =   840
      Width           =   6375
      Begin VB.CommandButton vdata 
         BackColor       =   &H00E0E0E0&
         Caption         =   "View Current User Results"
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
         Left            =   360
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   5640
         Width           =   2535
      End
      Begin VB.Frame userdetails 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "User Details:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2055
         Left            =   360
         TabIndex        =   9
         Top             =   2880
         Width           =   5775
         Begin VB.TextBox txtYr 
            DataField       =   "StudentYear"
            DataSource      =   "userView"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   17
            Top             =   1440
            Width           =   3375
         End
         Begin VB.TextBox txtAge 
            DataField       =   "StudentAge"
            DataSource      =   "userView"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   15
            Top             =   1080
            Width           =   3375
         End
         Begin VB.TextBox txtSname 
            DataField       =   "Studentsname"
            DataSource      =   "userView"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   13
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox txtFName 
            DataField       =   "StudentFname"
            DataSource      =   "userView"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   11
            Top             =   360
            Width           =   3375
         End
         Begin VB.Label lblYear 
            BackStyle       =   0  'Transparent
            Caption         =   "Year:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   360
            TabIndex        =   16
            Top             =   1440
            Width           =   3015
         End
         Begin VB.Label lblAge 
            BackStyle       =   0  'Transparent
            Caption         =   "Age:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   360
            TabIndex        =   14
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label lblSName 
            BackStyle       =   0  'Transparent
            Caption         =   "Surname:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   360
            TabIndex        =   12
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label lblFName 
            BackStyle       =   0  'Transparent
            Caption         =   "Forename:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame logindetails 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Login Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1935
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   5655
         Begin VB.TextBox txtPass 
            DataField       =   "StudentPassword"
            DataSource      =   "userView"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            TabIndex        =   8
            Top             =   720
            Width           =   3375
         End
         Begin VB.TextBox txtUser 
            DataField       =   "StudentUserName"
            DataSource      =   "userView"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2040
            MaxLength       =   15
            TabIndex        =   7
            Top             =   360
            Width           =   3375
         End
         Begin VB.ComboBox uLevel 
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmAdmin.frx":005B
            Left            =   2040
            List            =   "frmAdmin.frx":0065
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1080
            Width           =   3375
         End
         Begin VB.Label lblULevel 
            DataField       =   "UserLevel"
            DataSource      =   "userView"
            Height          =   255
            Left            =   1920
            TabIndex        =   37
            Top             =   1560
            Visible         =   0   'False
            Width           =   3255
         End
         Begin VB.Label lblID 
            BackStyle       =   0  'Transparent
            Caption         =   "Label7"
            DataField       =   "StudentID"
            DataSource      =   "userView"
            Height          =   375
            Left            =   5640
            TabIndex        =   18
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblUserLevel 
            BackStyle       =   0  'Transparent
            Caption         =   "User Type:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   480
            TabIndex        =   6
            Top             =   1080
            Width           =   3015
         End
         Begin VB.Label lblPassword 
            BackStyle       =   0  'Transparent
            Caption         =   "Password:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   480
            TabIndex        =   4
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label lblUser 
            BackStyle       =   0  'Transparent
            Caption         =   "Username:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   480
            TabIndex        =   3
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
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
         Left            =   3120
         TabIndex        =   1
         Top             =   5640
         Width           =   3015
      End
      Begin VB.Label lblUserName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   360
         TabIndex        =   36
         Top             =   240
         Width           =   5655
      End
   End
   Begin MSFlexGridLib.MSFlexGrid userNameGrid 
      Bindings        =   "frmAdmin.frx":007B
      Height          =   735
      Left            =   360
      TabIndex        =   21
      Top             =   10320
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   1296
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   27
      Top             =   7320
      Width           =   2055
   End
   Begin ComctlLib.ImageList imageList 
      Left            =   9240
      Top             =   9120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmAdmin.frx":0092
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmAdmin.frx":03AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label9 
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
      TabIndex        =   20
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label Label8 
      Caption         =   "   Manage Users and View Results"
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
      TabIndex        =   19
      Top             =   360
      Width           =   11655
   End
   Begin VB.Line Line1 
      X1              =   -120
      X2              =   8520
      Y1              =   600
      Y2              =   600
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
'
'
'
'
'
'
'
'
'
'
'
'

Private Sub ckAwards_Click()
If ckAwards.Value = 0 Then 'If not selected then...
    frmDev.Visible = False 'Disable Awards and Dev Frame
Else
    frmDev.Visible = True 'Enable Awards and Dev Frame
End If
End Sub

Private Sub ckFilter_Click()
If ckFilter.Value = 0 Then 'If not selected then...
    filter.Visible = False 'Disable Filter Frame
Else
    filter.Visible = True 'Enable Filter Frame
End If
End Sub

Private Sub cmdAdd_Click()
On Error Resume Next
userView.Recordset.AddNew 'Add New User
cmdAdd.Enabled = False 'Disable Add Button
UserList.Enabled = False 'Disable User List, so that new user is fully setup.
PopulateUsers 'Populate User List
cmdDel.Caption = "&Cancel"
End Sub

Private Sub cmdCert_Click()
frmCertificate.Show 'Load Cerficate Form
frmCertificate.lblName.Caption = Me.txtFName & " " & Me.txtSname 'Set Caption as name of student on certificate
frmCertificate.lblSubject.Caption = CurrentSubject & " | " & CurrentModule 'Set subject caption on certificate
frmCertificate.lblDate.Caption = Date 'Set date caption on certificate
End Sub

Private Sub CmdClose_Click()
If checkrecord = True Then Unload Me 'If record ok allow user to unload form
End Sub

Private Sub cmdDel_Click()
Dim SID As Long 'Subject ID
If cmdDel.Caption = "&Cancel" Then 'Check to see if in 'Add New User' or 'Normal' mode
    UserList.Enabled = True 'Disable User List
    userView.Recordset.MoveLast 'Move to end of record set
    userView.Recordset.Delete 'Delete record
    userView.Recordset.MoveFirst 'Move to start of record set
Else
    SID = lblID.Caption 'SID from database
    response = MsgBox("Are you sure you want to delete user? All their results will also be removed!", vbYesNo, "advance") 'Confirmation to user that all users results will be deleted.
    If response = vbYes Then 'If Yes then...
        userView.Recordset.Delete 'Delete Student
        LoadDB "Students", userView 'Reload Students Table
        PopulateUsers 'Populate users into User List
    End If
End If
End Sub


Private Sub cmdNosearch_Click()
If cboFilter.Text = "All Users" Then 'If All Users then
    LoadDB "Students", frmAdmin.userView 'Load Student Table
    PopulateUsers 'Populate Users into User List
    users.Caption = "User Administration" 'Set Caption
Else
    optionx = cboFilter.Text 'Set Filter Type to OptionX
    Select Case optionx
        Case "Forename" 'If Forename filter then...
            fieldname = "StudentFname" 'Set Database Field
        Case "Surname" 'If Surname filter then...
            fieldname = "StudentSName" 'Set Database Field
        Case "Username" 'If Username filter then...
            fieldname = "StudentUsername" 'Set Database Field
        Case "User Type" 'If User Type filter then...
            fieldname = "UserLevel" 'Set Database Field
        Case "Year" 'If Year type filter then...
            fieldname = "StudentYear" 'Set Database field
        Case "Age" 'If Age filter then...
        fieldname = "StudentAge" 'Set Database field
    End Select

    searchfor = InputBox(cboFilter.Text, "advance: User Search") 'Search Box
    If searchfor <> "" Then 'If Search String not equal "" then...
        LoadDB "SELECT * FROM Students WHERE [" & fieldname & "] ='" & searchfor & "';", frmAdmin.userView 'Search the Student Table
        users.Caption = "User Administration [for " & optionx & " :" & searchfor & "]" 'Set caption stating query
        PopulateUsers 'Populate users into User List
    End If
End If
End Sub

Private Sub cmdTopScore_Click()
frmTopScore.Show 'Load Top Score Form
End Sub

Private Sub cmdUpdate_Click()
'Check for full update
checkrecord
cmdDel.Caption = "&Delete"
End Sub
Private Sub Form_Load()
frmAdminChoice.Visible = False 'Admin Menu not visible
frmAdminChoice.Enabled = False 'Disable Admin Menu
StudentList 'Load Student Table
PopulateUsers 'Populate users into User List
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmAdminChoice.Visible = True 'Make Admin Menu visible
frmAdminChoice.Enabled = True 'Enable Admin Menu
End Sub

Private Sub txtFName_Change()
updateheader 'Update Header
End Sub

Private Sub txtFName_GotFocus()
lblFName.Font.Underline = True 'Underline Forename Label
End Sub

Private Sub txtFName_LostFocus()
lblFName.Font.Underline = False 'DeUnderline Forename Label
End Sub

Private Sub txtPass_Change()
updateheader 'Update Header
End Sub

Private Sub txtPass_GotFocus()
lblPassword.Font.Underline = True 'Underline Password Label
End Sub

Private Sub txtPass_LostFocus()
lblPassword.Font.Underline = False 'DeUnderline Password Label
End Sub

Private Sub txtSname_Change()
updateheader 'Update Header
End Sub

Private Sub txtSname_GotFocus()
lblSName.Font.Underline = True 'Underline Surname Label
End Sub

Private Sub txtSname_LostFocus()
lblSName.Font.Underline = False 'DeUnderline Surname Label
End Sub

Private Sub txtUser_Change()
updateheader 'Update Header
End Sub

Private Sub txtUser_GotFocus()
lblUser.Font.Underline = True  'Underline User Label
End Sub
Private Sub txtUser_LostFocus()
lblUser.Font.Underline = False 'DeUnderline User Label
End Sub

Private Sub uLevel_Click()
lblULevel.Caption = uLevel.Text 'Update Database from Combo
End Sub

Private Sub UserList_Click()
On Error Resume Next
If UserList.SelectedItem <> "" Then 'If User Selected = "" then...
    userView.Recordset.Index = "StudentID" 'Set Index to StudentID
    userView.Recordset.Seek "=", Val(Mid(UserList.SelectedItem.Key, 2)) 'Find Student in database from userlist
End If
End Sub

Private Sub userView_Reposition()

updateheader 'Update Header
If Me.txtUser.Text = StudentUserName Then 'If the currently logged in user is selected...
    cmdDel.Enabled = False 'Disable Delete Button
Else
    cmdDel.Enabled = True 'Enable Delete Button
End If

If lblULevel.Caption = "" Then 'If User level = "" then...
    lblULevel.Caption = "STUDENT" 'Set to Student default.
End If

uLevel.Text = lblULevel.Caption 'Update Combo Box

End Sub

Private Sub vDatA_Click()
VdataStudent = txtFName.Text & " " & txtSname.Text 'Student Information
StudentID = lblID.Caption 'Get Student ID
MainMenu.container.Caption = "Results for:" & VdataStudent 'Update Container Caption on main form
Unload Me 'Unload Form
MainMenu.Enabled = True
Unload frmAdminChoice
MainMenu.Show
End Sub

Sub PopulateUsers()
    'loop
    On Error Resume Next
    selindex1 = UserList.SelectedItem.Index 'Set Index
    UserList.ListItems.Clear 'Clear User List
For i = 1 To userView.Recordset.RecordCount 'loop through all users
    userNameGrid.Col = 2
    userNameGrid.Row = i
    StudentName = userNameGrid.Text 'Student Name
    userNameGrid.Col = 1
    CStudentID = userNameGrid.Text 'Student ID
    userNameGrid.Col = 6
    userTypez = userNameGrid.Text 'User Type
    
    If userTypez = "STUDENT" Then 'If a Student then...
        UserList.ListItems.Add i, "z" & CStudentID, StudentName, , 2 'Add user to userlist with Student Icon
    Else
        UserList.ListItems.Add i, "z" & CStudentID, StudentName, , 1 'Add user to userlist with Teacher Icon
    End If
Next i

End Sub

Sub updateheader()
lblUserName.Caption = txtFName.Text & " " & txtSname.Text & " [" & Me.txtUser.Text & "]" 'Update Caption
End Sub

Function checkrecord()
checkrecord = False 'Set Check record to false
'Check for full update
If Me.txtUser.Text = "" Then 'If the current record has no user name then...
    MsgBox "Please enter a username for the user.", vbInformation, "Information:" 'inform user
      Me.txtUser.SetFocus 'Set Focus to that field
ElseIf Me.txtPass.Text = "" Then 'If the current record has no password then...
    MsgBox "Please enter a password for the user.", vbInformation, "Information:" 'inform user
    Me.txtPass.SetFocus 'Set focus to that field
ElseIf txtFName.Text = "" Then 'If the current record has no forename then...
    MsgBox "Please enter a forename for the user.", vbInformation, "Information:" 'inform user
      Me.txtFName.SetFocus 'Set focus to that field
ElseIf txtSname.Text = "" Then 'If the current record has no surname then...
    MsgBox "Please enter a surname for the user.", vbInformation, "Information:" 'inform user
      Me.txtSname.SetFocus 'Set focus to that field
Else
    BookMarkrec = userView.Recordset.PercentPosition 'Save Book Mark

    On Error Resume Next
    If lblULevel.Caption = "" Then 'If userlevel = "" then...
        lblULevel.Caption = "STUDENT" 'Set to default Student
    End If

    userView.Recordset.Update 'Update Database
    StudentList 'Load Students Table
    PopulateUsers 'Populate users into User List
    cmdAdd.Enabled = True 'Allow user to add more users
    userView.Recordset.PercentPosition = BookMarkrec 'Goto saved Bookmark
    checkrecord = "TRUE" 'Check Record set to true
End If
End Function
