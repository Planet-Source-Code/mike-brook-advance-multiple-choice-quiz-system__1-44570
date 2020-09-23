VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form quizEngine 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "advance - Quiz"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7305
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame results 
      BackColor       =   &H00FFFFFF&
      Height          =   5295
      Left            =   0
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Frame frmRight 
         BackColor       =   &H00FFFFFF&
         Height          =   4695
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   6375
         Begin VB.Timer tmrUpdate 
            Enabled         =   0   'False
            Interval        =   150
            Left            =   5520
            Top             =   2520
         End
         Begin VB.Line Line5 
            X1              =   1080
            X2              =   5520
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Correct!"
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
            Height          =   495
            Left            =   960
            TabIndex        =   21
            Top             =   2040
            Width           =   4455
         End
         Begin VB.Label correctv 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   27.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   615
            Left            =   360
            TabIndex        =   20
            Top             =   1080
            Width           =   5535
         End
         Begin VB.Line Line4 
            X1              =   1080
            X2              =   5520
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label lblYouGot 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "You got"
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
            Height          =   495
            Left            =   960
            TabIndex        =   19
            Top             =   360
            Width           =   4455
         End
      End
      Begin VB.CommandButton cmdEnd 
         Caption         =   "&Close"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   17
         Top             =   4440
         Width           =   6015
      End
      Begin VB.Line Line6 
         BorderWidth     =   5
         X1              =   480
         X2              =   5880
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Grade:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   480
         TabIndex        =   26
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Label gradex 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   735
         Left            =   3120
         TabIndex        =   25
         Top             =   3600
         Width           =   2775
      End
      Begin VB.Line Line3 
         BorderWidth     =   5
         X1              =   480
         X2              =   5880
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblx2 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1695
         Left            =   3480
         TabIndex        =   16
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label percentv 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "99"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   99.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3015
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label lblx 
         BackStyle       =   0  'Transparent
         Caption         =   "You scored:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   720
         TabIndex        =   14
         Top             =   480
         Width           =   5775
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   4920
      Width           =   6855
      Begin MSComctlLib.ProgressBar progress 
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Stats 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   6120
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label lblAnswersR 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   480
         TabIndex        =   12
         Top             =   720
         Width           =   195
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   2160
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RIGHT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblPercent 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1440
         TabIndex        =   7
         Top             =   720
         Width           =   675
      End
   End
   Begin VB.Data questions 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid questionGrid 
      Bindings        =   "quizEngine.frx":0000
      Height          =   1095
      Left            =   840
      TabIndex        =   5
      Top             =   7200
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   1931
      _Version        =   393216
   End
   Begin VB.CommandButton answer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   3
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   3255
   End
   Begin VB.CommandButton answer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   2
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3480
      UseMaskColor    =   -1  'True
      Width           =   3255
   End
   Begin VB.CommandButton answer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   3720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   3255
   End
   Begin VB.CommandButton answer 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   3255
   End
   Begin VB.CommandButton cmdEndQuiz 
      Caption         =   "&Abandon Quiz"
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
      Left            =   5400
      TabIndex        =   24
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label lblMod 
      Caption         =   "  Maths : Basic Maths"
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
      TabIndex        =   23
      Top             =   360
      Width           =   9255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Advance Quiz"
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
      TabIndex        =   22
      Top             =   0
      Width           =   4215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   5
      X1              =   0
      X2              =   7200
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblQuestion 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   6975
   End
End
Attribute VB_Name = "quizEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NoofQRight As Long
Dim CurrentQuestion As Long
Dim RightAnswerID As Integer
Dim RecordC As Long
Dim Addx As Long



Private Sub answer_Click(Index As Integer)
Dim percentI As Integer
actualresponse = Index + 1

    If actualresponse = RightAnswerID Then 'If the selected answer matches the right answer then..
        NoofQRight = NoofQRight + 1 'Add one ont correct answers
    End If
    
    lblAnswersR.Caption = NoofQRight 'No of questions right
    percentValue = (NoofQRight / RecordC) * 100 'The percentage of the way through the quiz
    lblPercent.Caption = percentValue & "%" 'Set the caption on the form to show the inf
    progress.Max = 100 'set max progress to 100
    
    
    
    
       
    
    
    On Error Resume Next 'In case of percentage error
    progress.Value = (CurrentQuestion / RecordC) * 100 'Set progress
    PopulateBoxes 'Populate Boxes
End Sub

Private Sub cmdEnd_Click()
Unload Me 'Unload Form
End Sub

Private Sub cmdEndQuiz_Click()
response = MsgBox("Are you sure you want to abandon the quiz? Your results so far will be lost!", vbYesNo, "advance : Quiz") 'do you want to exit the quiz
If response = vbYes Then Unload Me 'If yes unload form

End Sub

Private Sub Form_Load()
lblMod.Caption = "  " & CurrentSubject & " : " & CurrentModule 'the tested subject and module
MainMenu.Visible = False 'Disable Main Menu
'Initialise Variables
NoofQRight = 0 'No of Questions Right
CurrentQuestion = 0 'Current Question
RightAnswerID = 0 'The right answer
RecordC = 0 'Record Total Count
progress.Value = 0 'Clear progress
modtotal = 0 'module total

'load the questions for the specific module/subject
LoadQuestions CurrentSubjectID, CurrentModuleID 'Load Questions
RecordC = questions.Recordset.RecordCount 'Total Number of Questions
modtotal = RecordC 'module total
'check to see if the module holds not questions
If RecordC = 0 And lblQuestion.Caption = "" Then  'If no records or if first question blank...
    Unload Me 'Unload Form
    MsgBox "No questions in this module.", vbInformation, "advance" 'error message
Else
'Fill the answers and question into the form
PopulateBoxes 'Populate the command boxes with the answers
End If
End Sub

Sub PopulateBoxes()

On Error GoTo FinishedQuiz 'If finished quiz

CurrentQuestion = CurrentQuestion + 1 ' increment to next question
questionGrid.Col = 1 'set column to 1 on flex grid
questionGrid.Row = CurrentQuestion 'Current Question
lblQuestion.Caption = questionGrid.Text 'Get the question
'loop to grab all the answers and populate the boxes
For i = 2 To 5 'loop though the responses
    questionGrid.Col = i 'for each response
    answer(i - 2).Caption = questionGrid.Text 'populate answer into cmd box
    If answer(i - 2).Caption = "" Then 'checking to see if question includes less that 4 answers
        answer(i - 2).Visible = False 'If so make button disabled
    Else
        answer(i - 2).Visible = True 'If not make it visible
    End If
Next i

questionGrid.Col = 6 'set correct answer column
RightAnswerID = questionGrid.Text 'the correct answer id

GoTo skip:

FinishedQuiz:
  Dim percentx As Integer
  Dim tmrIntx As Integer

tmrIntx = 250
  
  tmrUpdate.Interval = tmrIntx
  tmrUpdate.Enabled = True 'enable correct answer timer
  
    SaveStudentScore StudentID, NoofQRight, CurrentSubjectID, CurrentModuleID 'save student score
       results.Visible = True 'Show result
       percentx = (NoofQRight / RecordC) * 100 'calculate percentage
       percentv.Caption = percentx 'set caption

skip:
End Sub



Private Sub Form_Unload(Cancel As Integer)
MainMenu.Visible = True 'Enable Main Menu
End Sub

Private Sub tmrUpdate_Timer()
Dim percX As Integer

If Addx = NoofQRight + 1 Then
    Me.cmdEnd.Enabled = True 'close button enabled
    percX = percentv.Caption 'set percent
    gradex.Caption = grade(percX) 'calcuate grade


    If percX > 50 And percX < 100 Then 'If Percentage between 5-100 then..
        MsgBox "Well Done!!! Try Again to get a 100%.", vbInformation, "Advance" 'message.
    ElseIf percX <= 50 Then 'If Percentage less than 50...
        MsgBox "Ohhh, don't worry. Try again to get a better result.", vbInformation, "Advance" 'message.
    Else 'If higher
        MsgBox "Wow!!! That was fantastic, keep it up!", vbInformation, "Advance" 'message.
     End If
    

    frmRight.Visible = False  'Disable questions right caption
    tmrUpdate.Enabled = False 'stop timer
    correctv.Caption = 0
    Addx = 0
Else
    If Addx = NoofQRight Then 'If currently displayed result equals total right..
        tmrUpdate.Interval = 2500 'change interval
        Addx = Addx + 1 'add one to current question
    Else
        Addx = Addx + 1 'add one to current question
        correctv.Caption = Addx 'current displayed question updated.
    End If
End If
End Sub





