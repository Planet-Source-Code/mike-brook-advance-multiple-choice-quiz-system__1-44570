VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmResultView 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "advance - results ["
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   9255
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame gradeBox 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   7800
      TabIndex        =   18
      Top             =   0
      Width           =   1095
      Begin VB.Label gradex 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "A"
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
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   8655
      Begin MSChart20Lib.MSChart progress 
         Height          =   6375
         Left            =   120
         OleObjectBlob   =   "frmResultView.frx":0000
         TabIndex        =   2
         Top             =   240
         Width           =   8415
      End
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
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
      Left            =   2160
      TabIndex        =   15
      Top             =   9120
      Width           =   1695
   End
   Begin VB.TextBox txtAv 
      DataField       =   "AverageX"
      DataSource      =   "dbAvg"
      Height          =   285
      Left            =   4680
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data dbAvg 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
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
      Left            =   360
      TabIndex        =   8
      Top             =   9120
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Summary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   240
      TabIndex        =   3
      Top             =   7680
      Width           =   8655
      Begin VB.Frame frmSetting 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Population Analysis:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3840
         TabIndex        =   10
         Top             =   240
         Width           =   4575
         Begin VB.Label lblComment 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   600
            Width           =   4215
         End
         Begin VB.Label lblAP 
            BackStyle       =   0  'Transparent
            Height          =   255
            Left            =   2760
            TabIndex        =   12
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Student Average Percentage:"
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Label lblTrendr 
         BackStyle       =   0  'Transparent
         Caption         =   "Improving"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label lblTrend 
         BackStyle       =   0  'Transparent
         Caption         =   "Trend:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label lblAv 
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
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
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lblAve 
         BackStyle       =   0  'Transparent
         Caption         =   "Average Result:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   4455
      End
   End
   Begin VB.Data results 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   11520
      Width           =   5655
   End
   Begin MSFlexGridLib.MSFlexGrid resultGrid 
      Bindings        =   "frmResultView.frx":1995
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   10320
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   2355
      _Version        =   393216
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Results Analysis"
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
      TabIndex        =   17
      Top             =   0
      Width           =   4215
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "for X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label5 
      Caption         =   "  Summary of Results and History"
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
      TabIndex        =   16
      Top             =   360
      Width           =   9255
   End
End
Attribute VB_Name = "frmResultView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim resultsArr() As Long 'Results Array
Dim StudentAv As Integer 'Student Average

Private Sub CmdClose_Click()
Unload Me 'Unload Form
End Sub

Private Sub cmdPrint_Click()
cmdPrint.Visible = False 'Disable Print Button
CmdClose.Visible = False 'Disable Close Button
PrintForm 'Print the form
CmdClose.Visible = True 'After sending to printer reenable
cmdPrint.Visible = True 'After sending to printer reenable
End Sub


Private Sub Form_Load()

If VdataStudent = "" Then 'Check if being viewed by student or teacher
    lblName.Caption = "for " & StudentFname & " " & StudentSName 'Set Graph caption to Student name
Else
    lblName.Caption = "for " & VdataStudent 'Set graph to selected student name (In Teacher Mode)
End If

If UserLevel = "STUDENT" Then 'If a student..
    errMess = "No results in this subject module." 'set error message
ElseIf UserLevel = "TEACHER" And VdataStudent <> "" Then 'If a Teacher and No Student Selected...
    errMess = "No results in this subject module, for " & VdataStudent & "." 'set error message
Else
    errMess = "No student is selected, please select a student using 'Find Student'." 'set error message
End If

MainMenu.Enabled = False 'Disable Main Menu
Me.Caption = Me.Caption & CurrentSubject & "  |  " & CurrentModule & "]" 'set caption
resultsGrb CurrentSubjectID, CurrentModuleID 'grab the results for the user
If results.Recordset.RecordCount = 0 Then 'Check to see no results for the user in this module and subject
    MsgBox errMess, vbInformation, "advance" 'message to inform user
    Unload Me 'Unload Form
Else
    LoadDB "SELECT Avg([SubjectScore]) as AverageX from StudentScores WHERE [SubjectModuleID] =" & CurrentModuleID & ";", dbAvg 'grab the average score by student in this module and subject
    GraphIt 'Graph Results
    statsengine 'Calculate Stats
End If
End Sub


Sub GraphIt()
Dim cscore As Long 'Score
Dim Tscore As Long 'Total
Dim LabelDate As String 'The Date
progress.Title = "Results History for:" & CurrentSubject & "  |  " & CurrentModule 'set graph title
resultGrid.Col = 1 'set column to 1 on flex grid
progress.ColumnCount = 1 'graph setup
progress.RowCount = results.Recordset.RecordCount 'Total Number of Results to grid
ReDim resultsArr(results.Recordset.RecordCount) 'Re dim array to contain all results

For i = 1 To results.Recordset.RecordCount 'loop until all results records processed and added to graph

    resultGrid.Row = i 'set row as current record
    resultGrid.Col = 1
    cscore = resultGrid.Text 'Score
    resultGrid.Col = 2
    Tscore = resultGrid.Text 'Total
    resultGrid.Col = 3
    LabelDate = resultGrid.Text 'Date of Results
    progress.Row = i 'current row on Graph
    progress.Column = 1 'Column on Graph
    progress.RowLabel = LabelDate 'Date of Result
    resultsArr(i) = (cscore / Tscore) * 100 'Fill Array with percentage
    progress.Data = (cscore / Tscore) * 100 'Plot percentage on Graph
    
Next i

StudentAv = ((txtAv.Text / Tscore) * 100) 'Student Average in this module/subject
lblAP.Caption = StudentAv & "%" 'Update Caption




End Sub


Sub statsengine()
Dim ResultTotal As Long 'Results Total
Dim actresult As Integer 'The actual result by student
ResultTotal = 0 'Clear Total
For i = 1 To UBound(resultsArr) 'Loop to calculate total results
    ResultTotal = ResultTotal + resultsArr(i) 'Add result
Next i
    
actresult = (ResultTotal / UBound(resultsArr)) 'Work Out average result
lblAv.Caption = actresult & "%" 'Update Caption
    
'the fantastic trend engine
Dim firstresult As Long 'First Result
Dim midresult As Long 'Middle Result
Dim endresult As Long 'Last Result

firstresult = resultsArr(1) 'First Result
midresult = (resultsArr(CInt(UBound(resultsArr) / 2))) * 2 'Middle
endresult = resultsArr(UBound(resultsArr)) 'Last Result
        
        
diff1 = midresult - firstresult 'Difference between First and Middle Result
diff2 = endresult - midresult 'Difference between Middle and End Result
        
variancet = diff1 + diff2 'The variance between the two differences.
        
If variancet < 0 Then 'If variance less than 0 then..
        lblTrendr = "Deteriorating Results" 'trend info
ElseIf variancet > 0 Then 'If more that 0 then...
        lblTrendr = "Improving Results" 'trend info
ElseIf variancet = 0 Then 'If equal then..
        lblTrendr = "Consistent Results" 'trend info
End If
        
        
        ReDim resultsArr(0) 'Clear Array
        gradex.Caption = grade(actresult) 'Grade Average Result
         
        If StudentAv < actresult Then 'If Student Average more than population average then...
                lblComment.Caption = "Student is above average in year group." 'Comment.
        ElseIf StudentAv = actresult Then 'If Student Average the same as population average then...
                lblComment.Caption = "Student is average in year group." 'Comment.
        Else 'If Student Average less than poppulation average then...
                lblComment.Caption = "Student is below average in year group." 'Comment.
        End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
MainMenu.Enabled = True 'Enable Main Menu
End Sub

