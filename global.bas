Attribute VB_Name = "Module1"


'Global Variables
Global StudentUserName As String 'User Username
Global UserLevel As String 'User UserLevel
Global StudentAge As Long 'Students Age
Global StudentYear As String 'Student Yeat
Global StudentFname As String 'User Forename
Global StudentSName As String 'User Surname
Global StudentID As Long 'USer ID
Global CurrentSubjectID As Long 'Subject ID
Global CurrentModuleID As Long 'Module ID
Global appinfo As String 'Application Info (copyright etc)
Global CurrentSubject As String 'Current Subject String
Global CurrentModule As String 'Current Module String
Global VdataStudent As String 'Selected Student Information (in Teacher Mode)
Global modtotal As Long 'Total Mark Possible


Sub LoadQuestions(subjectid As Long, moduleid As Long)
'load questions for a specific subject and module
Dim db As Database, rs As Recordset
Set db = DBEngine.OpenDatabase(App.Path & "\main.mdb") 'load database
Set rs = db.OpenRecordset("SELECT Question,Response1,Response2,Response3,Response4,ResponseAnswer from Questions where [SubjectID] =" & subjectid & "AND [ModuleID] =" & moduleid & ";") 'Set record set to questions from selected subject and module
Set quizEngine.questions.Recordset = rs 'set quiz engine
End Sub

Sub SaveStudentScore(StudentID As Long, score As Long, subjectid As Long, moduleid As Long)
'save score of student
Dim db As Database, rs As Recordset
updateScore.Show 'load update form
Set db = DBEngine.OpenDatabase(App.Path & "\main.mdb") 'load database
Set rs = db.OpenRecordset("StudentScores") 'Set record set as StudentScores
Set updateScore.scores.Recordset = rs 'Pass record set to data control on updateScores form
updateScore.scores.Recordset.AddNew 'Add New Score Record
updateScore.txtSubjectID.Text = subjectid 'Subject ID
updateScore.txtModule.Text = moduleid 'Module ID
updateScore.txtStudentID.Text = StudentID 'Student ID
updateScore.txtScore.Text = score 'Score
updateScore.txtModt.Text = modtotal 'Total
updateScore.scores.Recordset.Update 'Update Record
Unload updateScore 'Unload Form
End Sub

Sub StudentList()
'list students
Dim db As Database, rs As Recordset
Set db = DBEngine.OpenDatabase(App.Path & "\main.mdb") 'Load Database
Set rs = db.OpenRecordset("Students") 'Set record set as Students
Set frmAdmin.userView.Recordset = rs 'Pass record to data control on userView form
End Sub

Sub resultsGrb(subjectid As Long, moduleid As Long)
'load questions for a specific subject and module
Dim db As Database, rs As Recordset
Set db = DBEngine.OpenDatabase(App.Path & "\main.mdb") 'Load database
Set rs = db.OpenRecordset("SELECT SubjectScore,ModuleTotal,ScoreDate from StudentScores where [SubjectAreaID] =" & subjectid & " AND [SubjectModuleID] =" & moduleid & "AND [StudentID] =" & StudentID & ";") 'Set record set to results of selected student in selected subject and module.
Set frmResultView.results.Recordset = rs 'Pass record to data control on resultsviewer form
End Sub


Sub EditSubjects()
'list subject areas
Dim db As Database, rs As Recordset
Set db = DBEngine.OpenDatabase(App.Path & "\main.mdb") 'Load database
Set rs = db.OpenRecordset("Subjects") 'Set record set as Subjects
Set frmEdit.dbSubjects.Recordset = rs 'Pass record to data control on Quiz Library Editor.
End Sub

Sub EditModules(subjectid As Long)
'list subject areas
Dim db As Database, rs As Recordset
Set db = DBEngine.OpenDatabase(App.Path & "\main.mdb") 'Load Database
Set rs = db.OpenRecordset("SELECT ModuleDescription, SubjectID,ModuleID from Modules where [SubjectID] =" & subjectid & ";") 'Set record as to load modules in selected subject
Set frmEdit.dbModules.Recordset = rs 'Pass record to data control on Quiz Library Editor
End Sub
Sub EditQuestions(subjectid As Long, moduleid As Long)
'list subject areas
Dim db As Database, rs As Recordset
Set db = DBEngine.OpenDatabase(App.Path & "\main.mdb") 'Load Database
Set rs = db.OpenRecordset("SELECT * FROM Questions where [ModuleID]=" & moduleid & ";") 'Set record to load questions in selected module
Set frmQuestions.dbQuestions.Recordset = rs 'Pass through to Question Editor

If frmQuestions.dbQuestions.Recordset.RecordCount <= 0 Then 'If record count =< 0 then
    frmQuestions.dbQuestions.Recordset.AddNew 'Add a new record
End If
End Sub




Sub LoadDB(sqlq As String, objx As Object, Optional Prune As Boolean, Optional PruneDay As Long)
Dim db As Database, rs As Recordset
If Prune = True Then 'If is Prune mode...
    Set db = DBEngine.OpenDatabase(App.Path & "\main.mdb") 'Load Database
    wherecond = "(DateDiff('d', Date(), [ScoreDate]) < -" & PruneDay & ");" 'Set Condition for deletion
    sqlStatementx = "DELETE StudentScores.* FROM StudentScores WHERE " & wherecond 'Process into SQL statement
    db.CreateQueryDef "prune", sqlStatementx 'Setup Query
    db.QueryDefs("prune").Execute 'Run Query
    db.QueryDefs.Delete ("prune") 'Remove Query
Else
    Set db = DBEngine.OpenDatabase(App.Path & "\main.mdb") 'Load Database
    Set rs = db.OpenRecordset(sqlq) 'Set record as specified when using LoadDatabase function (SQL)
    Set objx.Recordset = rs 'Pass through to a data control, as specified when using LoadDatabase function (SQL)
End If
End Sub


Sub LockForms(unlockfrm As Object)
'### LockForms ###
' locks all forms that are not in use, this allows processes within the program to be properly followed through.
On Error Resume Next
    frmAddModule.Enabled = False 'Disable Add Module Form
    frmAdmin.Enabled = False 'Disable Admin form
    frmEdit.Enabled = False 'Disable Quiz Editor Form
    frmQuestions.Enabled = False 'Disable Questions Editor
    frmResultView.Enabled = False 'Disable Result Viewer
    frmLoginFrm.Enabled = False 'Disable Login Form
    MainMenu.Enabled = False 'Disable Main Menu
    quizEngine.Enabled = False 'Disable Quiz Engine
    splash.Enabled = False 'Disable Splash Screen
    updateScore.Enabled = False 'Disable Update Score form
    
    'enable
    unlockfrm.Enabled = True 'Enable form pass through to function
End Sub

Sub UnlockForms()
On Error Resume Next
    frmAddModule.Enabled = True 'Enable Add Module Form
    frmAdmin.Enabled = True 'Enable Admin form
    frmEdit.Enabled = True 'Enable Quiz Editor Form
    frmQuestions.Enabled = True 'Enable Questions Editor
    frmResultView.Enabled = True 'Enable Result Viewer
    frmLoginFrm.Enabled = True 'Enable Login Form
    MainMenu.Enabled = True 'Enable Main Menu
    quizEngine.Enabled = True 'Enable Quiz Engine
    splash.Enabled = True 'Enable Splash Screen
    updateScore.Enabled = True 'Enable Update Score form

End Sub

Public Function grade(score As Integer)
    
'### grade function ###
' determines what grade a student has acheived.

Select Case score
    Case Is >= 80 'If score above 80...
        grade = "A" 'Give an A
    Case Is >= 70 'If score above 70...
        grade = "B" 'Give a B
    Case Is >= 60 'If score above 60...
        grade = "C" 'Give a C
    Case Is >= 50 'If score above 50...
        grade = "D" 'Give a D
    Case Is >= 40 'If score above 40...
        grade = "E" 'Give a E
    Case Is >= 30 'If score above 30...
        grade = "F" 'Give a F
    Case Is >= 20 'If score above 20...
        grade = "G" 'Give a G
    Case Else 'Else
        grade = "G" 'Give a G
End Select

End Function

