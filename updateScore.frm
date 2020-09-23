VERSION 5.00
Begin VB.Form updateScore 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Updating Scores...."
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   4605
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtModt 
      DataField       =   "ModuleTotal"
      DataSource      =   "scores"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3120
      Width           =   3495
   End
   Begin VB.TextBox txtScore 
      DataField       =   "SubjectScore"
      DataSource      =   "scores"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox txtModule 
      DataField       =   "SubjectModuleID"
      DataSource      =   "scores"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2160
      Width           =   3495
   End
   Begin VB.TextBox txtSubjectID 
      DataField       =   "SubjectAreaID"
      DataSource      =   "scores"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1680
      Width           =   3495
   End
   Begin VB.TextBox txtStudentID 
      DataField       =   "StudentID"
      DataSource      =   "scores"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Data scores 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Width           =   4020
   End
End
Attribute VB_Name = "updateScore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
