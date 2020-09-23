VERSION 5.00
Object = "{C8D25CD1-D346-4F11-9528-A624ECBDE057}#1.0#0"; "XMLParser.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2625
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   2625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Add New"
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Data"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2295
   End
   Begin XMLParser.Parser Parser1 
      Left            =   1560
      Top             =   2400
      _ExtentX        =   1244
      _ExtentY        =   1244
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
List1.Clear
Parser1.DataSource = App.Path + "\test.dat"
Parser1.OpenTable "friends"

For x = 0 To Parser1.RecordCount
List1.AddItem Parser1.Field("Name")
Parser1.MoveNext
Next
End Sub

Private Sub Command2_Click()
Parser1.DataSource = App.Path + "\test.dat"
Parser1.OpenTable "friends"

Parser1.AddNew
Parser1.UpdateField "Name", "Entry" & Parser1.RecordCount
Parser1.UpdateField "Age", "19"
Parser1.Update
Command1_Click
End Sub

