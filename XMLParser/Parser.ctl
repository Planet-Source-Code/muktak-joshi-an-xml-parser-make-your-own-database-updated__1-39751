VERSION 5.00
Begin VB.UserControl Parser 
   BackColor       =   &H00000000&
   ClientHeight    =   720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   720
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   720
   ScaleWidth      =   720
   ToolboxBitmap   =   "Parser.ctx":0000
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   700
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   700
   End
End
Attribute VB_Name = "Parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim sDataSource As String
Dim FileNum
Dim SelStart As Integer
Dim SelLength As Integer
Dim CurrentPos As Integer
Dim DataStruct() As String
Dim DataBase() As String
Dim RCount As Integer
Dim DataBack() As String

Private Function ResetVariables()
'SelStart = 0
'SelLength = 0
CurrentPos = 0
ReDim DataStruct(0) As String
ReDim DataBase(0) As String
RCount = 0
ReDim DataBack(0) As String
End Function
Public Property Let DataSource(ByVal Source As String)
ResetVariables
If Dir(Source) = "" Then
Err.Raise "1", UserControl.Name, "DataSource Property set to Missing File."
Else
sDataSource = Source
End If


End Property
Public Sub AddNew()
ReDim Preserve DataBase(UBound(DataBase) + 1) As String
CurrentPos = UBound(DataBase)
DataBase(UBound(DataBase)) = String(UBound(DataStruct), "#@#")

End Sub
Public Sub Update()
Dim sLine As String
Dim Count As Integer
Dim ArrTmp
Dim sOut As String

FileNum = FreeFile
Open sDataSource For Input As FileNum
Count = 0
While Not EOF(FileNum)
Line Input #FileNum, sLine
Count = Count + 1
ReDim Preserve DataBack(Count) As String
DataBack(Count) = sLine
Wend

Close FileNum


Open sDataSource For Output As FileNum
For z = 0 To UBound(DataBack)
If (z > SelStart) And (z < SelLength) Then
        For y = 0 To UBound(DataBase)
            ArrTmp = Split(DataBase(y), "#@#")
            sOut = sOut & vbTab & "<entry>" & vbCrLf
            For x = 0 To UBound(ArrTmp) - 1
                ArrTmp(x) = Replace(ArrTmp(x), "#@#", "")
                sOut = sOut & vbTab & vbTab & "<" & DataStruct(x + 1) & ">" & ArrTmp(x) & "</" & DataStruct(x + 1) & ">" & vbCrLf
            Next
            sOut = sOut & vbTab & "</entry>" & vbCrLf
        Next
        Print #FileNum, sOut
        z = SelLength + 2 + UBound(DataStruct)
Else
    
    If DataBack(z) <> "" Then
    Print #FileNum, DataBack(z)
    End If
End If
Next
Close FileNum
End Sub

Private Function GetColIndex(ColName As String) As Integer
Dim Found As Boolean
Dim ColInd As Integer
ColName = Replace(ColName, "<", "")
ColName = Replace(ColName, ">", "")
ColName = LCase(ColName)

For x = 0 To UBound(DataStruct)
DataStruct(x) = LCase(DataStruct(x))

If DataStruct(x) = ColName Then
    Found = True
    ColInd = x
End If
Next
If Found = True Then
GetColIndex = ColInd
Else
Err.Raise "2", UserControl.Name, "Field not Found in Table Structure."
End If
Found = False
End Function
Public Sub OpenTable(TableName As String)
Dim sLine As String
Dim sType As String
Dim iCount As Integer
Dim TableFound As Boolean
TableFound = False
FileNum = FreeFile
Open sDataSource For Input As FileNum
While Not EOF(FileNum)
Line Input #FileNum, sLine
sLine = Replace(sLine, vbTab, "")
sLine = Replace(sLine, vbCrLf, "")

sType = GetLeftWord(sLine, ">", True)
Select Case sType
Case "<table name=" & TableName & ">"
    
    
    TableFound = True
Case "</structure>"
Case "<structure>"
    If TableFound Then ReDim Preserve DataStruct(0) As String
Case "<column>"
    If TableFound Then
    ReDim Preserve DataStruct(UBound(DataStruct) + 1) As String
    DataStruct(UBound(DataStruct)) = Left$(sLine, Len(sLine) - 9)
    End If
Case "</column>"
Case "<data>"
    If TableFound Then
        SelStart = iCount + 1
    End If
Case "<entry>"
    If TableFound Then
    ReDim Preserve DataBase(RCount) As String
    End If
Case "</entry>"
    If TableFound Then
    RCount = RCount + 1
    End If
Case "</data>"
    If TableFound Then SelLength = iCount - UBound(DataStruct) - 2
Case "</table>"
    If TableFound Then
    TableFound = False
    End If
Case Else
    If TableFound Then
        If sType <> "" Then
        If sLine <> "" Then
            
            sLine = (Left$(sLine, Len(sLine) - Len(GetRightWord(sLine, "<", False)) - 1))
            'If DataBase(RCount) <> "" Then sLine = "," & sLine
            DataBase(RCount) = DataBase(RCount) & sLine & "#@#"
        End If
        End If
    End If
End Select
iCount = iCount + 1
Wend
iCount = 0
Close FileNum
CurrentPos = 0
TableFound = False
End Sub

Public Property Get RecordCount() As Integer
RecordCount = UBound(DataBase)
End Property



Public Function Field(FieldName As String) As String
Dim Col As Integer
Dim Count As Integer
Dim ArrTmp
Dim Output As String

Col = GetColIndex(FieldName)
Output = DataBase(CurrentPos)
ArrTmp = Split(Output, "#@#")
Output = ArrTmp(Col - 1)
Output = GetLeftWord(Output, "#@#", False)
Field = Output
End Function


Public Sub MoveNext()
If CurrentPos > UBound(DataBase) Then
Err.Raise "3", UserControl.Name, "Parser have reached to the End of the Database"
CurrentPos = 0
Else

CurrentPos = CurrentPos + 1
End If
End Sub
Public Sub MoveFirst()
CurrentPos = 0
End Sub
Public Sub MoveLast()
CurrentPos = UBound(DataBase)
End Sub
Public Sub MovePrevious()
If CurrentPos > 0 Then
CurrentPos = CurrentPos - 1
Else
Err.Raise "4", UserControl.Name, "Parser have reached to the Beginning of Database."
CurrentPos = 0
End If
End Sub

Public Sub UpdateField(Field As String, Value As String)
Dim ArrTmp
If InStr(1, DataBase(CurrentPos), "@", vbTextCompare) = 0 Then
DataBase(CurrentPos) = Replace(DataBase(CurrentPos), "#", "#@#")
End If
ArrTmp = Split(DataBase(CurrentPos), "#@#")
ArrTmp(GetColIndex(Field) - 1) = Value
DataBase(CurrentPos) = ""
For x = 0 To UBound(ArrTmp) - 1
DataBase(CurrentPos) = DataBase(CurrentPos) & ArrTmp(x) & "#@#"
Next
End Sub


Private Sub UserControl_Initialize()
UserControl.Height = 700
UserControl.Width = 700
End Sub

Private Sub UserControl_Resize()
UserControl.Height = 700
UserControl.Width = 700
End Sub
