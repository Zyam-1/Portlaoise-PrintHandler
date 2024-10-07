VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Fastings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolFastings As New Collection

Private mFastings As New Fastings

Public Function Add(ByVal Fx As Fasting) _
                    As Fasting

      Dim sql As String
      Dim tb As Recordset



10    On Error GoTo Add_Error

20    sql = "SELECT * FROM Fastings WHERE TestName = '" & Fx.TestName & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    With Fx
60      If tb.EOF Then
70        tb.AddNew
80      End If

90      tb!TestName = UCase$(.TestName)
100     tb!FastingLow = .FastingLow
110     tb!FastingHigh = .FastingHigh
120     tb!FastingText = .FastingText
130     tb.Update

140     mcolFastings.Add Fx, .TestName
150   End With

160   Set Add = Fx


170   Exit Function

Add_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "Fastings", "Add", intEL, strES, sql


End Function
Public Function Count() As Long
  
10    On Error GoTo Count_Error

20    Count = mcolFastings.Count

30    Exit Function

Count_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "Fastings", "Count", intEL, strES


End Function




Public Sub Delete(ByVal TestName As String)
  
      Dim sql As String



10    On Error GoTo Delete_Error

20    sql = "Delete FROM Fastings WHERE " & _
            "TestName = '" & TestName & "'"
30    Cnxn(0).Execute sql

40    Refresh



50    Exit Sub

Delete_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "Fastings", "Delete", intEL, strES, sql


End Sub

Public Function Item(ByVal TestName As String) As Fasting
Attribute Item.VB_UserMemId = 0

10    On Error GoTo Item_Error

20    Set Item = mcolFastings.Item(TestName)

30    Exit Function

Item_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "Fastings", "Item", intEL, strES


End Function



Private Sub Load()

      Dim tb As Recordset
      Dim sql As String
      Dim Fx As Fasting



10    On Error GoTo Load_Error

20    sql = "SELECT * FROM Fastings"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60      Set Fx = New Fasting
70      With Fx
80        .TestName = UCase$(tb!TestName & "")
90        .FastingLow = tb!FastingLow & ""
100       .FastingHigh = tb!FastingHigh & ""
110       .FastingText = tb!FastingText & ""
120       mcolFastings.Add Fx, .TestName
130     End With
140     tb.MoveNext
150   Loop


160   Exit Sub

Load_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "Fastings", "Load", intEL, strES


End Sub
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    On Error GoTo NewEnum_Error

20    Set NewEnum = mcolFastings.[_NewEnum]

30    Exit Function

NewEnum_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "Fastings", "NewEnum", intEL, strES


End Function


Public Property Get Fastings() As Fastings
  
10    On Error GoTo Fastings_Error

20    Set Fastings = mFastings

30    Exit Property

Fastings_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "Fastings", "Fastings", intEL, strES


End Property

Public Sub Refresh()

10    On Error GoTo Refresh_Error

20    Set mcolFastings = Nothing

30    Load

40    Exit Sub

Refresh_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "Fastings", "Refresh", intEL, strES


End Sub

Private Sub Class_Initialize()

10    On Error GoTo Class_Initialize_Error

20    Load

30    Exit Sub

Class_Initialize_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "Fastings", "Class_Initialize", intEL, strES


End Sub


