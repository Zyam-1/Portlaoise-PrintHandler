VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PRNs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolPRNs As New Collection

Private mPRNs As New PRNs

Public Sub Add(ByVal Px As PRN)

      Dim sql As String

10    On Error GoTo Add_Error

20    sql = "IF EXISTS(SELECT * FROM Printers WHERE " & _
            "          MappedTo = '" & Px.MappedTo & "') " & _
            "    Update Printers " & _
            "    SET PrinterName = '" & Px.PrinterName & "' " & _
            "    WHERE MappedTo = '" & Px.MappedTo & "' " & _
            "ELSE " & _
            "    INSERT INTO Printers " & _
            "    (MappedTo, PrinterName) VALUES " & _
            "    ('" & Px.MappedTo & "', " & _
            "     '" & Px.PrinterName & "')"
30    Cnxn(0).Execute sql

40    Exit Sub

Add_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "PRNs", "Add", intEL, strES, sql

End Sub

Public Function Count() As Long
  
10    On Error GoTo Count_Error

20    Count = mcolPRNs.Count

30    Exit Function

Count_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "PRNs", "Count", intEL, strES


End Function

Public Sub Delete(ByVal MappedTo As String)
  
      Dim sql As String
      Dim tb As Recordset



10    On Error GoTo Delete_Error

20    sql = "SELECT * FROM PRNs WHERE " & _
            "MappedTo = '" & MappedTo & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If Not tb.EOF Then
60      tb.Delete
70      mcolPRNs.Remove MappedTo
80    End If


90    Exit Sub

Delete_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "PRNs", "Delete", intEL, strES, sql


End Sub


Public Function Item(ByVal MappedTo As String) As PRN
Attribute Item.VB_UserMemId = 0
  
10    On Error GoTo Item_Error

20    On Error Resume Next

30    Set Item = mcolPRNs.Item(MappedTo)

40    Exit Function

Item_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "PRNs", "Item", intEL, strES


End Function

Private Sub Load()

      Dim tb As Recordset
      Dim sql As String
      Dim Px As PRN

10    On Error GoTo Load_Error

20    sql = "SELECT * FROM Printers"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    Do While Not tb.EOF
60      Set Px = New PRN
70      With Px
80        .MappedTo = UCase$(tb!MappedTo & "")
90        .PrinterName = UCase$(tb!PrinterName & "")
100       mcolPRNs.Add Px, .MappedTo
110     End With
120     tb.MoveNext
130   Loop

140   Exit Sub

Load_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "PRNs", "Load", intEL, strES, sql

End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    On Error GoTo NewEnum_Error

20    Set NewEnum = mcolPRNs.[_NewEnum]

30    Exit Function

NewEnum_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "PRNs", "NewEnum", intEL, strES


End Function

Public Property Get PRNs() As PRNs
  
10    On Error GoTo PRNs_Error

20    Set PRNs = mPRNs

30    Exit Property

PRNs_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "PRNs", "PRNs", intEL, strES


End Property

Public Sub Refresh()

10    On Error GoTo Refresh_Error

20    Set mcolPRNs = Nothing

30    Load

40    Exit Sub

Refresh_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "PRNs", "Refresh", intEL, strES


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
60    LogError "PRNs", "Class_Initialize", intEL, strES


End Sub


