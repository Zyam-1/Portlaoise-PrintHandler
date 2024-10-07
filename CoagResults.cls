VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CoagResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolCoagResults As New Collection
Private mCoagResults As New CoagResults

Public Function Add(CR As CoagResult) As CoagResult

10    On Error GoTo Add_Error

20    mcolCoagResults.Add CR

30    Set Add = CR

40    Exit Function

Add_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "CoagResults", "Add", intEL, strES


End Function

Public Property Get CoagResults() As CoagResults
  
10    On Error GoTo CoagResults_Error

20    Set CoagResults = mCoagResults

30    Exit Property

CoagResults_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "CoagResults", "CoagResults", intEL, strES


End Property

Public Sub Clear()

      Dim n As Long

10    On Error GoTo Clear_Error

20    For n = 1 To mcolCoagResults.Count
30      mcolCoagResults.Remove 1
40    Next

50    Exit Sub

Clear_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "CoagResults", "Clear", intEL, strES


End Sub

Public Function Count() As Long
  
10    On Error GoTo Count_Error

20    Count = mcolCoagResults.Count

30    Exit Function

Count_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "CoagResults", "Count", intEL, strES


End Function


Public Sub Delete(ByVal SampleID As String, _
                  ByVal Code As String, ByVal Units As String)

      Dim sql As String




10    On Error GoTo Delete_Error

20    sql = "DELETE FROM CoagResults WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "and Code = '" & Code & "' and units = '" & Units & "'"

30    Cnxn(0).Execute sql




40    Exit Sub

Delete_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "CoagResults", "Delete", intEL, strES, sql


End Sub

Public Function Item(ByVal X As Long) _
                     As CoagResult
Attribute Item.VB_UserMemId = 0


10    On Error GoTo Item_Error

20    Set Item = mcolCoagResults(X)

30    Exit Function

Item_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "CoagResults", "Item", intEL, strES


End Function


Public Function Load(ByVal SampleID As String, _
                     Optional ByVal Units As String) _
                     As CoagResults

      Dim CRs As New CoagResults
      Dim CR As CoagResult
      Dim sql As String
      Dim tb As New Recordset

10    On Error GoTo Load_Error

20    sql = "SELECT DISTINCT R.Code AS Expr2, R.*, D.PrintPriority " & _
            "FROM CoagResults R, CoagTestDefinitions D WHERE " & _
            "R.Code = D.Code " & _
            "AND R.SampleID = '" & SampleID & "' " & _
            "AND D.TestName NOT IN " & _
            "  ( SELECT Parameter FROM PrintInhibit WHERE " & _
            "    SampleID = '" & SampleID & "' And Discipline = 'Coa' ) " & _
            "ORDER BY PrintPriority"

30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    Do While Not tb.EOF
60      Set CR = New CoagResult
70      With CR
80        .SampleID = tb!SampleID & ""
90        .Code = tb!Code & ""
100       .Result = tb!Result & ""
110       .Rundate = Format$(tb!Rundate, "dd/mm/yyyy")
120       .RunTime = Format$(tb!RunTime, "dd/mm/yyyy hh:mm")
130       .Printed = tb!Printed
140       .Valid = tb!Valid
150       .Units = tb!Units & ""
160       .OperatorCode = tb!UserName & ""
170       CRs.Add CR
180     End With
190     tb.MoveNext
200   Loop

210   If CRs.Count <> 0 Then
220     Set Load = CRs
230   Else
240     Set Load = Nothing
250   End If

260   Exit Function

Load_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "CoagResults", "Load", intEL, strES, sql

End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    On Error GoTo NewEnum_Error

20    Set NewEnum = mcolCoagResults.[_NewEnum]

30    Exit Function

NewEnum_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "CoagResults", "NewEnum", intEL, strES


End Function



Public Sub Remove(ByVal Index As Integer)

10    On Error GoTo Remove_Error

20    mcolCoagResults.Remove (Index)

30    Exit Sub

Remove_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "CoagResults", "Remove", intEL, strES


End Sub

Public Sub Save(ByVal CRs As CoagResults)

      Dim sql As String
      Dim tb As New Recordset
      Dim CR As CoagResult
      Dim lngSampleID As Double
      Dim Rundate As String





10    On Error GoTo Save_Error

20    For Each CR In CRs
30      With CR
40        If IsNumeric(.Code) Then
50          lngSampleID = Val(.SampleID)
60          sql = "SELECT * FROM CoagResults WHERE " & _
                  "SampleID = '" & lngSampleID & "' " & _
                  "and  Code = '" & .Code & "' and units = '" & .Units & "'"
70          Set tb = New Recordset
80          RecOpenServer 0, tb, sql
90          If Not tb.EOF Then
100           sql = "SELECT * FROM CoagRepeats WHERE " & _
                    "SampleID = '" & lngSampleID & "' " & _
                    "and  Code = '" & CR.Code & "' and units = '" & .Units & "'"
110           Set tb = New Recordset
120         RecOpenServer 0, tb, sql
130         End If
140         tb.AddNew
150         tb!SampleID = lngSampleID
160         Rundate = Format$(.Rundate, "dd/mmm/yyyy")
170         tb!Rundate = Rundate
180         If Trim$(.RunTime) <> "" Then
190           tb!RunTime = Format$(.RunTime, "dd/mmm/yyyy hh:mm")
200         End If
210         tb!Code = Left$(.Code, 4)
220         If IsNumeric(.Result) Then
230           tb!Result = Left$(Format$(Val(.Result)), 6)
240         Else
250           tb!Result = Left$(.Result, 6)
260         End If
270         tb!Printed = IIf(.Printed, 1, 0)
280         tb!Valid = IIf(.Valid, 1, 0)
290         tb!Units = .Units
300         tb.Update
310         sql = "DELETE FROM CoagRequests WHERE " & _
                  "SampleID = '" & lngSampleID & "' " & _
                  "and Code = '" & .Code & "' and units = '" & .Units & "'"
320         Cnxn(0).Execute sql
330       End If
340     End With
350   Next


360   Set CR = Nothing




370   Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

380   intEL = Erl
390   strES = Err.Description
400   LogError "CoagResults", "Save", intEL, strES, sql


End Sub


Public Sub LogAsPrinted(ByVal SampleID As String)

      Dim sql As String

10    On Error GoTo LogAsPrinted_Error

20    sql = "UPDATE CoagResults set Printed = 1 WHERE " & _
            "SampleID = '" & SampleID & "'"
30    Cnxn(0).Execute sql

40    Exit Sub

LogAsPrinted_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "CoagResults", "LogAsPrinted", intEL, strES, sql

End Sub



