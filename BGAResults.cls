VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BGAResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolBGAResults As New Collection
Private mBGAResults As New BGAResults

Public Function Add(ByVal BG As BGAResult) As BGAResult

10    On Error GoTo Add_Error

20    mcolBGAResults.Add BG

30    Set Add = BG

40    Exit Function

Add_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "BGAResults", "Add", intEL, strES


End Function


Public Property Get BGAResults() As BGAResults
  
10    On Error GoTo BGAResults_Error

20    Set BGAResults = mBGAResults

30    Exit Property

BGAResults_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "BGAResults", "BGAResults", intEL, strES


End Property

Public Sub Clear()

      Dim n As Integer

10    On Error GoTo Clear_Error

20    For n = 1 To mcolBGAResults.Count
30      mcolBGAResults.Remove 1
40    Next

50    Exit Sub

Clear_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "BGAResults", "Clear", intEL, strES


End Sub




Public Function Count() As Long
  
10    On Error GoTo Count_Error

20    Count = mcolBGAResults.Count

30    Exit Function

Count_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "BGAResults", "Count", intEL, strES


End Function



Public Function Item(ByVal Index As Long) As BGAResult
Attribute Item.VB_UserMemId = 0

10    On Error GoTo Item_Error

20    Set Item = mcolBGAResults.Item(Index)

30    Exit Function

Item_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "BGAResults", "Item", intEL, strES


End Function


Public Function LoadResults(ByVal SampleID As String) _
                            As BGAResult

      Dim sn As Recordset
      Dim CR As BGAResult
      Dim sql As String

10    On Error GoTo LoadResults_Error

20    sql = "SELECT * FROM BGAresults WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
30    Set sn = New Recordset
40    RecOpenServer 0, sn, sql

50    If Not sn.EOF Then
60      Set CR = New BGAResult
70      With CR
80        .Rundate = sn!Rundate & ""
90        .RunDateTime = sn!RunDateTime & ""
100       .pH = sn!pH & ""
110       .PCO2 = sn!PCO2 & ""
120       .PO2 = sn!PO2 & ""
130       .HCO3 = sn!HCO3 & ""
140       .BE = sn!BE & ""
150       .O2SAT = sn!O2SAT & ""
160       .TotCO2 = sn!TotCO2 & ""
170       If Not IsNull(sn!Valid) Then
180         .Valid = IIf(sn!Valid, True, False)
190       Else
200         .Valid = False
210       End If
220       If Not IsNull(sn!Printed) Then
230         .Printed = IIf(sn!Printed, True, False)
240       Else
250         .Printed = False
260       End If
270       .NOPAS = sn!NOPAS & ""
280     End With
290     Set LoadResults = CR
300   Else
310     Set LoadResults = Nothing
320   End If




330   Exit Function

LoadResults_Error:

      Dim strES As String
      Dim intEL As Integer

340   intEL = Erl
350   strES = Err.Description
360   LogError "BGAResults", "LoadResults", intEL, strES, sql


End Function
Public Sub LogAsPrinted(ByVal SampleID As String)

      Dim sql As String


10    On Error GoTo LogAsPrinted_Error

20    sql = "Update BGAresults " & _
            "set Printed = 1 WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
30    Cnxn(0).Execute sql

40    Exit Sub

LogAsPrinted_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "BGAResults", "LogAsPrinted", intEL, strES, sql


End Sub

Public Sub SaveResults(ByVal CR As BGAResult)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo SaveResults_Error

20    sql = "SELECT * FROM BGAResults WHERE " & _
            "SampleID = '" & Format$(Val(CR.SampleID)) & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    tb.AddNew
60    tb!SampleID = Format$(Val(CR.SampleID))
70    tb!Rundate = Format(CR.Rundate, "dd/mmm/yyyy")
80    tb!RunDateTime = Format(CR.RunDateTime, "dd/mmm/yyyy hh:mm:ss")
90    tb!pH = CR.pH
100   tb!PCO2 = CR.PCO2
110   tb!PO2 = CR.PO2
120   tb!HCO3 = CR.HCO3
130   tb!BE = CR.BE
140   tb!O2SAT = CR.O2SAT
150   tb!TotCO2 = CR.TotCO2
160   tb!Valid = CR.Valid
170   tb!Printed = CR.Printed
180   tb!NOPAS = CR.NOPAS
190   tb.Update

200   Exit Sub

SaveResults_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "BGAResults", "SaveResults", intEL, strES, sql

End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    On Error GoTo NewEnum_Error

20    Set NewEnum = mcolBGAResults.[_NewEnum]

30    Exit Function

NewEnum_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "BGAResults", "NewEnum", intEL, strES


End Function

Public Sub Remove(ByVal Index As Long)

10    On Error GoTo Remove_Error

20    mcolBGAResults.Remove (Index)

30    Exit Sub

Remove_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "BGAResults", "Remove", intEL, strES


End Sub





