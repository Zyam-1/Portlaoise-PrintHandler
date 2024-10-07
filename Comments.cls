VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Comments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Load(ByVal lngSampleID As Double) As Comment

      Dim sql As String
      Dim tb As Recordset
      Dim Cx As New Comment
  
10    On Error GoTo Load_Error

20    If lngSampleID = 0 Then Exit Function

30    sql = "SELECT * FROM Comments WHERE " & _
            "SampleID = '" & lngSampleID & "'"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql

60    If tb.EOF Then
70      Set Load = Nothing
80    Else
90      Cx.lngSampleID = lngSampleID
100     Cx.Biochemistry = Trim$(tb!Biochemistry & "")
110     Cx.Haematology = Trim$(tb!Haematology & "")
120     Cx.Demographics = Trim$(tb!Demographic & "")
130     Cx.Coagulation = Trim$(tb!Coagulation & "")
140     Cx.Immunology = Trim$(tb!Immunology & "")
150     Cx.Endocrinology = Trim$(tb!Endocrinology & "")
160     Cx.BloodGas = Trim$(tb!BloodGas & "")
170     Cx.Semen = Trim$(tb!Semen & "")
180     Cx.MicroCS = tb!MicroCS & ""
190     Cx.MicroIdent = tb!MicroIdent & ""
200     Cx.MicroGeneral = tb!MicroGeneral & ""
210     Cx.MicroConsultant = tb!MicroConsultant & ""
220     Cx.Film = tb!Film & ""
230     Set Load = Cx
240   End If

250   Exit Function

Load_Error:

      Dim strES As String
      Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "Comments", "Load", intEL, strES, sql

End Function


Public Sub Save(ByVal Cx As Comment)

      Dim sql As String
      Dim tb As Recordset


10    On Error GoTo Save_Error

20    With Cx
30      If .lngSampleID = 0 Then Exit Sub

40      If Trim$(.Biochemistry & .Demographics & _
                 .Haematology & .Coagulation & _
                 .Semen & .MicroIdent & .MicroConsultant & _
                 .MicroGeneral & .MicroCS & .BloodGas & .Endocrinology) = "" Then
50        sql = "delete FROM comments WHERE " & _
                "SampleID = '" & .lngSampleID & "'"
60        Cnxn(0).Execute sql
70      Else
80        sql = "SELECT * FROM Comments WHERE " & _
                "SampleID = '" & .lngSampleID & "'"
90        Set tb = New Recordset
100       RecOpenClient 0, tb, sql
110       If tb.EOF Then
120         tb.AddNew
130         tb!SampleID = .lngSampleID
140       End If
150       tb!Biochemistry = .Biochemistry
160       tb!Demographic = .Demographics
170       tb!Haematology = .Haematology
180       tb!Coagulation = .Coagulation
190       tb!Immunology = .Immunology
200       tb!Endocrinology = .Endocrinology
210       tb!BloodGas = .BloodGas
220       tb!Semen = .Semen
      '    tb!MicroIdent = .MicroIdent
      '    tb!MicroGeneral = .MicroGeneral
      '    tb!MicroConsultant = .MicroConsultant
      '    tb!MicroCS = .MicroCS
230       tb!Film = .Film
240       tb.Update
250     End If
260   End With


270   Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

280   intEL = Erl
290   strES = Err.Description
300   LogError "Comments", "Save", intEL, strES, sql


End Sub


