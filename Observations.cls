VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Observations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolObservations As New Collection
Private mObservations As New Observations

Public Sub Clear()

      Dim n As Long

10    For n = 1 To mcolObservations.Count
20      mcolObservations.Remove 1
30    Next

End Sub

Public Property Get Observations() As Observations
  
10    Set Observations = mObservations

End Property


Public Function Item(ByVal X As Long) _
                     As Observation

10    Set Item = mcolObservations(X)

End Function



Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    Set NewEnum = mcolObservations.[_NewEnum]

End Function




Public Sub Remove(ByVal Index As Integer)

10    mcolObservations.Remove (Index)

End Sub


Public Function Load(ByVal SampleID As String, ParamArray Disc()) As Observations
      'Paramarray is one or a combination of :-
      'Demographic Biochemistry Haematology Coagulation Immunology BloodGas
      'Semen Microcs MicroIdent MicroGeneral MicroConsultant Film
      'Endocrinology Histology Cytology CSFFluid
      'ImmunologyA ImmunologyB ImmunologyC

      Dim OBS As New Observations
      Dim OB As Observation
      Dim sql As String
      Dim tb As New Recordset
      Dim n As Integer
      Dim s As String

10    On Error GoTo Load_Error

20    sql = "SELECT * FROM Observations WHERE " & _
            "SampleID = '" & Val(SampleID) & "' AND ("
30    s = ""
40    For n = LBound(Disc) To UBound(Disc)
50      s = s & "Discipline = '" & Disc(n) & "' OR "
60    Next
70    s = Left$(s, Len(s) - 3) & ")"
80    sql = sql & s

90    Set tb = New Recordset
100   RecOpenServer 0, tb, sql

110   Do While Not tb.EOF
120     Set OB = New Observation
130     With OB
140       .SampleID = tb!SampleID & ""
150       .Discipline = tb!Discipline & ""
160       .Comment = tb!Comment & ""
170       .DateTimeOfRecord = Format$(tb!DateTimeOfRecord, "dd/MM/yyyy HH:nn:ss")
180       .UserName = tb!UserName & ""
190       OBS.Add OB
200     End With
210     tb.MoveNext
220   Loop

230   If OBS.Count <> 0 Then
240     Set Load = OBS
250   Else
260     Set Load = Nothing
270   End If

280   Exit Function

Load_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "Observations", "Load", intEL, strES, sql


End Function
Public Function Save(ByVal SampleID As String, ByVal OverWrite As Boolean, ParamArray Disc()) As Observations

      'OverWrite = True : Any existing comment is replaced.
      'OverWrite = False : New comment is concatenated

      'Paramarray is one or a combination of :-
      'Demographic Biochemistry Haematology Coagulation Immunology BloodGas
      'Semen Microcs MicroIdent MicroGeneral MicroConsultant Film
      'Endocrinology Histology Cytology CSFFluid
      'ImmunologyA ImmunologyB ImmunologyC
      'eg "Haematology","Haem Comment","Biochemistry","Bio Comment" etc

      Dim sql As String
      Dim n As Integer
      Dim Discipline As String
      Dim Comment As String

10    On Error GoTo Save_Error

20    For n = LBound(Disc) To UBound(Disc) Step 2
30      Discipline = Disc(n)
40      Comment = AddTicks(Disc(n + 1))
50      If Comment = "" Then
60        sql = "DELETE FROM Observations " & _
                "WHERE SampleID = '" & SampleID & "' " & _
                "AND Discipline = '" & Discipline & "'"
70      Else
80        sql = "IF EXISTS (SELECT * FROM Observations " & _
                           "WHERE SampleID = '" & SampleID & "' " & _
                           "AND Discipline = '" & Discipline & "') " & _
                "  UPDATE Observations "
90        If OverWrite Then
100         sql = sql & "  SET Comment = '" & Comment & "' "
110       Else
120         sql = sql & "  SET Comment = Comment + ' " & Comment & "' "
130       End If
140       sql = sql & "  WHERE SampleID = '" & SampleID & "' " & _
                "  AND Discipline = '" & Discipline & "' " & _
                "ELSE " & _
                "  INSERT INTO Observations " & _
                "  (SampleID, Discipline, Comment, UserName ) " & _
                "  VALUES " & _
                "  ('" & SampleID & "', " & _
                "   '" & Discipline & "', " & _
                "   '" & Comment & "', " & _
                "   '" & AddTicks(UserName) & "')"
150     End If
160     Cnxn(0).Execute sql
170   Next

180   Exit Function

Save_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "Observations", "Save", intEL, strES, sql


End Function
Public Function Count() As Long
  
10    Count = mcolObservations.Count

20    Exit Function

End Function



Public Function Add(OB As Observation) As Observation
Attribute Add.VB_MemberFlags = "40"

10    mcolObservations.Add OB

20    Set Add = OB

End Function


