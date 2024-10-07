Attribute VB_Name = "basNewEXE"
Option Explicit

Public Function CheckNewEXE(ByVal NameOfExe As String) As String

      Dim FileName As String
      Dim Current As String
      Dim Found As Boolean
      Dim Path As String

10    On Error GoTo CheckNewEXE_Error

20    Found = False

30    Path = App.Path & "\"
40    Current = UCase$(NameOfExe) & ".EXE"
50    FileName = UCase$(Dir(Path & NameOfExe & "*.exe", vbNormal))

60    Do While FileName <> ""
70      If FileName > Current Then
80        Current = FileName
90        Found = True
100     End If
110     FileName = UCase$(Dir)
120   Loop

130   If Found And UCase$(App.EXEName) & ".EXE" <> Current Then
140     CheckNewEXE = Path & Current
150   Else
160     CheckNewEXE = ""
170   End If

180   Exit Function

CheckNewEXE_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "basNewEXE", "CheckNewEXE", intEL, strES

End Function
