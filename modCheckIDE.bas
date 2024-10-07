Attribute VB_Name = "modCheckIDE"
Option Explicit

'You can check this flag anywhere in your code
Public IsIDE As Boolean

Public Sub CheckIDE()
    
10    On Error GoTo CheckIDE_Error

20    IsIDE = False
    
      'This line is only executed if
      'running in the IDE and then
      'returns True
30    Debug.Assert CheckIfInIDE
    
      'Use the IsIDE flag anywhere
      'For example
      '   If IsIDE Then
      '       MsgBox ("Running under IDE")
      '   Else
      '       MsgBox ("Running as EXE")
      '   End If

40    Exit Sub

CheckIDE_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "modCheckIDE", "CheckIDE", intEL, strES

End Sub

Private Function CheckIfInIDE() As Boolean
    
      'This function will never be executed in an EXE
10    On Error GoTo CheckIfInIDE_Error

20    IsIDE = True        'set global flag

      'Set CheckIfInIDE or the Debug.Assert will Break
30    CheckIfInIDE = True

40    Exit Function

CheckIfInIDE_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "modCheckIDE", "CheckIfInIDE", intEL, strES

End Function
