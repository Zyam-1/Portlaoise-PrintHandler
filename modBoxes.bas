Attribute VB_Name = "modBoxes"
Option Explicit

Public Function iMsg(Optional ByVal Message As String, _
                     Optional ByVal t As Integer = 0, _
                     Optional ByVal Caption As String = "NetAcquire", _
                     Optional ByVal BckColour As Long = &HC0C000, _
                     Optional ByVal MsgFontSize As Long) _
                     As Integer

      Dim SafeMsgBox As New fcdrMsgBox

10    On Error GoTo iMsg_Error

20    With SafeMsgBox
30      .MsgFontSize = MsgFontSize
40      .BackColor = BckColour
50      .DisplayButtons = t And &H7
60      .DefaultButton = t And &H300
70      .ShowIcon = t And &H70
80      .Message = Message
90      .Caption = Caption
100     .Show vbModal
110     iMsg = .RetVal
120   End With

130   Unload SafeMsgBox
140   Set SafeMsgBox = Nothing

150   Exit Function

iMsg_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "modBoxes", "iMsg", intEL, strES

End Function

Public Function iBOX(ByVal Prompt As String, _
            Optional ByVal Title As String = "NetAcquire", _
            Optional ByVal Default As String, _
            Optional ByVal Pass As Boolean) As String

      Dim Box As New fcdrInputBox

10    On Error GoTo iBOX_Error

20    With Box
30      .PassWord = Pass
40      .Caption = Title
50      .lPrompt = Prompt
60      .tInput = Default
70      .Show vbModal
80      iBOX = .RetVal
90    End With

100   Unload Box
110   Set Box = Nothing

120   Exit Function

iBOX_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "modBoxes", "iBOX", intEL, strES

End Function



