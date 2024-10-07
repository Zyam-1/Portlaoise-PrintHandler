VERSION 5.00
Begin VB.Form fcdrInputBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   2685
   ClientTop       =   4485
   ClientWidth     =   5700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox tInput 
      Height          =   255
      Left            =   660
      TabIndex        =   0
      Top             =   1950
      Width           =   3375
   End
   Begin VB.CommandButton bOK 
      Caption         =   "O. K."
      Default         =   -1  'True
      Height          =   525
      Left            =   4200
      TabIndex        =   2
      Top             =   180
      Width           =   1245
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   525
      Left            =   4200
      TabIndex        =   1
      Top             =   1350
      Width           =   1245
   End
   Begin VB.Label lPrompt 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   1695
      Left            =   660
      TabIndex        =   3
      Top             =   180
      Width           =   3375
      WordWrap        =   -1  'True
   End
   Begin VB.Image i 
      Height          =   480
      Index           =   32
      Left            =   120
      Picture         =   "fcdrInputBox.frx":0000
      Top             =   720
      Width           =   480
   End
End
Attribute VB_Name = "fcdrInputBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ReturnValue As String
Private Pass As Boolean

Private Sub bcancel_Click()

10    On Error GoTo bcancel_Click_Error

20    ReturnValue = ""
30    Unload Me

40    Exit Sub

bcancel_Click_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "fcdrInputBox", "bcancel_Click", intEL, strES

End Sub

Private Sub bOK_Click()

10    On Error GoTo bOK_Click_Error

20    ReturnValue = tInput
30    Unload Me

40    Exit Sub

bOK_Click_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "fcdrInputBox", "bOK_Click", intEL, strES

End Sub

Public Property Get RetVal() As String

10    On Error GoTo Retval_Error

20    RetVal = ReturnValue

30    Exit Property

Retval_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "fcdrInputBox", "Retval", intEL, strES

End Property



Private Sub Form_Activate()

10    On Error GoTo Form_Activate_Error

20    tInput.PasswordChar = IIf(Pass, "*", "")

30    tInput.SelStart = 0
40    tInput.SelLength = Len(tInput)

50    Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "fcdrInputBox", "Form_Activate", intEL, strES

End Sub

Public Property Let PassWord(ByVal vNewValue As Boolean)

10    On Error GoTo PassWord_Error

20    Pass = vNewValue

30    Exit Property

PassWord_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "fcdrInputBox", "PassWord", intEL, strES

End Property
