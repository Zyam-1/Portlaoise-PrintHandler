VERSION 5.00
Begin VB.Form fHaemNoSexNormal 
   Caption         =   "NetAcquire"
   ClientHeight    =   5610
   ClientLeft      =   1635
   ClientTop       =   1155
   ClientWidth     =   6465
   Icon            =   "fHaemNoSexNormal.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   6465
   Begin VB.TextBox tNeut 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   15
      Top             =   4410
      Width           =   2265
   End
   Begin VB.TextBox tRDW 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   14
      Top             =   3420
      Width           =   2265
   End
   Begin VB.TextBox tMCV 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   13
      Top             =   2430
      Width           =   2265
   End
   Begin VB.TextBox tMCH 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   12
      Top             =   2760
      Width           =   2265
   End
   Begin VB.TextBox tRBC 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   11
      Top             =   1440
      Width           =   2265
   End
   Begin VB.TextBox tWBC 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   10
      Top             =   1110
      Width           =   2265
   End
   Begin VB.TextBox tMCHC 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   9
      Top             =   3090
      Width           =   2265
   End
   Begin VB.TextBox tHb 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   8
      Top             =   1770
      Width           =   2265
   End
   Begin VB.TextBox tPlt 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   7
      Top             =   3750
      Width           =   2265
   End
   Begin VB.TextBox tMPV 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   6
      Top             =   4080
      Width           =   2265
   End
   Begin VB.TextBox tLymp 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   5
      Top             =   4740
      Width           =   2265
   End
   Begin VB.TextBox tMono 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   4
      Top             =   5070
      Width           =   2265
   End
   Begin VB.TextBox tHct 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1620
      TabIndex        =   3
      Top             =   2100
      Width           =   2265
   End
   Begin VB.CommandButton bSave 
      Caption         =   "Save"
      Height          =   750
      Left            =   4650
      Picture         =   "fHaemNoSexNormal.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Save"
      Top             =   1815
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   705
      Left            =   4650
      Picture         =   "fHaemNoSexNormal.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cancel"
      Top             =   2910
      Width           =   1245
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "WBC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1140
      TabIndex        =   28
      Top             =   1140
      Width           =   435
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Hb"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1320
      TabIndex        =   27
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "MCV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1170
      TabIndex        =   26
      Top             =   2460
      Width           =   405
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "MCHC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1035
      TabIndex        =   25
      Top             =   3120
      Width           =   540
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Plt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1335
      TabIndex        =   24
      Top             =   3750
      Width           =   240
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Neutrophils"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   23
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Lymphocytes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   465
      TabIndex        =   22
      Top             =   4770
      Width           =   1110
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Monocytes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   645
      TabIndex        =   21
      Top             =   5130
      Width           =   930
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "MPV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1170
      TabIndex        =   20
      Top             =   4080
      Width           =   405
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "RDW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1110
      TabIndex        =   19
      Top             =   3450
      Width           =   465
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "MCH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1155
      TabIndex        =   18
      Top             =   2790
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Hct"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1260
      TabIndex        =   17
      Top             =   2130
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "RBC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1185
      TabIndex        =   16
      Top             =   1470
      Width           =   390
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "These Normal Ranges only apply when the Age/Sex Related option is Disabled."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   450
      TabIndex        =   0
      Top             =   180
      Width           =   5685
   End
End
Attribute VB_Name = "fHaemNoSexNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub bcancel_Click()

10    Unload Me

End Sub

Private Sub bSave_Click()

10    On Error GoTo bSave_Click_Error

20    SaveOptionSetting "PrintHandlerWBC", tWBC
30    SaveOptionSetting "PrintHandlerRBC", tRBC
40    SaveOptionSetting "PrintHandlerHB", tHb
50    SaveOptionSetting "PrintHandlerHCT", tHct
60    SaveOptionSetting "PrintHandlerMCV", tMCV
70    SaveOptionSetting "PrintHandlerMCH", tMCH
80    SaveOptionSetting "PrintHandlerMCHC", tMCHC
90    SaveOptionSetting "PrintHandlerRDW", tRDW
100   SaveOptionSetting "PrintHandlerPLT", tPlt
110   SaveOptionSetting "PrintHandlerMPV", tMPV
120   SaveOptionSetting "PrintHandlerNEUT", tNeut
130   SaveOptionSetting "PrintHandlerLYMP", tLymp
140   SaveOptionSetting "PrintHandlerMONO", tMono

150   bSave.Visible = False

160   Exit Sub

bSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "fHaemNoSexNormal", "bSave_Click", intEL, strES

End Sub

Private Sub Form_Load()

10    On Error GoTo Form_Load_Error

20    If TestSys = True Then Me.Caption = Me.Caption & " - TEST SYSTEM"

30    tWBC = GetOptionSetting("PrintHandlerWBC", "")
40    tRBC = GetOptionSetting("PrintHandlerRBC", "")
50    tHb = GetOptionSetting("PrintHandlerHB", "")
60    tHct = GetOptionSetting("PrintHandlerHCT", "")
70    tMCV = GetOptionSetting("PrintHandlerMCV", "")
80    tMCH = GetOptionSetting("PrintHandlerMCH", "")
90    tMCHC = GetOptionSetting("PrintHandlerMCHC", "")
100   tRDW = GetOptionSetting("PrintHandlerRDW", "")
110   tPlt = GetOptionSetting("PrintHandlerPLT", "")
120   tMPV = GetOptionSetting("PrintHandlerMPV", "")
130   tNeut = GetOptionSetting("PrintHandlerNEUT", "")
140   tLymp = GetOptionSetting("PrintHandlerLYMP", "")
150   tMono = GetOptionSetting("PrintHandlerMONO", "")

160   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "fHaemNoSexNormal", "Form_Load", intEL, strES

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    On Error GoTo Form_QueryUnload_Error

20    If bSave.Visible Then
30      If iMsg("Cancel without Saving?", vbYesNo) = vbNo Then
40        Cancel = True
50      End If
60    End If

70    Exit Sub

Form_QueryUnload_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "fHaemNoSexNormal", "Form_QueryUnload", intEL, strES

End Sub


Private Sub tHb_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tHct_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tLymp_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

20    Exit Sub

tLymp_KeyPress_Error:

      Dim strES As String
      Dim intEL As Integer

30    Screen.MousePointer = 0

40    intEL = Erl
50    strES = Err.Description
60    LogError "fHaemNoSexNormal", "tLymp_KeyPress", intEL, strES


End Sub


Private Sub tMCH_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tMCHC_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tMCV_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tMono_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tMPV_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tNeut_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tPlt_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tRBC_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tRDW_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


Private Sub tWBC_KeyPress(KeyAscii As Integer)

10    bSave.Visible = True

End Sub


