VERSION 5.00
Begin VB.Form frmSetRole 
   Caption         =   "NetAcquire- Set Print Handler Roles"
   ClientHeight    =   5568
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   4776
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5568
   ScaleWidth      =   4776
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Caption         =   "Micro Printing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1056
      Left            =   240
      TabIndex        =   11
      Top             =   3480
      Width           =   4215
      Begin VB.TextBox lblServer 
         Height          =   285
         Index           =   3
         Left            =   1380
         TabIndex        =   12
         Top             =   525
         Width           =   2640
      End
      Begin VB.Label Label4 
         Caption         =   "Server Name:"
         Height          =   255
         Left            =   150
         TabIndex        =   13
         Top             =   540
         Width           =   1035
      End
   End
   Begin VB.CommandButton bSave 
      Caption         =   "Save"
      Height          =   705
      Left            =   1725
      Picture         =   "frmSetRole.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Save"
      Top             =   4740
      Width           =   1245
   End
   Begin VB.Frame Frame3 
      Caption         =   "Faxing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1056
      Left            =   255
      TabIndex        =   5
      Top             =   2340
      Width           =   4215
      Begin VB.TextBox lblServer 
         Height          =   285
         Index           =   2
         Left            =   1380
         TabIndex        =   10
         Top             =   525
         Width           =   2640
      End
      Begin VB.Label Label3 
         Caption         =   "Server Name:"
         Height          =   255
         Left            =   150
         TabIndex        =   6
         Top             =   540
         Width           =   1035
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ward Printing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   996
      Left            =   255
      TabIndex        =   3
      Top             =   1260
      Width           =   4215
      Begin VB.TextBox lblServer 
         Height          =   285
         Index           =   1
         Left            =   1410
         TabIndex        =   9
         Top             =   495
         Width           =   2640
      End
      Begin VB.Label Label2 
         Caption         =   "Server Name:"
         Height          =   255
         Left            =   165
         TabIndex        =   4
         Top             =   465
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Laboratory Printing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   996
      Left            =   255
      TabIndex        =   1
      Top             =   180
      Width           =   4215
      Begin VB.TextBox lblServer 
         Height          =   285
         Index           =   0
         Left            =   1395
         TabIndex        =   8
         Top             =   465
         Width           =   2640
      End
      Begin VB.Label Label1 
         Caption         =   "Server Name:"
         Height          =   255
         Left            =   165
         TabIndex        =   2
         Top             =   495
         Width           =   975
      End
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   705
      Left            =   3240
      Picture         =   "frmSetRole.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cancel"
      Top             =   4740
      Width           =   1245
   End
End
Attribute VB_Name = "frmSetRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bcancel_Click()
10    Unload Me
End Sub

Private Sub bSave_Click()

Dim sql As String
Dim tb As Recordset

'Lab print handler pc name
On Error GoTo bSave_Click_Error

sql = "Select * from Options where Description = 'LAB_PRINT_HANDLER_PC_NAME' "
Set tb = New Recordset
RecOpenClient 0, tb, sql

If tb.EOF Then tb.AddNew
tb!Description = "LAB_PRINT_HANDLER_PC_NAME"
tb!Contents = lblServer(0)
tb.Update

'Ward print handler pc name
sql = "Select * from Options where Description = 'WARD_PRINT_HANDLER_PC_NAME' "
Set tb = New Recordset
RecOpenClient 0, tb, sql

If tb.EOF Then tb.AddNew
tb!Description = "WARD_PRINT_HANDLER_PC_NAME"
tb!Contents = lblServer(1)
tb.Update

'Fax print handler pc name
sql = "Select * from Options where Description = 'FAX_PRINT_HANDLER_PC_NAME' "
Set tb = New Recordset
RecOpenClient 0, tb, sql

If tb.EOF Then tb.AddNew
tb!Description = "FAX_PRINT_HANDLER_PC_NAME"
tb!Contents = lblServer(2)
tb.Update

'MICRO print handler pc name
sql = "Select * from Options where Description = 'MICRO_PRINT_HANDLER_PC_NAME' "
Set tb = New Recordset
RecOpenClient 0, tb, sql

If tb.EOF Then tb.AddNew
tb!Description = "MICRO_PRINT_HANDLER_PC_NAME"
tb!Contents = lblServer(3)
tb.Update

LoadSettings

Exit Sub

bSave_Click_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmSetRole", "bSave_Click", intEL, strES, sql

End Sub






Private Sub LoadSettings()

      Dim sql As String
      Dim tb As Recordset
      Dim n As Integer

      'Clear labels
10    On Error GoTo LoadSettings_Error

20    For n = 0 To lblServer.Count - 1
30        lblServer(n) = ""
      '40        lblServer(n).BackColor = &H8000000F
40    Next

      'Lab print handler pc name
50    sql = "Select contents from Options where Description = 'LAB_PRINT_HANDLER_PC_NAME' "
60    Set tb = New Recordset
70    RecOpenClient 0, tb, sql

80    If Not tb.EOF Then
90        lblServer(0) = tb!Contents & ""
100   End If
110   tb.Close

      'Ward print handler pc name
120   sql = "Select contents from Options where Description = 'WARD_PRINT_HANDLER_PC_NAME' "
130   Set tb = New Recordset
140   RecOpenClient 0, tb, sql

150   If Not tb.EOF Then
160       lblServer(1) = tb!Contents & ""
170   End If
180   tb.Close

      'Fax print handler pc name
190   sql = "Select contents from Options where Description = 'FAX_PRINT_HANDLER_PC_NAME' "
200   Set tb = New Recordset
210   RecOpenClient 0, tb, sql

220   If Not tb.EOF Then
230       lblServer(2) = tb!Contents & ""
240   End If
250   tb.Close


      'MICRO print handler pc name
260   sql = "Select contents from Options where Description = 'MICRO_PRINT_HANDLER_PC_NAME' "
270   Set tb = New Recordset
280   RecOpenClient 0, tb, sql

290   If Not tb.EOF Then
300       lblServer(3) = tb!Contents & ""
310   End If
320   tb.Close

330   Exit Sub

LoadSettings_Error:

      Dim strES As String
      Dim intEL As Integer

340   intEL = Erl
350   strES = Err.Description
360   LogError "frmSetRole", "LoadSettings", intEL, strES, sql

End Sub
Private Sub Form_Load()

10    LoadSettings


End Sub



Private Sub lblServer_KeyPress(Index As Integer, KeyAscii As Integer)
10    lblServer(Index).BackColor = vbYellow
End Sub
