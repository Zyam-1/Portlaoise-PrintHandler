VERSION 5.00
Begin VB.Form frmSetPHLocation 
   Caption         =   "NetAcquire - Set Print Handler Locations"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   4785
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton bSave 
      Caption         =   "Save"
      Height          =   705
      Left            =   1590
      Picture         =   "frmSetPHLocation.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Save"
      Top             =   5400
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Caption         =   "Different Network Domain Side"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   255
      TabIndex        =   6
      Top             =   2865
      Width           =   4215
      Begin VB.TextBox txtWardPHServer 
         Height          =   285
         Index           =   2
         Left            =   1215
         TabIndex        =   13
         Top             =   1665
         Width           =   2790
      End
      Begin VB.TextBox txtWardPHServer 
         Height          =   285
         Index           =   0
         Left            =   1215
         TabIndex        =   8
         Top             =   900
         Width           =   2790
      End
      Begin VB.TextBox txtWardPHServer 
         Height          =   285
         Index           =   1
         Left            =   1215
         TabIndex        =   7
         Top             =   1275
         Width           =   2790
      End
      Begin VB.Label Label5 
         Caption         =   "Enter server names where Print Handlers will be running in the Laboratory network domain."
         Height          =   405
         Left            =   225
         TabIndex        =   17
         Top             =   315
         Width           =   3795
      End
      Begin VB.Label Label7 
         Caption         =   "Server Name:"
         Height          =   255
         Left            =   165
         TabIndex        =   14
         Top             =   1710
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Server Name:"
         Height          =   255
         Left            =   165
         TabIndex        =   10
         Top             =   945
         Width           =   990
      End
      Begin VB.Label Label3 
         Caption         =   "Server Name:"
         Height          =   255
         Left            =   165
         TabIndex        =   9
         Top             =   1320
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Laboratory Side"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   240
      TabIndex        =   1
      Top             =   300
      Width           =   4215
      Begin VB.TextBox txtLabPHServer 
         Height          =   285
         Index           =   2
         Left            =   1215
         TabIndex        =   11
         Top             =   1740
         Width           =   2790
      End
      Begin VB.TextBox txtLabPHServer 
         Height          =   285
         Index           =   1
         Left            =   1215
         TabIndex        =   5
         Top             =   1365
         Width           =   2790
      End
      Begin VB.TextBox txtLabPHServer 
         Height          =   285
         Index           =   0
         Left            =   1215
         TabIndex        =   4
         Top             =   990
         Width           =   2790
      End
      Begin VB.Label Label8 
         Caption         =   "Enter server names where Print Handlers will be running in the Laboratory network domain."
         Height          =   405
         Left            =   210
         TabIndex        =   16
         Top             =   390
         Width           =   3795
      End
      Begin VB.Label Label6 
         Caption         =   "Server Name:"
         Height          =   255
         Left            =   165
         TabIndex        =   12
         Top             =   1785
         Width           =   990
      End
      Begin VB.Label Label2 
         Caption         =   "Server Name:"
         Height          =   255
         Left            =   165
         TabIndex        =   3
         Top             =   1410
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Server Name:"
         Height          =   255
         Left            =   165
         TabIndex        =   2
         Top             =   1035
         Width           =   1005
      End
   End
   Begin VB.CommandButton bCancel 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   705
      Left            =   3210
      Picture         =   "frmSetPHLocation.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cancel"
      Top             =   5400
      Width           =   1245
   End
End
Attribute VB_Name = "frmSetPHLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bcancel_Click()
10    Unload Me
End Sub

Private Function AreServerNamesUnique() As Boolean

      Dim n As Integer
      Dim h As Integer
      Dim varArray() As String
      Dim strTemp As String

10    AreServerNamesUnique = True
      ' arrays returned by Split are always zero-based
20    varArray() = Split(txtLabPHServer(0) & ";" & txtLabPHServer(1) & ";" & txtLabPHServer(2) & ";" & txtWardPHServer(0) & ";" & txtWardPHServer(1) & ";" & txtWardPHServer(2), ";")

30    For n = 0 To 5
40        strTemp = Trim$(varArray(n))
50        For h = 0 To 5
60            If n <> h Then
70                If strTemp = Trim$(varArray(h)) And strTemp <> "" Then
80                    AreServerNamesUnique = False
90                    Exit Function
100               End If
110           End If
120       Next
130   Next

End Function

Private Sub bSave_Click()

      Dim sql As String
      Dim n As Integer
      Dim rsRec As Recordset

10    On Error GoTo bSave_Click_Error

20    If Not AreServerNamesUnique Then
30        iMsg "Server names are not unique!"
40        Exit Sub
50    End If

60    sql = "Delete from Options WHERE Description = 'OPT_PRINTHANDLER_LAB_SIDE'"
70    Cnxn(0).Execute sql

80    For n = 0 To txtLabPHServer.Count - 1
90        If Len(Trim$(txtLabPHServer(n))) > 0 Then
100           sql = "Select * from Options WHERE Description = 'OPT_PRINTHANDLER_LAB_SIDE' and Contents = '" & txtLabPHServer(n) & "'"
  
110           Set rsRec = New Recordset
120           RecOpenClient 0, rsRec, sql
130           If rsRec.EOF Then
140               rsRec.AddNew
150               rsRec!Description = "OPT_PRINTHANDLER_LAB_SIDE"
160               rsRec!Contents = Trim$(txtLabPHServer(n))
170               rsRec.Update
180           End If
190       End If
200   Next

210   sql = "Delete from Options WHERE Description = 'OPT_PRINTHANDLER_WARD_SIDE'"
220   Cnxn(0).Execute sql

230   For n = 0 To txtWardPHServer.Count - 1
240       If Len(Trim$(txtWardPHServer(n))) > 0 Then
250           sql = "Select * from Options WHERE Description = 'OPT_PRINTHANDLER_WARD_SIDE' and Contents = '" & txtWardPHServer(n) & "'"
  
260           Set rsRec = New Recordset
270           RecOpenClient 0, rsRec, sql
280           If rsRec.EOF Then
290               rsRec.AddNew
300               rsRec!Description = "OPT_PRINTHANDLER_WARD_SIDE"
310               rsRec!Contents = Trim$(txtWardPHServer(n))
320               rsRec.Update
330           End If
340       End If
350   Next

360   Unload Me

370   Exit Sub

bSave_Click_Error:

      Dim strES As String
      Dim intEL As Integer

380   intEL = Erl
390   strES = Err.Description
400   LogError "frmSetPHLocation", "bSave_Click", intEL, strES

End Sub

Private Sub Form_Load()

      Dim sql As String
      Dim rsRec As Recordset

10    On Error GoTo Form_Load_Error

20    sql = "Select * from Options where Description = 'OPT_PRINTHANDLER_LAB_SIDE' "
    
30    Set rsRec = New Recordset
40    RecOpenClient 0, rsRec, sql

50    If Not rsRec.EOF Then
60        txtLabPHServer(0) = rsRec!Contents & ""
70        rsRec.MoveNext
80    End If

90    If Not rsRec.EOF Then
100       txtLabPHServer(1) = rsRec!Contents & ""
110       rsRec.MoveNext
120   End If

130   If Not rsRec.EOF Then
140       txtLabPHServer(2) = rsRec!Contents & ""
150       rsRec.MoveNext
160   End If
170   rsRec.Close

180   sql = "Select * from Options where Description = 'OPT_PRINTHANDLER_WARD_SIDE' "
    
190   Set rsRec = New Recordset
200   RecOpenClient 0, rsRec, sql

210   If Not rsRec.EOF Then
220       txtWardPHServer(0) = rsRec!Contents & ""
230       rsRec.MoveNext
240   End If

250   If Not rsRec.EOF Then
260       txtWardPHServer(1) = rsRec!Contents & ""
270       rsRec.MoveNext
280   End If

290   If Not rsRec.EOF Then
300       txtWardPHServer(2) = rsRec!Contents & ""
310       rsRec.MoveNext
320   End If
330   rsRec.Close

340   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

350   intEL = Erl
360   strES = Err.Description
370   LogError "frmSetPHLocation", "Form_Load", intEL, strES, sql

End Sub

