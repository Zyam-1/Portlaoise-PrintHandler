VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmForcePrinters 
   Caption         =   "Print Handler - Forced printer list"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9315
   ScaleWidth      =   9795
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   825
      Left            =   8715
      Picture         =   "frmForcePrinters.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3435
      Width           =   795
   End
   Begin VB.CommandButton cmdAdd 
      Appearance      =   0  'Flat
      Caption         =   "Add"
      Height          =   825
      Left            =   8715
      Picture         =   "frmForcePrinters.frx":0A02
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1755
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   210
      TabIndex        =   2
      Top             =   270
      Width           =   4455
      Begin VB.OptionButton optLoc 
         Caption         =   "Ward"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   4
         Top             =   315
         Width           =   1335
      End
      Begin VB.OptionButton optLoc 
         Alignment       =   1  'Right Justify
         Caption         =   "Laboratory"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1185
         TabIndex        =   3
         Top             =   315
         Value           =   -1  'True
         Width           =   1590
      End
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   825
      Left            =   8715
      Picture         =   "frmForcePrinters.frx":1404
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8205
      Width           =   795
   End
   Begin MSFlexGridLib.MSFlexGrid grdInstalledPrinters 
      Height          =   7275
      Left            =   195
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1755
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   12832
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483635
      BackColorFixed  =   -2147483647
      ForeColorFixed  =   -2147483624
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      GridLines       =   3
      GridLinesFixed  =   3
      ScrollBars      =   2
      FormatString    =   "<Printer Name                                                                                                       "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblGridTitle 
      Caption         =   "Printers available for force printing in the Laboratory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   225
      TabIndex        =   1
      Top             =   1305
      Width           =   4725
   End
End
Attribute VB_Name = "frmForcePrinters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bcancel_Click()
10    Unload Me
End Sub

Private Sub LoadPrintersForLocation()
      Dim tb As Recordset
      Dim sql As String
      Dim s As String


10    On Error GoTo LoadPrintersForLocation_Error

20    sql = "Select * from InstalledPrinters where Location = '" & IIf(optLoc(0), "LAB", "WARD") & "'"

30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql

50    grdInstalledPrinters.Rows = 2
60    grdInstalledPrinters.AddItem ""
70    grdInstalledPrinters.RemoveItem 1

80    Do While Not tb.EOF
90        s = tb!PrinterName & "" & vbTab & tb!RecordCounter
100       grdInstalledPrinters.AddItem s
110       tb.MoveNext
120   Loop

130   If grdInstalledPrinters.Rows > 2 Then
140     grdInstalledPrinters.RemoveItem 1
150   End If

160   Exit Sub

LoadPrintersForLocation_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "frmForcePrinters", "LoadPrintersForLocation", intEL, strES, sql

End Sub

Private Sub cmdAdd_Click()

10    If optLoc(0) Then
20        With frmAddForcePrinter
30            .optLoc(0) = True
40            .lblDesc = "Add printer to Laboratory force printer list."
50            .Show 1
60        End With
70    Else
80        With frmAddForcePrinter
90            .optLoc(1) = True
100           .lblDesc = "Add printer to Ward force printer list."
110           .Show 1
120       End With
130   End If

End Sub

Private Sub cmdDelete_Click()
      Dim n As Integer
      Dim sql As String
      Dim blnFound As Boolean


10    On Error GoTo cmdDelete_Click_Error

20    blnFound = False
30    For n = 1 To grdInstalledPrinters.Rows - 1
40        grdInstalledPrinters.Row = n
50        If grdInstalledPrinters.CellBackColor = vbYellow Then
60            blnFound = True
70            Exit For
80        End If
90    Next

100   If blnFound Then
110       If iMsg("Are you sure you wish to delete printer?", vbYesNo) = vbYes Then
120           sql = "Delete from InstalledPrinters where  RecordCounter = '" & grdInstalledPrinters.TextMatrix(n, 1) & "' "
  
130           Cnxn(0).Execute sql
140           LoadPrintersForLocation
150       End If
160   End If

170   Exit Sub

cmdDelete_Click_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "frmForcePrinters", "cmdDelete_Click", intEL, strES, sql

End Sub

Private Sub Form_Activate()
10    LoadPrintersForLocation
End Sub

Private Sub Form_Load()

10    grdInstalledPrinters.ColWidth(1) = 0
20    LoadPrintersForLocation

End Sub

Private Sub grdInstalledPrinters_Click()
      Dim n As Integer
      Dim intRowClicked As Integer

10    intRowClicked = grdInstalledPrinters.Row
20    grdInstalledPrinters.Col = 0
30    For n = 1 To grdInstalledPrinters.Rows - 1
40        grdInstalledPrinters.Row = n
50        grdInstalledPrinters.CellBackColor = &H80000018
60    Next
70    grdInstalledPrinters.Row = intRowClicked
80    grdInstalledPrinters.CellBackColor = vbYellow

End Sub

Private Sub optLoc_Click(Index As Integer)

10    If Index = 0 Then
20        lblGridTitle = "Printers available for force printing in the Laboratory."
30    Else
40        lblGridTitle = "Printers available for force printing on the Ward."
50    End If

60    LoadPrintersForLocation
End Sub
