VERSION 5.00
Begin VB.Form frmAddForcePrinter 
   Caption         =   "PrintHandler - Add Printer"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   8520
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton bSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   825
      Left            =   2910
      Picture         =   "frmAddForcePrinter.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3150
      Width           =   795
   End
   Begin VB.CommandButton bcancel 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   825
      Left            =   4425
      Picture         =   "frmAddForcePrinter.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3150
      Width           =   795
   End
   Begin VB.ComboBox cmbPrinter 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   195
      TabIndex        =   4
      Top             =   2235
      Width           =   8025
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
      Left            =   180
      TabIndex        =   0
      Top             =   420
      Width           =   4785
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
         Left            =   1410
         TabIndex        =   2
         Top             =   315
         Value           =   -1  'True
         Width           =   1320
      End
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
         Left            =   2835
         TabIndex        =   1
         Top             =   315
         Width           =   1035
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Available Printers"
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
      Left            =   210
      TabIndex        =   5
      Top             =   1890
      Width           =   4725
   End
   Begin VB.Label lblDesc 
      Caption         =   "Add printer to Laboratory Force printer list."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   210
      TabIndex        =   3
      Top             =   1380
      Width           =   4725
   End
End
Attribute VB_Name = "frmAddForcePrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub bcancel_Click()
10    Unload Me
End Sub



Private Sub bSave_Click()

10    If Len(Trim$(cmbPrinter)) > 0 Then
          'Check an see if it's assigned already
20        If PrinterAssignedAlready(cmbPrinter) Then
30            iMsg "Printer assigned already"
40        Else
50            iMsg "Printer added to " & IIf(optLoc(0), "Lab", "Ward") & " list."
60        End If
70        cmbPrinter = ""
80    End If

End Sub

Private Function PrinterAssignedAlready(ByVal strPrinter As String) As Boolean

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo PrinterAssignedAlready_Error

20    sql = "Select * from InstalledPrinters where PrinterName  = '" & strPrinter & "' " & _
      "and Location = '" & IIf(optLoc(0), "LAB", "WARD") & "'"

30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql

50    If tb.EOF Then
60        tb.AddNew
70        tb!PrinterName = strPrinter
80        tb!Location = IIf(optLoc(0), "LAB", "WARD")
90        tb.Update
100       PrinterAssignedAlready = False
110   Else
120       PrinterAssignedAlready = True
130   End If

140   Exit Function

PrinterAssignedAlready_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "frmAddForcePrinter", "PrinterAssignedAlready", intEL, strES, sql

End Function

Private Sub cmbPrinter_Click()
10    bSave.Enabled = True
End Sub

Private Sub cmbPrinter_KeyPress(KeyAscii As Integer)
10    KeyAscii = 0
End Sub

Private Sub Form_Load()
      Dim Px As Printer

10    cmbPrinter.Clear

20    For Each Px In Printers
30      cmbPrinter.AddItem Px.DeviceName
40    Next


End Sub

Private Sub optLoc_Click(Index As Integer)

10    cmbPrinter = ""

20    If Index = 0 Then
30        lblDesc = "Add printer to Laboratory force printer list."
40    Else
50        lblDesc = "Add printer to Ward force printer list."
60    End If

End Sub
