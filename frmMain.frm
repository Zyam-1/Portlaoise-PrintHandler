VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NetAcquire - Print Handler"
   ClientHeight    =   4605
   ClientLeft      =   3675
   ClientTop       =   4440
   ClientWidth     =   3885
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   3885
   Begin MSFlexGridLib.MSFlexGrid g 
      Height          =   4365
      Left            =   4050
      TabIndex        =   18
      Top             =   60
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   7699
      _Version        =   393216
      Rows            =   7
      Cols            =   7
      FixedRows       =   6
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   3
   End
   Begin VB.Frame Frame2 
      Caption         =   "Immunology"
      Height          =   645
      Left            =   90
      TabIndex        =   15
      Top             =   2250
      Width           =   3645
      Begin VB.Label lblImmMoreThan 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "18"
         Height          =   255
         Left            =   570
         TabIndex        =   17
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Immunology Results per Page"
         Height          =   225
         Left            =   1170
         TabIndex        =   16
         Top             =   270
         Width           =   2205
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdPending 
      Height          =   4380
      Left            =   9915
      TabIndex        =   11
      Top             =   45
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   7726
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      FormatString    =   "Sampleid   |D| Printer                 | FaxNumber     | Ward              | Clin                    |Gp                       "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid gDiff 
      Height          =   1815
      Left            =   0
      TabIndex        =   10
      Top             =   4830
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   3
   End
   Begin VB.Frame Frame1 
      Caption         =   "Biochemistry"
      Height          =   1095
      Left            =   90
      TabIndex        =   5
      Top             =   990
      Width           =   3645
      Begin VB.TextBox txtMoreThan 
         Height          =   255
         Left            =   1050
         TabIndex        =   9
         Text            =   "18"
         Top             =   210
         Width           =   525
      End
      Begin VB.OptionButton optSideBySide 
         Caption         =   "Print Side by Side"
         Height          =   195
         Left            =   1200
         TabIndex        =   8
         Top             =   780
         Width           =   1575
      End
      Begin VB.OptionButton optSecondPage 
         Caption         =   "Print on Second Page"
         Height          =   195
         Left            =   1200
         TabIndex        =   7
         Top             =   540
         Value           =   -1  'True
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "If more than xxxxxx  Biochemistry Results then "
         Height          =   225
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   3315
      End
   End
   Begin VB.OptionButton oNREnabled 
      Alignment       =   1  'Right Justify
      Caption         =   "Enabled"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   780
      TabIndex        =   4
      Top             =   720
      Value           =   -1  'True
      Width           =   1035
   End
   Begin VB.OptionButton oNRDisabled 
      Caption         =   "Disabled"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1830
      TabIndex        =   3
      Top             =   720
      Width           =   1065
   End
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000016&
      Height          =   3585
      Left            =   4920
      ScaleHeight     =   3525
      ScaleWidth      =   8805
      TabIndex        =   1
      Top             =   4560
      Width           =   8865
   End
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   225
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3120
      Top             =   300
   End
   Begin VB.Label lblPH_Roles 
      Height          =   195
      Left            =   105
      TabIndex        =   14
      Top             =   3465
      Width           =   3630
   End
   Begin VB.Label lblServer 
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   3750
      Width           =   3630
   End
   Begin VB.Label lblver 
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   4035
      Width           =   3630
   End
   Begin VB.Label lblHaem 
      Alignment       =   2  'Center
      Caption         =   "Haematology Age/Sex Related Normal Ranges are"
      Height          =   405
      Left            =   750
      TabIndex        =   2
      Top             =   300
      Width           =   2175
   End
   Begin VB.Menu mFile 
      Caption         =   "&File"
      Begin VB.Menu mPrinters 
         Caption         =   "&Mapped Printer setup"
      End
      Begin VB.Menu mnuForcePrinter 
         Caption         =   "&Force Printer Setup"
      End
      Begin VB.Menu mnuSetPHLoc 
         Caption         =   "&Set Print Handler Locations"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPHRoles 
         Caption         =   "Set Print Handler &Roles"
      End
      Begin VB.Menu mnuView 
         Caption         =   "View &Outstanding"
      End
      Begin VB.Menu mHaemNormal 
         Caption         =   "&Haem Normal Ranges"
      End
      Begin VB.Menu mNull 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Loading As Boolean

Private Sub ProcessPrintQueue(ByVal strConfig As String)

      Dim tb As Recordset
      Dim tbDem As Recordset
      Dim sql As String
      Dim SampleID As String
      Dim Initiator As String
      Dim ForcedPrintDone As Boolean
      Dim PrintThis As Boolean
      Dim Faxfound As Boolean
      Dim ForceFound As Boolean
      Dim Year As String
      Dim strRoles As String

10    On Error GoTo ProcessPrintQueue_Error

20    ForceFound = False
30    Faxfound = False
40    ForcedPrintDone = False

      '<Lab><Ward><Fax><MICR>
50    Select Case strConfig

      Case "0000":    '1
60        strRoles = "<NONE>"
70        lblPH_Roles = "Print Handler Roles: " & strRoles
80        Exit Sub
90    Case "1000":    '2  <LAB><><><>
100       sql = "SELECT * FROM PrintPending WHERE (COALESCE(WardPrint, '') = '') " & _
                "AND (COALESCE(FaxNumber, '') = '') " & _
                "AND (Department <> 'N') " & _
                "ORDER BY Department, ptime"
110       strRoles = "<LAB>"
120   Case "0100":    '3  <><WARD><><>
130       sql = "SELECT * FROM PrintPending WHERE WardPrint = 1 ORDER BY ptime "
140       strRoles = "<WARD>"
150   Case "0010":    '4  <><><FAX><>
160       sql = "SELECT * FROM PrintPending WHERE (COALESCE(FaxNumber, '') <> '') " & _
                "AND (COALESCE(WardPrint, '') = '') " & _
                "ORDER BY FaxNumber DESC, Department, ptime ASC"
170       strRoles = "<FAX>"
180   Case "0001":    '5  <><><><MICRO>
190       sql = "Select * from PrintPending WHERE (Department = 'N') " & _
                "AND COALESCE(FaxNumber, '') = '' " & _
                "AND (COALESCE(WardPrint, '') = '') ORDER BY ptime "
200       strRoles = "<MICRO>"
210   Case "1100"    '6   <LAB><WARD><><>
220       sql = "SELECT * FROM PrintPending WHERE " & _
                "(COALESCE(FaxNumber, '') = '') " & _
                "AND (Department <> 'N' OR COALESCE(WardPrint, '') <> '') " & _
                "ORDER BY Department, ptime"
230       strRoles = "<LAB><WARD>"
240   Case "1010"    '7   <LAB><><FAX><>
250       sql = "SELECT * FROM PrintPending WHERE " & _
                "(COALESCE(WardPrint, '') = '') " & _
                "AND (Department <> 'N') " & _
                "ORDER BY Department, ptime"
260       strRoles = "<LAB><FAX>"
270   Case "1001"    '8   <LAB><><><MICRO>
280       sql = "SELECT * FROM PrintPending WHERE " & _
                "(COALESCE(WardPrint, '') = '') " & _
                "AND (COALESCE(FaxNumber, '') = '') " & _
                "ORDER BY Department, ptime"
290       strRoles = "<LAB><MICRO>"
300   Case "0110"    '9   <><WARD><FAX><>
310       sql = "SELECT * FROM PrintPending WHERE " & _
                "(WardPrint = 1) OR COALESCE(FaxNumber, '') <> '' " & _
                "ORDER BY Department, ptime"
320       strRoles = "<WARD><FAX>"
330   Case "0101"    '10  <><WARD><><MICRO>
340       sql = "SELECT * FROM PrintPending WHERE " & _
                "(WardPrint =1) OR (Department = 'N') " & _
                "ORDER BY Department, ptime"
350       strRoles = "<WARD><MICRO>"
360   Case "0011"    '11  <><><FAX><MICRO>
370       sql = "SELECT * FROM PrintPending WHERE " & _
                "(COALESCE(WardPrint, '') = '') " & _
                "AND COALESCE(FaxNumber, '') <> '' " & _
                "OR (Department = 'N' AND COALESCE(WardPrint, '') = '') " & _
                "ORDER BY Department, ptime"
380       strRoles = "<FAX><MICRO>"
390   Case "1110"    '12  <LAB><WARD><FAX><>
400       sql = "SELECT * FROM PrintPending WHERE " & _
                "(Department <> 'N' OR COALESCE(WardPrint, '') <> '')" & _
                "ORDER BY Department, ptime"
410       strRoles = "<LAB><WARD><FAX>"
420   Case "1101"    '13  <LAB><WARD><><MICRO>
430       sql = "SELECT * FROM PrintPending WHERE " & _
                "COALESCE(FaxNumber, '') = '' " & _
                "ORDER BY Department, ptime"
440       strRoles = "<LAB><WARD><MICRO>"
450   Case "1011"    '14  <LAB><><FAX><MICRO>
460       sql = "SELECT * FROM PrintPending WHERE " & _
                "COALESCE(WardPrint, '') = '' " & _
                "ORDER BY Department, ptime"
470       strRoles = "<LAB><FAX><MICRO>"
480   Case "0111"    '15  <><WARD><FAX><MICRO>
490       sql = "SELECT * FROM PrintPending WHERE " & _
                "(WardPrint = 1) OR COALESCE(FaxNumber, '') <> '' " & _
                "OR (COALESCE(Department, '') = 'N') " & _
                "ORDER BY Department, ptime"
500       strRoles = "<WARD><FAX><MICRO>"
510   Case "1111"    '16  <LAB><WARD><FAX><MICRO>
520       sql = "SELECT * FROM PrintPending ORDER BY Department, ptime"
530       strRoles = "<LAB><WARD><FAX><MICRO>"
540   End Select

550   lblPH_Roles = "Print Handler Roles: " & strRoles

560   Set tb = New Recordset
570   RecOpenClient 0, tb, sql

580   Do While Not tb.EOF

590       frmRichText.rtb.Text = ""

600       LogEvent "ProcessPrintQueue - ForceFound", ForceFound
610       LogEvent "ProcessPrintQueue - tb!UsePrinter", tb!UsePrinter & ""
620       LogEvent "ProcessPrintQueue - SampleID", tb!SampleID & ""

630       If (ForceFound = False Or Trim$(tb!UsePrinter & "") = "") Then
640           If Trim(tb!SampleID & "") = "" Then
650               LogEvent "ProcessPrintQueue - Deleting blank SampleID", ""
660               sql = "DELETE FROM PrintPending WHERE SampleID = ''"
670               Cnxn(0).Execute sql
680               LogEvent "ProcessPrintQueue - Deleted blank SampleID", ""
690           Else
700               LogEvent "ProcessPrintQueue - Selecting SampleID", tb!SampleID
710               sql = "SELECT SampleDate, RunDate, PatName FROM Demographics WHERE " & _
                        "SampleID = '" & tb!SampleID & "'"
720               Set tbDem = New Recordset
730               RecOpenServer 0, tbDem, sql
740               If Not tb.EOF Then
                      'LogEvent "ProcessPrintQueue - Not tb!EOF", tb!SampleID
750                   If IsDate(tbDem!SampleDate) Then
760                       RP.SampleDate = tbDem!SampleDate
770                   End If
780                   If IsDate(tbDem!Rundate) Then
790                       RP.Rundate = tbDem!Rundate
800                   End If
810                   RP.PatientName = tbDem!PatName & ""
820               Else
830                   LogEvent "ProcessPrintQueue - tb!EOF", tb!SampleID
840               End If
850               RP.Year = Trim(tb!Hyear & "")
860               RP.SampleID = Trim(tb!SampleID & "")
870               RP.Initiator = Trim(tb!Initiator & "")
880               RP.Department = Trim(tb!Department & "")
890               RP.Ward = tb!Ward & ""
900               RP.Clinician = tb!Clinician & ""
910               RP.GP = tb!GP & ""
920               RP.FaxNumber = Trim(tb!FaxNumber & "")
930               RP.PTime = Format(tb!PTime, "dd/MMM/yyyy hh:mm:ss")
940               RP.SendCopyTo = ""
950               RP.NoOfCopies = IIf(IsNull(tb!NoOfCopies), 1, tb!NoOfCopies)
960               pForcePrintTo = Trim$(tb!UsePrinter & "")
970               RP.FinalInterim = IIf(tb!FinalInterim & "" = "", "F", tb!FinalInterim & "")
980               RP.WardPrint = IIf(IsNull(tb!WardPrint), False, tb!WardPrint)
990               RP.PrintAction = IIf(IsNull(tb!PrintAction), "", tb!PrintAction)    'Masood 19_Feb_2013

      '            LogEvent "ProcessPrintQueue - RP.SampleDate", RP.SampleDate
      '            LogEvent "ProcessPrintQueue - RP.RunDate", RP.Rundate
      '            LogEvent "ProcessPrintQueue - RP.PatientName", RP.PatientName
      '            LogEvent "ProcessPrintQueue - RP.Year", RP.Year
      '            LogEvent "ProcessPrintQueue - RP.SampleID", RP.SampleID
      '            LogEvent "ProcessPrintQueue - RP.Initiator", RP.Initiator
      '            LogEvent "ProcessPrintQueue - RP.Department", RP.Department
      '            LogEvent "ProcessPrintQueue - RP.Ward", RP.Ward
      '            LogEvent "ProcessPrintQueue - RP.Clinician", RP.Clinician
      '            LogEvent "ProcessPrintQueue - RP.GP", RP.GP
      '            LogEvent "ProcessPrintQueue - RP.FAXNumber", RP.FaxNumber
      '            LogEvent "ProcessPrintQueue - RP.PTime", RP.PTime
      '            LogEvent "ProcessPrintQueue - RP.NoOfCopies", RP.NoOfCopies
      '            LogEvent "ProcessPrintQueue - pForcePrintTo", pForcePrintTo
      '            LogEvent "ProcessPrintQueue - RP.FinalInterim", RP.FinalInterim
      '            LogEvent "ProcessPrintQueue - RP.WardPrint", RP.WardPrint

1000              If InStr(1, RP.PrintAction, "Print") > 0 Or RP.PrintAction = "" Then
1010                  If Not SetCurrentPrinter(pForcePrintTo) Then
1020                      pForcePrintTo = ""
1030                      If RP.SampleID <> "" Then
1040                          sql = "DELETE FROM PrintPending WHERE " & _
                                    "SampleID = '" & RP.SampleID & "' " & _
                                    "and COALESCE(PrintAction, '') = '" & RP.PrintAction & "' " & _
                                    "AND Department = '" & RP.Department & "'"
1050                          Cnxn(0).Execute sql
1060                      End If
1070                      Exit Sub
1080                  End If
1090              End If
1100              PrintThis = False

1110              If Trim$(tb!UsePrinter & "") = "" Then
1120                  PrintThis = True
1130              Else
1140                  If Not ForcedPrintDone Then
1150                      PrintThis = True
1160                      ForcedPrintDone = True
1170                  End If
1180              End If

1190              If PrintThis Then
1200                  If PrintRecord() = True Then
1210                      If RP.FaxNumber = "" Then
1220                          CheckCC
1230                      End If
1240                  End If
1250                  RemovePrintInihibitEntries RP.SampleID
1260                  sql = "DELETE FROM PrintPending WHERE " & _
                            "SampleID = '" & tb!SampleID & "' " & _
                            "and COALESCE(PrintAction, '') = '" & RP.PrintAction & "' " & _
                            "AND Department = '" & tb!Department & "'"
1270                   Cnxn(0).Execute sql
1280              End If

1290          End If
1300      End If

1310      Select Case RP.Department
          Case "B"
              '      sql = "Update BioResults set Printed = '1' WHERE " & _
                     '            "SampleID = '" & RP.SampleID & "' and valid = '1'"
              '      Cnxn(0).Execute sql
1320      Case "I"
1330          sql = "Update ImmResults set Printed = '1' WHERE " & _
                    "SampleID = '" & RP.SampleID & "' and valid = '1'"
1340          Cnxn(0).Execute sql
1350      Case "E"
              '      sql = "Update EndResults set Printed = '1' WHERE " & _
                     '            "SampleID = '" & RP.SampleID & "' and valid = '1'"
              '      Cnxn(0).Execute sql
1360      Case "J"
              '      sql = "Update ImmResults set Printed = '1' WHERE " & _
                     '            "SampleID = '" & RP.SampleID & "' and valid = '1'"
              '      Cnxn(0).Execute sql
1370      Case "H"
1380          sql = "Update HaemResults " & _
                    "set Printed = 1, Valid = 1 " & _
                    "WHERE SampleID = '" & RP.SampleID & "'"
1390          Cnxn(0).Execute sql
1400      Case "N"
1410          If SysOptShowIQ200(0) = True Then
1420              sql = "Update IQ200 " & _
                        "set Printed = 1 " & _
                        "WHERE SampleID = '" & RP.SampleID & "'"
1430              Cnxn(0).Execute sql
1440          End If
1450      End Select

1460      tb.MoveNext
1470  Loop

1480  Exit Sub

ProcessPrintQueue_Error:

      Dim strES As String
      Dim intEL As Integer

1490  intEL = Erl
1500  strES = Err.Description
1510  LogError "frmMain", "ProcessPrintQueue", intEL, strES, sql

1520  If RP.SampleID <> "" Then
1530      sql = "DELETE FROM PrintPending WHERE " & _
                "SampleID = '" & RP.SampleID & "' " & _
                "and COALESCE(PrintAction, '') = '" & RP.PrintAction & "' " & _
                "AND Department = '" & RP.Department & "'"
1540      Cnxn(0).Execute sql
1550  End If

End Sub


Private Sub ProcessPrintQueueOld(ByVal strConfig As String)

      Dim tb As Recordset
      Dim tbDem As Recordset
      Dim sql As String
      Dim SampleID As String
      Dim Initiator As String
      Dim ForcedPrintDone As Boolean
      Dim PrintThis As Boolean
      Dim Px As Printer
      Dim xFound As Boolean
      Dim Faxfound As Boolean
      Dim ForceFound As Boolean
      Dim Year As String
      Dim strRoles As String

10    On Error GoTo ProcessPrintQueue_Error

20    ForceFound = False
30    Faxfound = False
40    ForcedPrintDone = False

      '<Lab><Ward><Fax>

50    Select Case strConfig

      Case "000":
60        strRoles = "<NONE>"
70        lblPH_Roles = "Print Handler Roles: " & strRoles
80        Exit Sub
90    Case "100":    '<LAB><><>
100       sql = "SELECT * FROM PrintPending WHERE (WardPrint = '' OR WardPrint IS NULL) " & _
                "AND (FaxNumber = '' OR FaxNumber IS NULL) " & _
                "ORDER BY Department, ptime"
110       strRoles = "<LAB>"
120   Case "010":    '<><WARD><>
130       sql = "SELECT * FROM PrintPending WHERE WardPrint = 1"
140       strRoles = "<WARD>"
150   Case "001":    '<><><FAX>
160       sql = "Select * from PrintPending where " & _
                "faxnumber <> '' and (WardPrint = '' OR WardPrint IS NULL) order by faxnumber desc, department, ptime asc"
170       strRoles = "<FAX>"
180   Case "110"    '<LAB><WARD><>
190       sql = "SELECT * FROM PrintPending WHERE " & _
                "(FaxNumber = '' OR FaxNumber IS NULL) " & _
                "ORDER BY Department, ptime"
200       strRoles = "<LAB><WARD>"
210   Case "101"    '<LAB><><FAX>
220       sql = "SELECT * FROM PrintPending WHERE " & _
                "(WardPrint = '' OR WardPrint IS NULL)" & _
                "ORDER BY Department, ptime"
230       strRoles = "<LAB><FAX>"
240   Case "011"    '<><WARD><FAX>
250       sql = "SELECT * FROM PrintPending WHERE WardPrint = 1 AND faxnumber <> '' "
260       strRoles = "<WARD><FAX>"
270   Case "111"    '<LAB><WARD><FAX>
280       sql = "SELECT * FROM PrintPending ORDER BY Department, ptime"
290       strRoles = "<LAB><WARD><FAX>"
300   End Select

310   lblPH_Roles = "Print Handler Roles: " & strRoles

320   Set tb = New Recordset
330   RecOpenClient 0, tb, sql

340   Do While Not tb.EOF

350       frmRichText.rtb.Text = ""
360       If (ForceFound = False Or Trim$(tb!UsePrinter & "") = "") Then
370           If Trim(tb!SampleID & "") = "" Then
380               sql = "DELETE FROM PrintPending WHERE " & _
                        "SampleID = '" & Trim(tb!SampleID) & "'"
390               Cnxn(0).Execute sql
400           Else
410               sql = "SELECT SampleDate, RunDate, PatName FROM Demographics WHERE " & _
                        "SampleID = '" & tb!SampleID & "'"
420               Set tbDem = New Recordset
430               RecOpenServer 0, tbDem, sql
440               If Not tb.EOF Then
450                   If IsDate(tbDem!SampleDate) Then
460                       RP.SampleDate = tbDem!SampleDate
470                   End If
480                   If IsDate(tbDem!Rundate) Then
490                       RP.Rundate = tbDem!Rundate
500                   End If
510                   RP.PatientName = tbDem!PatName & ""
520               End If
530               RP.Year = Trim(tb!Hyear & "")
540               RP.SampleID = Trim(tb!SampleID & "")
550               RP.Initiator = Trim(tb!Initiator & "")
560               RP.Department = Trim(tb!Department & "")
570               RP.Ward = tb!Ward & ""
580               RP.Clinician = tb!Clinician & ""
590               RP.GP = tb!GP & ""
600               RP.FaxNumber = Trim(tb!FaxNumber & "")
610               RP.PTime = Format(tb!PTime, "dd/MMM/yyyy hh:mm:ss")

620               pForcePrintTo = Trim$(tb!UsePrinter & "")

630               If pForcePrintTo <> "" Then
640                   gOriginalPrinter = Printer.DeviceName
650                   xFound = False
660                   For Each Px In Printers
670                       If UCase$(Px.DeviceName) = UCase$(pForcePrintTo) Then
680                           Set Printer = Px
690                           xFound = True
700                           ForceFound = True
710                           Exit For
720                       End If
730                   Next
740                   If Not xFound Then
750                       pForcePrintTo = ""
760                       If RP.SampleID <> "" Then
770                           sql = "DELETE FROM PrintPending WHERE " & _
                                    "SampleID = '" & RP.SampleID & "' " & _
                                    "AND Department = '" & RP.Department & "'"
780                           Cnxn(0).Execute sql
790                       End If
800                       Exit Sub
810                   End If
820               End If

830               PrintThis = False

840               If Trim$(tb!UsePrinter & "") = "" Then
850                   PrintThis = True
860               Else
870                   If Not ForcedPrintDone Then
880                       PrintThis = True
890                       ForcedPrintDone = True
900                   End If
910               End If

920               If PrintThis Then
930                   If PrintRecord() = True Then
940                       RP.SampleID = tb!SampleID
950                       If RP.FaxNumber = "" Then CheckCC
960                       If RP.Department = "P" Or RP.Department = "Y" Then
970                           sql = "delete FROM PrintPending WHERE SampleID = '" & tb!SampleID & "' and department = '" & tb!Department & "' and ptime = '" & RP.PTime & "'"
980                       Else
990                           If RP.FaxNumber <> "" Then
1000                              sql = "delete FROM PrintPending WHERE SampleID = '" & tb!SampleID & "' and department = '" & tb!Department & "' and faxnumber = '" & Trim(tb!FaxNumber) & "'"
1010                          Else
1020                              sql = "delete FROM PrintPending WHERE SampleID = '" & tb!SampleID & "' and department = '" & tb!Department & "' and (faxnumber = '' or faxnumber is null)"
1030                          End If
1040                      End If
1050                      Cnxn(0).Execute sql
1060                  End If
1070              End If

1080          End If
1090      End If

1100      Select Case RP.Department
          Case "B"
              '      sql = "Update BioResults set Printed = '1' WHERE " & _
                     '            "SampleID = '" & RP.SampleID & "' and valid = '1'"
              '      Cnxn(0).Execute sql
1110      Case "I"
1120          sql = "Update ImmResults set Printed = '1' WHERE " & _
                    "SampleID = '" & RP.SampleID & "' and valid = '1'"
1130          Cnxn(0).Execute sql
1140      Case "E"
              '      sql = "Update EndResults set Printed = '1' WHERE " & _
                     '            "SampleID = '" & RP.SampleID & "' and valid = '1'"
              '      Cnxn(0).Execute sql
1150      Case "J"
              '      sql = "Update ImmResults set Printed = '1' WHERE " & _
                     '            "SampleID = '" & RP.SampleID & "' and valid = '1'"
              '      Cnxn(0).Execute sql
1160      Case "H"
1170          sql = "Update HaemResults " & _
                    "set Printed = 1, Valid = 1 " & _
                    "WHERE SampleID = '" & RP.SampleID & "'"
1180          Cnxn(0).Execute sql
1190      End Select

1200      tb.MoveNext
1210  Loop

1220  Exit Sub

ProcessPrintQueue_Error:

      Dim strES As String
      Dim intEL As Integer

1230  intEL = Erl
1240  strES = Err.Description
1250  LogError "frmMain", "ProcessPrintQueue", intEL, strES, sql

1260  If RP.SampleID <> "" Then
1270      sql = "DELETE FROM PrintPending WHERE " & _
                "SampleID = '" & RP.SampleID & "' " & _
                "AND Department = '" & RP.Department & "'"
1280      Cnxn(0).Execute sql
1290  End If

End Sub

Private Sub Form_Activate()

      Dim Path As String

10    On Error GoTo Form_Activate_Error

20    Path = CheckNewEXE("PrintHandler")    '<---Change this to your prog Name
        'Zyam
'30    If Path <> "" Then
'40        Shell App.Path & "\CustomStart.exe PrintHandler"    '<---Change this to your prog Name
'50        End
'60        Exit Sub
'70    End If
        'Zysm

80    Timer1.Enabled = True

90    lblImmMoreThan.Caption = GetOptionSetting("ImmMaxPrintLines", "14")

100   Exit Sub

Form_Activate_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "frmMain", "Form_Activate", intEL, strES

End Sub



Private Sub Form_Deactivate()
'
'10    On Error GoTo Form_Deactivate_Error
'
'20    Timer1.Enabled = False
'
'30    Exit Sub
'
'Form_Deactivate_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'40    intEL = Erl
'50    strES = Err.Description
'60    LogError "frmMain", "Form_Deactivate", intEL, strES
'
End Sub


Private Sub Form_Load()

      Dim strUseSecondPage As String

10    On Error GoTo Form_Load_Error

20    If App.PrevInstance Then End

30    lblver = "Print Handler Version : " & App.Major & "." & App.Minor
40    CheckIDE

50    ConnectToDatabase
'      colPRNs.Load
60    If colPRNs.Count = 0 Then MsgBox "Printers not Loaded!", vbExclamation
70    LoadOptions
80    CheckPrintHandlerLogInDb
90    LogEvent "Print Handler Started", ""
100   Loading = True
110   txtMoreThan = GetOptionSetting("PrintOptionsIfMoreThan", "18")

120   strUseSecondPage = GetOptionSetting("PrintOptionsUseSecondPage", "True")
130   If strUseSecondPage = "True" Then
140       optSecondPage = True
150   Else
160       optSideBySide = True
170   End If

180   Loading = False

190   If TestSys = True Then Me.Caption = "TEST SYSTEM - Print Handler"

200   lblServer = "Print Handler Server: " & vbGetComputerName()

210   Exit Sub

Form_Load_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "frmMain", "Form_Load", intEL, strES

End Sub





Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

10    If UCase(iBOX("Password required to close", , , True)) <> "TEMO" Then
20        Cancel = True
30        Exit Sub
40    End If

50    Unload frmRichText

End Sub



Private Sub lblImmMoreThan_Click()

      Dim NewValue As Integer

10    NewValue = Val(iBOX("Maximum number of Immunology Results" & vbCrLf & _
                          "per page?", , lblImmMoreThan))
20    If NewValue > 1 And NewValue < 30 Then

30        lblImmMoreThan.Caption = Format$(NewValue)
40        SaveOptionSetting "ImmMaxPrintLines", lblImmMoreThan

50    End If

End Sub

Private Sub mExit_Click()

10    On Error GoTo mExit_Click_Error

20    Unload frmRichText
30    Unload Me

40    Exit Sub

mExit_Click_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "frmMain", "mExit_Click", intEL, strES

End Sub

Private Sub mHaemNormal_Click()

10    On Error GoTo mHaemNormal_Click_Error

20    fHaemNoSexNormal.Show 1

30    Exit Sub

mHaemNormal_Click_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmMain", "mHaemNormal_Click", intEL, strES

End Sub



Private Sub mnuForcePrinter_Click()
10    frmForcePrinters.Show 1
End Sub

Private Sub mnuPHRoles_Click()
10    frmSetRole.Show 1
End Sub

'Private Sub mnuSetPHLoc_Click()
'10    frmSetPHLocation.Show 1
'20    SetPrintHandlerLocation
'End Sub

Private Sub mnuView_Click()

10    On Error GoTo mnuView_Click_Error

20    If mnuView.Caption = "Hide Outstanding" Then
30        mnuView.Caption = "View Outstanding"
40        Me.Width = 3900
50    Else
60        mnuView.Caption = "Hide Outstanding"
70        Me.Width = 13275
80    End If

90    Exit Sub

mnuView_Click_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "frmMain", "mnuView_Click", intEL, strES

End Sub

Private Sub mPrinters_Click()

10    fPrinters.Show 1

End Sub

Private Sub optSecondPage_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)

10    On Error GoTo optSecondPage_MouseUp_Error

20    SaveOptionSetting "PrintOptionsUseSecondPage", "True"

30    Exit Sub

optSecondPage_MouseUp_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmMain", "optSecondPage_MouseUp", intEL, strES

End Sub


Private Sub optSideBySide_Click()

10    On Error GoTo optSideBySide_Click_Error

20    SaveOptionSetting "PrintOptionsUseSecondPage", "False"

30    Exit Sub

optSideBySide_Click_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "frmMain", "optSideBySide_Click", intEL, strES

End Sub




Private Sub Timer1_Timer()

      Static Counter As Integer
      Dim Path As String

10    On Error GoTo Timer1_Timer_Error

20    LogEvent "Timer Tick", "pbar = " & pbar

30    If (pbar + 10) > pbar.Max Then
40        pbar = 0
50    End If

60    pbar = pbar + 10
70    If pbar = pbar.Max Then
80        LogEvent "Processing Print Queue", ""
90        ProcessPrintQueue (getPrintHandlerConfiguration)
100       pbar = 0
110   End If

120   Counter = Counter + 1
130   If Counter >= 120 Then    'timer fires every 500mS: 120 = 1 minute
140       LogEvent "Check for new exe", ""
150       Path = CheckNewEXE("PrintHandler")    '<---Change this to your prog Name
160       LogEvent "Path", Path
170       If Path <> "" Then
180           LogEvent "Shell to new exe", ""
190           Shell App.Path & "\CustomStart.exe PrintHandler"    '<---Change this to your prog Name
200           End
210           Exit Sub
220       End If

230       colFastings.Refresh
240       colPRNs.Refresh
250       Counter = 0
260   End If

270   Exit Sub

Timer1_Timer_Error:

      Dim strES As String
      Dim intEL As Integer

280   intEL = Erl
290   strES = Err.Description
300   LogError "frmMain", "Timer1_Timer", intEL, strES

310   Timer1.Enabled = True

End Sub

Private Sub txtMoreThan_Change()

On Error GoTo txtMoreThan_Change_Error
'Zyam 15-3-24
If Not Loading Then
    If Val(txtMoreThan) > 0 And Val(txtMoreThan) < 70 Then
        SaveOptionSetting "PrintOptionsIfMoreThan", txtMoreThan
    Else
        txtMoreThan = GetOptionSetting("PrintOptionsIfMoreThan", "18")
    End If
End If
'Zyam 15-3-24

Exit Sub

txtMoreThan_Change_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "frmMain", "txtMoreThan_Change", intEL, strES

End Sub


Private Sub CheckCC()

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo CheckCC_Error

20    sql = "SELECT * FROM SendCopyTo WHERE "
30    If RP.Department = "N" Then
40        sql = sql & "SampleID = '" & Val(RP.SampleID) & "'"
50    Else
60        sql = sql & "SampleID = '" & Val(RP.SampleID) & "'"
70    End If
80    Set tb = New Recordset
90    RecOpenServer 0, tb, sql
100   Do While Not tb.EOF

110       If Trim$(UCase$(RP.Ward)) <> Trim$(UCase$(tb!Ward & "")) And Trim$(tb!Ward & "") <> "" Then
120           RP.Ward = tb!Ward & " (Copy)"

130       End If

140       If Trim$(UCase$(RP.Clinician)) <> Trim$(UCase$(tb!Clinician & "")) And Trim$(tb!Clinician & "") <> "" Then
150           RP.Ward = tb!Ward & " (Copy)"
160       End If

170       If Trim$(UCase$(RP.GP)) <> Trim$(UCase$(tb!GP & "")) And Trim$(tb!GP & "") <> "" Then
180           RP.Ward = tb!Ward & " (Copy)"
190       End If

200       If Trim$(tb!Clinician & "") <> "" Then
210           RP.SendCopyTo = tb!Clinician
220       ElseIf Trim$(tb!GP & "") <> "" Then
230           RP.SendCopyTo = tb!GP
240       End If
250       SetCurrentPrinter (pForcePrintTo)
          '  If UCase(Trim$(tb!Device & "")) <> "PRINTER" Then
          '    If Not SetCurrentPrinter(tb!Device) Then
          '      Exit Sub
          '    End If
          '  End If

260       PrintRecord
270       tb.MoveNext

280   Loop

290   Exit Sub

CheckCC_Error:

      Dim strES As String
      Dim intEL As Integer

300   intEL = Erl
310   strES = Err.Description
320   LogError "frmMain", "CheckCC", intEL, strES, sql

End Sub

Private Sub txtMoreThan_LostFocus()
10    On Error GoTo txtMoreThan_LostFocus_Error

20    If Val(txtMoreThan) > 0 And Val(txtMoreThan) < 61 Then
30            SaveOptionSetting "PrintOptionsIfMoreThan", txtMoreThan
40        Else
50            txtMoreThan = GetOptionSetting("PrintOptionsIfMoreThan", "18")
60        End If

70    Exit Sub
txtMoreThan_LostFocus_Error:
         
80    LogError "frmMain", "txtMoreThan_LostFocus", Erl, Err.Description


End Sub
