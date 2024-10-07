Attribute VB_Name = "Other"
Option Explicit

Public Type ReportToPrint
    SampleID As String
    Department As String
    PatientName As String
    Ward As String
    DoB As String
    Chart As String
    Address0 As String
    Address1 As String
    Initiator As String
    Clinician As String
    GP As String
    Sex As String
    FaxNumber As String
    UsePrinter As String
    ThisIsCopy As Boolean
    SendCopyTo As String
    Year As String
    PTime As String
    Hospital As String
    SampleDate As String
    RecDate As String
    Rundate As String
    GpClin As String
    SampleType As String
    NoOfCopies As Integer
    FinalInterim As String
    WardPrint As Boolean
    PrintAction As String 'Masood 19_Feb_2013

End Type
Public RP As ReportToPrint

Private Type udtHead
    SampleID As String
    Dept As String
    Name As String
    Ward As String
    DoB As String
    Chart As String
    Clinician As String
    Address0 As String
    Address1 As String
    GP As String
    Sex As String
    Hospital As String
    SampleDate As String
    RecDate As String
    Rundate As String
    GpClin As String
    SampleType As String
    Notes As String
    DocumentNo As String
    AandE As String
End Type
Public udtHeading As udtHead

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public TestSys As Boolean

Public HospName(100) As String
Public dbConnect As String
Public dbConnectBB As String
Public dbSym As String
Public Hosp(100) As String
Public intOtherHospitalsInGroup As Integer
Public colFastings As New Fastings
Public colPRNs As New PRNs
Public colBgaResults As New BGAResults

Public Const MaxAgeToDays As Long = 43830

Public Cnxn(100) As Connection
Public CnxnBB As Connection
Public Cn As Integer

Public gData(1 To 365, 1 To 3) As Variant    '(n,1)=rundate, (n,2)=INR, (n,3)=Warfarin

Public LatestSampleID As String

Public LatestINR As String

Public CurrentDose As String
Public pLatest As String
Public pEarliest As String

Public pLowerTarget As String
Public pUpperTarget As String
Public pCondition As String

Public pForcePrintTo As String

Public Type PrintLine
    Analyte As String * 20
    Result As String * 8
    Flag As String * 3
    Units As String * 13
    NormalRange As String * 16
    Fasting As String * 9
    Reason As String * 23
    Comment As String * 54
End Type

Public Type ResultLine
    Analyte As String
    Result As String
    Flag As String
    Units As String
    NormalRange As String
    Fasting As String
    Reason As String
    Comment As String
    PrintBold As Boolean
    PrintItalic As Boolean
    PrintUnderline As Boolean
    
End Type

Public Enum PrintAlignContants
    AlignLeft = 0
    AlignCenter = 1
    AlignRight = 2
End Enum

Public strPrintHandlerLocation As String

Public Const UserName As String = "PrintHandler"    ' not used


Public DisplaySampleID As String


Public Function AddTicks(ByVal s As String) As String

10    AddTicks = Trim$(Replace(s, "'", "''"))

End Function


Public Sub ClearUdtHeading()

10    On Error GoTo ClearUdtHeading_Error

20    With udtHeading
30        .SampleID = ""
40        .Dept = ""
50        .Name = ""
60        .Ward = ""
70        .DoB = ""
80        .Chart = ""
90        .Clinician = ""
100       .Address0 = ""
110       .Address1 = ""
120       .GP = ""
130       .Sex = ""
140       .Hospital = ""
150       .SampleDate = ""
160       .RecDate = ""
170       .Rundate = ""
180       .GpClin = ""
190       .SampleType = ""
200       .Notes = ""
210   End With

220   Exit Sub

ClearUdtHeading_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "Other", "ClearUdtHeading", intEL, strES

End Sub

Public Function ListText(ByVal ListType As String, ByVal Code As String) As String

      Dim tb As New Recordset
      Dim sql As String

10    On Error GoTo ListText_Error

20    ListText = ""
30    Code = UCase$(Trim$(Code))

40    sql = "SELECT * FROM lists WHERE listtype = '" & ListType & "' and Code = '" & Code & "'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80        ListText = Trim(tb!Text)
90    End If

100   Exit Function

ListText_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "Other", "ListText", intEL, strES, sql

End Function


Public Sub FillCommentLines(ByVal FullComment As String, _
                            ByVal NumberOfLines As Integer, _
                            ByRef Comments() As String, _
                            Optional ByVal MaxLen As Integer = 80)

      Dim n As Integer
      Dim CurrentLine As Integer
      Dim X As Integer
      Dim ThisLine As String
      Dim SpaceFound As Boolean

10    On Error GoTo FillCommentLines_Error

20    For n = 1 To UBound(Comments)
30        Comments(n) = ""
40    Next

50    CurrentLine = 0
60    FullComment = Trim(FullComment)
70    FullComment = Replace(FullComment, vbCrLf, " ")
80    n = Len(FullComment)

90    For X = n - 1 To 1 Step -1
100       If Mid(FullComment, X, 1) = vbCr Or Mid(FullComment, X, 1) = vbLf Or Mid(FullComment, X, 1) = vbTab Then
110           Mid(FullComment, X, 1) = " "
120       End If
130   Next

140   For X = n - 3 To 1 Step -1
150       If Mid(FullComment, X, 2) = "  " Then
160           FullComment = Left(FullComment, X) & Mid(FullComment, X + 2)
170       End If
180   Next
190   n = Len(FullComment)

200   Do While n > MaxLen
210       SpaceFound = False
220       For X = MaxLen To 1 Step -1
230           If Mid(FullComment, X, 1) = " " Then
240               ThisLine = Left(FullComment, X - 1)
250               FullComment = Mid(FullComment, X + 1)

260               CurrentLine = CurrentLine + 1
270               If CurrentLine <= NumberOfLines Then
280                   Comments(CurrentLine) = ThisLine
290               End If
300               SpaceFound = True
310               Exit For
320           End If
330       Next
340       If Not SpaceFound Then
350           ThisLine = Left(FullComment, MaxLen)
360           FullComment = Mid(FullComment, MaxLen + 1)

370           CurrentLine = CurrentLine + 1
380           If CurrentLine <= NumberOfLines Then
390               Comments(CurrentLine) = ThisLine
400           End If
410       End If
420       n = Len(FullComment)
430   Loop

440   CurrentLine = CurrentLine + 1
450   If CurrentLine <= NumberOfLines Then
460       Comments(CurrentLine) = FullComment
470   End If

480   Exit Sub

FillCommentLines_Error:

      Dim strES As String
      Dim intEL As Integer

490   intEL = Erl
500   strES = Err.Description
510   LogError "Other", "FillCommentLines", intEL, strES

End Sub




Public Function Initial2Upper(ByVal s As String) As String

      Dim n As Integer

10    On Error GoTo Initial2Upper_Error

20    s = Trim$(s & "")
30    If s = "" Then
40        Initial2Upper = ""
50        Exit Function
60    End If

70    If InStr(UCase$(s), "MAC") > 0 Or InStr(UCase$(s), "MC") > 0 Or InStr(s, "'") > 0 Then
80        s = LCase$(s)
90        s = UCase$(Left$(s, 1)) & Mid(s, 2)

100       For n = 1 To Len(s) - 1
110           If Mid(s, n, 1) = " " Or Mid(s, n, 1) = "'" Then
120               s = Left$(s, n) & UCase$(Mid(s, n + 1, 1)) & Mid(s, n + 2)
130           End If
140           If n > 1 Then
150               If Mid(s, n, 1) = "c" And Mid(s, n - 1, 1) = "M" Then
160                   s = Left$(s, n) & UCase$(Mid(s, n + 1, 1)) & Mid(s, n + 2)
170               End If
180           End If
190       Next
200   Else
210       s = StrConv(s, vbProperCase)
220   End If
230   Initial2Upper = s

240   Exit Function

Initial2Upper_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "Other", "Initial2Upper", intEL, strES

End Function

Public Function InterpCoag(ByVal n As Integer, _
                           ByVal Sex As String, _
                           ByVal TestCode As String, _
                           ByVal Result As String, _
                           ByVal DaysOld As String) _
                           As String

    Dim tb As New Recordset
    Dim Low As String
    Dim High As String
    Dim sql As String
    Dim PlaHigh As String
    Dim PlaLow As String
    Dim LessSignFlag As Boolean
    Dim GreaterSignFlag As Boolean

    LessSignFlag = False
    GreaterSignFlag = False

10  On Error GoTo InterpCoag_Error
    'Zyam 04-10-24
    If InStr(1, Result, "<") Then
        LessSignFlag = True
        Result = Replace(Result, "<", "")
    End If

    If InStr(1, Result, ">") Then
        GreaterSignFlag = True
        Result = Replace(Result, ">", "")
    End If
    'Zyam 04-10-24

20  InterpCoag = ""
    
30  If Val(Result) = 0 Then Exit Function

40  sql = "SELECT * FROM coagtestdefinitions WHERE code = '" & TestCode & "' " & _
          "and agefromdays <= '" & DaysOld & "' and agetodays >= '" & DaysOld & "' " '& _
'          "and hospital = '" & HospName(0) & "'"
50  Set tb = New Recordset
60  RecOpenServer n, tb, sql

70  If tb.EOF Then
80      Exit Function
90  End If

100 Select Case UCase(Left(Sex, 1))
    Case "M":
110     Low = tb!MaleLow
120     High = tb!MaleHigh
130     PlaHigh = tb!PlausibleHigh
140     PlaLow = tb!PlausibleLow
    Case "F":
150     Low = tb!FemaleLow
160     High = tb!FemaleHigh
170     PlaHigh = tb!PlausibleHigh
180     PlaLow = tb!PlausibleLow
    Case Else:
190     Low = tb!FemaleLow
200     High = tb!MaleHigh
210     PlaHigh = tb!PlausibleHigh
220     PlaLow = tb!PlausibleLow
230 End Select

240 If Val(Result) > Val(PlaHigh) Then
250     InterpCoag = "X"
260 ElseIf Val(Result) < Val(PlaLow) Then
270     InterpCoag = "X"
        'Zyam Flag Issue
280 ElseIf Val(Result) > Val(High) And High <> 0 Then
290     InterpCoag = "H"
        'Zyam Flag Issue
300 ElseIf Val(Result) < Val(Low) Then
310     InterpCoag = "L"
320 End If
    'Zyam 04-10-24
    If LessSignFlag Then
        
        If Val(Result) <= Val(Low) Then
            InterpCoag = "L"
        End If
    End If

    If GreaterSignFlag Then
        
        If Val(Result) >= Val(High) Then
            InterpCoag = "H"
        End If
    End If
    'Zyam 04-10-24



330 Exit Function

InterpCoag_Error:

    Dim strES As String
    Dim intEL As Integer

340 intEL = Erl
350 strES = Err.Description
360 LogError "Other", "InterpCoag", intEL, strES, sql

End Function

Public Function InterpH(ByVal strValue As String, _
                        ByVal Analyte As String, _
                        ByVal Sex As String, _
                        ByVal DoB As String, _
                        ByVal SampleDate As String) _
                        As String

      Dim sql As String
      Dim tb As Recordset
      Dim DaysOld As Long
      Dim SexSQL As String
      Dim X As Long
      Dim Value As Single

10    On Error GoTo InterpH_Error

20    If Trim(Sex) = "" Or (Not IsDate(DoB)) Or (strValue) = "" Then       'QMS Ref #818581
30        InterpH = " "
40        Exit Function
50    End If

60    Value = Val(strValue)

70    If InStr(strValue, ">") Then
80        X = InStr(strValue, ">")
90        Value = Val(Mid(strValue, X + 1))
100   End If

110   Select Case Left(UCase(Sex), 1)
          Case "M"
120           SexSQL = "MaleLow as Low, MaleHigh as High "
130       Case "F"
140           SexSQL = "FemaleLow as Low, FemaleHigh as High "
150       Case Else
160           SexSQL = "FemaleLow as Low, MaleHigh as High "
170   End Select

180   If IsDate(DoB) Then

190       DaysOld = Abs(DateDiff("d", SampleDate, DoB))

200       sql = "SELECT top 1 PlausibleLow, PlausibleHigh, " & _
                SexSQL & _
                "FROM HaemTestDefinitions WHERE " & _
                "AnalyteName = '" & Analyte & "' and AgeFromDays <= '" & DaysOld & "' " & _
                "and AgeToDays >= '" & DaysOld & "' " & _
                "order by AgeFromDays desc, AgeToDays asc"
210   Else
220       sql = "SELECT top 1 PlausibleLow, PlausibleHigh, " & _
                SexSQL & _
                "FROM HaemTestDefinitions WHERE " & _
                "AnalyteName = '" & Analyte & "' and " & _
                "AgeFromDays <= '9125' " & _
                "and AgeToDays >= '9125'"
230   End If

240   Set tb = New Recordset
250   RecOpenClient 0, tb, sql
260   If Not tb.EOF Then

270       If Value > tb!PlausibleHigh Then
280           InterpH = "X"
290           Exit Function
300       ElseIf Value < tb!PlausibleLow Then
310           InterpH = "X"
320           Exit Function
330       End If

340       If Value > tb!High And tb!High <> 0 Then
350           InterpH = "H"
360       ElseIf Value < tb!Low Then
370           InterpH = "L"
380       Else
390           InterpH = " "
400       End If
410   Else
420       InterpH = " "
430   End If

440   Exit Function

InterpH_Error:

      Dim strES As String
      Dim intEL As Integer

450   intEL = Erl
460   strES = Err.Description
470   LogError "Other", "InterpH", intEL, strES, sql

End Function
Public Sub RecOpenClient(ByVal n As Integer, ByVal RecSet As Recordset, ByVal sql As String)

10    With RecSet
20        .CursorLocation = adUseClient
30        .CursorType = adOpenDynamic
40        .LockType = adLockOptimistic
50        .ActiveConnection = Cnxn(n)
60        .Source = sql
70        .Open
80    End With

End Sub



Public Sub RecOpenServer(ByVal n As Integer, ByVal RecSet As Recordset, ByVal sql As String)

10    With RecSet
20        .CursorLocation = adUseServer
30        .CursorType = adOpenDynamic
40        .LockType = adLockOptimistic
50        .ActiveConnection = Cnxn(n)
60        .Source = sql
70        .Open
80    End With

End Sub



Public Function PrintRecord() As Boolean

      Dim i          As Integer

10    On Error GoTo PrintRecord_Error

20    PrintRecord = False

30    Select Case RP.Department
          Case "B":
40            DisplaySampleID = RP.SampleID
50            If frmMain.optSecondPage Then
60                If GetOptionSetting("BiochemistryPrintA4Enabled", "0") = "1" Then
70                    PrintResultBioWin (GetOptionSetting("BiochemistryPrintA4", "0"))
80                Else
90                    PrintResultBioWin
100               End If

110           Else
120               PrintResultBioSideBySide
130           End If
140       Case "C", "D":
150           DisplaySampleID = RP.SampleID
160           If GetOptionSetting("CoagulationPrintA4Enabled", "0") = "1" Then
170               PrintResultCoagA4 (GetOptionSetting("CoagulationPrintA4", "0"))
180           Else
190               PrintResultCoag
200           End If

210       Case "E":
220           DisplaySampleID = RP.SampleID
230           If GetOptionSetting("EndocrinologyPrintA4Enabled", "0") = "1" Then
240               PrintResultEndWin (GetOptionSetting("EndocrinologyPrintA4", "False"))
250           Else
260               PrintResultEndWin1
270           End If

280       Case "F":
290           DisplaySampleID = RP.SampleID - SysOptMicroOffset(0)
300           PrintFaeces
310       Case "G":
320           DisplaySampleID = RP.SampleID
330           PrintGTT
340       Case "H", "K":
350           DisplaySampleID = RP.SampleID
360           If GetOptionSetting("HaematologyPrintA4Enabled", "0") = "1" Then
370               PrintResultHaem GetOptionSetting("HaematologyPrintA4", "0")
380           Else
390               If PrintResultHaemAdvia() = False Then PrintRecord = False
400           End If
410       Case "I", "W":
420           DisplaySampleID = RP.SampleID
430           PrintResultImmWin
440       Case "J":
450           DisplaySampleID = RP.SampleID
460           PrintResultImmWin
470       Case "M":
480           DisplaySampleID = RP.SampleID
490           PrintComposit
500       Case "N":
510           DisplaySampleID = RP.SampleID - SysOptMicroOffset(0)
520           If RP.FaxNumber <> "" Then
530               If GetOptionSetting("MicrobiologyPrintA4Enabled", "0") = "1" Then
540                   DoPrint GetOptionSetting("MicrobiologyPrintA4", "0")
550               Else
560                   DoPrint
570               End If

580           Else
590               For i = 1 To RP.NoOfCopies
600                   If GetOptionSetting("MicrobiologyPrintA4Enabled", "0") = "1" Then
610                       DoPrint GetOptionSetting("MicrobiologyPrintA4", "0")
620                   Else
630                       DoPrint
640                   End If
650               Next i
660           End If
              'PrintResultMicroRTF
670       Case "P":
680           DisplaySampleID = RP.SampleID
690           PrintHistology RP.NoOfCopies
700       Case "Q":
710           DisplaySampleID = RP.SampleID
720           PrintResultBloodGas
730       Case "R", "T":
740           DisplaySampleID = RP.SampleID
750           PrintCreatinine
760       Case "S":
770           DisplaySampleID = RP.SampleID
780           PrintGlucoseSeries
790       Case "U":
800           DisplaySampleID = RP.SampleID - SysOptMicroOffset(0)
810           DoPrint
              'PrintUrine
820       Case "V":
830           DisplaySampleID = RP.SampleID
840           PrintExternalMicro
850       Case "X":
860           DisplaySampleID = RP.SampleID
870           PrintExternal
880       Case "Y":
890           DisplaySampleID = RP.SampleID
900           PrintCytology RP.NoOfCopies
910       Case "Z":
920           DisplaySampleID = RP.SampleID - SysOptSemenOffset(0)
930           DoPrint    ' PrintSAReport
940   End Select

950   PrintRecord = True

960   Exit Function

PrintRecord_Error:

      Dim strES      As String
      Dim intEL      As Integer

970   intEL = Erl
980   strES = Err.Description
990   LogError "Other", "PrintRecord", intEL, strES

End Function

Public Function PrinterName(ByVal strMappedTo As String) As String

      Dim tb As Recordset
      Dim sql As String
      Dim RetVal As String

10    On Error GoTo PrinterName_Error

20    RetVal = ""

30    sql = "SELECT PrinterName FROM Printers WHERE " & _
            "MappedTo = '" & strMappedTo & "'"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    If Not tb.EOF Then
70        RetVal = Trim$(UCase$(tb!PrinterName & ""))
80    End If

90    PrinterName = RetVal

100   Exit Function

PrinterName_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "Module1", "PrinterName", intEL, strES, sql

End Function


Public Function GetPrinterFromWard(ByVal Ward As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GetPrinterFromWard_Error

20    sql = "SELECT PrinterAddress FROM Wards WHERE " & _
            "[Text] = '" & AddTicks(Ward) & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If Not tb.EOF Then
60        GetPrinterFromWard = Trim$(tb!PrinterAddress & "")
70    Else
80        GetPrinterFromWard = ""
90    End If

100   Exit Function

GetPrinterFromWard_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "Other", "GetPrinterFromWard", intEL, strES

End Function

Public Function Row_Count(ByVal TestCount As Integer) As Integer

      Dim n As Integer

10    On Error GoTo Row_Count_Error

20    Select Case TestCount
          Case 1: n = 9
30        Case 2: n = 8
40        Case 3: n = 8
50        Case 4: n = 7
60        Case 5: n = 7
70        Case 6: n = 6
80        Case 7: n = 6
90        Case 8: n = 5
100       Case 9: n = 5
110       Case 10: n = 4
120       Case 11: n = 4
130       Case 12: n = 3
140       Case 13: n = 3
150       Case 14: n = 2
160       Case 15: n = 2
170       Case Else: n = 1
180   End Select

190   Row_Count = n

200   Exit Function

Row_Count_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "Other", "Row_Count", intEL, strES

End Function


Public Sub SendFax(ByVal FaxNumber As String, ByVal SampleID As String, ByVal FaxFile As String)

      Dim lngSend As Long
      Dim strComputer As String
      Dim oFaxServer As FAXCOMLib.FaxServer
      Dim oFaxDoc As FAXCOMLib.FaxDoc

10    On Error GoTo SendFax_Error

20    strComputer = GetOptionSetting("FAXServer", "")
30    If strComputer <> "" Then
      '20    strComputer = SysOptFaxServer(0)
40      Set oFaxServer = New FAXCOMLib.FaxServer
50      With oFaxServer
60          .Connect strComputer
70          .ServerCoverpage = 0
80          .Retries = 3
90          .RetryDelay = 5
100     End With
110     Set oFaxDoc = oFaxServer.CreateDocument(SysOptFax(0) & SampleID & "Send.doc")
120     With oFaxDoc
130         .FileName = FaxFile
140         .FaxNumber = FaxNumber
150         .DisplayName = HospName(0) & " - Fax Server"
160         lngSend = .Send()
170     End With
180     Set oFaxDoc = Nothing
190     oFaxServer.Disconnect
200     Set oFaxServer = Nothing
210   End If

220   Exit Sub

SendFax_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "Other", "SendFax", intEL, strES

End Sub



Public Sub LogError(ByVal ModuleName As String, _
                    ByVal ProcedureName As String, _
                    ByVal ErrorLineNumber As Integer, _
                    ByVal ErrorDescription As String, _
                    Optional ByVal SQLStatement As String, _
                    Optional ByVal EventDesc As String)

      Dim sql As String
      Dim MyMachineName As String
      Dim Vers As String
      Dim UID As String

10    On Error Resume Next

20    UID = AddTicks(UserName)

30    SQLStatement = AddTicks(SQLStatement)

40    ErrorDescription = Replace(ErrorDescription, "[Microsoft][ODBC SQL Server Driver][SQL Server]", "[MSSQL]")
50    ErrorDescription = Replace(ErrorDescription, "[Microsoft][ODBC SQL Server Driver]", "[SQL]")
60    ErrorDescription = AddTicks(ErrorDescription)

70    Vers = App.Major & "-" & App.Minor & "-" & App.Revision

80    MyMachineName = vbGetComputerName()

90    sql = "IF NOT EXISTS " & _
      "    (SELECT * FROM ErrorLog WHERE " & _
      "     ModuleName = '" & ModuleName & "' " & _
      "     AND ProcedureName = '" & ProcedureName & "' " & _
      "     AND ErrorLineNumber = '" & ErrorLineNumber & "' " & _
      "     AND AppName = '" & App.EXEName & "' " & _
      "     AND AppVersion = '" & Vers & "' ) " & _
      "  INSERT INTO ErrorLog (" & _
      "    ModuleName, ProcedureName, ErrorLineNumber, SQLStatement, " & _
      "    ErrorDescription, UserName, MachineName, Eventdesc, AppName, AppVersion, EventCounter, eMailed) " & _
      "  VALUES  ('" & ModuleName & "', " & _
      "           '" & ProcedureName & "', " & _
      "           '" & ErrorLineNumber & "', " & _
      "           '" & SQLStatement & "', " & _
      "           '" & ErrorDescription & "', " & _
      "           '" & UID & "', " & _
      "           '" & MyMachineName & "', " & _
      "           '" & AddTicks(EventDesc) & "', " & _
      "           '" & App.EXEName & "', " & _
      "           '" & Vers & "', " & _
      "           '1', '0') " & _
      "ELSE "
100   sql = sql & "  UPDATE ErrorLog " & _
      "  SET SQLStatement = '" & SQLStatement & "', " & _
      "  ErrorDescription = '" & ErrorDescription & "', " & _
      "  MachineName = '" & MyMachineName & "', " & _
      "  DateTime = getdate(), " & _
      "  UserName = '" & UID & "', " & _
      "  EventCounter = COALESCE(EventCounter, 0) + 1 " & _
      "  WHERE ModuleName = '" & ModuleName & "' " & _
      "  AND ProcedureName = '" & ProcedureName & "' " & _
      "  AND ErrorLineNumber = '" & ErrorLineNumber & "' " & _
      "  AND AppName = '" & App.EXEName & "' " & _
      "  AND AppVersion = '" & Vers & "'"

110   Cnxn(0).Execute sql

End Sub
Public Sub LogEvent(ByVal Description As String, _
                    ByVal OptionalParameter As String)

      Dim sql As String

10    On Error GoTo LogEvent_Error

20    Description = AddTicks(Description)
30    OptionalParameter = AddTicks(OptionalParameter)

      '40    sql = "IF NOT EXISTS " & _
      '      "    (SELECT * FROM PrintHandlerLog WHERE " & _
      '      "     Description = '" & Description & "' ) " & _
      '      "  INSERT INTO PrintHandlerLog (Description, OptionalParameter) VALUES " & _
      '      "  ('" & Description & "', " & _
      '      "   '" & OptionalParameter & "') " & _
      '      "ELSE " & _
      '      "  UPDATE PrintHandlerLog " & _
      '      "  SET Description = '" & Description & "', " & _
      '      "  OptionalParameter = '" & OptionalParameter & "', " & _
      '      "  DateTimeOfRecord = getdate() " & _
      '      "  WHERE Description = '" & Description & "'"
40    sql = "INSERT INTO PrintHandlerLog (Description, OptionalParameter) VALUES " & _
            "('" & Description & "', " & _
            " '" & OptionalParameter & "') "
50    Cnxn(0).Execute sql

60    sql = "DELETE FROM PrintHandlerLog " & _
            "WHERE DATEDIFF(minute, DateTimeOfRecord, getdate()) > 30"
70    Cnxn(0).Execute sql

80    Exit Sub

LogEvent_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "Other", "LogEvent", intEL, strES, sql

End Sub

Public Function vbGetComputerName() As String

      'Gets the name of the machine
      Const MAXSIZE As Integer = 256
      Dim sTmp As String * MAXSIZE
      Dim lLen As Long

10    On Error GoTo vbGetComputerName_Error

20    lLen = MAXSIZE - 1
30    If (GetComputerName(sTmp, lLen)) Then
40        vbGetComputerName = Left$(sTmp, lLen)
50    Else
60        vbGetComputerName = ""
70    End If

80    Exit Function

vbGetComputerName_Error:

      Dim strES As String
      Dim intEL As Integer

90    intEL = Erl
100   strES = Err.Description
110   LogError "Other", "vbGetComputerName", intEL, strES

End Function


Public Function getPrintHandlerConfiguration() As String

      Dim sql As String
      Dim tb As Recordset
      Dim RetVal As String

      'Lab print handler pc name
10    On Error GoTo getPrintHandlerConfiguration_Error

20    sql = "Select contents from Options where contents = '" & vbGetComputerName() & "' " & _
      "and Description = 'LAB_PRINT_HANDLER_PC_NAME' "
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql

50    If tb.EOF Then
60        RetVal = "0"
70    Else
80        RetVal = "1"
90    End If



      'Ward print handler pc name
100   sql = "Select contents from Options where contents = '" & vbGetComputerName() & "' " & _
      "and Description = 'WARD_PRINT_HANDLER_PC_NAME' "
110   Set tb = New Recordset
120   RecOpenClient 0, tb, sql

130   If tb.EOF Then
140       RetVal = RetVal & "0"
150   Else
160       RetVal = RetVal & "1"
170   End If


      'Fax print handler pc name
180   sql = "Select contents from Options where contents = '" & vbGetComputerName() & "' " & _
      "and Description = 'FAX_PRINT_HANDLER_PC_NAME' "
190   Set tb = New Recordset
200   RecOpenClient 0, tb, sql

210   If tb.EOF Then
220       RetVal = RetVal & "0"
230   Else
240       RetVal = RetVal & "1"
250   End If


      'MICRO print handler pc name
260   On Error GoTo getPrintHandlerConfiguration_Error

270   sql = "Select contents from Options where contents = '" & vbGetComputerName() & "' " & _
            "and Description = 'MICRO_PRINT_HANDLER_PC_NAME' "
280   Set tb = New Recordset
290   RecOpenClient 0, tb, sql

300   If tb.EOF Then
310       RetVal = RetVal & "0"
320   Else
330       RetVal = RetVal & "1"
340   End If

350   LogEvent "getPrintHandlerConfiguration", RetVal

360   getPrintHandlerConfiguration = RetVal

370   Exit Function

getPrintHandlerConfiguration_Error:

      Dim strES As String
      Dim intEL As Integer

380   intEL = Erl
390   strES = Err.Description
400   LogError "Other", "getPrintHandlerConfiguration", intEL, strES, sql

End Function


Public Function FaxPrintHandlerOnly() As Boolean

10    If getPrintHandlerConfiguration = "001" Then
20        FaxPrintHandlerOnly = True
30    Else
40        FaxPrintHandlerOnly = False
50    End If

End Function

Public Sub LogTestAsPrinted(Disp As String, SampleID As String, Code As String)

      Dim sql As String

10    On Error GoTo LogTestAsPrinted_Error

20    sql = "Update " & Disp & "Results " & _
            "Set Printed = 1 Where " & _
            "SampleID = '" & SampleID & "' " & _
            "And Code = '" & Code & "'"
30    Cnxn(0).Execute sql

40    Exit Sub

LogTestAsPrinted_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "modImmunology", "LogTestAsPrinted", intEL, strES, sql

End Sub

Public Function FormatString(strDestString As String, _
                             intNumChars As Integer, _
                             Optional strSeperator As String = "", _
                             Optional intAlign As PrintAlignContants = AlignLeft) As String

      '**************intAlign = 0 --> Left Align
      '**************intAlign = 1 --> Center Align
      '**************intAlign = 2 --> Right Align
      Dim intPadding As Integer

10    On Error GoTo FormatString_Error

20    intPadding = 0

30    If Len(strDestString) > intNumChars Then
40        FormatString = Mid(strDestString, 1, intNumChars) & strSeperator
50    ElseIf Len(strDestString) < intNumChars Then
          Dim i As Integer
          Dim intStringLength As String
60        intStringLength = Len(strDestString)
70        intPadding = intNumChars - intStringLength

80        If intAlign = PrintAlignContants.AlignLeft Then
90            strDestString = strDestString & String(intPadding, " ")  '& " "
100       ElseIf intAlign = PrintAlignContants.AlignCenter Then
110           If (intPadding Mod 2) = 0 Then
120               strDestString = String(intPadding / 2, " ") & strDestString & String(intPadding / 2, " ")
130           Else
140               strDestString = String((intPadding - 1) / 2, " ") & strDestString & String((intPadding - 1) / 2 + 1, " ")
150           End If
160       ElseIf intAlign = PrintAlignContants.AlignRight Then
170           strDestString = String(intPadding, " ") & strDestString
180       End If

190       strDestString = strDestString & strSeperator
200       FormatString = strDestString
210   Else
220       strDestString = strDestString & strSeperator
230       FormatString = strDestString
240   End If

250   Exit Function

FormatString_Error:

      Dim strES As String
      Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "Other", "FormatString", intEL, strES

End Function

Public Function PrintTextRTB(rtb As RichTextBox, ByVal Text As String, _
                             Optional FontSize As Integer = 9, Optional FontBold As Boolean = False, _
                             Optional FontItalic As Boolean = False, Optional FontUnderline As Boolean = False, _
                             Optional FontColor As ColorConstants = vbBlack, _
                             Optional SuperScript As Boolean = False)

      '---------------------------------------------------------------------------------------
      ' Procedure : PrintText
      ' DateTime  : 05/06/2008 11:40
      ' Author    : Babar Shahzad
      ' Note      : Printer object needs to be set first before calling this function.
      '             Portrait mode (width X height) = 11800 X 16500
      '---------------------------------------------------------------------------------------
      Dim ChrCnt As Integer
      Dim UnitPart As String

10    On Error GoTo PrintTextRTB_Error

20    With rtb
30    .SelFontSize = FontSize
40            .SelBold = FontBold
50            .SelItalic = FontItalic
60            .SelUnderline = FontUnderline
70            .SelColor = FontColor
          
80        If SuperScript Then
90            If InStr(1, Text, "^") > 0 And InStr(1, Text, "/") > 0 And (InStr(1, Text, "/") - InStr(1, Text, "^")) > 0 Then
100               ChrCnt = 0
110               UnitPart = Left$(Text, InStr(1, Text, "^") - 1)
120               ChrCnt = ChrCnt + Len(UnitPart)
130               .SelText = UnitPart

140               UnitPart = FormatString(Mid$(Text, InStr(1, Text, "^") + 1, InStr(1, Text, "/") - InStr(1, Text, "^") - 1), 2, , AlignLeft)
150               ChrCnt = ChrCnt + Len(UnitPart)
160               .SelCharOffset = 40
170               .SelFontSize = FontSize - 3
180               .SelText = UnitPart
                  
190               .SelCharOffset = 0
200               .SelFontSize = FontSize
210               UnitPart = Mid$(Text, InStr(1, Text, "/"), Len(Text))
220               ChrCnt = ChrCnt + Len(UnitPart)
230               .SelText = UnitPart
240           Else
250               .SelText = Text
260           End If
270       Else
280           .SelCharOffset = 0
              
290           .SelText = Text
300       End If
          
310   End With

320   Exit Function

PrintTextRTB_Error:

      Dim strES      As String
      Dim intEL      As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "Other", "PrintTextRTB", intEL, strES

End Function



'---------------------------------------------------------------------------------------
' Procedure : GetGPAddress (CHanged)
' DateTime  : 25/02/2011 12:01
' Author    : Babar Shahzad
' Purpose   : (default) 0 = pick both addresses
'                       1 = address(0)
'                       2 = address(1)
'---------------------------------------------------------------------------------------
'
Public Function GetGPAddress(GPName As String, Optional AddressType As Integer = 0) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GetGPAddress_Error

20    Select Case AddressType
      Case 0: sql = "Select Addr0 + ' ' + Addr1 As Address From Gps Where Text = '%gpname' AND InUse = '1'"
30    Case 1: sql = "Select Addr0 As Address From Gps Where Text = '%gpname' AND InUse = '1'"
40    Case 2: sql = "Select Addr1 As Address From Gps Where Text = '%gpname' AND InUse = '1'"

50    End Select
60    sql = Replace(sql, "%gpname", GPName)
70    Set tb = New Recordset
80    RecOpenClient 0, tb, sql

90    If tb.EOF Then
100       GetGPAddress = ""
110   Else
120       GetGPAddress = tb!Address & ""
130   End If


140   Exit Function

GetGPAddress_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "Other", "GetGPAddress", intEL, strES, sql


End Function

Public Function GetSampleType(Department As String, SampleID As String) As String

          Dim tb As Recordset
          Dim sql As String
          Dim Disp As String
          Dim ST As String
          'HospName(0) = "PORTLAOISE"
10        On Error GoTo GetSampleType_Error

20        Select Case Department
              Case "B", "M", "G", "S": Disp = "Bio"
30            Case "C", "D": Disp = "Coag"
40            Case "E": Disp = "End"
50            Case "H", "K": Disp = "Haem"
60            Case "I", "J", "W": Disp = "Imm"
70            Case "Q": Disp = "Bga"
80        End Select


90        Select Case Disp
              Case "Bio", "End", "Imm", "Bga":
100               sql = "Select Distinct SampleType From %dispResults Where SampleID = '%sampleid'"
110               sql = Replace(sql, "%disp", Disp)
120               sql = Replace(sql, "%sampleid", SampleID)
130               Set tb = New Recordset
140               RecOpenClient 0, tb, sql
150               If tb.EOF Then
160                   GetSampleType = ""
170               Else
180                   ST = ListText("ST", tb!SampleType & "")
190                   If UCase(ST) = "SERUM/PLASMA" Then ST = "Serum"
200                   If InStr(1, ST, "Plasma") > 0 Then ST = ST & " (Pl)"
210                   GetSampleType = ST
220                   tb.MoveNext
230                   While Not tb.EOF
240                       ST = ListText("ST", tb!SampleType & "")
250                       If UCase(ST) = "SERUM/PLASMA" Then ST = "Serum"
260                       If InStr(1, ST, "Plasma") > 0 Then ST = ST & " (Pl)"
270                       GetSampleType = GetSampleType & ", " & ST
280                       tb.MoveNext
290                   Wend
300               End If
310           Case "Coag":
                  'Trevor 18/11/15
320               If UCase(HospName(0)) = "PORTLAOISE" Or UCase(HospName(0)) = "MULLINGAR" Then
330                   GetSampleType = "Citrated Plasma"
340               Else
350                   GetSampleType = "Plasma"
360               End If
                  
370           Case "Haem":
                      'Comment
                      'HospName(0) = "PORTLAOISE"
                      'Comment
380               If UCase(HospName(0)) = "PORTLAOISE" Then
390                   If EsrExists(SampleID) Then
400                       GetSampleType = "EDTA Blood"
410                   Else
420                       GetSampleType = "EDTA Blood"
430                   End If
                      'Trevor 18/11/15
440               ElseIf UCase(HospName(0)) = "MULLINGAR" Then
450                   GetSampleType = "EDTA  Whole Blood"
                  
460               Else
470                   GetSampleType = "Blood"
480               End If
490       End Select
500       If Department = "M" Then
510           GetSampleType = GetSampleType & "Blood"
520       End If

530       Exit Function

GetSampleType_Error:

          Dim strES As String
          Dim intEL As Integer

540       intEL = Erl
550       strES = Err.Description
560       LogError "Other", "GetSampleType", intEL, strES, sql

End Function

Public Function GetAuthorisedBy(UsernameOrCode As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GetAuthorisedBy_Error

20    sql = "SELECT %usernameorcode AS AuthorisedBy FROM Users Where Name = '%criteria' OR Code = '%criteria'"

30    If GetOptionSetting("PrintAuthorisedByCode", 0) = 0 Then
40        sql = Replace(sql, "%usernameorcode", "Name")
50    Else
60        sql = Replace(sql, "%usernameorcode", "Code")
70    End If
80    sql = Replace(sql, "%criteria", UsernameOrCode)

90    Set tb = New Recordset
100   RecOpenClient 0, tb, sql
110   If tb.EOF Then
120       GetAuthorisedBy = ""
130   Else
140       GetAuthorisedBy = tb!AuthorisedBy & ""
150   End If

160   Exit Function

GetAuthorisedBy_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "Other", "GetAuthorisedBy", intEL, strES, sql

End Function

Public Function GetLatestAuthorisedBy(Disp As String, SampleID As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GetLatestAuthorisedBy_Error

20    sql = "SELECT TOP 1 Operator FROM " & Disp & "Results WHERE sampleid= '" & SampleID & "' ORDER BY Runtime DESC"


90    Set tb = New Recordset
100   RecOpenClient 0, tb, sql
110   If tb.EOF Then
120       GetLatestAuthorisedBy = ""
130   Else
140       GetLatestAuthorisedBy = GetAuthorisedBy(tb!Operator & "")
150   End If

160   Exit Function

GetLatestAuthorisedBy_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "Other", "GetLatestAuthorisedBy", intEL, strES, sql

End Function
Public Function GetLatestRunDateTime(ByVal Disp As String, ByVal SampleID As String, ByVal RunDateTime As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GetLatestRunDateTime_Error

20    If GetOptionSetting("GetLatestRunDate", "0") = 0 Then
30        GetLatestRunDateTime = RunDateTime
40        Exit Function
50    End If

60    sql = "SELECT TOP 1 RunTime FROM " & Disp & "Results WHERE sampleid= '" & SampleID & "' ORDER BY Runtime"


70    Set tb = New Recordset
80    RecOpenClient 0, tb, sql
90    If tb.EOF Then
100       GetLatestRunDateTime = ""
110   Else
120       GetLatestRunDateTime = tb!RunTime
130   End If

140   Exit Function

GetLatestRunDateTime_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "Other", "GetLatestRunDateTime", intEL, strES, sql

End Function
'Public Function GeneralValidatedBy(SampleID As String, ByVal Department As String) As String
'
'      Dim tb As Recordset
'      Dim sql As String
'
'10    On Error GoTo GeneralValidatedBy_Error
'
'20    sql = "Select Top 1 ValidatedBy From PrintValidLog Where SampleID = '%sampleid' And Valid = 1 " & _
 '            "Order By ValidatedDateTime Desc"
'30    sql = Replace(sql, "%sampleid", SampleID)
'40    Set tb = New Recordset
'50    RecOpenClient 0, tb, sql
'60    If tb.EOF Then
'70        GeneralValidatedBy = ""
'80    Else
'90        GeneralValidatedBy = tb!ValidatedBy & ""
'100   End If
'
'110   Exit Function
'
'GeneralValidatedBy_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'120   intEL = Erl
'130   strES = Err.Description
'140   LogError "modNewMicro", "GeneralValidatedBy", intEL, strES, sql
'
'End Function
Public Function GetAuthorisedByConsultant(SampleID As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GetAuthorisedByconsultant_Error

20    sql = "SELECT UserName FROM ConsultantList Where sampleid = '" & SampleID & "'"


30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If tb.EOF Then
60        GetAuthorisedByConsultant = ""
70    Else
80        GetAuthorisedByConsultant = tb!UserName & ""
90    End If

100   Exit Function

GetAuthorisedByconsultant_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "Other", "GetAuthorisedByConsultant", intEL, strES, sql

End Function
Public Function GetAuthorisedStatus(SampleID As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GetAuthorisedStatus_Error

20    sql = "SELECT status FROM ConsultantList Where sampleid = '" & SampleID & "'"


30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If tb.EOF Then
60        GetAuthorisedStatus = ""
70    Else
80        GetAuthorisedStatus = tb!status & ""
90    End If

100   Exit Function

GetAuthorisedStatus_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "Other", "GetAuthorisedStatus", intEL, strES, sql

End Function
Public Sub RemovePrintInihibitEntries(SampleID As String)

      Dim sql As String
      Dim Discipline As String

10    On Error GoTo RemovePrintInihibitEntries_Error


20    Select Case RP.Department
          Case "B": Discipline = "Bio"
30        Case "C": Discipline = "Coag"
40        Case "E": Discipline = "End"
50        Case "I": Discipline = "Imm"
60        Case Else: Discipline = ""
70    End Select

80    If Discipline = "" Then Exit Sub

90    sql = "DELETE FROM PrintInhibit WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "AND Discipline = '" & Discipline & "'"
100   Cnxn(0).Execute sql

110   Exit Sub

RemovePrintInihibitEntries_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "Other", "RemovePrintInihibitEntries", intEL, strES, sql

End Sub

Public Function GetLastValidatedBy(BRs As BIEResults) As String

      Dim LastOperator As String
      Dim LastRunDate As Date
      Dim br As BIEResult

10    On Error GoTo GetLastestValidatedBy_Error

20    LastOperator = BRs(1).Operator
30    LastRunDate = BRs(1).RunTime

40    For Each br In BRs
50        If (DateDiff("n", LastRunDate, br.RunTime) > 0) And Trim(br.Operator) <> "" Then
60            LastOperator = br.Operator
70            LastRunDate = br.RunTime
80        End If
90    Next

100   GetLastValidatedBy = LastOperator


110   Exit Function

GetLastestValidatedBy_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "Other", "GetLastestValidatedBy", intEL, strES

End Function

Public Function EsrExists(SampleID As String) As Boolean

10    On Error GoTo EsrExists_Error

      Dim tb As Recordset
      Dim sql As String

20    sql = "SELECT COALESCE(ESR,'') ESR FROM HaemResults WHERE SampleID = '" & SampleID & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    If tb.EOF Then
60        EsrExists = False
70    Else
80        If tb!ESR & "" = "" Then
90            EsrExists = False
100       Else
110           EsrExists = True
120       End If
130   End If


140   Exit Function

EsrExists_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "Other", "EsrExists", intEL, strES, sql

End Function
Public Function ApplyPrintRule(ResultLine As String) As String

      Dim sql As String
      Dim tb As Recordset
      Dim StartPos As Integer
      Dim EndPos As Integer
      Dim RulesToApply As String
      Dim RulesToCancel As String


10    On Error GoTo ApplyPrintRule_Error


20    ApplyPrintRule = ResultLine
30    sql = "SELECT * FROM PrintingRules"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then

70        While Not tb.EOF
80            If InStr(1, ResultLine, tb!TestName & "") > 0 Then
90                StartPos = InStr(1, ResultLine, Trim(tb!Criteria & ""))
100               EndPos = StartPos + Len(Trim(tb!Criteria & ""))
110               If tb!Bold = True Then
120                   RulesToApply = "^Bold+^"
130                   RulesToCancel = "^Bold-^"
140               End If
150               If tb!Italic = True Then
160                   RulesToApply = RulesToApply & "^Italic+^"
170                   RulesToCancel = RulesToCancel & "^Italic-^"
180               End If
190               If tb!Underline = True Then
200                   RulesToApply = RulesToApply & "^Underline+^"
210                   RulesToCancel = RulesToCancel & "^Underline-^"
220               End If
230               ApplyPrintRule = Mid(ResultLine, 1, StartPos - 1) & _
                                   RulesToApply & _
                                   tb!Criteria & "" & _
                                   RulesToCancel & _
                                   Mid(ResultLine, EndPos)
240           End If

250           tb.MoveNext
260       Wend
270   End If
280   Exit Function

ApplyPrintRule_Error:

      Dim strES As String
      Dim intEL As Integer

290   intEL = Erl
300   strES = Err.Description
310   LogError "Other", "ApplyPrintRule", intEL, strES, sql

End Function

Public Function CheckDisablePrinting(ByVal GPName As String, Department As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo CheckDisablePrinting_Error

20    CheckDisablePrinting = False
30    If RP.WardPrint = True Then Exit Function

40    sql = "SELECT * from DisablePrinting WHERE " & _
            "Department = '" & Department & "' " & _
            "AND GPName = '" & AddTicks(GPName) & "'"
50    Set tb = New Recordset
60    RecOpenClient 0, tb, sql
70    If Not tb.EOF Then
80        CheckDisablePrinting = True
90    End If

100   Exit Function

CheckDisablePrinting_Error:
      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "Other", "CheckDisablePrinting", intEL, strES, sql

End Function

Public Function IsItemInList(Item As String, ListType As String) As Boolean
          Dim tb As Recordset
          Dim sql As String
          Dim RetVal As Boolean

10        On Error GoTo IsItemInList_Error

20        sql = "SELECT Code, ListType, Text From Lists " & _
                "WHERE InUse = 1 AND ListType = '" & ListType & " '  AND Text = '" & Item & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql
50        If (tb.EOF And tb.BOF) Then
60            RetVal = False
70        ElseIf (Trim$(tb!Text & "") = "") Then
80            RetVal = False
90        Else
100           RetVal = True
110       End If

120       IsItemInList = RetVal

130       Exit Function

IsItemInList_Error:

          Dim strES As String
          Dim intEL As Integer

140       intEL = Erl
150       strES = Err.Description
160       LogError "BasShared", "IsItemInList", intEL, strES, sql

End Function

Public Sub AddCommentToLP(ByRef udtPrintLine() As ResultLine, ByRef lp() As String, ByRef lpc As Integer, FullComment As String, Optional CommentTitle As String = "")

      Dim TempString As String
      Dim TempInteger As Integer
      Dim n          As Integer
      Dim Comments() As String

10    On Error GoTo AddCommentToLP_Error

20    TempString = Trim(FullComment)
30    TempInteger = Int(Len(TempString) / 80) + IIf((Len(TempString) / 80) > 0, 1, 0)
40    ReDim Comments(1 To TempInteger) As String
50    FillCommentLines TempString, TempInteger, Comments(), 90

60    If CommentTitle <> "" Then
70        lp(lpc) = CommentTitle
80        udtPrintLine(lpc).Analyte = "*HEADING*"
90        lpc = lpc + 1
100   End If

110   For n = 1 To TempInteger
120       If Trim(Comments(n) & "") <> "" Then
130           lp(lpc) = Trim(Comments(n) & "")
140           udtPrintLine(lpc).Analyte = "*COMMENT*"
150           lpc = lpc + 1
160       End If
170   Next

180   Exit Sub
AddCommentToLP_Error:

190   LogError "modHaematology", "AddCommentToLP", Erl, Err.Description




End Sub

Public Sub AddResultToLP(ByRef udtPrintLine() As ResultLine, ByRef lp() As String, ByRef lpc As Integer, _
                ByVal Analyte As String, ByVal Result As String, Optional ByVal Units As String = "", Optional ByVal NormalRange As String = "", _
                Optional ByVal Flag As String = "", Optional ByVal Comment As String = "", _
                Optional ByVal Fasting As String = "", Optional ByVal Reason As String = "", _
                Optional ByVal PrintBold As Boolean = False, Optional ByVal PrintItalic As Boolean = False, Optional ByVal PrintUnderline As Boolean = False)

10    With udtPrintLine(lpc)
20        .Analyte = Analyte
30        .Comment = Comment
40        .Fasting = Fasting
50        .Flag = Flag
60        .NormalRange = NormalRange
70        .Reason = Reason
80        .Result = IIf(Flag = "X", "XXXXX", Result)
90        .Units = Units
100       .PrintBold = PrintBold
110       .PrintItalic = PrintItalic
120       .PrintUnderline = PrintUnderline
          
130       lp(lpc) = .Analyte & " " & .Result & " " & .Units & " " & .Flag & " " & .NormalRange & " " & .Comment
          
140       lpc = lpc + 1
150   End With

End Sub

Public Sub PrintReport(ByRef udtPrintLine() As ResultLine, ByRef lp() As String, ByRef lpc As Integer, ByVal Disc As String, ByVal PrintA4 As Boolean, _
                       ByVal SampleDate As String, ByVal Rundate As String, ByVal AuthorisedBy As String, ByVal PrintTime As String, ByVal resultCount As String, ByVal CmtCount As Long, _
                       Optional ByVal SampleType As String = "", Optional ExternalTestingNote As String = "")

      Dim i          As Integer
      Dim n          As Integer
      Dim f          As Integer
      Dim PageNumber As Integer
      Dim sql        As String
      Dim tb         As Recordset

      Dim Fontz1     As Integer
      Dim Fontz2     As Integer
      Dim Fontz3     As Integer
      Dim Fontz4     As Integer
      Dim FontBold   As Boolean
      Dim FontItalic   As Boolean
      Dim FontUnderline   As Boolean


      Dim TotalLines As Integer
      Dim CommentLines As Integer
      Dim PerPageLines As Integer
      Dim BodyLines  As Integer
      Dim FooterLines As Integer
      Dim LineNoStartComment As Integer
      Dim TotalPages As Integer
      Dim Discipline As String
      Dim FileName   As String

      Dim LengthAnalyte As Integer
      Dim LengthResult As Integer
      Dim LengthUnit As Integer
      Dim LengthFlag As Integer
      Dim LengthNR As Integer
      Dim LengthComment As Integer


10    On Error GoTo PrintReport_Error

20    Select Case UCase(Disc)
          Case "BIO"
30            Discipline = "Biochemistry"
40            FileName = "BIO"
50        Case "END"
60            Discipline = "Endocrinology"
70            FileName = "END1"
80        Case "COAG"
90            Discipline = "Coagulation"
100           FileName = "COAG"
110       Case "HAEM"
120           Discipline = "Haematology"
130           FileName = "HAEM"
140   End Select

150   If PrintA4 Then
160       BodyLines = 55
170       LineNoStartComment = 70
180   Else
          'Zyam Commented this fix 2 pages report for Biochemistry 24-09-24
190       'BodyLines = 18
          BodyLines = Val(GetOptionSetting("PrintOptionsIfMoreThan", 18)) + Val(CmtCount)
          
          'Zyam added this fix 2 pages report in Biochemistry^^^^ 24-09-24
200       LineNoStartComment = 33
210   End If

220   If RP.FaxNumber <> "" Then
230       Fontz1 = 9
240       Fontz2 = 12
250   Else
260       Fontz1 = 10
270       Fontz2 = 14
280   End If
MsgBox CmtCount
MsgBox BodyLines
290   If Val(resultCount) <= BodyLines Then
300       TotalPages = 1
310   Else
320       TotalPages = Int(lpc / BodyLines) + IIf((lpc Mod BodyLines) > 0, 1, 0)
330   End If

340   With frmRichText
350       i = 0
360       PageNumber = 1
370       For n = 0 To lpc
380           If i = 0 Then
390               If RP.FaxNumber <> "" Then
400                   PrintHeadingRTBFax
410               Else
420                   PrintHeadingRTB ("Page " & PageNumber & " of " & TotalPages)
430               End If

440           End If

450           If Trim(lp(n)) <> "" Or Trim(lp(n)) = "" Then
460               i = i + 1
470               PrintTextRTB .rtb, " ", Fontz1, True, , , vbBlack            'this line is important to keep printing line height to font 9 bold
480               If udtHeading.Dept = "Microbiology" Then
490                   LengthAnalyte = 44
500                   LengthResult = 15
510                   LengthUnit = 8
520                   LengthFlag = 0
530                   LengthNR = 8
540                   LengthComment = 0
                      
550               Else
      '                LengthAnalyte = Len(udtPrintLine(n).Analyte)
      '                LengthResult = Len(udtPrintLine(n).Result)
      '                LengthUnit = Len(udtPrintLine(n).Units)
      '                LengthFlag = Len(udtPrintLine(n).Flag)
      '                LengthNR = Len(udtPrintLine(n).NormalRange)
      '                LengthComment = 35
                      
560                   LengthAnalyte = 20
570                   LengthResult = 8
580                   LengthUnit = 13
590                   LengthFlag = 5
600                   LengthNR = 16
610                   LengthComment = 35
                      
620               End If
                  
630               FontBold = IIf(InStr(1, lp(n), "<BOLD>") > 0, True, False)
640               lp(n) = Replace(lp(n), "<BOLD>", "")
650               If InStr(1, udtPrintLine(n).Analyte, "*") > 0 Then

660                   If Trim(udtPrintLine(n).Analyte) = "*HEADING*" Then
670                       PrintTextRTB .rtb, lp(n) & vbCrLf, Fontz1, True, , , vbBlack
680                   ElseIf Trim(udtPrintLine(n).Analyte) = "*COMMENT*" Then
690                       PrintTextRTB .rtb, lp(n) & vbCrLf, 9, FontBold, , , vbBlack
700                   ElseIf Trim(udtPrintLine(n).Analyte) = "*NRCOMMENT*" Then
710                       PrintTextRTB .rtb, lp(n) & vbCrLf, Fontz1, FontBold, , , vbBlack
720                   ElseIf Trim(udtPrintLine(n).Analyte) = "*LINE*" Then
730                       PrintTextRTB .rtb, lp(n) & vbCrLf, 1, FontBold, , , vbBlack
740                   Else
750                       PrintTextRTB .rtb, lp(n) & vbCrLf, Fontz1, FontBold, , , vbBlack
760                   End If

770               Else
780                   FontBold = udtPrintLine(n).PrintBold
790                   FontItalic = udtPrintLine(n).PrintItalic
800                   FontUnderline = udtPrintLine(n).PrintUnderline
810                   If InStr(lp(n), " L ") Or InStr(lp(n), " H ") Or (InStr(lp(n), ">")) Or (InStr(lp(n), "Positive") And SampleType = "T") Then
820                       FontBold = True
830                   Else
                          'FontBold = False
840                   End If

850                   PrintTextRTB .rtb, "   ", Fontz1, True, , , vbBlack
860                   PrintTextRTB .rtb, FormatString(udtPrintLine(n).Analyte, LengthAnalyte, " "), Fontz1, FontBold, FontItalic, FontUnderline, vbBlack
870                   PrintTextRTB .rtb, FormatString(udtPrintLine(n).Result, LengthResult, " "), Fontz1, FontBold, FontItalic, FontUnderline, vbBlack
880                   PrintTextRTB .rtb, FormatString(udtPrintLine(n).Units, LengthUnit, " "), Fontz1, FontBold, FontItalic, FontUnderline, vbBlack
                      'PrintTextRTB .rtb, "R", Fontz1, FontBold, FontItalic, FontUnderline, vbBlack

890                   PrintTextRTB .rtb, FormatString(udtPrintLine(n).Flag, LengthFlag, " "), Fontz1, FontBold, FontItalic, FontUnderline, vbBlack
900                   PrintTextRTB .rtb, FormatString(udtPrintLine(n).NormalRange, LengthNR, " "), Fontz1, FontBold, FontItalic, FontUnderline, vbBlack
910                   PrintTextRTB .rtb, FormatString(Trim(Trim(udtPrintLine(n).Fasting) & " " & Trim(udtPrintLine(n).Reason) & " " & Trim(udtPrintLine(n).Comment)), LengthComment), 7, False, , , vbBlack
920                   PrintTextRTB .rtb, vbCrLf
930               End If
940           End If
950           CrCnt = CrCnt + 1
960           If (i > BodyLines) Or (n = lpc) Then
970               While CrCnt < LineNoStartComment
980                   PrintTextRTB .rtb, vbCrLf, 10, True, , , vbBlack
990                   CrCnt = CrCnt + 1

1000              Wend
1010              If RP.FaxNumber <> "" Then
1020                  If UCase(GetOptionSetting("GetLatestAuthorisedBy", "")) = UCase("True") And Disc <> "Haem" And Disc <> "Coag" Then
1030                      PrintFooterRTB GetLatestAuthorisedBy(Disc, RP.SampleID), SampleDate, GetLatestRunDateTime(Disc, RP.SampleID, Rundate)
1040                  Else
1050                      PrintFooterRTB AuthorisedBy, SampleDate, GetLatestRunDateTime(Disc, RP.SampleID, Rundate)
1060                  End If
1070                  f = FreeFile
1080                  Open SysOptFax(0) & RP.SampleID & FileName & ".doc" For Output As f
1090                  Print #f, .rtb.TextRTF
1100                  Close f
1110                  SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & FileName & ".doc"
1120              Else
1130                  If UCase(GetOptionSetting("GetLatestAuthorisedBy", "")) = UCase("True") And Disc <> "Haem" And Disc <> "Coag" Then
1140                      PrintFooterRTB GetLatestAuthorisedBy(Disc, RP.SampleID), SampleDate, GetLatestRunDateTime(Disc, RP.SampleID, Rundate), ExternalTestingNote
1150                  Else
1160                      PrintFooterRTB AuthorisedBy, SampleDate, GetLatestRunDateTime(Disc, RP.SampleID, Rundate), ExternalTestingNote
1170                  End If
1180                  .rtb.SelStart = 0
                      'Do not print if Doctor is disabled in DisablePrinting
                      '*******************************************************************
1190                  If CheckDisablePrinting(RP.Ward, Discipline) Then

1200                  ElseIf CheckDisablePrinting(RP.GP, Discipline) Then
1210                  Else
1220                      .rtb.SelPrint Printer.hdc
1230                  End If

1240              End If
1250              sql = "SELECT * FROM Reports WHERE 0 = 1"
1260              Set tb = New Recordset
1270              RecOpenServer 0, tb, sql
1280              tb.AddNew
1290              tb!SampleID = RP.SampleID
1300              tb!Name = udtHeading.Name
1310              tb!Dept = UCase(Left(Disc, 1))
1320              tb!Initiator = RP.Initiator
1330              tb!PrintTime = PrintTime
1340              tb!RepNo = "0" & UCase(Left(Disc, 1)) & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
1350              tb!PageNumber = PageNumber - 1
1360              tb!Report = .rtb.TextRTF
1370              tb!Printer = Printer.DeviceName
1380              tb.Update

1390              PageNumber = PageNumber + 1
1400              i = 0
1410          End If
1420      Next
1430  End With

1440  ResetPrinter

1450  Exit Sub
PrintReport_Error:

1460  LogError "Other", "PrintReport", Erl, Err.Description, sql


End Sub

