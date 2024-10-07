Attribute VB_Name = "modCoagulation"
Option Explicit

Public Function CoagCodeFor(ByVal TestName As String) _
                        As String

      Dim tb As New Recordset
      Dim sql As String

10    On Error GoTo CoagCodeFor_Error

20    CoagCodeFor = "???"

30    sql = "SELECT * FROM Coagtestdefinitions WHERE testname = '" & Trim(TestName) & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql

60    If Not tb.EOF Then
70      CoagCodeFor = tb!Code
80    End If

90    Exit Function

CoagCodeFor_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modCoagulation", "CoagCodeFor", intEL, strES, sql

End Function

Public Function CoagNameFor(ByVal Code As String) _
                        As String

      Dim tb As New Recordset
      Dim sql As String

10    On Error GoTo CoagNameFor_Error

20    CoagNameFor = Code

30    sql = "SELECT * FROM Coagtestdefinitions WHERE Code = '" & Trim(Code) & "'"

40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql

60    If Not tb.EOF Then
70        CoagNameFor = Trim(tb!TestName)
80    End If

90    Exit Function

CoagNameFor_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modCoagulation", "CoagNameFor", intEL, strES, sql

End Function

Public Function CoagPrintFormat(ByVal Code As String) _
                            As Integer

      Dim tb As New Recordset
      Dim sql As String

10    On Error GoTo CoagPrintFormat_Error

20    CoagPrintFormat = 1

30    sql = "SELECT * FROM Coagtestdefinitions WHERE Code = '" & Trim(Code) & "'"

40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql

60    If Not tb.EOF Then
70      CoagPrintFormat = tb!DP
80    End If

90    Exit Function

CoagPrintFormat_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modCoagulation", "CoagPrintFormat", intEL, strES, sql

End Function

Public Function CoagUnitsFor(ByVal Code As String) _
                        As String

      Dim tb As New Recordset
      Dim sql As String

10    On Error GoTo CoagUnitsFor_Error

20    CoagUnitsFor = "???"

30    sql = "SELECT * FROM Coagtestdefinitions WHERE Code = '" & Trim(Code) & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql

60    If Not tb.EOF Then
70      CoagUnitsFor = tb!Units & ""
80    End If

90    Exit Function

CoagUnitsFor_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modCoagulation", "CoagUnitsFor", intEL, strES, sql

End Function
Public Function nrCoag(ByVal TestCode As String, _
                       ByVal Sex As String, _
                       ByVal DoB As String, _
                       ByVal SampleDate As String) As String

      Dim l As String * 5
      Dim h As String * 5
      Dim fMat As String
      Dim DaysOld As Long
      '      Dim TestedLow As Long
      '      Dim TestedHigh As Long
      Dim PF As Integer
      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo nrCoag_Error


20    If Not IsDate(DoB) Or Trim(Sex) = "" Then       'QMS Ref #818581
30        nrCoag = "              "
40        Exit Function
50    End If
60    nrCoag = "(     -      )"

70    sql = "SELECT * FROM coagtestdefinitions WHERE code = '" & TestCode & "'"
80    Set tb = New Recordset
90    RecOpenServer 0, tb, sql
100   If tb.EOF Then Exit Function
110   PF = tb!DP
120   Select Case PF
        Case 0: fMat = "0"
130     Case 1: fMat = "0.0"
140     Case 2: fMat = "0.00"
150     Case 3: fMat = "0.000"
160   End Select

170   If IsDate(DoB) Then
180     DaysOld = Abs(DateDiff("d", SampleDate, DoB))

      '150     TestedLow = MaxAgeToDays
      '160     TestedHigh = 0
      '
      '170     sql = "SELECT * FROM coagtestdefinitions WHERE code = '" & TestCode & "'"
      '180     Set tb = New Recordset
      '190     RecOpenServer 0, tb, sql
      '200     Do While Not tb.EOF
      '210       If tb!Code = TestCode Then
      '220         If tb!AgeFromDays <= TestedLow And tb!AgeToDays >= TestedHigh Then
      '230           If DaysOld >= tb!AgeFromDays And DaysOld <= tb!AgeToDays Then
      '240             TestedLow = tb!AgeFromDays
      '250             TestedHigh = tb!AgeFromDays
      '260           End If
      '270         End If
      '280       End If
      '290     tb.MoveNext
      '300     Loop
190   Else
200     DaysOld = 25 * 365
210   End If

      '340   If TestedHigh = 0 Then
220     sql = "SELECT * FROM coagtestdefinitions WHERE code = '" & Trim$(TestCode) & "' " & _
             "and agefromdays <= '" & DaysOld & "' and agetodays >= '" & DaysOld & "'"
230     Set tb = New Recordset
240     RecOpenServer 0, tb, sql
        'Zyam
        If tb!MaleLow = 0 And tb!MaleHigh = 0 And tb!FemaleLow = 0 And tb!FemaleHigh = 0 Then
            nrCoag = "           "
        Else

250     Select Case Sex
          Case "M":
260         If tb!MaleHigh = 999 Then
270           nrCoag = "             "
280           Exit Function
290         End If
300         RSet l = Format(tb!MaleLow, fMat)
310         Mid(nrCoag, 2, 5) = l
320         h = Format(tb!MaleHigh, fMat)
330         Mid(nrCoag, 9, 5) = h
340       Case "F":
350         If tb!FemaleHigh = 999 Then
360           nrCoag = "           "
370           Exit Function
380         End If
390         RSet l = Format(tb!FemaleLow, fMat)
400         Mid(nrCoag, 2, 5) = l
410         LSet h = Format(tb!FemaleHigh, fMat)
420         Mid(nrCoag, 9, 5) = h
430       Case Else:
440         If tb!MaleHigh = 999 Then
450           nrCoag = "             "
460           Exit Function
470         End If
480         RSet l = Format(tb!FemaleLow, fMat)
490         Mid(nrCoag, 2, 5) = l
500         LSet h = Format(tb!MaleHigh, fMat)
510         Mid(nrCoag, 9, 5) = h
520     End Select
660         End If

        'Zyam
      

530   Exit Function

nrCoag_Error:

      Dim strES As String
      Dim intEL As Integer

540   intEL = Erl
550   strES = Err.Description
560   LogError "modCoagulation", "nrCoag", intEL, strES, sql

End Function




Public Sub PrintResultCoagA4(Optional ByVal PrintA4 As Boolean = False)

      Dim tb         As Recordset
      Dim tbH        As Recordset
      Dim n          As Integer
      Dim Sex        As String
      Dim DaysOld    As String
      Dim DoB        As String
      Dim CRs        As New CoagResults
      Dim CR         As CoagResult
      Dim sql        As String
      '      Dim Cx As Comment
      '      Dim Cxs As New Comments
      Dim OB         As Observation
      Dim OBS        As New Observations
10    ReDim Comments(1 To 4) As String
      Dim SampleDate As String
      Dim Rundate    As String
      Dim f          As Integer
      Dim Fontz1     As Integer
      Dim Fontz2     As Integer
      Dim PrintTime  As String
      Dim AuthorisedBy As String
      Dim udtPrintLine() As ResultLine
      Dim TotalLines As Integer
      Dim CommentLines As Integer
      Dim PerPageLines As Integer
      Dim BodyLines  As Integer
      Dim FooterLines As Integer
      Dim LineNoStartComment As Integer
      Dim TotalPages As Integer
      Dim lpc        As Integer
      Dim i          As Integer
      Dim PageNumber As Integer
      Dim FontBold   As Boolean



20    On Error GoTo PrintResultCoagA4_Error

30    If PrintA4 Then
40        TotalLines = 100
50        CommentLines = 10
60        PerPageLines = 77
70        FooterLines = 3
80    Else
90        TotalLines = 100
100       CommentLines = 4
110       PerPageLines = 35
120       FooterLines = 3
130   End If

140   ReDim lp(0 To TotalLines) As String
150   ReDim udtPrintLine(0 To TotalLines) As ResultLine
160   ReDim Comments(1 To CommentLines) As String


      'Clear All
170   For n = 0 To TotalLines
180       udtPrintLine(n).Analyte = ""
190       udtPrintLine(n).Result = ""
200       udtPrintLine(n).Flag = ""
210       udtPrintLine(n).Units = ""
220       udtPrintLine(n).NormalRange = ""
230       udtPrintLine(n).Fasting = ""
240       udtPrintLine(n).Reason = ""
250   Next



260   PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

270   Set CRs = CRs.Load(RP.SampleID, Trim(SysOptExp(0)))

280   sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
290   Set tb = New Recordset
300   RecOpenClient 0, tb, sql
310   If tb.EOF Then Exit Sub

320   If IsDate(tb!SampleDate) Then
330       SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
340   Else
350       SampleDate = ""
360   End If
370   If IsDate(tb!Rundate) Then
380       Rundate = Format(tb!Rundate, "dd/mmm/yyyy hh:mm")
390   Else
400       Rundate = ""
410   End If

420   sql = "SELECT * FROM HaemResults WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
430   Set tbH = New Recordset
440   RecOpenClient 0, tbH, sql

450   If IsDate(tb!DoB) Then
460       DoB = Format(tb!DoB & "", "Short Date")
470   End If

480   Select Case Left(UCase(tb!Sex & ""), 1)
          Case "M": Sex = "M"
490       Case "F": Sex = "F"
500       Case Else: Sex = ""
510   End Select


520   ClearUdtHeading
530   With udtHeading
540       .SampleID = RP.SampleID
550       .Dept = "Coagulation"
560       .Name = tb!PatName & ""
570       .Ward = RP.Ward
580       .DoB = DoB
590       .Chart = tb!Chart & ""
600       .Clinician = RP.Clinician
610       .Address0 = tb!Addr0 & ""
620       .Address1 = tb!Addr1 & ""
630       .GP = RP.GP
640       .Sex = tb!Sex & ""
650       .Hospital = tb!Hospital & ""
660       .SampleDate = tb!SampleDate & ""
670       .RecDate = tb!RecDate & ""
680       .Rundate = tb!Rundate & ""
690       .GpClin = ""
700       .SampleType = ""
710       .DocumentNo = GetOptionSetting("CoagMainDocumentNo", "")
720       .AandE = tb!AandE & ""
730   End With


740   AddResultToLP udtPrintLine, lp, lpc, "Test", "Result", "Unit", "Ref. Range", "Flag", , , , True, , True
750   CrCnt = CrCnt + 1

760   For Each CR In CRs
770       If AuthorisedBy = "" Then AuthorisedBy = GetAuthorisedBy(CR.OperatorCode)

780       Rundate = Format(CR.RunTime, "dd/MMM/yyyy hh:mm")

790       If DoB <> "" And Len(DoB) > 9 Then DaysOld = Abs(DateDiff("d", SampleDate, DoB)) Else DaysOld = 12783
800       If DaysOld = 0 Then DaysOld = 1

810       sql = "SELECT * FROM coagtestdefinitions WHERE code = '" & Trim(CR.Code) & "' " & _
                "and agefromdays <='" & DaysOld & "' and agetodays >= '" & DaysOld & "'"
820       Set tb = New Recordset
830       RecOpenServer 0, tb, sql
840       If Not tb.EOF Then
              'If Trim(CR.OperatorCode) <> "" Then RP.Initiator = CR.OperatorCode
850           If tb!Printable Then

860               udtPrintLine(lpc).Analyte = Left$("     " & IIf(tb!PrintName & "" = "", tb!TestName, tb!PrintName) & Space(19), 19)
870               If InterpCoag(0, Sex, CR.Code, CR.Result, DaysOld) = "X" Then
880                   udtPrintLine(lpc).Result = Left$("*****" & Space(9), 9)
890               Else
900                   Select Case CoagPrintFormat(CR.Code)
                          Case 0: udtPrintLine(lpc).Result = Left$(Format(CR.Result, "#0") & Space(9), 9)
910                       Case 1: udtPrintLine(lpc).Result = Left$(Format(CR.Result, "0.0") & Space(9), 9)
920                       Case 2: udtPrintLine(lpc).Result = Left$(Format(CR.Result, "0.00") & Space(9), 9)
930                       Case 3: udtPrintLine(lpc).Result = Left$(Format(CR.Result, "0.000") & Space(9), 9)
940                       Case Else: udtPrintLine(lpc).Result = Left$(CR.Result & Space(9), 9)
950                   End Select
960               End If
970               If Trim(CR.Units) = "ÆG/ML" Then udtPrintLine(lpc).Units = Left$("ug/ML" & Space(12), 12)
                  
980               If UCase(HospName(0)) = "MULLINGAR" And tb!TestName <> "INR" Then
990                   If Trim(CR.Units) <> "INR" Then udtPrintLine(lpc).Units = Left$(CR.Units & Space(12), 12)
1000                  If Sex <> "" Then
1010                      If Trim(CR.Units) <> "INR" Then udtPrintLine(lpc).NormalRange = Left$(InterpCoag(0, Sex, CR.Code, CR.Result, DaysOld) & Space(4), 4)
1020                      If Trim(CR.Units) <> "INR" Then udtPrintLine(lpc).NormalRange = nrCoag(CR.Code, Sex, DoB, SampleDate)
1030                  End If
1040              Else
1050                  If Trim(CR.Units) <> "INR" Then udtPrintLine(lpc).Units = Left$(CR.Units & Space(12), 12)
1060                  If Sex <> "" Then
1070                      If Trim(CR.Units) <> "INR" Then udtPrintLine(lpc).NormalRange = Left$(InterpCoag(0, Sex, CR.Code, CR.Result, DaysOld) & Space(4), 4)
1080                      If Trim(CR.Units) <> "INR" Then udtPrintLine(lpc).NormalRange = nrCoag(CR.Code, Sex, DoB, SampleDate)
1090                  End If
1100              End If
1110          End If
1120          With udtPrintLine(lpc)
                  
1130              lp(lpc) = .Analyte & " " & .Result & .Units & " " & .Flag & " " & .NormalRange & " " & .Comment
1140          End With
1150          lpc = lpc + 1
1160          CrCnt = CrCnt + 1
1170      End If
1180      LogTestAsPrinted "Coag", CR.SampleID, CR.Code
1190  Next

      'add blank line before comment
1200  AddResultToLP udtPrintLine, lp, lpc, "", ""

      'comments
1210  Set OBS = OBS.Load(RP.SampleID, "Coagulation", "Demographic")
1220  If Not OBS Is Nothing Then
1230      For Each OB In OBS
1240          Select Case UCase$(OB.Discipline)
                  Case "COAGULATION"
1250                  AddCommentToLP udtPrintLine, lp, lpc, OB.Comment, ""
1260                  CrCnt = CrCnt + 1
1270              Case "DEMOGRAPHIC"

1280                  AddCommentToLP udtPrintLine, lp, lpc, OB.Comment, ""
1290                  CrCnt = CrCnt + 1
1300          End Select
1310      Next
1320  End If


1330  If Not IsDate(DoB) Or Trim(Sex) = "" Then
1340      AddCommentToLP udtPrintLine, lp, lpc, "**** No Sex/DoB given. No reference range applied! ****"
1350      CrCnt = CrCnt + 1
1360  End If


1370  PrintReport udtPrintLine, lp, lpc, "Coag", PrintA4, SampleDate, Rundate, AuthorisedBy, PrintTime, "", ""

1380  Exit Sub

PrintResultCoagA4_Error:

      Dim strES      As String
      Dim intEL      As Integer

1390  sql = "Delete FROM printpending WHERE SampleID = '" & RP.SampleID & "' and department = '" & RP.Department & "'"
1400  Cnxn(0).Execute sql

1410  intEL = Erl
1420  strES = Err.Description
1430  LogError "modCoagulation", "PrintResultCoagA4", intEL, strES, sql

End Sub


Public Sub PrintResultCoag()

      Dim tb As Recordset
      Dim tbH As Recordset
      Dim n As Integer
      Dim Sex As String
      Dim DaysOld As String
      Dim DoB As String
      Dim CRs As New CoagResults
      Dim CR As CoagResult
      Dim sql As String
      '      Dim Cx As Comment
      '      Dim Cxs As New Comments
      Dim OB As Observation
      Dim OBS As New Observations
10    ReDim Comments(1 To 4) As String
      Dim SampleDate As String
      Dim Rundate As String
      Dim f As Integer
      Dim Fontz1 As Integer
      Dim Fontz2 As Integer
      Dim PrintTime As String
      Dim AuthorisedBy As String

20    On Error GoTo PrintResultCoag_Error

30    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

40    Set CRs = CRs.Load(RP.SampleID, Trim(SysOptExp(0)))

50    sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
60    Set tb = New Recordset
70    RecOpenClient 0, tb, sql
80    If tb.EOF Then Exit Sub

90    If IsDate(tb!SampleDate) Then
100       SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
110   Else
120       SampleDate = ""
130   End If
140   If IsDate(tb!Rundate) Then
150       Rundate = Format(tb!Rundate, "dd/mmm/yyyy hh:mm")
160   Else
170       Rundate = ""
180   End If

190   sql = "SELECT * FROM HaemResults WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
200   Set tbH = New Recordset
210   RecOpenClient 0, tbH, sql

220   If IsDate(tb!DoB) Then
230       DoB = Format(tb!DoB & "", "Short Date")
240   End If

250   Select Case Left(UCase(tb!Sex & ""), 1)
      Case "M": Sex = "M"
260   Case "F": Sex = "F"
270   Case Else: Sex = ""
280   End Select


290   ClearUdtHeading
300   With udtHeading
310       .SampleID = RP.SampleID
320       .Dept = "Coagulation"
330       .Name = tb!PatName & ""
340       .Ward = RP.Ward
350       .DoB = DoB
360       .Chart = tb!Chart & ""
370       .Clinician = RP.Clinician
380       .Address0 = tb!Addr0 & ""
390       .Address1 = tb!Addr1 & ""
400       .GP = RP.GP
410       .Sex = tb!Sex & ""
420       .Hospital = tb!Hospital & ""
430       .SampleDate = tb!SampleDate & ""
440       .RecDate = tb!RecDate & ""
450       .Rundate = tb!Rundate & ""
460       .GpClin = ""
470       .SampleType = ""
480       .DocumentNo = GetOptionSetting("CoagMainDocumentNo", "")
490       .AandE = tb!AandE & ""
500   End With

510   If RP.FaxNumber <> "" Then
520       PrintHeadingRTBFax
530   Else
540       PrintHeadingRTB ("Page 1 of 1")
550   End If

560   If RP.FaxNumber <> "" Then Fontz1 = 8 Else Fontz1 = 10
570   If RP.FaxNumber <> "" Then Fontz2 = 10 Else Fontz2 = 12

580   With frmRichText.rtb
590       .SelFontName = "Courier New"
600       .SelFontSize = Fontz1

610       .SelText = vbCrLf
620       CrCnt = CrCnt + 2

630       For Each CR In CRs
640           If AuthorisedBy = "" Then AuthorisedBy = GetAuthorisedBy(CR.OperatorCode)
650           .SelFontName = "Courier New"
660           .SelFontSize = Fontz2
670           .SelBold = True
680           Rundate = Format(CR.RunTime, "dd/MMM/yyyy hh:mm")

690           If DoB <> "" And Len(DoB) > 9 Then DaysOld = Abs(DateDiff("d", SampleDate, DoB)) Else DaysOld = 12783
700           If DaysOld = 0 Then DaysOld = 1

710           sql = "SELECT * FROM coagtestdefinitions WHERE code = '" & Trim(CR.Code) & "' " & _
                    "and agefromdays <='" & DaysOld & "' and agetodays >= '" & DaysOld & "'"
720           Set tb = New Recordset
730           RecOpenServer 0, tb, sql
740           If Not tb.EOF Then
                  'If Trim(CR.OperatorCode) <> "" Then RP.Initiator = CR.OperatorCode
750               If tb!Printable Then
760                   .SelText = Left$("     " & IIf(tb!PrintName & "" = "", tb!TestName, tb!PrintName) & Space(19), 19)
770                   If InterpCoag(0, Sex, CR.Code, CR.Result, DaysOld) = "X" Then
780                       .SelText = Left$("*****" & Space(9), 9)
790                   Else
800                       Select Case CoagPrintFormat(CR.Code)
                          Case 0: .SelText = Left$(Format(CR.Result, "#0") & Space(9), 9)
810                       Case 1: .SelText = Left$(Format(CR.Result, "0.0") & Space(9), 9)
820                       Case 2: .SelText = Left$(Format(CR.Result, "0.00") & Space(9), 9)
830                       Case 3: .SelText = Left$(Format(CR.Result, "0.000") & Space(9), 9)
840                       Case Else: .SelText = Left$(CR.Result & Space(9), 9)
850                       End Select
860                   End If
870                   If Trim(CR.Units) = "ÆG/ML" Then .SelText = Left$("ug/ML" & Space(12), 12)

880                   If UCase(HospName(0)) = "MULLINGAR" And tb!TestName <> "INR" Then
890                       If Trim(CR.Units) <> "INR" Then .SelText = Left$(CR.Units & Space(12), 12)
900                       If Sex <> "" Then
910                           If Trim(CR.Units) <> "INR" Then .SelText = Left$(InterpCoag(0, Sex, CR.Code, CR.Result, DaysOld) & Space(4), 4)
920                           If Trim(CR.Units) <> "INR" Then .SelText = nrCoag(CR.Code, Sex, DoB, SampleDate)
930                       End If
940                   Else
950                       If Trim(CR.Units) <> "INR" Then .SelText = Left$(CR.Units & Space(12), 12)
960                       If Sex <> "" Then
970                           If Trim(CR.Units) <> "INR" Then .SelText = Left$(InterpCoag(0, Sex, CR.Code, CR.Result, DaysOld) & Space(4), 4)
980                           If Trim(CR.Units) <> "INR" Then .SelText = nrCoag(CR.Code, Sex, DoB, SampleDate)
990                       End If
1000                  End If
1010              End If
1020              .SelText = vbCrLf
1030              CrCnt = CrCnt + 1
1040          End If
1050          LogTestAsPrinted "Coag", CR.SampleID, CR.Code
1060      Next

1070      CrCnt = CrCnt + 3

1080      .SelFontName = "Courier New"
1090      .SelFontSize = Fontz1

1100      Do While CrCnt < 28
1110          .SelText = vbCrLf
1120          CrCnt = CrCnt + 1
1130      Loop

          '  Set Cx = Cxs.Load(RP.SampleID)
1140      Set OBS = OBS.Load(RP.SampleID, "Coagulation", "Demographic")
1150      If Not OBS Is Nothing Then
1160          For Each OB In OBS
1170              Select Case UCase$(OB.Discipline)
                  Case "COAGULATION"
1180                  FillCommentLines OB.Comment, 4, Comments(), 80
1190                  For n = 1 To 4
1200                      If Trim(Comments(n) & "") <> "" Then
1210                          .SelText = " " & Comments(n)
1220                          .SelText = vbCrLf
1230                          CrCnt = CrCnt + 1
1240                      End If
1250                  Next
1260              Case "DEMOGRAPHIC"
1270                  FillCommentLines OB.Comment, 2, Comments(), 80
1280                  For n = 1 To 2
1290                      If Trim(Comments(n) & "") <> "" Then
1300                          .SelText = " " & Comments(n)
1310                          .SelText = vbCrLf
1320                          CrCnt = CrCnt + 1
1330                      End If
1340                  Next
1350              End Select
1360          Next
1370      End If


1380      If Not IsDate(DoB) Or Trim(Sex) = "" Then
1390          .SelColor = vbBlack
1400          .SelText = "**** No Sex/DoB given. No reference range applied! ****" & vbCrLf
1410          .SelText = vbCrLf
1420          CrCnt = CrCnt + 1
              '    ElseIf Not IsDate(DoB) Then
              '        .SelColor = vbBlack
              '        .SelText = "*** No Dob. Adult Age 25 used for Normal Ranges! ***" & vbCrLf
              '        .SelText = vbCrLf
              '        CrCnt = CrCnt + 1
              '    ElseIf Trim(Sex) = "" Then
              '        .SelColor = vbBlack
              '        .SelText = "No Sex given. No reference range applied" & vbCrLf
              '        .SelText = vbCrLf
              '        CrCnt = CrCnt + 1
1430      End If
1440      .SelColor = vbBlack

1450      If RP.FaxNumber <> "" Then
1460          PrintFooterRTBFax RP.Initiator, SampleDate, Rundate
1470          f = FreeFile
1480          Open SysOptFax(0) & RP.SampleID & "COAG.doc" For Output As f
1490          Print #f, .TextRTF
1500          Close f
1510          SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "COAG.doc"
1520      Else
1530          PrintFooterRTB AuthorisedBy, SampleDate, Rundate
1540          .SelStart = 0
              'Do not print if Doctor is disabled in DisablePrinting
              '*******************************************************************
1550          If CheckDisablePrinting(RP.Ward, "Coagulation") Then

1560          ElseIf CheckDisablePrinting(RP.GP, "Coagulation") Then
1570          Else
1580              .SelPrint Printer.hdc
1590          End If
              '*******************************************************************
              '.SelPrint Printer.hDC
1600      End If

1610      sql = "SELECT * FROM Reports WHERE 0 = 1"
1620      Set tb = New Recordset
1630      RecOpenServer 0, tb, sql
1640      tb.AddNew
1650      tb!SampleID = RP.SampleID
1660      tb!Name = udtHeading.Name
1670      tb!Dept = "C"
1680      tb!Initiator = RP.Initiator
1690      tb!PrintTime = PrintTime
1700      tb!RepNo = "0C" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
1710      tb!PageNumber = 0
1720      tb!Report = .TextRTF
1730      tb!Printer = Printer.DeviceName
1740      tb.Update
1750  End With

1760  ResetPrinter

1770  Exit Sub

PrintResultCoag_Error:

      Dim strES As String
      Dim intEL As Integer

1780  sql = "Delete FROM printpending WHERE SampleID = '" & RP.SampleID & "' and department = '" & RP.Department & "'"
1790  Cnxn(0).Execute sql

1800  intEL = Erl
1810  strES = Err.Description
1820  LogError "modCoagulation", "PrintResultCoag", intEL, strES, sql

End Sub


