Attribute VB_Name = "modBiochemistry"
Option Explicit



Sub LogBioAsPrinted(ByVal SampleID As String, _
                    ByVal TestCode As String)

      Dim sql As String

10    On Error GoTo LogBioAsPrinted_Error

20    sql = "update BioResults " & _
            "set valid = 1, printed = 1 WHERE " & _
            "SampleID = '" & RP.SampleID & "' " & _
            "and code = '" & TestCode & "'"
30    Cnxn(0).Execute sql

40    Exit Sub

LogBioAsPrinted_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "modBiochemistry", "LogBioAsPrinted", intEL, strES, sql

End Sub


Public Sub PrintCreatinine()

      Dim tb As Recordset
      Dim tc As Recordset
      Dim sql As String
      Dim Sex As String
      Dim n As Integer
      'Dim Cx As Comment
      'Dim Cxs As New Comments
      Dim OB As Observation
      Dim OBS As New Observations
10    ReDim Comments(1 To 4) As String
      Dim SampleDate As String
      Dim Rundate As String
      Dim DoB As String
      Dim RunTime As String
      Dim SorU As String
      Dim PrintTime As String
      Dim AuthorisedBy As String

20    On Error GoTo PrintCreatinine_Error

30    AuthorisedBy = GetAuthorisedBy(RP.Initiator)
40    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

50    If RP.Department = "R" Then
60        SorU = "Urine"
70    Else
80        SorU = "Serum"
90    End If

100   sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
110   Set tb = New Recordset
120   RecOpenClient 0, tb, sql

130   If tb.EOF Then
140       Exit Sub
150   End If

160   If IsDate(tb!DoB) Then
170       DoB = Format(tb!DoB, "dd/mmm/yyyy")
180   Else
190       DoB = ""
200   End If

210   sql = "SELECT * FROM Creatinine WHERE " & _
            SorU & "Number = '" & RP.SampleID & "'"
220   Set tc = New Recordset
230   RecOpenServer n, tc, sql

240   ClearUdtHeading
250   With udtHeading
260       .SampleID = RP.SampleID
270       .Dept = "Biochemistry"
280       .Name = tb!PatName & ""
290       .Ward = RP.Ward
300       .DoB = DoB
310       .Chart = tb!Chart & ""
320       .Clinician = RP.Clinician
330       .Address0 = tb!Addr0 & ""
340       .Address1 = tb!Addr1 & ""
350       .GP = RP.GP
360       .Sex = tb!Sex & ""
370       .Hospital = tb!Hospital & ""
380       .SampleDate = tb!SampleDate & ""
390       .RecDate = tb!RecDate & ""
400       .Rundate = tb!Rundate & ""
410       .GpClin = ""
420       .SampleType = "S"
430       .DocumentNo = GetOptionSetting("BioCreatDocumentNo", "")
440       .AandE = tb!AandE & ""
450       .AandE = tb!AandE & ""
460   End With

470   PrintHeadingRTB

480   With frmRichText.rtb
490       .SelFontSize = 10
500       .SelText = vbCrLf
510       .SelText = "Creatinine Clearance Test" & vbCrLf
520       .SelText = vbCrLf
530       .SelText = "     Volume Collected: " & Left(" " & Space(15), 15) & Format(tc!UrineVolume & "", "#####") & " mL" & vbCrLf
540       .SelText = vbCrLf
550       .SelText = "    Plasma Creatinine: " & Left(" " & Space(15), 15) & Format(tc!SerumCreat & "", "#####.0") & " umol/L" & vbCrLf
560       .SelText = vbCrLf
570       .SelText = "   Urinary Creatinine: " & Left(" " & Space(15), 15) & Format(tc!UrineCreat & "", "0.000") & " umol/L" & vbCrLf
580       .SelText = vbCrLf
590       .SelText = "            Clearance: " & Left(" " & Space(15), 15) & Format(tc!CCl & "", "####") & " ml/min" & vbCrLf
600       .SelText = vbCrLf
610       If Trim(tc!UrineProL & "") <> "" Then
620           .SelText = "Protein Concentration: " & Left(" " & Space(15), 15) & Format(tc!UrineProL & "", "0.000") & " mg/dl" & vbCrLf
630           .SelText = vbCrLf
640           .SelText = Left(" " & Space(38), 38) & Format(tc!UrinePro24Hr & "", "#0.000") & " g/24Hr" & vbCrLf
650       End If
660       .SelText = vbCrLf
670       .SelText = "          Report Date: " & Left(" " & Space(15), 15) & Format(tb!Rundate, "dd/mm/yyyy") & vbCrLf


680       CrCnt = CrCnt + 16

690       Set OBS = OBS.Load(RP.SampleID, "Biochemistry", "Demographic")
700       If Not OBS Is Nothing Then
710           For Each OB In OBS
720               Select Case UCase$(OB.Discipline)
                      Case "BIOCHEMISTRY"
730                       FillCommentLines OB.Comment, 4, Comments(), 87
740                       For n = 1 To 4
750                           If Trim(Comments(n) & "") <> "" Then
760                               .SelFontSize = 10
770                               .SelText = Comments(n)
780                               .SelText = vbCrLf
790                               CrCnt = CrCnt + 1
800                           End If
810                       Next
820                   Case "DEMOGRAPHIC"
830                       FillCommentLines OB.Comment, 2, Comments(), 87
840                       For n = 1 To 2
850                           If Trim(Comments(n) & "") <> "" Then
860                               .SelFontSize = 10
870                               .SelText = Comments(n)
880                               .SelText = vbCrLf
890                               CrCnt = CrCnt + 1
900                           End If
910                       Next
920               End Select
930           Next
940       End If

950       If IsDate(tb!SampleDate) Then
960           SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy")
970       Else
980           SampleDate = ""
990       End If
1000      If IsDate(RunTime) Then
1010          Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
1020      Else
1030          If IsDate(tb!Rundate) Then
1040              Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
1050          Else
1060              Rundate = ""
1070          End If
1080      End If

'1090      PrintFooterRTB AuthorisedBy, SampleDate, Rundate
3770    If UCase(GetOptionSetting("GetLatestAuthorisedBy", "")) = UCase("True") Then
3780        PrintFooterRTB GetLatestAuthorisedBy("Bio", RP.SampleID), SampleDate, GetLatestRunDateTime("Bio", RP.SampleID, Rundate)
3790    Else
3800        PrintFooterRTB AuthorisedBy, SampleDate, GetLatestRunDateTime("Bio", RP.SampleID, Rundate)
3810    End If
1100      .SelStart = 0
1110      .SelPrint Printer.hdc

1120      sql = "SELECT * FROM Reports WHERE 0 = 1"
1130      Set tb = New Recordset
1140      RecOpenServer 0, tb, sql
1150      tb.AddNew
1160      tb!SampleID = RP.SampleID
1170      tb!Name = udtHeading.Name
1180      tb!Dept = "R"
1190      tb!Initiator = RP.Initiator
1200      tb!PrintTime = PrintTime
1210      tb!RepNo = "0R" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
1220      tb!PageNumber = 0
1230      tb!Report = .TextRTF
1240      tb!Printer = Printer.DeviceName
1250      tb.Update
1260  End With

1270  Exit Sub

PrintCreatinine_Error:

      Dim strES As String
      Dim intEL As Integer

1280  intEL = Erl
1290  strES = Err.Description
1300  LogError "modBiochemistry", "PrintCreatinine", intEL, strES, sql

End Sub
Public Sub PrintGlucoseSeries()

      Dim tb As Recordset
      Dim tbDems As Recordset
      Dim tbRes As Recordset
      Dim sql As String
      Dim DoB As String
      Dim NameToFind As String
      Dim SampleDate As String
      Dim Rundate As String
      Dim Code As Integer
      Dim PrintTime As String
      Dim AuthorisedBy As String

10    On Error GoTo PrintGlucoseSeries_Error

20    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

30    Code = SysOptBioCodeForGlucose(0)

40    sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
50    Set tb = New Recordset
60    RecOpenClient 0, tb, sql
70    If tb.EOF Then Exit Sub

80    If IsDate(tb!SampleDate) Then
90        SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
100   Else
110       SampleDate = ""
120   End If
130   If IsDate(tb!Rundate) Then
140       Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
150   Else
160       Rundate = ""
170   End If

180   NameToFind = tb!PatName & ""

190   If IsDate(tb!DoB) Then
200       DoB = Format(tb!DoB, "dd/mmm/yyyy")
210   Else
220       DoB = ""
230   End If

240   ClearUdtHeading
250   With udtHeading
260       .SampleID = RP.SampleID
270       .Dept = "Biochemistry"
280       .Name = tb!PatName & ""
290       .Ward = RP.Ward
300       .DoB = DoB
310       .Chart = tb!Chart & ""
320       .Clinician = RP.Clinician
330       .Address0 = tb!Addr0 & ""
340       .Address1 = tb!Addr1 & ""
350       .GP = RP.GP
360       .Sex = tb!Sex & ""
370       .Hospital = tb!Hospital & ""
380       .SampleDate = tb!SampleDate & ""
390       .RecDate = tb!RecDate & ""
400       .Rundate = tb!Rundate & ""
410       .GpClin = ""
420       .SampleType = "S"
430       .DocumentNo = GetOptionSetting("BioGluDocumentNo", "")
440       .AandE = tb!AandE & ""
450   End With

460   sql = "SELECT * FROM demographics WHERE " & _
            "patname = '" & NameToFind & "' " & _
            "and rundate = '" & Format(tb!Rundate, "dd/mmm/yyyy") & "' "
470   If IsDate(DoB) Then
480       sql = sql & "and DoB = '" & Format(DoB, "dd/mmm/yyyy") & "' "
490   End If
500   sql = sql & "order by SampleID"
510   Set tbDems = New Recordset
520   RecOpenClient 0, tbDems, sql

530   If tbDems.EOF Then
540       Exit Sub
550   End If

560   PrintHeadingRTB
570   With frmRichText.rtb
580       .SelText = ""

590       .SelFontSize = 10

600       .SelText = vbCrLf & vbCrLf & vbCrLf

610       .SelColor = vbGreen
620       .SelFontSize = 12
630       .SelText = Left(" " & Space(30), 30) & "Glucose Series" & vbCrLf
640       .SelColor = vbBlack
650       .SelFontSize = 10
660       .SelText = vbCrLf & vbCrLf
670       .SelText = Left(" " & Space(15), 15) & Left("Test Name " & Space(45), 45) & "Serum mmol/L" & vbCrLf & vbCrLf    '  Urine mmol/L" & vbCrLf &  vbcrlf
680       CrCnt = CrCnt + 8

690       Do While Not tbDems.EOF
        
700           sql = "SELECT * FROM BioResults WHERE " & _
                    "SampleID = '" & tbDems!SampleID & "' " & _
                    "and (Code = '" & SysOptBioCodeForFastGlucose(0) & _
                    "' or code = '" & SysOptBioCodeForGlucose1(0) & _
                    "' or code = '" & SysOptBioCodeForGlucose2(0) & _
                    "' or code = '" & SysOptBioCodeForGlucose3(0) & _
                    "' or Code = '" & SysOptBioCodeForFastGlucoseP(0) & _
                    "' or code = '" & SysOptBioCodeForGlucose1P(0) & _
                    "' or code = '" & SysOptBioCodeForGlucose2P(0) & _
                    "' or code = '" & SysOptBioCodeForGlucose3P(0) & "')"
710           Set tbRes = New Recordset
720           RecOpenClient 0, tbRes, sql

730           If Not tbRes.EOF Then
740               If AuthorisedBy = "" Then AuthorisedBy = GetAuthorisedBy(tb!Operator & "")
750               LogBioAsPrinted tbDems!SampleID & "", Code
760               .SelText = Left(" " & Space(15), 15) & Left(LongNameforCode(tbRes!Code) & "" & Space(22), 22)
770               .SelText = Left(" " & Space(27), 27)
780               .SelText = Format(tbRes!Result, "0.0") & vbCrLf
790               .SelText = vbCrLf
800               CrCnt = CrCnt + 2
810           End If
820           tbDems.MoveNext
830       Loop

'840       PrintFooterRTB AuthorisedBy, SampleDate, Rundate
840     If UCase(GetOptionSetting("GetLatestAuthorisedBy", "")) = UCase("True") Then
850         PrintFooterRTB GetLatestAuthorisedBy("Bio", RP.SampleID), SampleDate, GetLatestRunDateTime("Bio", RP.SampleID, Rundate)
860     Else
870         PrintFooterRTB AuthorisedBy, SampleDate, GetLatestRunDateTime("Bio", RP.SampleID, Rundate)
880     End If
890       .SelStart = 0
900       .SelPrint Printer.hdc

910       sql = "SELECT * FROM Reports WHERE 0 = 1"
920       Set tb = New Recordset
930       RecOpenServer 0, tb, sql
940       tb.AddNew
950       tb!SampleID = RP.SampleID
960       tb!Name = udtHeading.Name
970       tb!Dept = "S"
980       tb!Initiator = RP.Initiator
990       tb!PrintTime = PrintTime
1000      tb!RepNo = "0S" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
1010      tb!PageNumber = 0
1020      tb!Report = .TextRTF
1030      tb!Printer = Printer.DeviceName
1040      tb.Update
1050  End With

1060  ResetPrinter

1070  Exit Sub

PrintGlucoseSeries_Error:

      Dim strES As String
      Dim intEL As Integer

1080  intEL = Erl
1090  strES = Err.Description
1100  LogError "modBiochemistry", "PrintGlucoseSeries", intEL, strES, sql

End Sub

Public Sub PrintGTT()

      Dim tb As Recordset
      Dim tbDems As Recordset
      Dim tbRes As Recordset
      Dim sql As String
      Dim DoB As String
      Dim NameToFind As String
      Dim SampleDate As String
      Dim Rundate As String
      Dim Code As Integer
      Dim CodeP As String
      Dim PrintTime As String
      Dim AuthorisedBy As String
      Dim CodeSerum As String
      Dim CodePlasma As String
      Dim i As Integer
      Dim BRs As New BIEResults
      Dim BRres As BIEResults
      Dim br As BIEResult

10    On Error GoTo PrintGTT_Error

20    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

30    sql = "SELECT * FROM Demographics WHERE " & _
              "SampleID = '" & RP.SampleID & "'"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    If tb.EOF Then Exit Sub

70    If IsDate(tb!SampleDate) Then
80        SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
90    Else
100       SampleDate = ""
110   End If
120   If IsDate(tb!Rundate) Then
130       Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
140   Else
150       Rundate = ""
160   End If

170   NameToFind = AddTicks(tb!PatName & "")

180   If IsDate(tb!DoB) Then
190       DoB = Format(tb!DoB, "dd/mmm/yyyy")
200   Else
210       DoB = ""
220   End If



230   ClearUdtHeading
240   With udtHeading
250       .SampleID = RP.SampleID
260       .Dept = "Biochemistry"
270       .Name = tb!PatName & ""
280       .Ward = RP.Ward
290       .DoB = DoB
300       .Chart = tb!Chart & ""
310       .Clinician = RP.Clinician
320       .Address0 = tb!Addr0 & ""
330       .Address1 = tb!Addr1 & ""
340       .GP = RP.GP
350       .Sex = tb!Sex & ""
360       .Hospital = tb!Hospital & ""
370       .SampleDate = tb!SampleDate & ""
380       .RecDate = tb!RecDate & ""
390       .Rundate = tb!Rundate & ""
400       .GpClin = ""
410       .SampleType = "S"
420       .DocumentNo = GetOptionSetting("BioGgtDocumentNo", "")
430       .AandE = tb!AandE & ""
440   End With

450   PrintHeadingRTB
460   With frmRichText.rtb
470       .SelFontSize = 10

480       .SelText = vbCrLf & vbCrLf & vbCrLf

490       .SelColor = vbGreen
500       .SelFontSize = 12
510       .SelText = Left(" " & Space(30), 30) & "Glucose Tolerance Test" & vbCrLf
520       .SelColor = vbBlack
530       .SelFontSize = 10
540       .SelText = vbCrLf & vbCrLf
550       .SelText = Left(" " & Space(15), 15) & Left("Test Name " & Space(45), 45) & "Result Unit" & vbCrLf & vbCrLf    '  Urine mmol/L" & vbCrLf & vbcrlf
560       CrCnt = CrCnt + 8

          'now we know patient details. pick up all glucose tests for this patient
570       sql = "SELECT d.SampleID , b.Result, b.Code, b.Units, b.Operator from demographics d " & _
                  "left join BioResults b on d.SampleID = b.SampleId " & _
                  "WHERE d.patname = '" & NameToFind & "' and d.RunDate = '" & Format$(Rundate, "dd/mmm/yyyy") & "' " & _
                  "and (Code = '" & SysOptBioCodeForFastGlucose(0) & "' " & _
                  " or Code = '" & SysOptBioCodeForGlucose1(0) & "'" & _
                  " or Code = '" & SysOptBioCodeForGlucose2(0) & "'" & _
                  " or Code = '" & SysOptBioCodeForGlucose3(0) & "'" & _
                  " or Code = '" & SysOptBioCodeForFastGlucoseP(0) & "' " & _
                  " or Code = '" & SysOptBioCodeForGlucose1P(0) & "'" & _
                  " or Code = '" & SysOptBioCodeForGlucose2P(0) & "'" & _
                  " or Code = '" & SysOptBioCodeForGlucose3P(0) & "')"
580       If IsDate(DoB) Then
590           sql = sql & "and d.DoB = '" & Format$(DoB, "dd/mmm/yyyy") & "' "
600       End If

610       Set tb = New Recordset
620       RecOpenClient 0, tb, sql
630       If Not tb.EOF Then
640           For i = 1 To 4

650               CodeSerum = Choose(i, SysOptBioCodeForFastGlucose(0), SysOptBioCodeForGlucose1(0), SysOptBioCodeForGlucose2(0), SysOptBioCodeForGlucose3(0))
660               CodePlasma = Choose(i, SysOptBioCodeForFastGlucoseP(0), SysOptBioCodeForGlucose1P(0), SysOptBioCodeForGlucose2P(0), SysOptBioCodeForGlucose3P(0))
670               tb.MoveFirst
680               While Not tb.EOF
690                   If tb!Code & "" = CodeSerum Or tb!Code & "" = CodePlasma Then

700                       If AuthorisedBy = "" Then AuthorisedBy = GetAuthorisedBy(tb!Operator & "")
710                       LogBioAsPrinted tb!SampleID & "", CodeSerum
720                       LogBioAsPrinted tb!SampleID & "", CodePlasma

730                       .SelText = Left(" " & Space(15), 15) & Left(LongNameforCode(tb!Code & "") & "" & Space(22), 22)
740                       .SelText = Left("" & Space(23), 23)
750                       .SelText = Left(Format(tb!Result & "", "0.0") & Space(6), 6) & " " & tb!Units & "" & vbCrLf
760                       .SelText = vbCrLf
770                       CrCnt = CrCnt + 2

780                   End If
790                   tb.MoveNext
800               Wend
810               tb.MoveFirst
820           Next i
830       End If


840       If UCase$(HospName(0)) = "PORTLAOISE" Then
850           .SelText = vbCrLf
860           .SelText = vbCrLf
870           .SelText = "GESTATIONAL GTT REFERENCE RANGES." & vbCrLf
880           .SelText = "Fasting    3.5 -  5.1   mmol/L " & vbCrLf
890           .SelText = "1Hr        3.5 - 10.0 mmol/L" & vbCrLf
900           .SelText = "2Hr        3.5 -  8.5   mmol/L" & vbCrLf
              '        .SelText = "3Hr        3.5 -  8.2   mmol/L." & vbCrLf
910           .SelText = "One or more of the venous Plasma concentrations must be met" & vbCrLf
920           .SelText = "or exceeded for a Positive Diagnosis." & vbCrLf
930           CrCnt = CrCnt + 8
940       End If

'950       PrintFooterRTB AuthorisedBy, SampleDate, Rundate
950     If UCase(GetOptionSetting("GetLatestAuthorisedBy", "")) = UCase("True") Then
960         PrintFooterRTB GetLatestAuthorisedBy("Bio", RP.SampleID), SampleDate, GetLatestRunDateTime("Bio", RP.SampleID, Rundate)
970     Else
980         PrintFooterRTB AuthorisedBy, SampleDate, GetLatestRunDateTime("Bio", RP.SampleID, Rundate)
990     End If
1000      .SelStart = 0
1010      .SelPrint Printer.hdc

1020      sql = "SELECT * FROM Reports WHERE 0 = 1"
1030      Set tb = New Recordset
1040      RecOpenServer 0, tb, sql
1050      tb.AddNew
1060      tb!SampleID = RP.SampleID
1070      tb!Name = udtHeading.Name
1080      tb!Dept = "G"
1090      tb!Initiator = RP.Initiator
1100      tb!PrintTime = PrintTime
1110      tb!RepNo = "0G" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
1120      tb!PageNumber = 0
1130      tb!Report = .TextRTF
1140      tb!Printer = Printer.DeviceName
1150      tb.Update
1160  End With

1170  ResetPrinter

1180  Exit Sub

PrintGTT_Error:

      Dim strES As String
      Dim intEL As Integer

1190  intEL = Erl
1200  strES = Err.Description
1210  LogError "modBiochemistry", "PrintGTT", intEL, strES, sql

End Sub


'Public Sub PrintGTT()
'
'Dim tb As Recordset
'Dim tbDems As Recordset
'Dim tbRes As Recordset
'Dim sql As String
'Dim DoB As String
'Dim NameToFind As String
'Dim SampleDate As String
'Dim Rundate As String
'Dim Code As Integer
'Dim CodeP As String
'Dim PrintTime As String
'Dim AuthorisedBy As String
'
'On Error GoTo PrintGTT_Error
'
'PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
'
'Code = SysOptBioCodeForGlucose(0)
'CodeP = SysOptBioCodeForGlucoseP(0)
'
'sql = "SELECT * FROM Demographics WHERE " & _
 '      "SampleID = '" & RP.SampleID & "'"
'Set tb = New Recordset
'RecOpenClient 0, tb, sql
'If tb.EOF Then Exit Sub
'
'If IsDate(tb!SampleDate) Then
'    SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
'Else
'    SampleDate = ""
'End If
'If IsDate(tb!Rundate) Then
'    Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
'Else
'    Rundate = ""
'End If
'
'NameToFind = tb!PatName & ""
'
'If IsDate(tb!DoB) Then
'    DoB = Format(tb!DoB, "dd/mmm/yyyy")
'Else
'    DoB = ""
'End If
'
'ClearUdtHeading
'With udtHeading
'    .SampleID = RP.SampleID
'    .Dept = "Biochemistry"
'    .Name = tb!PatName & ""
'    .Ward = RP.Ward
'    .DoB = DoB
'    .Chart = tb!Chart & ""
'    .Clinician = RP.Clinician
'    .Address0 = tb!Addr0 & ""
'    .Address1 = tb!Addr1 & ""
'    .GP = RP.GP
'    .Sex = tb!Sex & ""
'    .Hospital = tb!Hospital & ""
'    .SampleDate = tb!SampleDate & ""
'    .RecDate = tb!RecDate & ""
'    .Rundate = tb!Rundate & ""
'    .GpClin = ""
'    .SampleType = "S"
'    .DocumentNo = GetOptionSetting("BioGgtDocumentNo", "")
'End With
'
'sql = "SELECT * FROM demographics WHERE " & _
 '      "patname = '" & NameToFind & "' " & _
 '      "and rundate = '" & Format(tb!Rundate, "dd/mmm/yyyy") & "' "
'If IsDate(DoB) Then
'    sql = sql & "and DoB = '" & Format(DoB, "dd/mmm/yyyy") & "' "
'End If
'sql = sql & "order by SampleID"
'Set tbDems = New Recordset
'RecOpenClient 0, tbDems, sql
'
'If tbDems.EOF Then
'    Exit Sub
'End If
'
'PrintHeadingRTB
'With frmRichText.rtb
'    .SelFontSize = 10
'
'    .SelText = vbCrLf & vbCrLf & vbCrLf
'
'    .SelColor = vbGreen
'    .SelFontSize = 12
'    .SelText = Left(" " & Space(30), 30) & "Glucose Tolerance Test" & vbCrLf
'    .SelColor = vbBlack
'    .SelFontSize = 10
'    .SelText = vbCrLf & vbCrLf
'    .SelText = Left(" " & Space(15), 15) & Left("Test Name " & Space(45), 45) & "Serum mmol/L" & vbCrLf & vbCrLf    '  Urine mmol/L" & vbCrLf & vbcrlf
'    CrCnt = CrCnt + 8
'
'    Do While Not tbDems.EOF
'        sql = "SELECT * FROM BioResults WHERE " & _
         '              "SampleID = '" & tbDems!SampleID & "' " & _
         '              "and (Code = '" & SysOptBioCodeForFastGlucose(0) & _
         '              "' or code = '" & SysOptBioCodeForGlucose1(0) & _
         '              "' or code = '" & SysOptBioCodeForGlucose2(0) & _
         '              "' or Code = '" & SysOptBioCodeForFastGlucoseP(0) & _
         '              "' or code = '" & SysOptBioCodeForGlucose1P(0) & _
         '              "' or code = '" & SysOptBioCodeForGlucose2P(0) & _
         '              "' or code = '" & SysOptBioCodeForGlucose3P(0) & "')"
'        Set tbRes = New Recordset
'        RecOpenClient 0, tbRes, sql
'
'        If Not tbRes.EOF Then
'            If AuthorisedBy = "" Then AuthorisedBy = GetAuthorisedBy(tbRes!Operator & "")
'            LogBioAsPrinted tbDems!SampleID & "", tbRes!Code
'            .SelText = Left(" " & Space(15), 15) & Left(LongNameforCode(tbRes!Code) & "" & Space(22), 22)
'            .SelText = Left("" & Space(27), 27)
'            .SelText = Format(tbRes!Result, "0.0") & vbCrLf
'            .SelText = vbCrLf
'            CrCnt = CrCnt + 2
'        End If
'        tbDems.MoveNext
'    Loop
'
'    If UCase$(HospName(0)) = "PORTLAOISE" Then
'        .SelText = vbCrLf
'        .SelText = vbCrLf
'        .SelText = "GESTATIONAL GTT REFERENCE RANGES." & vbCrLf
'        .SelText = "Fasting    3.5 -  5.1   mmol/L " & vbCrLf
'        .SelText = "1Hr        3.5 - 10.0 mmol/L" & vbCrLf
'        .SelText = "2Hr        3.5 -  8.5   mmol/L" & vbCrLf
''        .SelText = "3Hr        3.5 -  8.2   mmol/L." & vbCrLf
'        .SelText = "One or more of the venous Plasma concentrations must be met" & vbCrLf
'        .SelText = "or exceeded for a Positive Diagnosis." & vbCrLf
'        CrCnt = CrCnt + 8
'    End If
'
'    PrintFooterRTB AuthorisedBy, SampleDate, Rundate
'    .SelStart = 0
'    .SelPrint Printer.hDC
'
'    sql = "SELECT * FROM Reports WHERE 0 = 1"
'    Set tb = New Recordset
'    RecOpenServer 0, tb, sql
'    tb.AddNew
'    tb!SampleID = RP.SampleID
'    tb!Name = udtHeading.Name
'    tb!Dept = "G"
'    tb!Initiator = RP.Initiator
'    tb!PrintTime = PrintTime
'    tb!RepNo = "0G" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
'    tb!PageNumber = 0
'    tb!Report = .TextRTF
'    tb!Printer = Printer.DeviceName
'    tb.Update
'End With
'
'ResetPrinter
'
'Exit Sub
'
'PrintGTT_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'intEL = Erl
'strES = Err.Description
'LogError "modBiochemistry", "PrintGTT", intEL, strES, sql
'
'End Sub

Public Sub PrintResultBioSideBySide()

Dim f As Integer
Dim tb As Recordset
Dim tbUN As Recordset
Dim sql As String
Dim Sex As String
Dim lpc As Integer
Dim cUnits As String
Dim Flag As String
Dim n As Integer
Dim v As String
Dim Low As Single
Dim High As Single
Dim strLow As String * 4
Dim strHigh As String * 4
Dim BRs As New BIEResults
Dim br As BIEResult
Dim TestCount As Integer
Dim SampleType As String
Dim ResultsPresent As Boolean
'Dim Cx As Comment
'Dim Cxs As New Comments
Dim OB As Observation
Dim OBS As New Observations
10    ReDim Comments(1 To 4) As String
Dim SampleDate As String
Dim Rundate As String
Dim DoB As String
Dim RunTime As String
Dim Fasting As String
Dim Fx As Fasting
Dim CodeGLU As String
Dim CodeCHO As String
Dim CodeTRI As String
Dim CodeGLUP As String
Dim CodeCHOP As String
Dim CodeTRIP As String
Dim udtPrintLine(0 To 60) As PrintLine    'max 30 result lines per page
Dim strFormat As String
Dim MultiColumn As Boolean
Dim str As String
Dim PrintTime As String
Dim AuthorisedBy As String

20    On Error GoTo PrintResultBioSideBySide_Error

30    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

40    For n = 0 To 60
50    udtPrintLine(n).Analyte = ""
60    udtPrintLine(n).Result = ""
70    udtPrintLine(n).Flag = ""
80    udtPrintLine(n).Units = ""
90    udtPrintLine(n).NormalRange = ""
100   udtPrintLine(n).Fasting = ""
110   udtPrintLine(n).Reason = ""
120   Next

130   sql = "SELECT * FROM Demographics WHERE " & _
          "SampleID = '" & RP.SampleID & "'"
140   Set tb = New Recordset
150   RecOpenClient 0, tb, sql

160   If tb.EOF Then
170   Exit Sub
180   End If

190   If Not IsNull(tb!Fasting) Then
200   Fasting = tb!Fasting
210   Else
220   Fasting = False
230   End If

240   CodeGLU = SysOptBioCodeForGlucose(0)
250   CodeCHO = SysOptBioCodeForChol(0)
260   CodeTRI = SysOptBioCodeForTrig(0)
270   CodeGLUP = SysOptBioCodeForGlucoseP(0)
280   CodeCHOP = SysOptBioCodeForCholP(0)
290   CodeTRIP = SysOptBioCodeForTrigP(0)

300   If IsDate(tb!DoB) Then
310   DoB = Format(tb!DoB, "dd/mmm/yyyy")
320   Else
330   DoB = ""
340   End If

350   ResultsPresent = False
360   Set BRs = BRs.Load("Bio", RP.SampleID, "Results", 0, "", "")
370   If Not BRs Is Nothing Then
380   TestCount = BRs.Count
390   If TestCount <> 0 Then
400     ResultsPresent = True
410     SampleType = BRs(1).SampleType
420     If Trim(SampleType) = "" Then SampleType = "S"
430   End If
440   End If

450   lpc = 0
460   If ResultsPresent Then
470   For Each br In BRs
480     If AuthorisedBy = "" Then AuthorisedBy = GetAuthorisedBy(br.Operator)
490     RunTime = br.RunTime
500     v = br.Result

510     If br.Code = CodeGLU Or br.Code = CodeCHO Or br.Code = CodeTRI Or _
           br.Code = CodeGLUP Or br.Code = CodeCHOP Or br.Code = CodeTRIP Then
520         If Fasting Then
530             Set Fx = Nothing
540             If br.Code = CodeGLU Or br.Code = CodeGLUP Then
550                 Set Fx = colFastings("GLU")
560             ElseIf br.Code = CodeCHO Or br.Code = CodeCHOP Then
570                 Set Fx = colFastings("CHO")
580             ElseIf br.Code = CodeTRI Or br.Code = CodeTRIP Then
590                 Set Fx = colFastings("TRI")
600             End If
610             If Not Fx Is Nothing Then
620                 High = Fx.FastingHigh
630                 Low = Fx.FastingLow
640             Else
650                 High = Val(br.High)
660                 Low = Val(br.Low)
670             End If
680         Else
690             High = Val(br.High)
700             Low = Val(br.Low)
710         End If
720     Else
730         High = Val(br.High)
740         Low = Val(br.Low)
750     End If

760     If Low < 10 Then
770         strLow = Format(Low, "0.00")
780     ElseIf Low < 100 Then
790         strLow = Format(Low, "##.0")
800     Else
810         strLow = Format(Low, " ###")
820     End If
830     If High < 10 Then
840         strHigh = Format(High, "0.00")
850     ElseIf High < 100 Then
860         strHigh = Format(High, "##.0")
870     Else
880         strHigh = Format(High, "### ")
890     End If

900     If IsNumeric(v) Then
910         If Val(v) > br.PlausibleHigh Then
920             udtPrintLine(lpc).Flag = " X "
930             udtPrintLine(lpc).Result = "***"
940             Flag = " X"
950         ElseIf Val(v) < br.PlausibleLow Then
960             udtPrintLine(lpc).Flag = " X "
970             udtPrintLine(lpc).Result = "***"
980             Flag = " X"
990         ElseIf Val(v) > High And High <> 0 Then
1000            udtPrintLine(lpc).Flag = " H "
1010            Flag = " H"
1020        ElseIf Val(v) < Low Then
1030            udtPrintLine(lpc).Flag = " L "
1040            Flag = " L"
1050        Else
1060            udtPrintLine(lpc).Flag = "   "
1070            Flag = "  "
1080        End If
1090    Else
1100        udtPrintLine(lpc).Flag = "   "
1110        Flag = "  "
1120    End If
1130    udtPrintLine(lpc).Analyte = Left(br.LongName & Space(16), 16)

1140    If TestAffected(br) = False Then
1150        If IsNumeric(v) Then
1160            Select Case br.Printformat
                Case 0: strFormat = "######"
1170            Case 1: strFormat = "###0.0"
1180            Case 2: strFormat = "##0.00"
1190            Case 3: strFormat = "#0.000"
1200            End Select
1210            If udtPrintLine(lpc).Result <> "***" Then udtPrintLine(lpc).Result = Format(v, strFormat)
1220        Else
1230            If udtPrintLine(lpc).Result <> "***" Then udtPrintLine(lpc).Result = v
1240        End If
1250    Else
1260        udtPrintLine(lpc).Result = "XXXXXX"
1270    End If

1280    sql = "SELECT * FROM Lists WHERE " & _
              "ListType = 'UN' and Code = '" & br.Units & "'"
1290    Set tbUN = Cnxn(0).Execute(sql)
1300    If Not tbUN.EOF Then
1310        cUnits = Left(tbUN!Text & Space(10), 10)
1320    Else
1330        cUnits = Left(br.Units & Space(10), 10)
1340    End If
        'Zyam 11-15-23
        If Val(strLow) = 0 And Val(strHigh) = 0 Then
            udtPrintLine(lpc).NormalRange = " "
        Else
1350    udtPrintLine(lpc).NormalRange = "(" & strLow & "-" & strHigh & ") "
        End If
        'Zyam 11-15-23
1360    udtPrintLine(lpc).Units = cUnits

1370    udtPrintLine(lpc).Fasting = ""
1380    If Fasting Then
1390        If br.Code = CodeGLU Or br.Code = CodeCHO Or br.Code = CodeTRI Or _
               br.Code = CodeGLUP Or br.Code = CodeCHOP Or br.Code = CodeTRIP Then
1400            udtPrintLine(lpc).Fasting = "(Fasting)"
1410        End If
1420    End If

1430    If TestAffected(br) = True Then
1440        udtPrintLine(lpc).Reason = ReasonAffect(br)
1450    Else
1460        udtPrintLine(lpc).Reason = ""
1470    End If
1480    LogTestAsPrinted "Bio", br.SampleID, br.Code
1490    lpc = lpc + 1
1500  Next
1510  End If

1520  ClearUdtHeading
1530  With udtHeading
1540  .SampleID = RP.SampleID
1550  .Dept = "Biochemistry"
1560  .Name = tb!PatName & ""
1570  .Ward = RP.Ward
1580  .DoB = DoB
1590  .Chart = tb!Chart & ""
1600  .Clinician = RP.Clinician
1610  .Address0 = tb!Addr0 & ""
1620  .Address1 = tb!Addr1 & ""
1630  .GP = RP.GP
1640  .Sex = tb!Sex & ""
1650  .Hospital = tb!Hospital & ""
1660  .SampleDate = tb!SampleDate & ""
1670  .RecDate = tb!RecDate & ""
1680  .Rundate = tb!Rundate & ""
1690  .GpClin = ""
1700  .SampleType = SampleType
1710  .DocumentNo = GetOptionSetting("BioSbsDocumentNo", "")
1720  .AandE = tb!AandE & ""
1730  End With
1740  PrintHeadingRTB

1750  Sex = tb!Sex & ""

1760  With frmRichText.rtb
1770  If TestCount <= Val(frmMain.txtMoreThan) Then
1780    MultiColumn = False
        '  Printer.CurrentY = 2500 + (20 - TestCount) * 100
1790    For n = 1 To Val(frmMain.txtMoreThan) - TestCount / 2
1800        .SelText = vbCrLf
1810    Next
1820  Else
1830    MultiColumn = True
        '  Printer.CurrentY = 2500
1840  End If

1850  .SelFontSize = 10

1860  If Not IsDate(DoB) Or Trim(udtHeading.Sex) = "" Then        'QMS Ref #818581 (And changed to or)
1870    .SelBold = True
1880    .SelText = "               " & "**** No Sex/DoB given. No reference range applied! ****" & vbCrLf
1890  End If

1900  If MultiColumn Then
1910    For n = 0 To Val(frmMain.txtMoreThan) - 1
1920        .SelColor = vbBlack
1930        If Trim(Sex) <> "" Then      'QMS Ref #817982
1940            If Trim(udtPrintLine(n).Flag) = "L" Then .SelColor = vbBlue
1950            If Trim(udtPrintLine(n).Flag) = "H" Then .SelColor = vbRed
1960        End If
1970        .SelBold = False
1980        .SelText = udtPrintLine(n).Analyte

1990        If udtPrintLine(n).Flag <> "   " And Trim(Sex) <> "" Then       'QMS Ref #817982
2000            .SelBold = True
2010        End If
2020        .SelText = udtPrintLine(n).Result
2030        If Trim(Sex) <> "" Then      'QMS Ref #817982
2040            .SelText = udtPrintLine(n).Flag
2050        End If
2060        .SelBold = False
2070        .SelFontSize = 8
2080        .SelText = udtPrintLine(n).Units
2090        If Trim(Sex) <> "" Then      'QMS Ref #817982
2100            .SelText = udtPrintLine(n).NormalRange
2110        End If
2120        .SelFontSize = 10
            'Now Right Hand Column
2130        .SelText = "     "
2140        .SelText = udtPrintLine(n + Val(frmMain.txtMoreThan)).Analyte
2150        If udtPrintLine(n + Val(frmMain.txtMoreThan)).Flag <> "   " And Trim(Sex) <> "" Then       'QMS Ref #817982
2160            .SelBold = True
2170        End If
2180        .SelText = udtPrintLine(n + Val(frmMain.txtMoreThan)).Result
2190        If Trim(Sex) <> "" Then      'QMS Ref #817982
2200            .SelText = udtPrintLine(n + Val(frmMain.txtMoreThan)).Flag
2210        End If
2220        .SelBold = False
2230        .SelFontSize = 8
2240        .SelText = udtPrintLine(n + Val(frmMain.txtMoreThan)).Units
2250        If Trim(Sex) <> "" Then      'QMS Ref #817982
2260            .SelText = udtPrintLine(n + Val(frmMain.txtMoreThan)).NormalRange
2270        End If
2280        .SelText = vbCrLf
2290        .SelFontSize = 10
2300    Next
2310    If Fasting Then
2320        .SelText = "(All above relate to Normal Fasting Ranges.)"
2330        .SelText = vbCrLf
2340    End If
2350  Else
2360    For n = 0 To 35
2370        If Trim(udtPrintLine(n).Analyte) <> "" Then
2380            .SelColor = vbBlack
2390            If Trim(Sex) <> "" Then      'QMS Ref #817982
2400                If Trim(udtPrintLine(n).Flag) = "L" Then .SelColor = vbBlue
2410                If Trim(udtPrintLine(n).Flag) = "H" Then .SelColor = vbRed
2420            End If
2430            .SelText = Space(20)
2440            .SelBold = False
2450            .SelText = udtPrintLine(n).Analyte
2460            If udtPrintLine(n).Flag <> "   " And Trim(Sex) <> "" Then      'QMS Ref #817982
2470                .SelBold = True
2480            End If
2490            .SelText = udtPrintLine(n).Result
2500            If Trim(Sex) <> "" Then      'QMS Ref #817982
2510                .SelText = udtPrintLine(n).Flag
2520            End If
2530            .SelBold = False
2540            .SelText = udtPrintLine(n).Units
2550            If Trim(Sex) <> "" Then      'QMS Ref #817982
2560                .SelText = udtPrintLine(n).NormalRange
2570            End If
2580            .SelText = udtPrintLine(n).Fasting
2590            .SelText = vbCrLf
2600        End If
2610    Next
2620  End If
    '  Set Cx = Cxs.Load(RP.SampleID)
2630  Set OBS = OBS.Load(RP.SampleID, "Biochemistry", "Demographic")
2640  If Not OBS Is Nothing Then
2650    For Each OB In OBS
2660        Select Case UCase$(OB.Discipline)
            Case "BIOCHEMISTRY"
2670            FillCommentLines OB.Comment, 4, Comments(), 97
2680            For n = 1 To 4
2690                .SelText = Comments(n) & vbCrLf
2700            Next
2710        Case "DEMOGRAPHIC"
2720            FillCommentLines OB.Comment, 2, Comments(), 97
2730            For n = 1 To 4
2740                .SelText = Comments(n) & vbCrLf
2750            Next
2760        End Select
2770    Next
2780  End If

2790  .SelText = vbCrLf
    '    If Not IsDate(tb!DoB) And Trim(Sex) = "" Then
    '        .SelColor = vbBlue
    '        .SelText = Space(24) & "No Sex/DoB given. Normal ranges may not be relevant"
    '    ElseIf Not IsDate(tb!DoB) Then
    '        .SelColor = vbBlue
    '        .SelText = Space(24) & "No DoB given. Normal ranges may not be relevant"
    '    ElseIf Trim(Sex) = "" Then
    '        .SelColor = vbBlue
    '        .SelText = Space(24) & "No Sex given. No Reference range applied"
    '    End If
    '    .SelText = vbCrLf

2800  .SelColor = vbBlack

2810  If IsDate(tb!SampleDate) Then
2820    SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
2830  Else
2840    SampleDate = ""
2850  End If
2860  If IsDate(RunTime) Then
2870    Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
2880  Else
2890    If IsDate(tb!Rundate) Then
2900        Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
2910    Else
2920        Rundate = ""
2930    End If
2940  End If

2950  .SelBold = False


2960  If RP.FaxNumber <> "" Then
2970    PrintFooterRTBFax RP.Initiator, SampleDate, Rundate
2980    f = FreeFile
2990    Open SysOptFax(0) & RP.SampleID & "BIO2.doc" For Output As f
3000    .SelStart = 0
3010    Print #f, .TextRTF
3020    Close f
3030    SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "BIOX.doc"
3040  Else
        '3050          PrintFooterRTB AuthorisedBy, SampleDate, Rundate
3050    If UCase(GetOptionSetting("GetLatestAuthorisedBy", "")) = UCase("True") Then
3060        PrintFooterRTB GetLatestAuthorisedBy("Bio", RP.SampleID), SampleDate, GetLatestRunDateTime("Bio", RP.SampleID, Rundate)
3070    Else
3080        PrintFooterRTB AuthorisedBy, SampleDate, GetLatestRunDateTime("Bio", RP.SampleID, Rundate)
3090    End If
3100    .SelStart = 0
3110    .SelPrint Printer.hdc
3120  End If
3130  sql = "SELECT * FROM Reports WHERE 0 = 1"
3140  Set tb = New Recordset
3150  RecOpenServer 0, tb, sql
3160  tb.AddNew
3170  tb!SampleID = RP.SampleID
3180  tb!Name = udtHeading.Name
3190  tb!Dept = "B"
3200  tb!Initiator = RP.Initiator
3210  tb!PrintTime = PrintTime
3220  tb!RepNo = "0B" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
3230  tb!PageNumber = 0
3240  tb!Report = .TextRTF
3250  tb!Printer = Printer.DeviceName
3260  tb.Update
3270  End With

3280  ResetPrinter

3290  sql = "Update BioResults set Printed = '1' WHERE " & _
           "SampleID = '" & RP.SampleID & "'"
3300  Cnxn(0).Execute sql

3310  Exit Sub

PrintResultBioSideBySide_Error:

Dim strES As String
Dim intEL As Integer

3320  intEL = Erl
3330  strES = Err.Description
3340  LogError "modBiochemistry", "PrintResultBioSideBySide", intEL, strES, sql

End Sub



Public Sub PrintResultBioWin(Optional ByVal PrintA4 As Boolean = False)



      Dim C          As Integer
      Dim tb         As Recordset
      Dim tbUN       As Recordset
      Dim bc         As Recordset
      Dim sql        As String
      Dim Sex        As String

      Dim lpc        As Integer
      Dim cUnits     As String
      Dim TempUnits  As String
      Dim Flag       As String
      Dim n          As Integer
      Dim i          As Integer
      Dim v          As String
      Dim Low        As Single
      Dim High       As Single
      Dim strLow     As String * 5
      Dim strHigh    As String * 5
      Dim BRs        As New BIEResults
      Dim br         As BIEResult
      Dim SampleType As String
      Dim ResultsPresent As Boolean
      Dim TempString As String
      Dim TempInteger As Integer

      Dim OB         As Observation
      Dim OBS        As New Observations

      Dim SampleDate As String
      Dim Rundate    As String
      Dim DoB        As String
      Dim RunTime    As String
      Dim Fasting    As String
      Dim Fx         As Fasting
      Dim CodeGLU    As String
      Dim CodeCHO    As String
      Dim CodeTRI    As String
      Dim CodeGLUP   As String
      Dim CodeCHOP   As String
      Dim CodeTRIP   As String

      Dim strFormat  As String
      Dim xT         As Long
      Dim copies     As Long
      Dim Clin       As String
      Dim d          As Long
      Dim f          As Integer
      Dim Fontz1     As Integer
      Dim Fontz2     As Integer
      Dim Fontz3     As Integer
      Dim Fontz4     As Integer
      Dim FontBold   As Boolean
      Dim EGFrFound  As Boolean
      Dim PrintTime  As String
      Dim AuthorisedBy As String
      Dim PageNumber As Integer
      Dim TestPerformedAt As String
      Dim ExternalTestingNote As String
      Dim udtPrintLine() As ResultLine
      Dim resultCount As Integer
      Dim CmtCount As Integer


      Dim TotalLines As Integer
      Dim CommentLines As Integer
      Dim PerPageLines As Integer
      Dim BodyLines  As Integer
      Dim FooterLines As Integer
      Dim LineNoStartComment As Integer
      Dim TotalPages As Integer

10    On Error GoTo PrintResultBioWin_Error

20    If PrintA4 Then
30        TotalLines = 100
40        CommentLines = 10
          PerPageLines = 77
50        'PerPageLines = 50
60        FooterLines = 3
70    Else
80        TotalLines = 100
90        CommentLines = 4
'100       PerPageLines = 35
          PerPageLines = Val(GetOptionSetting("PrintOptionsIfMoreThan", 18))
110       FooterLines = 3
120   End If

130   ReDim lp(0 To TotalLines) As String
140   ReDim udtPrintLine(0 To TotalLines) As ResultLine
150   ReDim Comments(1 To CommentLines) As String



      'Clear All
160   For n = 0 To TotalLines
170       udtPrintLine(n).Analyte = ""
180       udtPrintLine(n).Result = ""
190       udtPrintLine(n).Flag = ""
200       udtPrintLine(n).Units = ""
210       udtPrintLine(n).NormalRange = ""
220       udtPrintLine(n).Fasting = ""
230       udtPrintLine(n).Reason = ""
240   Next
250   ClearUdtHeading


      'Initailise Variables
260   xT = 5
270   EGFrFound = False
280   PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
290   CodeGLU = SysOptBioCodeForGlucose(0)
300   CodeCHO = SysOptBioCodeForChol(0)
310   CodeTRI = SysOptBioCodeForTrig(0)
320   CodeGLUP = SysOptBioCodeForGlucoseP(0)
330   CodeCHOP = SysOptBioCodeForCholP(0)
340   CodeTRIP = SysOptBioCodeForTrigP(0)


350   sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
360   Set tb = New Recordset
370   RecOpenClient 0, tb, sql

380   If tb.EOF Then
390       Exit Sub
400   End If

410   Rundate = tb!Rundate

420   If Not IsNull(tb!Fasting) Then
430       Fasting = tb!Fasting
440   Else
450       Fasting = False
460   End If

470   If IsDate(tb!DoB) Then
480       DoB = Format(tb!DoB, "dd/mmm/yyyy")
490   Else
500       DoB = ""
510   End If

520   With udtHeading
530       .SampleID = RP.SampleID
540       .Dept = "Biochemistry"
550       .Name = tb!PatName & ""
560       .Ward = RP.Ward
570       .DoB = DoB
580       .Chart = tb!Chart & ""
590       .Clinician = RP.Clinician
600       .Address0 = tb!Addr0 & ""
610       .Address1 = tb!Addr1 & ""
620       .GP = RP.GP
630       .Sex = tb!Sex & ""
640       .Hospital = tb!Hospital & ""
650       .SampleDate = tb!SampleDate & ""
660       .RecDate = tb!RecDate & ""
670       .Rundate = tb!Rundate & ""
680       .GpClin = Clin
690       .SampleType = SampleType
700       .DocumentNo = GetOptionSetting("BioMainDocumentNo", "")
710       .AandE = tb!AandE & ""
720   End With

730   ResultsPresent = False
740   Set BRs = BRs.Load("Bio", RP.SampleID, "Results", 0, "", "")
750   If Not BRs Is Nothing Then
760       If BRs.Count <> 0 Then
770           ResultsPresent = True
780           SampleType = BRs(1).SampleType
790           If Trim(SampleType) = "" Then SampleType = "S"
800       End If
810   End If

820   If IsDate(tb!SampleDate) Then
830       SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
840   Else
850       SampleDate = ""
860   End If
870   If IsDate(RunTime) Then
880       Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
890   Else
900       If IsDate(tb!Rundate) Then
910           Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
920       Else
930           Rundate = ""
940       End If
950   End If

960   Sex = tb!Sex & ""


970   lpc = 0
      'Zyam 25-7-24 Changed as due to Portlaoise requested
980   'AddResultToLP udtPrintLine, lp, lpc, "Test", "Result", "Unit", "Ref. Range", "Flag", , , , True, , True
      'Zyam 25-07-24
990   'CrCnt = CrCnt + 1

1000  If Not IsDate(DoB) Or Trim(udtHeading.Sex) = "" Then
1010      lp(lpc) = "               " & "**** No Sex/DoB given. No reference range applied! ****"
1020      udtPrintLine(lpc).Analyte = "*COMMENT*"
1030      lpc = lpc + 1

1040  End If

1050  If ResultsPresent Then
1060      For Each br In BRs
1070          If br.Printable = True Then
1080              AuthorisedBy = GetAuthorisedBy(br.Operator)
                  'If Trim(Br.Operator) <> "" Then RP.Initiator = Br.Operator
1090              If InStr(1, SampleType, br.SampleType) = 0 Then
1100                  SampleType = SampleType & ", " & br.SampleType
1110              End If
1120              RunTime = br.RunTime
1130              Rundate = br.Rundate
1140              If Rundate = "" Then
1150                  Rundate = RunTime
1160              End If
1170              If br.Code = SysOptBioCodeForRF(0) Then
1180                  If IsNumeric(br.Result) And Val(br.Result) < 7 Then
1190                      br.Result = "< 7"
1200                  End If
1210              End If
1220              v = br.Result
1230              lp(lpc) = ""
1240              If br.Code = CodeGLU Or br.Code = CodeCHO Or br.Code = CodeTRI Or _
                     br.Code = CodeGLUP Or br.Code = CodeCHOP Or br.Code = CodeTRIP Then
1250                  If Fasting Then
1260                      Set Fx = Nothing
1270                      If br.Code = CodeGLU Or br.Code = CodeGLUP Then
1280                          Set Fx = colFastings("GLU")
1290                      ElseIf br.Code = CodeCHO Or br.Code = CodeCHOP Then
1300                          Set Fx = colFastings("CHO")
1310                      ElseIf br.Code = CodeTRI Or br.Code = CodeTRIP Then
1320                          Set Fx = colFastings("TRI")
1330                      End If
1340                      If Not Fx Is Nothing Then
1350                          High = Fx.FastingHigh
1360                          Low = Fx.FastingLow
1370                      Else
1380                          High = Val(br.High)
1390                          Low = Val(br.Low)
1400                      End If
1410                  Else
1420                      High = Val(br.High)
1430                      Low = Val(br.Low)
1440                  End If
1450              Else
1460                  High = Val(br.High)
1470                  Low = Val(br.Low)
1480              End If

1490              If Low < 10 Then
1500                  strLow = Format(Low, "0.00")
1510              ElseIf Low < 100 Then
1520                  strLow = Format(Low, "##.0")
1530              ElseIf Low > 99 And Low < 1000 Then
1540                  strLow = Format(Low, " ###0")
1550              Else
1560                  strLow = Format(Low, "####")
1570              End If
1580              If High < 10 Then
1590                  strHigh = Format(High, "0.00")
1600              ElseIf High < 100 Then
1610                  strHigh = Format(High, "##.0")
1620              Else
1630                  strHigh = Format(High, "#### ")
1640              End If

1650              If IsNumeric(v) Then
1660                  If Val(v) > br.PlausibleHigh Then
1670                      udtPrintLine(lpc).Flag = " X "
1680                      udtPrintLine(lpc).Result = "***"
1690                      Flag = " X"
1700                  ElseIf Val(v) < br.PlausibleLow Then
1710                      udtPrintLine(lpc).Flag = " X "
1720                      udtPrintLine(lpc).Result = "***"
1730                      Flag = " X"
1740                  ElseIf Val(v) > High And High <> 0 Then
1750                      udtPrintLine(lpc).Flag = " H "
1760                      Flag = " H"
1770                  ElseIf Val(v) < Low Then
1780                      udtPrintLine(lpc).Flag = " L "
1790                      Flag = " L"
1800                  Else
1810                      udtPrintLine(lpc).Flag = "   "
1820                      Flag = "  "
1830                  End If
1840              Else
1850                  If InStr(v, ">") Then
1860                      udtPrintLine(lpc).Flag = " H "
1870                      Flag = " H"
1880                  Else
1890                      udtPrintLine(lpc).Flag = "   "
1900                      Flag = "  "
1910                  End If
1920              End If
                  'If SysOptBioCodeForEGFR(0) = Br.Code Then

1930               If InStr(br.Code, SysOptBioCodeForEGFR(0)) Then

1940                  If UCase(HospName(0)) = "TULLAMORE" Then
1950                      If InStr(Flag, "H") > 0 Then Flag = "   ": udtPrintLine(lpc).Flag = ""
1960                      If InStr(br.Result, ">") > 0 Then
1970                          Flag = "   ": udtPrintLine(lpc).Flag = ""
1980                      Else
1990                          Flag = " L ": udtPrintLine(lpc).Flag = " L "
2000                      End If
2010                  Else
2020                      Flag = "   ": udtPrintLine(lpc).Flag = ""
2030                  End If
2040              End If
2050              If br.Code = "418" Or br.Code = "746" Or Trim(udtHeading.Sex) = "" Or _
                     udtHeading.DoB = "" Then         'QMS Ref #817982, #812404, #818581   'suppress ref range for Gentamicin
2060                  udtPrintLine(lpc).Flag = "   "
2070                  Flag = "  "
2080              End If


2090              TestPerformedAt = ""
2100              If UCase(HospName(0)) <> UCase(br.Hospital) Then
2110                  TestPerformedAt = Left(UCase(br.Hospital), 1)
2120                  If InStr(ExternalTestingNote, UCase(br.Hospital)) = 0 Then
2130                      ExternalTestingNote = ExternalTestingNote & TestPerformedAt & " = Test Analysed at " & UCase(br.Hospital) & " "
2140                  End If
2150                  TestPerformedAt = "(" & TestPerformedAt & ")"
2160              End If


2170              lp(lpc) = lp(lpc) & Left(br.LongName & TestPerformedAt & Space(20), 20)
2180              udtPrintLine(lpc).Analyte = Left(br.LongName & Space(16), 16)

2190              If TestAffected(br) = False Then
2200                  If IsNumeric(v) Then
2210                      Select Case br.Printformat
                              Case 0: strFormat = "########"
2220                          Case 1: strFormat = "#####0.0"
2230                          Case 2: strFormat = "####0.00"
2240                          Case 3: strFormat = "###0.000"
2250                      End Select
2260                      If Trim(udtPrintLine(lpc).Result) <> "***" Then
2270                          lp(lpc) = lp(lpc) & " " & Right(Space(8) & Format(v, strFormat), 8)
2280                      Else
2290                          lp(lpc) = lp(lpc) & "  ******* "
2300                      End If
2310                      If Trim(udtPrintLine(lpc).Result) <> "***" Then
2320                          udtPrintLine(lpc).Result = Format(v, strFormat)
2330                      End If
2340                  Else
2350                      If Trim(udtPrintLine(lpc).Result) <> "***" Then
2360                          lp(lpc) = lp(lpc) & " " & Right(Space(8) & Format(v, strFormat), 8)
2370                      Else
2380                          lp(lpc) = lp(lpc) & "  ******* "
2390                      End If
2400                      If Trim(udtPrintLine(lpc).Result) <> "***" Then
2410                          udtPrintLine(lpc).Result = Format(v, strFormat)
2420                      End If
2430                  End If

2440                  lp(lpc) = lp(lpc) & Flag & " "

2450              Else

2460                  lp(lpc) = lp(lpc) & " " & Right$(Space(8) & SysOptBioMaskSym(0), 8) & "   "
2470              End If

2480              sql = "SELECT * FROM Lists WHERE " & _
                        "ListType = 'UN' and Code = '" & br.Units & "'"
2490              Set tbUN = Cnxn(0).Execute(sql)
2500              If Not tbUN.EOF Then
                  'Change here to accommodate 13 characters in eGFR-----
2510                  cUnits = Left(tbUN!Text & Space(13), 13)
2520              Else
                  'Change here to accommodate 13 characters in eGFR-----
2530                  cUnits = Left(br.Units & Space(13), 13)
2540              End If
2550              udtPrintLine(lpc).Units = cUnits
2560              If br.Code = "418" Or br.Code = "746" Or _
                     Trim(udtHeading.Sex) = "" Or udtHeading.DoB = "" Then       'QMS Ref #817982, #812404, #818581
2570                  lp(lpc) = lp(lpc) & "                     " & cUnits
2580              ElseIf br.Code = SysOptBioCodeBNP(0) Then
2590                  lp(lpc) = lp(lpc) & "   (< 50) Normal"
2600                  lpc = lpc + 1
2610                  lp(lpc) = lp(lpc) & "                                             (50 - 100) Equivocal"
2620                  lpc = lpc + 1
2630                  lp(lpc) = lp(lpc) & "                                             (> 100) Abnormal"
2640              ElseIf SysOptBioCodeForEGFR(0) = br.Code Then
2650                  lp(lpc) = lp(lpc) & "                 " & br.Units
2660                  EGFrFound = True
2670              Else
2680                  If (Val(strLow) = 0 And Val(strHigh) = 0) Or (Val(strLow) = 0 And Val(strHigh) = 999) Or (Val(strLow) = 0 And Val(strHigh) = 9999) Then
2690                      lp(lpc) = lp(lpc) & "                   " & cUnits
2700                      udtPrintLine(lpc).NormalRange = "             "
2710                  Else
2720                      If br.ShowLessThan = True And strLow = 0 Then
2730                          lp(lpc) = lp(lpc) & "     < " & strHigh & "          " & cUnits
2740                      ElseIf br.ShowMoreThan = True And strHigh = 9999 Then
2750                          lp(lpc) = lp(lpc) & "     > " & strLow & "          " & cUnits
2760                      Else

                              'lp(lpc) = lp(lpc) & "   (" & strLow & " - " & strHigh & ")   " & cUnits
                              'Zyam
2770                          lp(lpc) = lp(lpc) & IIf(Val(strLow) = 0 And Val(strHigh) = 0, " ", " (" & strLow & " - " & strHigh & ")    ") & cUnits
                              'Zyam

2780                      End If
                          '11-15-23 Zyam
2790                      If Val(strLow) = 0 And Val(strHigh) = 0 Then
2800                        udtPrintLine(lpc).NormalRange = " "
2810                      Else
2820                        udtPrintLine(lpc).NormalRange = "(" & strLow & " - " & strHigh & ")"
2830                      End If
                          '11-15-23 Zyam
2840
2850                  End If
2860              End If

2870              udtPrintLine(lpc).Fasting = ""
2880              If Not IsNull(tb!Fasting) Then
2890                  If tb!Fasting Then
2900                      If tb!Fasting Then
2910                          If br.Code = CodeGLU Or br.Code = CodeCHO Or br.Code = CodeTRI Or _
                                 br.Code = CodeGLUP Or br.Code = CodeCHOP Or br.Code = CodeTRIP Then
2920                              udtPrintLine(lpc).Fasting = "(Fasting)"
2930                              lp(lpc) = lp(lpc) & "(Fasting)"
2940                          End If
2950                      End If
2960                  End If
2970              End If

2980              udtPrintLine(lpc).Reason = ""
2990              If TestAffected(br) = True Then
3000                  lp(lpc) = lp(lpc) & " " & ReasonAffect(br)
3010                  udtPrintLine(lpc).Reason = Trim(ReasonAffect(br))
3020              End If

3030              udtPrintLine(lpc).Comment = ""
3040              If br.Comment <> "" Then
3050                  lp(lpc) = lp(lpc) & " " & Trim(br.Comment)
3060                  xT = 0
3070                  udtPrintLine(lpc).Comment = Trim(br.Comment)
3080              End If
3090              LogTestAsPrinted "Bio", br.SampleID, br.Code
3100              lpc = lpc + 1
                  resultCount = resultCount + 1
3110          End If
3120      Next
3130  End If

      'add blank line before comment
3140  AddResultToLP udtPrintLine, lp, lpc, "", ""


      'comments
3150  If EGFrFound = True Then
3160      TempString = GetOptionSetting("EGFRComm1" & " ", "") & _
                       GetOptionSetting("EGFRComm2" & " ", "") & _
                       GetOptionSetting("EGFRComm3" & " ", "") & _
                       GetOptionSetting("EGFRComm4" & " ", "")

3170      AddCommentToLP udtPrintLine, lp, lpc, GetOptionSetting("EGFRComm1", ""), "EGFR Comment : "
3180      AddCommentToLP udtPrintLine, lp, lpc, GetOptionSetting("EGFRComm2", "")
3190      AddCommentToLP udtPrintLine, lp, lpc, GetOptionSetting("EGFRComm3", "")
3200      AddCommentToLP udtPrintLine, lp, lpc, GetOptionSetting("EGFRComm4", "")
3210  End If

3220  Set OBS = OBS.Load(RP.SampleID, "Biochemistry", "Demographic")
3230  If Not OBS Is Nothing Then
3240      For Each OB In OBS
3250          Select Case UCase$(OB.Discipline)
                  Case "BIOCHEMISTRY"
3260                  AddCommentToLP udtPrintLine, lp, lpc, OB.Comment, ""
3270                  CrCnt = CrCnt + 1
                      
3280              Case "DEMOGRAPHIC"
3290                  AddCommentToLP udtPrintLine, lp, lpc, OB.Comment, ""
3300                  CrCnt = CrCnt + 1
3310          End Select
              CmtCount = CmtCount + 1
3320      Next
3330  End If


3340  PrintReport udtPrintLine, lp, lpc, "Bio", PrintA4, SampleDate, Rundate, AuthorisedBy, PrintTime, resultCount, CmtCount, SampleType, ExternalTestingNote

3350  Exit Sub

PrintResultBioWin_Error:

      Dim strES      As String
      Dim intEL      As Integer

3360  intEL = Erl
3370  strES = Err.Description

3380  LogError "modBiochemistry", "PrintResultBioWin", intEL, strES, sql, Printer.DeviceName

3390  sql = "DELETE FROM PrintPending WHERE " & _
            "SampleID = '" & RP.SampleID & "' " & _
            "AND Department = '" & RP.Department & "'"
3400  Cnxn(0).Execute sql

End Sub


Function ReasonAffect(ByVal br As BIEResult) As String

      Dim TestName As String
      Dim sql As String
      Dim tb As Recordset
      Dim sn As Recordset

10    On Error GoTo ReasonAffect_Error

20    sql = "SELECT * FROM masks WHERE SampleID = '" & br.SampleID & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    ReasonAffect = ""
60    TestName = Trim(br.LongName)

70    If tb.EOF Then Exit Function

80    sql = "SELECT * FROM biotestdefinitions WHERE code = '" & br.Code & "' and shortname = '" & br.ShortName & "'"
90    Set sn = New Recordset
100   RecOpenServer 0, sn, sql
110   Do While Not sn.EOF
120       If sn!g And tb!g Then
130           ReasonAffect = "Grossly Haemolysed"
140           Exit Do
150       End If
160       If sn!h And tb!h Then
170           ReasonAffect = "Haemolysed"
180           Exit Do
190       End If
200       If sn!s And tb!s Then
210           ReasonAffect = "Slightly Haemolysed"
220           Exit Do
230       End If
240       If sn!l And tb!l Then
250           ReasonAffect = "Lipaemic"
260           Exit Do
270       End If
280       If sn!J And tb!J Then
290           ReasonAffect = "Icteric"
300           Exit Do
310       End If
320       If sn!o And tb!o Then
330           ReasonAffect = "Aged Sample"
340           Exit Do
350       End If
360       sn.MoveNext
370   Loop

380   Exit Function

ReasonAffect_Error:

      Dim strES As String
      Dim intEL As Integer

390   intEL = Erl
400   strES = Err.Description
410   LogError "modBiochemistry", "ReasonAffect", intEL, strES, sql

End Function

Function TestAffected(ByVal br As BIEResult) As Boolean

      Dim TestName As String
      Dim tb As Recordset
      Dim sql As String
      Dim sn As Recordset

10    On Error GoTo TestAffected_Error

20    TestAffected = False
30    TestName = Trim(br.LongName)

40    sql = "SELECT * FROM masks WHERE SampleID = '" & br.SampleID & "'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql

70    If tb.EOF Then Exit Function

80    sql = "SELECT * FROM biotestdefinitions WHERE code = '" & br.Code & "' and shortname = '" & br.ShortName & "'"

90    Set sn = New Recordset
100   RecOpenServer 0, sn, sql
110   Do While Not sn.EOF
120       If sn!h And tb!h Then
130           TestAffected = True
140           Exit Do
150       End If
160       If sn!s And tb!s Then
170           TestAffected = True
180           Exit Do
190       End If
200       If sn!l And tb!l Then
210           TestAffected = True
220           Exit Do
230       End If
240       If sn!o And tb!o Then
250           TestAffected = True
260           Exit Do
270       End If
280       If sn!g And tb!g Then
290           TestAffected = True
300           Exit Do
310       End If
320       If sn!J And tb!J Then
330           TestAffected = True
340           Exit Do
350       End If
360       sn.MoveNext
370   Loop

380   Exit Function

TestAffected_Error:

      Dim strES As String
      Dim intEL As Integer

390   intEL = Erl
400   strES = Err.Description
410   LogError "modBiochemistry", "TestAffected", intEL, strES, sql

End Function

Public Function LongNameforCode(ByVal Code As String) _
       As String

      Dim tb As New Recordset
      Dim sql As String

10    On Error GoTo LongNameforCode_Error

20    LongNameforCode = "???"

30    sql = "SELECT TOP 1 LongName FROM BioTestDefinitions WHERE " & _
            "Code = '" & Code & "'"

40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70        LongNameforCode = Trim(tb!LongName)
80    End If

90    Exit Function

LongNameforCode_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modBiochemistry", "LongNameforCode", intEL, strES, sql

End Function


Public Function Printable(ByVal Code As String) As Boolean

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo Printable_Error

20    Printable = False

30    sql = "SELECT TOP 1 Printable FROM BioTestDefinitions WHERE " & _
            "Code = '" & Code & "' " & _
            "AND InUse = 1"

40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70        Printable = tb!Printable
80        tb.MoveNext
90    End If

100   Exit Function

Printable_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modBiochemistry", "Printable", intEL, strES, sql

End Function

Public Function CheckBioFlag(ByVal Code As String, ByVal Res As String, _
                             ByVal DaysOld As Integer, _
                             ByVal Sex As String, _
                             ByVal DefIndex As Integer, ByVal Fasting As String) As String
      Dim sql As String
      Dim tb As Recordset
      Dim s As String
      Dim StrRes As Double
      Dim X As Integer

10    On Error GoTo CheckBioFlag_Error

20    CheckBioFlag = " "
30    StrRes = Val(Res)

40    X = InStr(Res, "<")
50    If X > 0 Then
60        StrRes = Mid(Res, X + 1)
70    End If

80    X = InStr(Res, ">")
90    If X > 0 Then
100       StrRes = Mid(Res, X + 1)
110   End If

120   If StrRes = 0 Then Exit Function

130   sql = "SELECT * FROM biotestdefinitions WHERE code = '" & Code & "'"
140   If DefIndex <> 0 Then
150       sql = sql & " and defindex = '" & DefIndex & "'"
160   Else
170       sql = sql & " and inuse = 1"
180   End If

190   Set tb = New Recordset
200   RecOpenServer 0, tb, sql
210   Do While Not tb.EOF
220       If DaysOld >= tb!AgeFromDays And DaysOld <= tb!AgeToDays Then
230           If Left(Sex, 1) = "M" Then
240               If StrRes > tb!MaleHigh And tb!MaleHigh <> 0 Then
250                   s = "H"
260               ElseIf StrRes < tb!MaleLow Then
270                   s = "L"
280               End If
290           ElseIf Left(Sex, 1) = "F" Then
300               If StrRes > tb!FemaleHigh And tb!FemaleHigh <> 0 Then
310                   s = "H"
320               ElseIf StrRes < tb!FemaleLow Then
330                   s = "L"
340               End If
350           Else
360               If StrRes > tb!MaleHigh And tb!MaleHigh <> 0 Then
370                   s = "H"
380               ElseIf StrRes < tb!FemaleLow Then
390                   s = "L"
400               End If
410           End If
420       End If
430       If StrRes > Val(tb!PlausibleHigh) Then
440           s = "X"
450       ElseIf StrRes < Val(tb!PlausibleLow) Then
460           s = "X"
470       ElseIf Code = SysOptBioCodeForGlucose(0) Or _
                 Code = SysOptBioCodeForChol(0) Or _
                 Code = SysOptBioCodeForTrig(0) Then
480           If Fasting Then
490               If Code = SysOptBioCodeForGlucose(0) Or Code = SysOptBioCodeForGlucoseP(0) Then
500                   sql = "SELECT * FROM fastings WHERE testname = '" & "GLU" & "'"
510               ElseIf Code = SysOptBioCodeForChol(0) Or Code = SysOptBioCodeForCholP(0) Then
520                   sql = "SELECT * FROM fastings WHERE testname = '" & "CHO" & "'"
530               ElseIf Code = SysOptBioCodeForTrig(0) Or Code = SysOptBioCodeForTrigP(0) Then
540                   sql = "SELECT * FROM fastings WHERE testname = '" & "TRI" & "'"
550               End If
560               Set tb = New Recordset
570               RecOpenServer 0, tb, sql
580               If Not tb.EOF Then
590                   If StrRes > tb!FastingHigh And tb!FastingHigh <> 0 Then
600                       s = "H"
610                   ElseIf StrRes < tb!FastingLow Then
620                       s = "L"
630                   End If
640               End If
650           End If
660       End If
670       tb.MoveNext
680   Loop

690   CheckBioFlag = s

700   Exit Function

CheckBioFlag_Error:

      Dim strES As String
      Dim intEL As Integer

710   intEL = Erl
720   strES = Err.Description
730   LogError "modBiochemistry", "CheckBioFlag", intEL, strES, sql

End Function

Public Function CheckBioNR(ByVal Code As String, _
                           ByVal DaysOld As Integer, _
                           ByVal Sex As String, _
                           ByVal DefIndex As Integer, ByVal Fasting As String) As String
      Dim sql As String
      Dim tb As Recordset
      Dim Nr As String

10    On Error GoTo CheckBioNR_Error

20    sql = "SELECT * FROM biotestdefinitions WHERE code = '" & Code & "'"
30    If DefIndex <> 0 Then
40        sql = sql & " and defindex = '" & DefIndex & "'"
50    Else
60        sql = sql & " and inuse = 1"
70    End If

80    Set tb = New Recordset
90    RecOpenServer 0, tb, sql
100   Do While Not tb.EOF
110       If DaysOld >= tb!AgeFromDays And DaysOld <= tb!AgeToDays Then
120           If Left(Sex, 1) = "M" Then
130               Nr = tb!MaleLow & " - " & tb!MaleHigh
140           ElseIf Left(Sex, 1) = "F" Then
150               Nr = tb!FemaleLow & " - " & tb!FemaleHigh
160           Else
170               Nr = tb!FemaleLow & " - " & tb!MaleHigh
180           End If
190       End If
200       If Code = SysOptBioCodeForGlucose(0) Or _
             Code = SysOptBioCodeForChol(0) Or _
             Code = SysOptBioCodeForTrig(0) Then
210           If Fasting Then
220               If Code = SysOptBioCodeForGlucose(0) Or Code = SysOptBioCodeForGlucoseP(0) Then
230                   sql = "SELECT * FROM fastings WHERE testname = '" & "GLU" & "'"
240               ElseIf Code = SysOptBioCodeForChol(0) Or Code = SysOptBioCodeForCholP(0) Then
250                   sql = "SELECT * FROM fastings WHERE testname = '" & "CHO" & "'"
260               ElseIf Code = SysOptBioCodeForTrig(0) Or Code = SysOptBioCodeForTrigP(0) Then
270                   sql = "SELECT * FROM fastings WHERE testname = '" & "TRI" & "'"
280               End If
290               Set tb = New Recordset
300               RecOpenServer 0, tb, sql
310               If Not tb.EOF Then
320                   Nr = tb!FastingText
330               End If
340           End If
350       End If
360       tb.MoveNext
370   Loop

380   CheckBioNR = Nr

390   Exit Function

CheckBioNR_Error:

      Dim strES As String
      Dim intEL As Integer

400   intEL = Erl
410   strES = Err.Description
420   LogError "modBiochemistry", "CheckBioNR", intEL, strES, sql

End Function



