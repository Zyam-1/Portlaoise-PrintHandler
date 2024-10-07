Attribute VB_Name = "modEndocrinology"
Option Explicit





Sub LogEndAsPrinted(ByVal SampleID As String, _
                    ByVal TestCode As String)

          Dim sql As String

10        On Error GoTo LogEndAsPrinted_Error

20        sql = "update EndResults " & _
                "set valid = 1, printed = 1 WHERE " & _
                "SampleID = '" & RP.SampleID & "' " & _
                "and code = '" & TestCode & "'"
30        Cnxn(0).Execute sql

40        Exit Sub

LogEndAsPrinted_Error:

          Dim strES As String
          Dim intEL As Integer

50        intEL = Erl
60        strES = Err.Description
70        LogError "modEndocrinology", "LogEndAsPrinted", intEL, strES, sql

End Sub

Function EndTestAffected(ByVal br As BIEResult) As Boolean

          Dim TestName As String
          Dim tb As Recordset
          Dim sql As String
          Dim sn As Recordset

10        On Error GoTo EndTestAffected_Error

20        EndTestAffected = False
30        TestName = Trim(br.LongName)

40        sql = "SELECT * FROM EndMasks WHERE SampleID = '" & br.SampleID & "'"
50        Set tb = New Recordset
60        RecOpenServer 0, tb, sql

70        If tb.EOF Then Exit Function

80        sql = "SELECT * FROM EndTestDefinitions WHERE Code = '" & br.Code & "' AND Shortname = '" & br.ShortName & "'"

90        Set sn = New Recordset
100       RecOpenServer 0, sn, sql
110       Do While Not sn.EOF
120           If sn!h And tb!h Then
130               EndTestAffected = True
140               Exit Do
150           End If
160           If sn!s And tb!s Then
170               EndTestAffected = True
180               Exit Do
190           End If
200           If sn!l And tb!l Then
210               EndTestAffected = True
220               Exit Do
230           End If
240           If sn!o And tb!o Then
250               EndTestAffected = True
260               Exit Do
270           End If
280           If sn!g And tb!g Then
290               EndTestAffected = True
300               Exit Do
310           End If
320           If sn!J And tb!J Then
330               EndTestAffected = True
340               Exit Do
350           End If
360           sn.MoveNext
370       Loop

380       Exit Function

EndTestAffected_Error:

          Dim strES As String
          Dim intEL As Integer

390       intEL = Erl
400       strES = Err.Description
410       LogError "modEndocrinology", "EndTestAffected", intEL, strES, sql

End Function


Function EndReasonAffect(ByVal br As BIEResult) As String

          Dim TestName As String
          Dim sql As String
          Dim tb As Recordset
          Dim sn As Recordset

10        On Error GoTo EndReasonAffect_Error

20        sql = "SELECT * FROM EndMasks WHERE SampleID = '" & br.SampleID & "'"
30        Set tb = New Recordset
40        RecOpenServer 0, tb, sql

50        EndReasonAffect = ""
60        TestName = Trim(br.LongName)

70        If tb.EOF Then Exit Function

80        sql = "SELECT * FROM EndMasks WHERE SampleID = '" & br.SampleID & "'"
90        Set sn = New Recordset
100       RecOpenServer 0, sn, sql
110       Do While Not sn.EOF
120           If sn!g And tb!g Then
130               EndReasonAffect = "Grossly Haemolysed"
140               Exit Do
150           End If
160           If sn!h And tb!h Then
170               EndReasonAffect = "Haemolysed"
180               Exit Do
190           End If
200           If sn!s And tb!s Then
210               EndReasonAffect = "Slightly Haemolysed"
220               Exit Do
230           End If
240           If sn!l And tb!l Then
250               EndReasonAffect = "Lipaemic"
260               Exit Do
270           End If
280           If sn!J And tb!J Then
290               EndReasonAffect = "Icteric"
300               Exit Do
310           End If
320           If sn!o And tb!o Then
330               EndReasonAffect = "Old Sample"
340               Exit Do
350           End If
360           sn.MoveNext
370       Loop

380       Exit Function

EndReasonAffect_Error:

          Dim strES As String
          Dim intEL As Integer

390       intEL = Erl
400       strES = Err.Description
410       LogError "modEndocrinology", "EndReasonAffect", intEL, strES, sql

End Function

Public Sub PrintResultEndWin1()

      Dim bc As Recordset
      Dim tb As Recordset
      Dim tbUN As Recordset
      Dim sql As String
      Dim Sex As String
10    ReDim lp(0 To 70) As String
      Dim lpc As Integer
      Dim cUnits As String
      Dim TempUnits As String
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
      Dim DualRep As Boolean
      '    Dim Cx As Comment
      '    Dim Cxs As New Comments
      Dim OB As Observation
      Dim OBS As New Observations
20    ReDim Comments(1 To 4) As String
      Dim SampleDate As String
      Dim Rundate As String
      Dim DoB As String
      Dim RunTime As String
      Dim Fasting As String
      Dim Fx As Fasting
      Dim CodeGLU As String
      Dim CodeCHO As String
      Dim CodeTRI As String
      Dim udtPrintLine(0 To 70) As PrintLine
      Dim strFormat As String
      Dim Cat As String
      Dim copies As Long
      Dim d As Long
      Dim Clin As String
      Dim C As Integer
      Dim f As Integer
      Dim Fontz1 As Integer
      Dim Fontz2 As Integer
      Dim PrintTime As String
      Dim TResult As String
      Dim InconclusiveFound As Boolean
      Dim Analyser As String    'user to detect AxSYM Virology
      Dim CodeCort As String
      Dim AuthorisedBy As String
      Dim PageNumber As String
      Dim TestPerformedAt As String
      Dim ExternalTestingNote As String


30    On Error GoTo PrintResultEndWin1_Error

40    CodeCort = GetOptionSetting("EndCodeForCortisol", "COR")

50    InconclusiveFound = False
60    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
70    DualRep = False

80    For n = 0 To 70
90        udtPrintLine(n).Analyte = ""
100       udtPrintLine(n).Result = ""
110       udtPrintLine(n).Flag = ""
120       udtPrintLine(n).Units = ""
130       udtPrintLine(n).NormalRange = ""
140       udtPrintLine(n).Fasting = ""
150   Next

160   sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
170   Set tb = New Recordset
180   RecOpenClient 0, tb, sql

190   If tb.EOF = False Then
200       Cat = Trim(tb!Category & "")
210   End If

220   If Cat = "" Then Cat = "Default"

230   If IsDate(tb!DoB) Then
240       DoB = Format(tb!DoB, "dd/mmm/yyyy")
250   Else
260       DoB = ""
270   End If

280   ClearUdtHeading
290   With udtHeading
300       .SampleID = RP.SampleID
310       If UCase$(Analyser) = "VIROLOGY" Then
320           .Dept = "Microbiology"
330           .DocumentNo = GetOptionSetting("EndVirologyDocumentNo", "")
340       Else
350           .Dept = "Endocrinology"
360           .DocumentNo = GetOptionSetting("EndMainDocumentNo", "")
370       End If
380       .Name = tb!PatName & ""
390       .Ward = RP.Ward
400       .DoB = DoB
410       .Chart = tb!Chart & ""
420       .Clinician = RP.Clinician
430       .Address0 = tb!Addr0 & ""
440       .Address1 = tb!Addr1 & ""
450       .GP = RP.GP
460       .Sex = tb!Sex & ""
470       .Hospital = tb!Hospital & ""
480       .SampleDate = tb!SampleDate & ""
490       .RecDate = tb!RecDate & ""
500       .Rundate = tb!Rundate & ""
510       .GpClin = Clin
520       .SampleType = SampleType
530       .AandE = tb!AandE & ""

540   End With

550   ResultsPresent = False
560   If pForcePrintTo = "Fax" Then
570       Set BRs = BRs.Load("End", RP.SampleID, "Results", "0", Cat, "")
580   Else
590       Set BRs = BRs.Load("End", RP.SampleID, "Results", "0", Cat, "")
600   End If
610   If Not BRs Is Nothing Then
620       TestCount = BRs.Count
630       If TestCount <> 0 Then
640           ResultsPresent = True
650       End If
660   End If


670   lpc = 0
680   Analyser = ""
      'Analyser = BRs(1).Analyser
690   If ResultsPresent Then
700       Analyser = BRs(1).Analyser
710       If UCase$(Analyser) = "VIROLOGY" Then
              'Virology report
720           For Each br In BRs
730               If AuthorisedBy = "" Then AuthorisedBy = GetAuthorisedBy(br.Operator)
740               If IsNumeric(br.Result) Then
750                   TResult = TranslateEndResultVirology(br.Code, br.Result)
760                   If TResult = "Inconclusive *" Then InconclusiveFound = True
770               Else
780                   TResult = br.Result
790                   If TResult = "Inconclusive *" Then InconclusiveFound = True
800               End If
810               If br.Code = "118" Then
820                   lp(lpc) = FormatString(br.LongName & " (" & br.ShortName & ")", 50, , AlignLeft) & _
                                FormatString(TResult, 15, , AlignLeft) & _
                                FormatString(br.Result, 8, , AlignLeft) & _
                                FormatString(br.Units, 7, , AlignLeft)
830               Else
840                   lp(lpc) = FormatString(br.LongName & " (" & br.ShortName & ")", 50, , AlignLeft) & _
                                FormatString(TResult, 15, , AlignLeft)
850               End If
860               LogTestAsPrinted "End", br.SampleID, br.Code
870               lpc = lpc + 1
880           Next

890       Else
              'Endocrinology Report
900           For Each br In BRs
910               If AuthorisedBy = "" Then AuthorisedBy = GetAuthorisedBy(br.Operator)
920               If Analyser = "" Then
930                   Analyser = Trim$(br.Analyser)
940               End If
950               SampleType = br.SampleType
960               If InStr(1, SampleType, br.SampleType) = 0 Then SampleType = SampleType & " "
                  'If Trim(Br.Operator) <> "" Then RP.Initiator = Br.Operator
970               RunTime = br.RunTime
980               v = br.Result

990               High = Val(br.High)
1000              Low = Val(br.Low)

1010              If Low = 0 Then
1020                  strLow = Low
1030              ElseIf Low < 10 Then
1040                  strLow = Format(Low, "0.00")
1050              ElseIf Low < 100 Then
1060                  strLow = Format(Low, "##.0")
1070              Else
1080                  strLow = Format(Low, " ###")
1090              End If

1100              If High < 10 Then
1110                  strHigh = Format(High, "0.00")
1120              ElseIf High < 100 Then
1130                  strHigh = Format(High, "##.0")
1140              Else
1150                  strHigh = Format(High, "### ")
1160              End If
1170              Flag = "  "
1180              sql = "SELECT * FROM endtestdefinitions WHERE code = '" & br.Code & "'"
1190              Set bc = New Recordset
1200              RecOpenServer 0, bc, sql
1210              If Not bc.EOF Then
1220                  If Trim(bc!forfert & "") <> 1 And udtHeading.Sex <> "" And IsDate(udtHeading.DoB) Then
1230                      If IsNumeric(v) Then
1240                          If Val(v) > br.PlausibleHigh Then
1250                              udtPrintLine(lpc).Flag = " X "
1260                              lp(lpc) = "  "
1270                              Flag = " X"
1280                          ElseIf Val(v) < br.PlausibleLow Then
1290                              udtPrintLine(lpc).Flag = " X "
1300                              lp(lpc) = "  "
1310                              Flag = " X"
1320                          ElseIf Val(v) > High And Val(High) > 0 Then
1330                              udtPrintLine(lpc).Flag = " H "
1340                              lp(lpc) = "  "    'bold
1350                              Flag = " H"
1360                          ElseIf Val(v) < Low Then
1370                              udtPrintLine(lpc).Flag = " L "
1380                              lp(lpc) = "  "    'bold
1390                              Flag = " L"
1400                          Else
1410                              udtPrintLine(lpc).Flag = "   "
1420                              lp(lpc) = "  "    'unbold
1430                              Flag = "  "
1440                          End If
1450                      Else
1460                          If Left(v, 1) = "<" Then
1470                              udtPrintLine(lpc).Flag = " L "
1480                              lp(lpc) = "  "    'bold
1490                              Flag = " L"
1500                          ElseIf Left(v, 1) = ">" Then
1510                              udtPrintLine(lpc).Flag = " H "
1520                              lp(lpc) = "  "    'bold
1530                              Flag = " H"
1540                          Else
1550                              udtPrintLine(lpc).Flag = "   "
1560                              lp(lpc) = "  "    'unbold
1570                              Flag = "  "
1580                          End If
1590                      End If
1600                  Else
1610                      lp(lpc) = "  "    'unbold
1620                  End If
1630              Else
1640                  lp(lpc) = "  "    'unbold
1650              End If
1660              lp(lpc) = lp(lpc) & "    "





1670              TestPerformedAt = ""
1680              If UCase(HospName(0)) <> UCase(br.Hospital) Then
1690                  TestPerformedAt = Left(UCase(br.Hospital), 1)
1700                  If InStr(ExternalTestingNote, UCase(br.Hospital)) = 0 Then
1710                      ExternalTestingNote = ExternalTestingNote & TestPerformedAt & " = Test Analysed at " & UCase(br.Hospital) & " "
1720                  End If

1730                  TestPerformedAt = "(" & TestPerformedAt & ")"
1740              End If



1750              lp(lpc) = lp(lpc) & Left(br.LongName & TestPerformedAt & Space(26), 26)
1760              udtPrintLine(lpc).Analyte = Left(br.LongName & Space(26), 26)

1770              If EndTestAffected(br) = False Then
1780                  If IsNumeric(v) Then
1790                      Select Case br.Printformat
                          Case 0: strFormat = "#########"
1800                      Case 1: strFormat = "######0.0"
1810                      Case 2: strFormat = "#####0.00"
1820                      Case 3: strFormat = "####0.000"
1830                      End Select
1840                      lp(lpc) = lp(lpc) & " " & Right(Space(9) & Format(v, strFormat), 9)
1850                      udtPrintLine(lpc).Result = Format(v, strFormat)
1860                  Else
1870                      lp(lpc) = lp(lpc) & " " & Right(Space(9) & v, 9)
1880                      udtPrintLine(lpc).Result = v
1890                  End If
1900              Else
1910                  lp(lpc) = lp(lpc) & " " & Right(Space(9) & "XXXXXXX", 9)
1920              End If
1930              lp(lpc) = lp(lpc) & Flag & " "

1940              sql = "SELECT * FROM Lists WHERE " & _
                        "ListType = 'UN' and Code = '" & br.Units & "'"
1950              Set tbUN = Cnxn(0).Execute(sql)
1960              If Not tbUN.EOF Then
1970                  cUnits = Left(tbUN!Text & Space(8), 8)
1980              Else
1990                  cUnits = Left(br.Units & Space(8), 8)
2000              End If
2010              udtPrintLine(lpc).Units = cUnits
2020              lp(lpc) = lp(lpc) & " " & Right(Space(6) & cUnits, 8)
2030              If udtHeading.Sex = "" Or (Not IsDate(udtHeading.DoB)) Then
2040                  lp(lpc) = lp(lpc) & " "
2050              Else
2060                  If UCase(br.Code) = UCase(CodeCort) Then
2070                      lp(lpc) = lp(lpc) & "   (100-500) am Sample"
2080                      If br.Comment <> "" Then
2090                          lp(lpc) = lp(lpc) & " " & br.Comment
2100                      End If
2110                      lpc = lpc + 1
2120                      lp(lpc) = lp(lpc) & "                                                         (72 - 371) pm Sample"
2130                  ElseIf UCase(br.Code) = UCase(SysOptEndCodeHBA1C(0)) Then
                          'lp(lpc) = lp(lpc) & "   (4.9 - 5.9 %) Normal (Non-diabetic)"
2140                      lp(lpc) = lp(lpc) & "   " & SysOptEndHbA1cComment(0)
2150                      If br.Comment <> "" Then
2160                          lp(lpc) = lp(lpc) & " " & br.Comment
2170                      End If
2180                      lpc = lpc + 1
2190                      If UCase(HospName(0)) = "MULLINGAR" Then
2200                          lp(lpc) = lp(lpc) & "                                                     Diabetic Goal < 7 %"
2210                      End If
2220                  ElseIf UCase(br.Code) = UCase(SysOptEndCodeCalcA1C(0)) Then
                          'lp(lpc) = lp(lpc) & "   (30 - 41 mmol/mol) Normal (Non-diabetic)"
2230                      lp(lpc) = lp(lpc) & "   " & SysOptEndCalcA1cComment(0)
2240                      If br.Comment <> "" Then
2250                          lp(lpc) = lp(lpc) & " " & br.Comment
2260                      End If
2270                      lpc = lpc + 1
2280                      If UCase(HospName(0)) = "MULLINGAR" Then
2290                          lp(lpc) = lp(lpc) & "                                                     Diabetic Goal < 53 mmol/mol"
2300                      End If
2310                      DualRep = False
2320                  ElseIf UCase(br.Code) = UCase(SysOptEndCodeBNP(0)) Then
2330                      lp(lpc) = lp(lpc) & "   (< 50) Normal"
2340                      If br.Comment <> "" Then
2350                          lp(lpc) = lp(lpc) & " " & br.Comment
2360                      End If
2370                      lpc = lpc + 1
2380                      lp(lpc) = lp(lpc) & "                                                         (50 - 100) Equivocal"
2390                      lpc = lpc + 1
2400                      lp(lpc) = lp(lpc) & "                                                         (> 100) Abnormal"
2410                  ElseIf UCase(br.Code = SysOptEndCodeB12(0)) Then
2420                      lp(lpc) = lp(lpc) & "   (< 160) Deficient"
2430                      If br.Comment <> "" Then
2440                          lp(lpc) = lp(lpc) & " " & br.Comment
2450                      End If
2460                      lpc = lpc + 1
2470                      lp(lpc) = lp(lpc) & "                                                         (160 - 239) Indeterminate"
2480                      lpc = lpc + 1
2490                      lp(lpc) = lp(lpc) & "                                                         (240 - 911) Normal"
2500                  ElseIf UCase(br.Code) = UCase(SysOptEndCodeB12New(0)) Then
2510                      lp(lpc) = lp(lpc) & "   (< 140) Deficient"
2520                      If br.Comment <> "" Then
2530                          lp(lpc) = lp(lpc) & " " & br.Comment
2540                      End If
2550                      lpc = lpc + 1
2560                      lp(lpc) = lp(lpc) & "                                                         (140 - 699) Normal"
2570                  ElseIf UCase(br.Code) = UCase(SysOptEndCodeCo(0)) Then
2580                      lp(lpc) = lp(lpc) & "   (185-624) AM Sample"
2590                      If br.Comment <> "" Then
2600                          lp(lpc) = lp(lpc) & " " & br.Comment
2610                      End If
2620                      lpc = lpc + 1
2630                      lp(lpc) = lp(lpc) & "                                                         (< 276) PM Sample"



2640                  ElseIf UCase(br.Code) = UCase(SysOptEndCodeVITD(0)) Then
2650                      lp(lpc) = lp(lpc) & "    Sufficient: 50 nmol/l"
2660                      If br.Comment <> "" Then
2670                          lp(lpc) = lp(lpc) & " " & br.Comment
2680                      End If
                          '                    lpc = lpc + 1
                          '                    lp(lpc) = lp(lpc) & "                                                          Sufficient: 50 nmol/l"
2690                      lpc = lpc + 1
2700                      lp(lpc) = lp(lpc) & "                                                          Insufficient: <30 nmol/l"
2710                      lpc = lpc + 1
2720                      lp(lpc) = lp(lpc) & "                                                          Upper Safty Limit: >125 nmol/l"





2730                  ElseIf UCase(br.Code) = UCase(SysOptEndCodeFSH(0)) And udtHeading.Sex = "F" Then
2740                      lp(lpc) = lp(lpc) & "   (3.9-8.8) Follicular"
2750                      If br.Comment <> "" Then
2760                          lp(lpc) = lp(lpc) & " " & br.Comment
2770                      End If
2780                      lpc = lpc + 1
2790                      lp(lpc) = lp(lpc) & "                                                         (4.5-22.5) Mid-Cycle"
2800                      lpc = lpc + 1
2810                      lp(lpc) = lp(lpc) & "                                                         (1.8-5.1) Luteal"
2820                      lpc = lpc + 1
2830                      lp(lpc) = lp(lpc) & "                                                         (16.7-114) Menopause"
2840                  ElseIf UCase(br.Code) = UCase(SysOptEndCodeLH(0)) And udtHeading.Sex = "F" Then
2850                      lp(lpc) = lp(lpc) & "   (2.1-10.8) Follicular"
2860                      If br.Comment <> "" Then
2870                          lp(lpc) = lp(lpc) & " " & br.Comment
2880                      End If
2890                      lpc = lpc + 1
2900                      lp(lpc) = lp(lpc) & "                                                         (19.1-103) Mid-Cycle"
2910                      lpc = lpc + 1
2920                      lp(lpc) = lp(lpc) & "                                                         (1.2-12.8) Luteal"
2930                      lpc = lpc + 1
2940                      lp(lpc) = lp(lpc) & "                                                         (10.8-58.6) Menopause"
2950                  ElseIf UCase(br.Code) = UCase(SysOptEndCodePRO(0)) And udtHeading.Sex = "F" Then
2960                      lp(lpc) = lp(lpc) & "   (0.98-4.83) Follicular"
2970                      If br.Comment <> "" Then
2980                          lp(lpc) = lp(lpc) & " " & br.Comment
2990                      End If
3000                      lpc = lpc + 1
3010                      lp(lpc) = lp(lpc) & "                                                         (16.40-95.76) Luteal"
3020                      lpc = lpc + 1
3030                      lp(lpc) = lp(lpc) & "                                                         (0.25-2.48) Menopause"
3040                      lpc = lpc + 1
3050                      lp(lpc) = lp(lpc) & "                                                         (15-161) Pregnancy 1St Trimester"
3060                      lpc = lpc + 1
3070                      lp(lpc) = lp(lpc) & "                                                         (74-144) Pregnancy 2nd Trimester"
3080                  ElseIf UCase(br.Code) = UCase(SysOptEndCodeOES(0)) And udtHeading.Sex = "F" Then
3090                      lp(lpc) = lp(lpc) & "   (92-422) Follicular"
3100                      If br.Comment <> "" Then
3110                          lp(lpc) = lp(lpc) & " " & br.Comment
3120                      End If
3130                      lpc = lpc + 1
3140                      lp(lpc) = lp(lpc) & "                                                         (118-1898) Mid-Cycle"
3150                      lpc = lpc + 1
3160                      lp(lpc) = lp(lpc) & "                                                         (134-903) Luteal"
3170                      lpc = lpc + 1
3180                      lp(lpc) = lp(lpc) & "                                                         (55-92) Menopause"
3190                  ElseIf UCase(br.Code) = UCase(SysOptEndCodePRL(0)) And udtHeading.Sex = "F" Then
3200                      lp(lpc) = lp(lpc) & "   (71-567) Premenopausal"
3210                      If br.Comment <> "" Then
3220                          lp(lpc) = lp(lpc) & " " & br.Comment
3230                      End If
3240                      lpc = lpc + 1
3250                      lp(lpc) = lp(lpc) & "                                                         (58-416) Postmenopausal"
3260                  ElseIf br.Code = SysOptEndCodeTHC(0) And udtHeading.Sex = "F" Then
3270                      lp(lpc) = lp(lpc) & "   (<5.0) Non-Pregnant"
3280                      If br.Comment <> "" Then
3290                          lp(lpc) = lp(lpc) & " " & br.Comment
3300                      End If
3310                      lpc = lpc + 1
3320                      lp(lpc) = lp(lpc) & "                                                         (5-50) 1 Week"
3330                      lpc = lpc + 1
3340                      lp(lpc) = lp(lpc) & "                                                         (62.5-625) 1-2 Weeks"
3350                      lpc = lpc + 1
3360                      lp(lpc) = lp(lpc) & "                                                         (125-6250) 2-3 Weeks"
3370                      lpc = lpc + 1
3380                      lp(lpc) = lp(lpc) & "                                                         (625-12500) 3-4 Weeks"
3390                      lpc = lpc + 1
3400                      lp(lpc) = lp(lpc) & "                                                         (1250-62500) 4-5 Weeks"
3410                      lpc = lpc + 1
3420                      lp(lpc) = lp(lpc) & "                                                         (12500-125000) 5-6 Weeks"
3430                      lpc = lpc + 1
3440                      lp(lpc) = lp(lpc) & "                                                         (18750-250000) 6-8 Weeks"
3450                      lpc = lpc + 1
3460                      lp(lpc) = lp(lpc) & "                                                         (12500-125000) 8-12 Weeks"


3470                  ElseIf UCase(br.Code) = UCase(SysOptEndCodeTRO(0)) Then
3480                      lp(lpc) = lp(lpc) & "   (<0.4) Normal"
3490                      If br.Comment <> "" Then
3500                          lp(lpc) = lp(lpc) & " " & br.Comment
3510                      End If
3520                      lpc = lpc + 1
3530                      lp(lpc) = lp(lpc) & "                                                         (0.04) Cut Off for AMI"
3540                  Else
3550                      If Left(Trim(tb!Sex), 1) = "F" Then
3560                          If UCase(HospName(0)) = "MULLINGAR" Then
3570                              sql = "SELECT  * FROM endtestdefinitions WHERE code = '" & br.Code & "' and forFert = 1 order by printpriority, category"
3580                          Else
3590                              sql = "SELECT  * FROM endtestdefinitions WHERE code = '" & br.Code & "' and forFert = 1 and category <> 'Default' order by category"
3600                          End If
3610                          Set bc = New Recordset
3620                          RecOpenClient 0, bc, sql
3630                          If Not bc.EOF Then
3640                              Do While Not bc.EOF
3650                                  n = 1
3660                                  If Left(Trim(tb!Sex), 1) = "F" Then
3670                                      If (Not IsNull(bc!ShowLessThan) And bc!ShowLessThan <> 0) _
                                             And Val(bc!FemaleLow) = 0 And (Val(bc!FemaleHigh) <> 999 Or Val(bc!FemaleHigh) <> 9999) Then
3680                                          lp(lpc) = lp(lpc) & "    " & " <" & Format(bc!FemaleHigh, strFormat) & " "
3690                                      ElseIf (Not IsNull(bc!ShowMoreThan) And bc!ShowMoreThan <> 0) _
                                                 And Val(bc!FemaleLow) <> 0 And (Val(bc!FemaleHigh) = 999 Or Val(bc!FemaleHigh) = 9999) Then
3700                                          lp(lpc) = lp(lpc) & "    " & " >" & Format(bc!FemaleLow, strFormat) & " "
3710                                      Else
3720                                          lp(lpc) = lp(lpc) & "   (" & Format(bc!FemaleLow, strFormat) & "-" & Format(bc!FemaleHigh, strFormat) & ") "
3730                                      End If
3740                                  End If
3750                                  If n < Val(bc.RecordCount) Then
3760                                      If Trim(bc!Category) <> "Default" Then
3770                                          lp(lpc) = lp(lpc) & Right(bc!Category, 16)
3780                                          If br.LongName = "ThCG" And UCase(HospName(0)) = "MULLINGAR" Then lp(lpc) = lp(lpc) & " Days"
3790                                      ElseIf br.LongName = "ThCG" And UCase(HospName(0)) = "MULLINGAR" Then
3800                                          If br.LongName = "ThCG" Then
3810                                              lp(lpc) = lp(lpc) & "< 7 Days"
3820                                          End If
3830                                      End If
3840                                      If Trim(Left(lp(lpc), 14)) <> "" Then
3850                                          If EndTestAffected(br) = True Then
3860                                              lp(lpc) = lp(lpc) & " " & EndReasonAffect(br)
3870                                          End If
3880                                          If br.Comment <> "" Then
3890                                              lp(lpc) = lp(lpc) & " " & br.Comment
3900                                          End If
3910                                      End If
3920                                      lpc = lpc + 1
3930                                      lp(lpc) = "                                          "
3940                                      udtPrintLine(lpc).NormalRange = udtPrintLine(lpc).NormalRange & vbCrLf
3950                                  Else
3960                                      lpc = lpc + 1
3970                                      lp(lpc) = "@"
3980                                  End If
3990                                  bc.MoveNext
4000                              Loop
4010                          Else
4020                              If Val(strHigh) <> 0 Then
                                      'Zyam
4030                                  lp(lpc) = lp(lpc) & IIf(Val(strLow) = 0 And Val(strHigh) = 0, " ", "   (")
4040                                  lp(lpc) = lp(lpc) & IIf(Val(strLow) = 0 And Val(strHigh) = 0, " ", Format(strLow, strFormat) & "-")
4050                                  lp(lpc) = lp(lpc) & IIf(Val(strLow) = 0 And Val(strHigh) = 0, " ", Format(strHigh, strFormat) & ")")
                                      'Zyam
4060                                  If EndTestAffected(br) = True Then
4070                                      lp(lpc) = lp(lpc) & " " & EndReasonAffect(br)
4080                                  End If
4090                                  If br.Comment <> "" Then
4100                                      lp(lpc) = lp(lpc) & " " & br.Comment
4110                                  End If
4120                              End If
                                  '11-15-23 Zyam
                                  If Val(strLow) = 0 And Val(strHigh) = 0 Then
4130                                udtPrintLine(lpc).NormalRange = " "
                                  Else
                                    udtPrintLine(lpc).NormalRange = "(" & strLow & "-" & strHigh & ")"
                                  End If
                                  '11-15-23 Zyam
4140                          End If
4150                      Else
4160                          If Val(strHigh) <> 0 Then
4170                              If (Not IsNull(bc!ShowLessThan) And bc!ShowLessThan <> 0) _
                                     And Val(strLow) = 0 And (Val(strHigh) <> 999 Or Val(strHigh) <> 9999) Then
4180                                  lp(lpc) = lp(lpc) & "    "
4190                                  lp(lpc) = lp(lpc) & " <"
4200                                  lp(lpc) = lp(lpc) & Format(strHigh, strFormat) & " "
4210                              ElseIf (Not IsNull(bc!ShowMoreThan) And bc!ShowMoreThan <> 0) _
                                         And Val(strLow) <> 0 And (Val(strHigh) = 999 Or Val(strHigh) = 9999) Then
4220                                  lp(lpc) = lp(lpc) & "    "
4230                                  lp(lpc) = lp(lpc) & " >"
4240                                  lp(lpc) = lp(lpc) & Format(strLow, strFormat) & " "
4250                              Else
                                      'Zyam
4260                                  lp(lpc) = lp(lpc) & IIf(Val(strLow) = 0 And Val(strHigh) = 0, " ", "   (")
4270                                  lp(lpc) = lp(lpc) & IIf(Val(strLow) = 0 And Val(strHigh) = 0, " ", Format(strLow, strFormat) & "-")
4280                                  lp(lpc) = lp(lpc) & IIf(Val(strLow) = 0 And Val(strHigh) = 0, " ", Format(strHigh, strFormat) & ")")
                                      'Zyam
4290                              End If
4300                              If EndTestAffected(br) = True Then
4310                                  lp(lpc) = lp(lpc) & " " & EndReasonAffect(br)
4320                              End If
4330                              If br.Comment <> "" Then
4340                                  lp(lpc) = lp(lpc) & " " & br.Comment
4350                              End If
4360                          End If
                              '11-15-23 Zyam
                                  If Val(strLow) = 0 And Val(strHigh) = 0 Then
4131                                udtPrintLine(lpc).NormalRange = " "
                                  Else
                                    udtPrintLine(lpc).NormalRange = "(" & Format(strLow, strFormat) & "-" & Format(strHigh, strFormat) & ")"
                                  End If
                                  '11-15-23 Zyam
4370
4380                      End If
4390                  End If
4400              End If
4410              LogTestAsPrinted "End", br.SampleID, br.Code
4420              lpc = lpc + 1
4430          Next
4440      End If
4450  End If



4460  frmRichText.rtb.Text = ""
4470  If lpc > Val(frmMain.txtMoreThan) Then
4480      PageNumber = "Page 1 of 2"
4490  Else
4500      PageNumber = "Page 1 of 1"
4510  End If
4520  If RP.FaxNumber <> "" Then
4530      PrintHeadingRTBFax
4540  Else
4550      PrintHeadingRTB (PageNumber)
4560  End If

4570  Sex = tb!Sex & ""

4580  If RP.FaxNumber <> "" Then
4590      Fontz1 = 9
4600      Fontz2 = 12
4610  Else
4620      Fontz1 = 10
4630      Fontz2 = 14
4640  End If

4650  With frmRichText.rtb
4660      .SelFontSize = Fontz1
          '    If lpc > Val(frmMain.txtMoreThan) Then
          '        .SelText = Space(35) & "Page 1 of 2" & vbCrLf
          '    Else
          '        .SelText = Space(35) & "Page 1 of 1" & vbCrLf
          '    End If
          '    .SelText = vbCrLf
4670      CrCnt = CrCnt + 1

4680      If SysOptPrintMiddle(0) And Analyser <> "X" Then
4690          n = Row_Count(lpc)
4700          For C = 1 To n
4710              .SelText = vbCrLf
4720              CrCnt = CrCnt + 1
4730          Next
4740      End If

4750      .SelFontSize = Fontz1

4760      If UCase$(Analyser) = "VIROLOGY" Then
4770          PrintTextRTB frmRichText.rtb, Space(3)
4780          PrintTextRTB frmRichText.rtb, String(407, "-"), 2
4790          PrintTextRTB frmRichText.rtb, vbCrLf, 2
4800          PrintTextRTB frmRichText.rtb, Space(3)
4810          PrintTextRTB frmRichText.rtb, FormatString("Test", 50, , AlignLeft), 10, True
4820          PrintTextRTB frmRichText.rtb, FormatString("Result", 15, , AlignLeft), 10, True
4830          PrintTextRTB frmRichText.rtb, FormatString("Value", 8, , AlignLeft), 10, True
4840          PrintTextRTB frmRichText.rtb, FormatString("Unit", 7, , AlignLeft), 10, True
4850          PrintTextRTB frmRichText.rtb, vbCrLf, 2
4860          PrintTextRTB frmRichText.rtb, Space(3)
4870          PrintTextRTB frmRichText.rtb, String(407, "-"), 2
4880          PrintTextRTB frmRichText.rtb, vbCrLf
4890          CrCnt = CrCnt + 3
4900      End If
4910      For n = 0 To 19
4920          .SelFontSize = Fontz1
4930          If Trim(lp(n)) <> "" Then
4940              If Left(lp(n), 1) = "@" Then lp(n) = ""
4950              If Trim$(udtPrintLine(n).Analyte) = EndLongNameFor(SysOptEndCodeB12(0)) And InStr(lp(n), " H ") Then
4960                  .SelColor = vbBlack
4970                  .SelText = "   " & lp(n)
4980                  .SelText = vbCrLf
4990                  CrCnt = CrCnt + 1
5000              Else
5010                  If InStr(lp(n), " L ") Or InStr(lp(n), " H ") Then
5020                      .SelColor = vbBlack
5030                      .SelBold = True
5040                      .SelText = "   " & Left(lp(n), 33)
5050                      .SelText = Mid(lp(n), 34, 3)
5060                      .SelText = Mid(lp(n), 37)
5070                      .SelBold = False
5080                      .SelText = vbCrLf
5090                      CrCnt = CrCnt + 1
5100                  Else
5110                      .SelColor = vbBlack
5120                      .SelText = "   " & lp(n)
5130                      .SelText = vbCrLf
5140                      CrCnt = CrCnt + 1
5150                  End If
5160              End If

5170          End If
5180      Next
5190      If UCase$(Analyser) = "VIROLOGY" And InconclusiveFound = True Then
5200          PrintTextRTB frmRichText.rtb, vbCrLf
5210          PrintTextRTB frmRichText.rtb, Space(3)
5220          PrintTextRTB frmRichText.rtb, "Comment: Inconclusive * = Sample must be sent to the VRL for further testing. Result not confirmed.", 8
5230          CrCnt = CrCnt + 2
5240      End If

5250      Do While CrCnt < 28
5260          .SelFontSize = Fontz1
5270          .SelText = vbCrLf
5280          CrCnt = CrCnt + 1
5290      Loop

          '        Set Cx = Cxs.Load(RP.SampleID)
5300      Set OBS = OBS.Load(RP.SampleID, "Endocrinology", "Demographic")
5310      If Not OBS Is Nothing Then
5320          For Each OB In OBS
5330              Select Case UCase$(OB.Discipline)
                  Case "ENDOCRINOLOGY"
5340                  FillCommentLines OB.Comment, 4, Comments(), 97
5350                  For n = 1 To 4
5360                      If Trim(Comments(n)) <> "" Then
5370                          .SelFontSize = Fontz1
5380                          .SelText = "     " & Comments(n) & vbCrLf
5390                          CrCnt = CrCnt + 1
5400                      End If
5410                  Next
5420              Case "DEMOGRAPHIC"
5430                  FillCommentLines OB.Comment, 2, Comments(), 97
5440                  For n = 1 To 2
5450                      If Trim(Comments(n)) <> "" Then
5460                          .SelFontSize = Fontz1
5470                          .SelText = "     " & Comments(n) & vbCrLf
5480                          CrCnt = CrCnt + 1
5490                      End If
5500                  Next
5510              End Select
5520          Next
5530      End If

5540      .SelFontSize = Fontz1
5550      If Not IsDate(DoB) Or Trim(udtHeading.Sex) = "" Then
5560          .SelColor = vbBlack
5570          .SelText = "**** No Sex/DoB given. No reference range applied ****" & vbCrLf
5580          .SelText = vbCrLf
5590          CrCnt = CrCnt + 1
              '    ElseIf Not IsDate(DoB) Then
              '        .SelColor = vbBlack
              '        .SelText = "*** No Dob. Adult Age 25 used for Normal Ranges! ***" & vbCrLf
              '        .SelText = vbCrLf
              '        CrCnt = CrCnt + 1
              '    ElseIf Trim(udtHeading.Sex) = "" Then
              '        .SelColor = vbBlack
              '        .SelText = "No Sex given. No reference range applied" & vbCrLf
              '        .SelText = vbCrLf
              '        CrCnt = CrCnt + 1
5600      End If

5610      If DualRep = True Then
5620          .SelColor = vbBlack
5630          .SelText = "Please note dual reporting commenced on 1st July 2010 and will cease on 31st December 2011." & vbCrLf
5640          .SelText = vbCrLf
5650          CrCnt = CrCnt + 1
5660      End If
5670      .SelColor = vbBlack

5680      If IsDate(tb!SampleDate) Then
5690          SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
5700      Else
5710          SampleDate = ""
5720      End If
5730      If IsDate(RunTime) Then
5740          Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
5750      Else
5760          If IsDate(tb!Rundate) Then
5770              Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
5780          Else
5790              Rundate = ""
5800          End If
5810      End If

5820      If RP.FaxNumber <> "" Then
              '5830          PrintFooterRTBFax RP.Initiator, SampleDate, Rundate
5830          If UCase(GetOptionSetting("GetLatestAuthorisedBy", "")) = UCase("True") Then
5840              PrintFooterRTB GetLatestAuthorisedBy("End", RP.SampleID), SampleDate, GetLatestRunDateTime("End", RP.SampleID, Rundate)
5850          Else
5860              PrintFooterRTB AuthorisedBy, SampleDate, GetLatestRunDateTime("End", RP.SampleID, Rundate)
5870          End If
5880          f = FreeFile
5890          Open SysOptFax(0) & RP.SampleID & "END1.doc" For Output As f
5900          Print #f, .TextRTF
5910          Close f
5920          SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "END1.doc"
5930      Else
              '5900          PrintFooterRTB AuthorisedBy, SampleDate, Rundate, ExternalTestingNote
5940          If UCase(GetOptionSetting("GetLatestAuthorisedBy", "")) = UCase("True") Then
5950              PrintFooterRTB GetLatestAuthorisedBy("End", RP.SampleID), SampleDate, GetLatestRunDateTime("End", RP.SampleID, Rundate), ExternalTestingNote
5960          Else
5970              PrintFooterRTB AuthorisedBy, SampleDate, GetLatestRunDateTime("End", RP.SampleID, Rundate), ExternalTestingNote
5980          End If
5990          .SelStart = 0
              'Do not print if Doctor is disabled in DisablePrinting
              '*******************************************************************
6000          If CheckDisablePrinting(RP.Ward, "Endocrinology") Then

6010          ElseIf CheckDisablePrinting(RP.GP, "Endocrinology") Then
6020          Else
6030              .SelPrint Printer.hdc
6040          End If
              '*******************************************************************
              '.SelPrint Printer.hDC
6050      End If
6060      sql = "SELECT * FROM Reports WHERE 0 = 1"
6070      Set tb = New Recordset
6080      RecOpenServer 0, tb, sql
6090      tb.AddNew
6100      tb!SampleID = RP.SampleID
6110      tb!Name = udtHeading.Name
6120      tb!Dept = "E"
6130      tb!Initiator = RP.Initiator
6140      tb!PrintTime = PrintTime
6150      tb!RepNo = "0E" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
6160      tb!PageNumber = 0
6170      tb!Report = .TextRTF
6180      tb!Printer = Printer.DeviceName
6190      tb.Update
6200  End With


      '##########################
      'Print Second Page
6210  If lpc > 20 Then
6220      With frmRichText.rtb
6230          If RP.FaxNumber <> "" Then
6240              PrintHeadingRTBFax
6250          Else
6260              PrintHeadingRTB ("Page 2 of 2")
6270          End If

              '        .SelFontSize = Fontz1
              '        .SelText = Space(35) & "Page 2 of 2"
              '        .SelText = vbCrLf
              '        .SelText = vbCrLf
6280          CrCnt = CrCnt + 1

6290          For n = 20 To 35
6300              .SelFontSize = Fontz1
6310              If Trim(lp(n)) <> "" Then
6320                  If Left(lp(n), 1) = "@" Then lp(n) = ""
6330                  If InStr(lp(n), " L ") Or InStr(lp(n), " H ") Then
6340                      .SelColor = vbBlack
6350                      .SelBold = True
6360                      .SelText = "   " & Left(lp(n), 33)
6370                      .SelText = Mid(lp(n), 34, 3)
6380                      .SelText = Mid(lp(n), 37)
6390                      .SelBold = False
6400                      .SelText = vbCrLf
6410                      CrCnt = CrCnt + 1
6420                  Else
6430                      .SelColor = vbBlack
6440                      .SelText = "   " & lp(n)
6450                      .SelText = vbCrLf
6460                      CrCnt = CrCnt + 1
6470                  End If
6480              End If
6490          Next

6500          Do While CrCnt < 28
6510              .SelFontSize = Fontz1
6520              .SelText = vbCrLf
6530              CrCnt = CrCnt + 1
6540          Loop
              '            Set Cx = Cxs.Load(RP.SampleID)
6550          Set OBS = OBS.Load(RP.SampleID, "Endocrinology", "Demographic")
6560          If Not OBS Is Nothing Then
6570              For Each OB In OBS
6580                  Select Case UCase$(OB.Discipline)
                      Case "ENDOCRINOLOGY"
6590                      FillCommentLines OB.Comment, 4, Comments(), 97
6600                      For n = 1 To 4
6610                          .SelFontSize = Fontz1
6620                          If Trim(Comments(n)) <> "" Then
6630                              .SelText = "     " & Comments(n) & vbCrLf
6640                          End If
6650                          CrCnt = CrCnt + 1
6660                      Next
6670                  Case "DEMOGRAPHIC"
6680                      FillCommentLines OB.Comment, 2, Comments(), 97
6690                      For n = 1 To 2
6700                          .SelFontSize = Fontz1
6710                          If Trim(Comments(n)) <> "" Then .SelText = "     " & Comments(n) & vbCrLf
6720                          CrCnt = CrCnt + 1
6730                      Next
6740                  End Select
6750              Next
6760          End If

6770          .SelFontSize = Fontz1
6780          If Not IsDate(DoB) Or Trim(udtHeading.Sex) = "" Then
6790              .SelColor = vbBlack
6800              .SelText = "**** No Sex/DoB given. No reference range applied! ****" & vbCrLf
6810              CrCnt = CrCnt + 1
                  '        ElseIf Not IsDate(DoB) Then
                  '            .SelColor = vbBlack
                  '            .SelText = "*** No Dob. Adult Age 25 used for Normal Ranges! ***" & vbCrLf
                  '            .SelText = vbCrLf
                  '            CrCnt = CrCnt + 1
                  '        ElseIf Trim(udtHeading.Sex) = "" Then
                  '            .SelColor = vbBlack
                  '            .SelText = "No Sex given. No Reference range applied" & vbCrLf
                  '            .SelText = vbCrLf
                  '            CrCnt = CrCnt + 1
6820          End If
6830          If DualRep = True Then
6840              .SelColor = vbBlack
6850              .SelText = "Please note dual reporting commenced on 1st July 2010 and will cease on 31st December 2011." & vbCrLf
6860              .SelText = vbCrLf
6870              CrCnt = CrCnt + 1
6880          End If
6890          .SelColor = vbBlack

6900          If RP.FaxNumber <> "" Then
6910              PrintFooterRTBFax RP.Initiator, SampleDate, Rundate
6920              .SelStart = 0
6930              f = FreeFile
6940              Open SysOptFax(0) & RP.SampleID & "END2.doc" For Output As f
6950              Print #f, .TextRTF
6960              Close f
6970              SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "END2.doc"
6980          Else
                  '6970        PrintFooterRTB AuthorisedBy, SampleDate, Rundate
6990              If UCase(GetOptionSetting("GetLatestAuthorisedBy", "")) = UCase("True") Then
7000                  PrintFooterRTB GetLatestAuthorisedBy("End", RP.SampleID), SampleDate, GetLatestRunDateTime("End", RP.SampleID, Rundate)
7010              Else
7020                  PrintFooterRTB AuthorisedBy, SampleDate, GetLatestRunDateTime("End", RP.SampleID, Rundate)
7030              End If
7040              .SelStart = 0

                  'Do not print if Doctor is disabled in DisablePrinting
                  '*******************************************************************
7050              If CheckDisablePrinting(RP.Ward, "Endocrinology") Then

7060              ElseIf CheckDisablePrinting(RP.GP, "Endocrinology") Then
7070              Else
7080                  .SelPrint Printer.hdc
7090              End If
                  '*******************************************************************
7100          End If
7110          sql = "SELECT * FROM Reports WHERE 0 = 1"
7120          Set tb = New Recordset
7130          RecOpenServer 0, tb, sql
7140          tb.AddNew
7150          tb!SampleID = RP.SampleID
7160          tb!Name = udtHeading.Name
7170          tb!Dept = "E"
7180          tb!Initiator = RP.Initiator
7190          tb!PrintTime = PrintTime
7200          tb!RepNo = "1E" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
7210          tb!PageNumber = 1
7220          tb!Report = .TextRTF
7230          tb!Printer = Printer.DeviceName
7240          tb.Update
7250      End With
7260  End If

7270  ResetPrinter

7280  Exit Sub

PrintResultEndWin1_Error:

      Dim strES As String
      Dim intEL As Integer

7290  intEL = Erl
7300  strES = Err.Description

7310  sql = "Delete FROM printpending WHERE SampleID = '" & RP.SampleID & "' and department = '" & RP.Department & "'"
7320  Cnxn(0).Execute sql

7330  LogError "modEndocrinology", "PrintResultEndWin1", intEL, strES, sql

End Sub

Public Sub PrintResultEndWin(Optional ByVal PrintA4 As Boolean = True)

      Dim tb         As Recordset
      Dim sql        As String
      Dim Sex        As String
      Dim cUnits     As String
      Dim Flag       As String
      Dim v          As String
      Dim SampleType As String
      Dim SampleDate As String
      Dim Rundate    As String
      Dim DoB        As String
      Dim RunTime    As String
      Dim Cat        As String
      Dim PrintTime  As String
      Dim TResult    As String
      Dim InconclusiveFound As Boolean
      Dim Analyser   As String    'user to detect AxSYM Virology
      Dim AuthorisedBy As String
      Dim TestPerformedAt As String
      Dim ExternalTestingNote As String

      Dim lpc        As Integer
      Dim n          As Integer
      Dim TestCount  As Integer
      Dim strFormat  As String
      Dim f          As Integer
      Dim Fontz1     As Integer
      Dim Fontz2     As Integer

      Dim PageNumber As String
      Dim TotalLines As Integer
      Dim CommentLines As Integer
      Dim BodyLines  As Integer
      Dim LineNoStartFooter As Integer
      Dim TotalPages As Integer
      Dim i          As Integer
      Dim FontBold   As Boolean

      Dim BRs        As New BIEResults
      Dim br         As BIEResult
      Dim OB         As Observation
      Dim OBS        As New Observations
      Dim udtPrintLine() As ResultLine


10    On Error GoTo PrintResultEndWin_Error


20    If PrintA4 Then
30        TotalLines = 100
40        CommentLines = 10
50        BodyLines = GetOptionSetting("EndocrinologyPrintA4BodyLines", "60")
60        LineNoStartFooter = BodyLines + 13      'body lines + header lines
70    Else
80        TotalLines = 100
90        CommentLines = 4
100       BodyLines = GetOptionSetting("EndocrinologyPrintA5BodyLines", "19")
110       LineNoStartFooter = BodyLines + 13      'body lines + header lines
120   End If


130   ReDim lp(0 To TotalLines) As String
140   ReDim udtPrintLine(0 To TotalLines) As ResultLine
150   ReDim Comments(1 To CommentLines) As String

160   For n = 0 To TotalLines - 1
170       udtPrintLine(n).Analyte = ""
180       udtPrintLine(n).Result = ""
190       udtPrintLine(n).Flag = ""
200       udtPrintLine(n).Units = ""
210       udtPrintLine(n).NormalRange = ""
220       udtPrintLine(n).Fasting = ""
230       udtPrintLine(n).Reason = ""
240       udtPrintLine(n).Comment = ""
250   Next

260   InconclusiveFound = False
270   lpc = 0
280   Analyser = ""

290   PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

300   sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
310   Set tb = New Recordset
320   RecOpenClient 0, tb, sql
330   If tb.EOF Then Exit Sub

340   Cat = Trim(tb!Category & "")
350   If Cat = "" Then Cat = "Default"
360   If IsDate(tb!DoB) Then
370       DoB = Format(tb!DoB, "dd/mmm/yyyy")
380   Else
390       DoB = ""
400   End If
410   If IsDate(tb!SampleDate) Then
420       SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
430   Else
440       SampleDate = ""
450   End If
460   If IsDate(RunTime) Then
470       Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
480   Else
490       If IsDate(tb!Rundate) Then
500           Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
510       Else
520           Rundate = ""
530       End If
540   End If
550   Sex = tb!Sex & ""

560   Set BRs = BRs.Load("End", RP.SampleID, "Results", "0", Cat, "")
570   If BRs.Count = 0 Then Exit Sub

580   Analyser = BRs(1).Analyser

590   ClearUdtHeading
600   With udtHeading
610       .SampleID = RP.SampleID
620       If UCase$(Analyser) = "VIROLOGY" Then
630           .Dept = "Microbiology"
640           .DocumentNo = GetOptionSetting("EndVirologyDocumentNo", "")
650       Else
660           .Dept = "Endocrinology"
670           .DocumentNo = GetOptionSetting("EndMainDocumentNo", "")
680       End If
690       .Name = tb!PatName & ""
700       .Ward = RP.Ward
710       .DoB = DoB
720       .Chart = tb!Chart & ""
730       .Clinician = RP.Clinician
740       .Address0 = tb!Addr0 & ""
750       .Address1 = tb!Addr1 & ""
760       .GP = RP.GP
770       .Sex = tb!Sex & ""
780       .Hospital = tb!Hospital & ""
790       .SampleDate = tb!SampleDate & ""
800       .RecDate = tb!RecDate & ""
810       .Rundate = tb!Rundate & ""
820       .GpClin = ""
830       .SampleType = SampleType
840       .AandE = tb!AandE & ""

850   End With



860   If UCase$(Analyser) = "VIROLOGY" Then

      '    lp(lpc) = Space(3) & _
      '              FormatString("Test", 50, , AlignLeft) & _
      '              FormatString("Result", 15, , AlignLeft) & _
      '              FormatString("Value", 8, , AlignLeft) & _
      '              FormatString("Unit", 7, , AlignLeft)
      '    udtPrintLine(lpc).Analyte = "*HEADING*"
      '    lpc = lpc + 1
870       AddResultToLP udtPrintLine, lp, lpc, "Test", "Result ", "Value", "Unit", "", , , , True, , True
880       CrCnt = CrCnt + 1
          

      '    lp(lpc) = String(81, "-")
      '    udtPrintLine(lpc).Analyte = "*LINE*"
      '    lpc = lpc + 1

          'Virology report
890       For Each br In BRs
900           If AuthorisedBy = "" Then AuthorisedBy = GetAuthorisedBy(br.Operator)
910           If IsNumeric(br.Result) Then
920               TResult = TranslateEndResultVirology(br.Code, br.Result)
930               If TResult = "Inconclusive *" Then InconclusiveFound = True
940           Else
950               TResult = br.Result
960               If TResult = "Inconclusive *" Then InconclusiveFound = True
970           End If
              
980           AddResultToLP udtPrintLine, lp, lpc, br.LongName & " (" & br.ShortName & ")", TResult, br.Result, br.Units
              
      '        If Br.Code = "118" Then
      ''            AddResultToLP udtPrintLine, lp, lpc, FormatString(Br.LongName & " (" & Br.ShortName & ")", 50, , AlignLeft), _
      ''                     FormatString(TResult, 15, , AlignLeft), _
      ''                     FormatString(Br.Result, 8, , AlignLeft), _
      ''                     FormatString(Br.Units, 7, , AlignLeft)
      '            AddResultToLP udtPrintLine, lp, lpc, Br.LongName & " (" & Br.ShortName & ")", TResult, Br.Result, , Br.Units
      ''            lp(lpc) = FormatString(Br.LongName & " (" & Br.ShortName & ")", 50, , AlignLeft) & _
      ''                      FormatString(TResult, 15, , AlignLeft) & _
      ''                      FormatString(Br.Result, 8, , AlignLeft) & _
      ''                      FormatString(Br.Units, 7, , AlignLeft)
      ''            udtPrintLine(lpc).Analyte = Br.LongName & " (" & Br.ShortName & ")"
      ''            udtPrintLine(lpc).Result = FormatString(TResult, 15, , AlignLeft) & FormatString(Br.Result, 8, , AlignLeft)
      ''            udtPrintLine(lpc).Units = Br.Units
      '        Else
      '            AddResultToLP udtPrintLine, lp, lpc, Br.LongName & " (" & Br.ShortName & ")", _
      '                    FormatString(TResult, 15, , AlignLeft) & FormatString(Br.Result, 8, , AlignLeft), _
      '                    Br.Units
      ''            lp(lpc) = FormatString(Br.LongName & " (" & Br.ShortName & ")", 50, , AlignLeft) & _
      ''                      FormatString(TResult, 15, , AlignLeft)
      ''            udtPrintLine(lpc).Analyte = Br.LongName & " (" & Br.ShortName & ")"
      ''            udtPrintLine(lpc).Result = FormatString(TResult, 15, , AlignLeft) & FormatString(Br.Result, 8, , AlignLeft)
      '        End If
990           LogTestAsPrinted "End", br.SampleID, br.Code
1000      Next

1010  Else
1020      AddResultToLP udtPrintLine, lp, lpc, "Test", "Result", "Unit", "Ref. Range", "Flag", , , , True, , True
1030      CrCnt = CrCnt + 1
          'Endocrinology Report
1040      For Each br In BRs
1050          If AuthorisedBy = "" Then AuthorisedBy = GetAuthorisedBy(br.Operator)
1060          If Analyser = "" Then
1070              Analyser = Trim$(br.Analyser)
1080          End If
1090          SampleType = br.SampleType
1100          If InStr(1, SampleType, br.SampleType) = 0 Then SampleType = SampleType & " "
1110          RunTime = br.RunTime
1120          v = br.Result

1130          TestPerformedAt = ""
1140          If UCase(HospName(0)) <> UCase(br.Hospital) Then
1150              TestPerformedAt = Left(UCase(br.Hospital), 1)
1160              If InStr(ExternalTestingNote, UCase(br.Hospital)) = 0 Then
1170                  ExternalTestingNote = ExternalTestingNote & TestPerformedAt & " = Test Analysed at " & UCase(br.Hospital) & " "
1180              End If

1190              TestPerformedAt = "(" & TestPerformedAt & ")"
1200          End If

              'Test Name
1210          lp(lpc) = lp(lpc) & Left(br.LongName & TestPerformedAt & Space(26), 26)
1220          udtPrintLine(lpc).Analyte = Left(br.LongName & TestPerformedAt & Space(26), 26)

              'Result
1230          If EndTestAffected(br) = False Then
1240              If IsNumeric(v) Then
1250                  Select Case br.Printformat
                          Case 0: strFormat = "#########"
1260                      Case 1: strFormat = "######0.0"
1270                      Case 2: strFormat = "#####0.00"
1280                      Case 3: strFormat = "####0.000"
1290                  End Select
1300                  lp(lpc) = lp(lpc) & " " & Right(Space(9) & Format(v, strFormat), 9)
1310                  udtPrintLine(lpc).Result = Format(v, strFormat)
1320              Else
1330                  lp(lpc) = lp(lpc) & " " & Right(Space(9) & v, 9)
1340                  udtPrintLine(lpc).Result = v
1350              End If
1360          Else
1370              lp(lpc) = lp(lpc) & " " & Right(Space(9) & "XXXXXXX", 9)
1380          End If

              'Flag
1390          udtPrintLine(lpc).Flag = InterpFlag(br, v)
1400          lp(lpc) = lp(lpc) & udtPrintLine(lpc).Flag & " "

              'Unit
1410          cUnits = ListText("UN", br.Units)
1420          If cUnits = "" Then cUnits = br.Units
1430          udtPrintLine(lpc).Units = cUnits
1440          lp(lpc) = lp(lpc) & " " & Right(Space(6) & cUnits, 8)

              'Normal Range
1450          udtPrintLine(lpc).NormalRange = InterpNormalRange(br, tb!Sex & "")
1460          lp(lpc) = lp(lpc) & " " & udtPrintLine(lpc).NormalRange

              'Reason
1470          If EndTestAffected(br) Then
1480              udtPrintLine(lpc).Reason = LTrim(RTrim(EndReasonAffect(br)))
1490              lp(lpc) = lp(lpc) & " " & udtPrintLine(lpc).Reason
1500          Else
1510              udtPrintLine(lpc).Reason = ""
1520          End If

              'Result Comment
1530          udtPrintLine(lpc).Comment = LTrim(RTrim(br.Comment))
1540          lp(lpc) = lp(lpc) & " " & udtPrintLine(lpc).Comment


1550          lpc = lpc + 1

              'if there are any ref range comments add them as seperate lines current test
1560          AddNRComment br.Code, udtPrintLine, lp, lpc


1570          LogTestAsPrinted "End", br.SampleID, br.Code

1580      Next
1590  End If

      'add blank line before comment
1600  AddResultToLP udtPrintLine, lp, lpc, "", ""

      'comments
1610  Set OBS = OBS.Load(RP.SampleID, "Endocrinology", "Demographic")
1620  If Not OBS Is Nothing Then
1630      For Each OB In OBS
1640          Select Case UCase$(OB.Discipline)
                  Case "ENDOCRINOLOGY"
1650                  AddCommentToLP udtPrintLine, lp, lpc, OB.Comment, ""
1660              Case "DEMOGRAPHIC"
1670                  AddCommentToLP udtPrintLine, lp, lpc, OB.Comment, ""
1680          End Select
1690      Next
1700  End If

1710  If Not IsDate(DoB) Or Trim(udtHeading.Sex) = "" Then
1720      lp(lpc) = "               " & "**** No Sex/DoB given. No reference range applied! ****"
1730      udtPrintLine(lpc).Analyte = "*COMMENT*"
1740      lpc = lpc + 1
1750  End If


      'START PRINTING (ALL DATA GATHERED IN udtPrintLine Structure
      'bring lpc index back by one position
1760  lpc = lpc - 1

1770  PrintReport udtPrintLine, lp, lpc, "End", PrintA4, SampleDate, Rundate, AuthorisedBy, PrintTime, SampleType, ""

1780  Exit Sub

PrintResultEndWin_Error:

      Dim strES      As String
      Dim intEL      As Integer

1790  intEL = Erl
1800  strES = Err.Description

1810  sql = "Delete FROM printpending WHERE SampleID = '" & RP.SampleID & "' and department = '" & RP.Department & "'"
1820  Cnxn(0).Execute sql

1830  LogError "modEndocrinology", "PrintResultEndWin", intEL, strES, sql

End Sub

Private Function InterpNormalRange(ByVal br As BIEResult, ByVal Sex As String) As String

      Dim Low        As Single
      Dim High       As Single
      Dim strLow     As String * 4
      Dim strHigh    As String * 4
      Dim n          As Integer
      Dim strFormat  As String

      Dim tb         As Recordset
      Dim sql        As String

10    On Error GoTo InterpNormalRange_Error

20    High = Val(br.High)
30    Low = Val(br.Low)

40    If Low = 0 And (High = 0 Or High = 999 Or High = 9999) Then
50        InterpNormalRange = ""
60    Else
70        If Low = 0 Then
80            strLow = Low
90        ElseIf Low < 10 Then
100           strLow = Format(Low, "0.00")
110       ElseIf Low < 100 Then
120           strLow = Format(Low, "##.0")
130       Else
140           strLow = Format(Low, " ###")
150       End If

160       If High < 10 Then
170           strHigh = Format(High, "0.00")
180       ElseIf High < 100 Then
190           strHigh = Format(High, "##.0")
200       Else
210           strHigh = Format(High, "### ")
220       End If

230       sql = "SELECT * FROM endtestdefinitions WHERE code = '" & br.Code & "'"
240       Set tb = New Recordset
250       RecOpenClient 0, tb, sql
260       If Not tb.EOF Then

270           If (Not IsNull(tb!ShowLessThan) And tb!ShowLessThan <> 0) And Val(strLow) = 0 And (Val(strHigh) <> 999 Or Val(strHigh) <> 9999) Then
280               InterpNormalRange = "<" & Format(strHigh, strFormat)
290           ElseIf (Not IsNull(tb!ShowMoreThan) And tb!ShowMoreThan <> 0) And Val(strLow) <> 0 And (Val(strHigh) = 999 Or Val(strHigh) = 9999) Then
300               InterpNormalRange = ">" & Format(strLow, strFormat) & " "
310           Else
                'Zyam
                  If Val(strLow) = 0 And Val(strHigh) = 0 Then
                    InterpNormalRange = " "
                  Else
                    InterpNormalRange = "(" & Format(strLow, strFormat) & "-" & Format(strHigh, strFormat) & ")"
                  End If
                  'Zyam
320
330           End If


340       Else
                'Zyam
              If Val(strLow) = 0 And Val(strHigh) = 0 Then
                InterpNormalRange = " "
              
              Else
                InterpNormalRange = "(" & Format(strLow, strFormat) & "-" & Format(strHigh, strFormat) & ")"
              End If
              'Zyam
350
360       End If
370   End If


380   Exit Function
InterpNormalRange_Error:

390   LogError "modEndocrinology", "InterpNormalRange", Erl, Err.Description, sql


End Function

'Private Function InterpNormalRange(ByVal Br As BIEResult, ByVal Sex As String) As String
'
'      Dim Low        As Single
'      Dim High       As Single
'      Dim strLow     As String * 4
'      Dim strHigh    As String * 4
'      Dim n          As Integer
'      Dim strFormat  As String
'
'      Dim tb         As Recordset
'      Dim sql        As String
'
'10    On Error GoTo InterpNormalRange_Error
'
'20    High = Val(Br.High)
'30    Low = Val(Br.Low)
'
'40    If Low = 0 Then
'50        strLow = Low
'60    ElseIf Low < 10 Then
'70        strLow = Format(Low, "0.00")
'80    ElseIf Low < 100 Then
'90        strLow = Format(Low, "##.0")
'100   Else
'110       strLow = Format(Low, " ###")
'120   End If
'
'130   If High < 10 Then
'140       strHigh = Format(High, "0.00")
'150   ElseIf High < 100 Then
'160       strHigh = Format(High, "##.0")
'170   Else
'180       strHigh = Format(High, "### ")
'190   End If
'
'
'200   sql = "SELECT * FROM endtestdefinitions WHERE code = '" & Br.Code & "'"
'210   Set tb = New Recordset
'220   RecOpenServer 0, tb, sql
'
'230   If (Not IsNull(tb!ShowLessThan) And tb!ShowLessThan <> 0) _
 '         And Val(tb!FemaleLow) = 0 And (Val(tb!FemaleHigh) <> 999 Or Val(tb!FemaleHigh) <> 9999) Then
'240       InterpNormalRange = "<" & Format(tb!FemaleHigh, strFormat)
'250   ElseIf (Not IsNull(tb!ShowMoreThan) And tb!ShowMoreThan <> 0) _
 '             And Val(tb!FemaleLow) <> 0 And (Val(tb!FemaleHigh) = 999 Or Val(tb!FemaleHigh) = 9999) Then
'260       InterpNormalRange = ">" & Format(tb!FemaleLow, strFormat) & " "
'270   Else
'280       InterpNormalRange = "(" & Format(tb!FemaleLow, strFormat) & "-" & Format(tb!FemaleHigh, strFormat) & ")"
'290   End If
'
'
'
'300   Exit Function
'InterpNormalRange_Error:
'
'310   LogError "modEndocrinology", "InterpNormalRange", Erl, Err.Description, sql
'
'
'End Function

Private Sub AddNRComment(ByVal Code As String, udtPrintLine() As ResultLine, lp() As String, lpc As Integer)

      Dim NRComment As String
      Dim NRCommentComponents() As String
      Dim i As Integer

10    On Error GoTo InterpNormalRange_Error



20    If udtHeading.Sex = "" Or (Not IsDate(udtHeading.DoB)) Then Exit Sub


30    If UCase(Code) = UCase(GetOptionSetting("EndCodeForCortisol", "COR")) Then
40        NRComment = GetOptionSetting("ENDCORTRANGE", "")
50    ElseIf UCase(Code) = UCase(SysOptEndCodeHBA1C(0)) Then
60        NRComment = GetOptionSetting("ENDHBARANGE", "")
70    ElseIf UCase(Code) = UCase(SysOptEndCodeCalcA1C(0)) Then
80        NRComment = GetOptionSetting("ENDCALCA1CRANGE", "")
90    ElseIf UCase(Code) = UCase(SysOptEndCodeBNP(0)) Then
100       NRComment = GetOptionSetting("ENDBNPRANGE", "")
110   ElseIf UCase(Code = SysOptEndCodeB12(0)) Then
120       NRComment = GetOptionSetting("ENDB12RANGE", "")
130   ElseIf UCase(Code) = UCase(SysOptEndCodeB12New(0)) Then
140       NRComment = GetOptionSetting("ENDB12NEWRange", "")
150   ElseIf UCase(Code) = UCase(SysOptEndCodeCo(0)) Then
160       NRComment = GetOptionSetting("ENDCORange", "")
170   ElseIf UCase(Code) = UCase(SysOptEndCodeVITD(0)) Then
180       NRComment = GetOptionSetting("ENDVITDRange", "")
190   ElseIf UCase(Code) = UCase(SysOptEndCodeFSH(0)) And udtHeading.Sex = "F" Then
200       NRComment = GetOptionSetting("ENDFSHRange", "")
210   ElseIf UCase(Code) = UCase(SysOptEndCodeLH(0)) And udtHeading.Sex = "F" Then
220       NRComment = GetOptionSetting("ENDLHRange", "")
230   ElseIf UCase(Code) = UCase(SysOptEndCodePRO(0)) And udtHeading.Sex = "F" Then
240       NRComment = GetOptionSetting("ENDPRORange", "")
250   ElseIf UCase(Code) = UCase(SysOptEndCodeOES(0)) And udtHeading.Sex = "F" Then
260       NRComment = GetOptionSetting("ENDOESRange", "")
270   ElseIf UCase(Code) = UCase(SysOptEndCodePRL(0)) And udtHeading.Sex = "F" Then
280       NRComment = GetOptionSetting("ENDPRLRange", "")
290   ElseIf Code = SysOptEndCodeTHC(0) And udtHeading.Sex = "F" Then
300       NRComment = GetOptionSetting("ENDTHCRange", "")
310   ElseIf UCase(Code) = UCase(SysOptEndCodeTRO(0)) Then
320       NRComment = GetOptionSetting("ENDTRORange", "")
330   End If

340   If NRComment = "" Then Exit Sub
350   NRComment = Replace(NRComment, "&lt;", "<")
360   NRCommentComponents = Split(NRComment, "<escape V="".br""/>")

370   For i = LBound(NRCommentComponents) To UBound(NRCommentComponents)
380       udtPrintLine(lpc).Analyte = "*NRCOMMENT*"
390       udtPrintLine(lpc).NormalRange = NRCommentComponents(i)
400       lp(lpc) = Space(3) & _
                      Space(20 + 1) & _
                      Space(8 + 1) & _
                      Space(5 + 1) & _
                      Space(13 + 1) & _
                      LTrim(RTrim(NRCommentComponents(i)))
410       lpc = lpc + 1
420   Next i



430   Exit Sub
InterpNormalRange_Error:

440   LogError "modEndocrinology", "InterpNormalRange", Erl, Err.Description


End Sub

Private Function InterpFlag(ByVal br As BIEResult, ByVal v As String) As String

      Dim tb         As Recordset
      Dim sql        As String
      Dim Flag       As String

10    On Error GoTo InterpFlag_Error


20    Flag = "  "
30    sql = "SELECT * FROM endtestdefinitions WHERE code = '" & br.Code & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70        If Trim(tb!forfert & "") <> 1 And udtHeading.Sex <> "" And IsDate(udtHeading.DoB) Then
80            If IsNumeric(v) Then
90                If Val(v) > br.PlausibleHigh Then
100                   Flag = " X"
110               ElseIf Val(v) < br.PlausibleLow Then
120                   Flag = " X"
130               ElseIf Val(v) > br.High And Val(br.High) > 0 Then
140                   Flag = " H"
150               ElseIf Val(v) < br.Low Then
160                   Flag = " L"
170               Else
180                   Flag = "  "
190               End If
200           Else
210               If Left(v, 1) = "<" Then
220                   Flag = " L"
230               ElseIf Left(v, 1) = ">" Then
240                   Flag = " H"
250               Else
260                   Flag = "  "
270               End If
280           End If

290       End If

300   End If
310   InterpFlag = Flag


320   Exit Function
InterpFlag_Error:

330   LogError "modEndocrinology", "InterpFlag", Erl, Err.Description, sql


End Function





Public Function TranslateEndResultVirology(Code As String, Result As String) As String

10        On Error GoTo TranslateEndResultVirology_Error

20        If Result = "Negative" Or Result = "Positive" Or Result = "Inconclusive *" Then
30            TranslateEndResultVirology = Result
40        Else

50            Select Case Code
              Case "106":    'HBsAg
60                If Result < 1 Then
70                    Result = "Negative"
80                ElseIf Result >= 1 Then
90                    Result = "Inconclusive *"
100               End If
110           Case "118":    'AUSAB
120               If Result < 10 Then
130                   Result = "Negative"
140               ElseIf Result >= 10 Then
150                   Result = "Positive"
160               End If
170           Case "126":    'HepBCo
180               If Result >= 1.001 And Result <= 3 Then
190                   Result = "Negative"
200               ElseIf Result >= 0 And Result <= 1 Then
210                   Result = "Inconclusive *"
220               End If
230           Case "841":    'HCV
240               If Result < 1 Then
250                   Result = "Negative"
260               ElseIf Result >= 1 Then
270                   Result = "Inconclusive *"
280               End If
290           Case "817":    'HIV
300               If Result < 0.9 Then
310                   Result = "Negative"
320               ElseIf Result >= 0.9 Then
330                   Result = "Inconclusive *"
340               End If
350           End Select
360           TranslateEndResultVirology = Result
370       End If
          

380       Exit Function

TranslateEndResultVirology_Error:

          Dim strES As String
          Dim intEL As Integer

390       intEL = Erl
400       strES = Err.Description
410       LogError "modEndocrinology", "TranslateEndResultVirology", intEL, strES

End Function




Public Function EndLongNameFor(ByVal Code As String) As String

      Dim tb As New Recordset
      Dim sql As String


10    On Error GoTo EndLongNameFor_Error

20    EndLongNameFor = "???"

30    sql = "SELECT * from EndTestDefinitions WHERE Code = '" & Code & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70        EndLongNameFor = Trim(tb!LongName)
80    End If




90    Exit Function

EndLongNameFor_Error:

      Dim strES As String
      Dim intEL As Integer



100   intEL = Erl
110   strES = Err.Description
120   LogError "basEndocrinology", "EndLongNameFor", intEL, strES, sql


End Function
