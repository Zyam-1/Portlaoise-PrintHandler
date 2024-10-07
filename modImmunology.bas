Attribute VB_Name = "modImmunology"
Option Explicit

Sub LogImmAsPrinted(ByVal TestCode As String)

      Dim sql As String

10    On Error GoTo LogImmAsPrinted_Error

20    sql = "Update ImmResults " & _
            "set valid = 1, printed = 1 WHERE " & _
            "SampleID = '" & RP.SampleID & "' " & _
            "and code = '" & TestCode & "'"
30    Cnxn(0).Execute sql

40    Exit Sub

LogImmAsPrinted_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "modImmunology", "LogImmAsPrinted", intEL, strES, sql

End Sub

Public Sub PrintResultImmWin()

          Dim tb As Recordset
          Dim tbUN As Recordset
          Dim sql As String
          Dim Sex As String
10        ReDim lp(0 To 35) As String
20        ReDim lc(0 To 35) As String
          Dim lpc As Integer
          Dim cUnits As String
          Dim TempUnits As String
          Dim Flag As String
          Dim n As Integer
          Dim v As String
          Dim Low As Single
          Dim High As Single
          Dim strLow As String * 4
          Dim z As Long
          Dim s As String
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
30        ReDim Comments(1 To 8) As String
          Dim SampleDate As String
          Dim Rundate As String
          Dim DoB As String
          Dim RunTime As String
          Dim Fasting As String
          Dim Fx As Fasting
          Dim CodeGLU As String
          Dim CodeCHO As String
          Dim CodeTRI As String
          Dim udtPrintLine(0 To 35) As PrintLine
          Dim strFormat As String
          Dim SerumPrn As Boolean
          Dim UrinePrn As Boolean
          Dim C As Integer
          Dim f As Integer
          Dim Fontz1 As Integer
          Dim Fontz2 As Integer
          Dim PrintTime As String
          Dim AuthorisedBy As String
          Dim PageNumber As String

          Dim TestPerformedAt As String
          Dim ExternalTestingNote As String

40        On Error GoTo PrintResultImmWin_Error

50        PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

60        SerumPrn = False
70        UrinePrn = False

80        For n = 0 To 35
90            udtPrintLine(n).Analyte = ""
100           udtPrintLine(n).Result = ""
110           udtPrintLine(n).Flag = ""
120           udtPrintLine(n).Units = ""
130           udtPrintLine(n).NormalRange = ""
140           udtPrintLine(n).Fasting = ""
150       Next

160       sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & RP.SampleID & "'"
170       Set tb = New Recordset
180       RecOpenClient 0, tb, sql

190       If tb.EOF Then
200           Exit Sub
210       End If

220       If IsDate(tb!DoB) Then
230           DoB = Format(tb!DoB, "dd/mmm/yyyy")
240       Else
250           DoB = ""
260       End If

270       ClearUdtHeading
280       With udtHeading
290           .SampleID = RP.SampleID
300           .Dept = "Immunology"
310           .Name = tb!PatName & ""
320           .Ward = RP.Ward
330           .DoB = DoB
340           .Chart = tb!Chart & ""
350           .Clinician = RP.Clinician
360           .Address0 = tb!Addr0 & ""
370           .Address1 = tb!Addr1 & ""
380           .GP = RP.GP
390           .Sex = tb!Sex & ""
400           .Hospital = tb!Hospital & ""
410           .SampleDate = tb!SampleDate & ""
420           .RecDate = tb!RecDate & ""
430           .Rundate = tb!Rundate & ""
440           .GpClin = ""
450           .SampleType = SampleType
460           .DocumentNo = GetOptionSetting("ImmMainDocumentNo", "")
470           .AandE = tb!AandE & ""
480       End With

490       ResultsPresent = False
500       Set BRs = BRs.Load("Imm", RP.SampleID, "Results", 0, "Default", "")
510       If Not BRs Is Nothing Then

520           TestCount = BRs.Count
530           If TestCount <> 0 Then
540               If IsAllergy(BRs) Then
550                   PrintResultAllergy BRs
560                   Exit Sub
570               ElseIf IsHemochromatosis(BRs) Then
580                   PrintResultHemochromatosis BRs
590                   Exit Sub
600               Else
610                   ResultsPresent = True
620                   SampleType = BRs(1).SampleType
630                   If Trim(SampleType) = "" Then SampleType = "S"
640               End If
650           End If
660       End If



670       lpc = 0
680       If ResultsPresent Then
690           If AuthorisedBy = "" Then
700               AuthorisedBy = GetAuthorisedBy(GetLastValidatedBy(BRs))
710           End If
720           For Each br In BRs

730               If br.Printable = True Then
                      'If Br.Operator <> "" Then RP.Initiator = Br.Operator
740                   Rundate = br.Rundate
750                   If br.SampleType = "S" And SerumPrn = False Then
                          '                lp(lpc) = "SERUM"              'QMS Ref #817963
                          'lpc = lpc + 1
760                       SerumPrn = True
770                   End If
780                   If br.SampleType = "U" And UrinePrn = False Then
790                       lp(lpc) = ""
800                       lpc = lpc + 1
                          '                lp(lpc) = "URINE"              'QMS Ref #817963
                          '                lpc = lpc + 1
810                       UrinePrn = True
820                   End If

830                   If br.Pc = "P" Then
840                       lc(lpc) = br.LongName & " - Phoned. "
850                   ElseIf br.Pc = "C" Then
860                       lc(lpc) = br.LongName & " - Checked. "
870                   ElseIf br.Pc = "PC" Then
880                       lc(lpc) = br.LongName & " - Phoned & Checked. "
890                   End If

900                   lc(lpc) = lc(lpc) & br.Comment
910                   RunTime = br.RunTime
920                   v = br.Result

930                   High = Val(br.High)
940                   Low = Val(br.Low)

950                   If Low < 10 Then
960                       strLow = Format(Low, "0.00")
970                   ElseIf Low < 100 Then
980                       strLow = Format(Low, "##.0")
990                   ElseIf Low > 99 And Low < 1000 Then
1000                      strLow = Format(Low, " ###0")
1010                  Else
1020                      strLow = Format(Low, "####")
1030                  End If

1040                  If High < 10 Then
1050                      strHigh = Format(High, "0.00")
1060                  ElseIf High < 100 Then
1070                      strHigh = Format(High, "##.0")
1080                  Else
1090                      strHigh = Format(High, "### ")
1100                  End If

1110                  If IsNumeric(v) And udtHeading.Sex <> "" And IsDate(udtHeading.DoB) Then
1120                      If Val(v) > br.PlausibleHigh Then
1130                          udtPrintLine(lpc).Flag = " X "
1140                          udtPrintLine(lpc).Result = "***"
1150                          lp(lpc) = "  "
1160                          Flag = " X"
1170                      ElseIf Val(v) < br.PlausibleLow Then
1180                          udtPrintLine(lpc).Flag = " X "
1190                          udtPrintLine(lpc).Result = "***"
1200                          lp(lpc) = "  "
1210                          Flag = " X"
                          'Zyam
1220                      ElseIf Val(v) > High And High <> 0 Then
1230                          udtPrintLine(lpc).Flag = " H "
1240                          lp(lpc) = "  "    'bold
1250                          Flag = " H"
                          'Zyam
1260                      ElseIf Val(v) < Low Then
1270                          udtPrintLine(lpc).Flag = " L "
1280                          lp(lpc) = "  "    'bold
1290                          Flag = " L"
1300                      Else
1310                          udtPrintLine(lpc).Flag = " E "
1320                          lp(lpc) = "  "    'unbold
1330                          Flag = "  "
1340                      End If
1350                  Else
1360                      If Left(v, 1) = "<" Or Left(v, 1) = ">" Then
1370                          If Left(v, 1) = "<" And Trim(Mid(v, 2)) <= Low Then
1380                              udtPrintLine(lpc).Flag = " L "
1390                              lp(lpc) = "  "    'bold
1400                              Flag = " L"

1410                          ElseIf Left(v, 1) = ">" And Trim(Mid(v, 2)) >= High And High <> 0 Then
1420                              udtPrintLine(lpc).Flag = " H "
1430                              lp(lpc) = "  "    'bold
1440                              Flag = " H"
1450                          End If
1460                      Else
1470                          udtPrintLine(lpc).Flag = " C "
1480                          lp(lpc) = "  "    'unbold
1490                          Flag = "  "
1500                      End If
1510                  End If


1520                  TestPerformedAt = ""
1530                  If UCase(HospName(0)) <> UCase(br.Hospital) Then
1540                      TestPerformedAt = Left(UCase(br.Hospital), 1)
1550                      If InStr(ExternalTestingNote, UCase(br.Hospital)) = 0 Then
1560                          ExternalTestingNote = ExternalTestingNote & TestPerformedAt & " = Test Analysed at " & UCase(br.Hospital) & " "
1570                      End If
1580                      TestPerformedAt = "(" & TestPerformedAt & ")"
1590                  End If


                      '                lp(lpc) = "Allergy Test               Result"
                      '                lp(lpc) = lp(lpc) & "                Units"
                      '                lp(lpc) = lp(lpc) & "        Reference Range"
                      '                lpc = lpc + 1
1600                  If UCase$(HospName(0)) = "MULINGAR" Then
1610                      If br.Code = "IT" Then
1620                          lp(lpc) = lp(lpc) & Left("Serum Immunofixation." & Space(25), 25)
1630                      ElseIf br.Code = "NDNA" Then
1640                          lp(lpc) = lp(lpc) & Left("DNA Abs" & Space(25), 25)
1650                      ElseIf br.Code = "TTG" Then
1660                          lp(lpc) = lp(lpc) & Left("IgA tissue Transglutaminase Abs" & Space(25), 25)
1670                      Else
1680                          lp(lpc) = lp(lpc) & Left(br.LongName & TestPerformedAt & Space(25), 25)
1690                      End If
1700                  Else

1710                      lp(lpc) = lp(lpc) & Left(br.LongName & TestPerformedAt & Space(25), 25)
1720                  End If


1730                  udtPrintLine(lpc).Analyte = Left(br.LongName & Space(16), 16)
1740                  If ImmTestAffected(br) = False Then
1750                      If IsNumeric(v) Then
1760                          Select Case br.Printformat
                              Case 0: strFormat = "######"
1770                          Case 1: strFormat = "###0.0"
1780                          Case 2: strFormat = "##0.00"
1790                          Case 3: strFormat = "#0.000"
1800                          End Select


1810                          If Trim(udtPrintLine(lpc).Result) <> "***" Then
1820                              If br.Code = "RUB" And Val(v) > 50 Then
1830                                  lp(lpc) = lp(lpc) & " " & Right(Space(6) & "> 50", 6)
1840                              Else
1850                                  lp(lpc) = lp(lpc) & " " & Right(Space(6) & Format(v, strFormat), 6)
1860                              End If
1870                          Else
1880                              lp(lpc) = lp(lpc) & "  ***** "
1890                          End If
1900                          If Trim(udtPrintLine(lpc).Result) <> "***" Then udtPrintLine(lpc).Result = Format(v, strFormat)
1910                      Else
1920                          If Trim(udtPrintLine(lpc).Result) <> "***" Then
1930                              lp(lpc) = lp(lpc) & " " & v
1940                          Else
1950                              lp(lpc) = lp(lpc) & "  ***** "
1960                          End If
1970                          If Trim(udtPrintLine(lpc).Result) <> "***" Then udtPrintLine(lpc).Result = Format(v, strFormat)
1980                      End If
1990                  Else
2000                      lp(lpc) = lp(lpc) & "XXXXXX "
2010                  End If

2020                  If br.Code = "RUB" And IsNumeric(udtPrintLine(lpc).Result) = True Then
2030                      lp(lpc) = lp(lpc) & " IU/ml "
2040                      If Val(udtPrintLine(lpc).Result) >= 0 And Val(udtPrintLine(lpc).Result) <= 9.9 Then
2050                          lp(lpc) = lp(lpc) & "Non Immune"
2060                      ElseIf Val(udtPrintLine(lpc).Result) > 9.9 And Val(udtPrintLine(lpc).Result) <= 14.9 Then
2070                          lp(lpc) = lp(lpc) & "Immune"
2080                      ElseIf Val(udtPrintLine(lpc).Result) > 14.9 Then
2090                          lp(lpc) = lp(lpc) & "Immune"
2100                      End If
2110                  Else
2120                      lp(lpc) = lp(lpc) & Flag & " "
2130                      If IsNumeric(v) Then
2140                          sql = "SELECT * FROM Lists WHERE " & _
                                    "ListType = 'UN' and Code = '" & br.Units & "'"
2150                          Set tbUN = Cnxn(0).Execute(sql)
2160                          If Not tbUN.EOF Then
2170                              cUnits = Left(tbUN!Text & Space(10), 10)
2180                          Else
2190                              cUnits = Left(br.Units & Space(10), 10)
2200                          End If
2210                          udtPrintLine(lpc).Units = cUnits
2220                          If br.PrnRR = True And udtHeading.Sex <> "" And IsDate(udtHeading.DoB) And br.Code <> "RUB" Then
2230                              If (Val(strLow) = 0 And Val(strHigh) = 0) Or (Val(strLow) = 0 And Val(strHigh) = 999) Or (Val(strLow) = 0 And Val(strHigh) = 9999) Then
2240                                  lp(lpc) = lp(lpc) & "              " & cUnits
2250                                  udtPrintLine(lpc).NormalRange = "             "
2260                              Else
                                      'Zyam
2270                                  lp(lpc) = lp(lpc) & "              " & cUnits
2280                                  lp(lpc) = lp(lpc) & IIf(Val(strLow) = 0 And Val(strHigh) = 0, " ", "   (")
2290                                  lp(lpc) = lp(lpc) & IIf(Val(strLow) = 0 And Val(strHigh) = 0, " ", strLow & "-")
2300                                  lp(lpc) = lp(lpc) & IIf(Val(strLow) = 0 And Val(strHigh) = 0, " ", strHigh & ")   ")
                                      'Zyam
                                      If Val(strLow) = 0 And Val(strHigh) = 0 Then
                                        udtPrintLine(lpc).NormalRange = " "
                                      Else
                                        udtPrintLine(lpc).NormalRange = "(" & strLow & "-" & strHigh & ")"
                                      End If
                                      'Zyam
2310
2320                              End If
2330                          Else
2340                              lp(lpc) = lp(lpc) & "                 " & cUnits
2350                          End If
2360                      Else
2370                          If br.LongName = ("IgG") Or br.LongName = ("IgA") Or br.LongName = ("IgM") And br.Result = "Paraprotein" Then

2380                          Else
2390                              lp(lpc) = lp(lpc) & "                " & br.Units
2400                              If br.PrnRR = True Then
2410                                  'Zyam
                                      lp(lpc) = lp(lpc) & IIf(Val(strLow) = 0 And Val(strHigh) = 0, " ", "   (")
2299                                  lp(lpc) = lp(lpc) & IIf(Val(strLow) = 0 And Val(strHigh) = 0, " ", strLow & "-")
2499                                  lp(lpc) = lp(lpc) & IIf(Val(strLow) = 0 And Val(strHigh) = 0, " ", strHigh & ")   ")
                                      '11-15-23 Zyam
                                      If Val(strLow) = 0 And Val(strHigh) = 0 Then
                                        udtPrintLine(lpc).NormalRange = " "
                                      Else
                                        udtPrintLine(lpc).NormalRange = "(" & strLow & "-" & strHigh & ")"
                                      End If
                                      '11-15-23 Zyam
2440
2450                              End If
2460                          End If
2470                      End If
2480                      If ImmTestAffected(br) = True Then
2490                          lp(lpc) = lp(lpc) & " " & ImmReasonAffect(br)
2500                      End If
2510                  End If
2520                  LogTestAsPrinted "Imm", br.SampleID, br.Code
2530                  lpc = lpc + 1
2540              End If
2550          Next
2560      End If

2570      If TestCount > Val(frmMain.lblImmMoreThan) Then
2580          PageNumber = "Page 1 og 2"
2590      Else
2600          PageNumber = "Page 1 of 1"
2610      End If

2620      If RP.FaxNumber <> "" Then
2630          PrintHeadingRTBFax
2640      Else
2650          PrintHeadingRTB (PageNumber)
2660      End If

2670      Sex = tb!Sex & ""

2680      If RP.FaxNumber <> "" Then
2690          Fontz1 = 8
2700          Fontz2 = 12
2710      Else
2720          Fontz1 = 10
2730          Fontz2 = 14
2740      End If

2750      With frmRichText.rtb
2760          .SelFontSize = Fontz1

              '2390      If TestCount > Val(frmMain.lblImmMoreThan) Then
              '2400          .SelText = Space$(40) & "Page 1 of 2" & vbCrLf
              '2410      Else
              '2420          .SelText = Space$(40) & "Page 1 of 1" & vbCrLf
              '2430      End If
              '2440      CrCnt = CrCnt + 1

2770          n = Row_Count(lpc)

              '2210    For C = 1 To n
              '2220      .SelText = vbCrLf
              '2230      CrCnt = CrCnt + 1
              '2240    Next

2780          .SelFontName = "Courier New"
2790          .SelFontSize = Fontz1

2800          For n = 0 To Val(frmMain.lblImmMoreThan)
2810              If Trim(lp(n)) <> "" Then

2820                  If Trim(lp(n)) = "SERUM" Or Trim(lp(n)) = "URINE" Then
2830                      .SelBold = True
2840                      .SelFontSize = Fontz2
2850                  Else
2860                      .SelBold = False
2870                      .SelFontName = "Courier New"
2880                      .SelFontSize = Fontz1
2890                  End If
2900                  If InStr(lp(n), " L ") Or InStr(lp(n), " H ") Or InStr(UCase(lp(n)), "POSITIVE ") Then
2910                      If InStr(lp(n), " L ") Then
2920                          .SelColor = vbBlue
2930                      ElseIf InStr(lp(n), " H ") Then
2940                          .SelColor = vbBlack
2950                      End If
2960                      .SelBold = True
2970                      .SelText = lp(n) & vbCrLf
2980                      .SelBold = False
                          '2730                  .SelText = Left(lp(n), 34)
                          '
                          '2750                  .SelText = Mid(lp(n), 35, 3)
                          '2760                  .SelBold = False
                          '2770                  .SelText = Mid(lp(n), 38) & vbCrLf
2990                      CrCnt = CrCnt + 1
3000                  Else
3010                      .SelColor = vbBlack
3020                      If Len(lp(n)) > 80 Then
3030                          .SelText = Left(lp(n), 28)
3040                          FillCommentLines Mid(lp(n), 29, Len(lp(n)) - 28), 4, Comments(), 60
3050                          For z = 1 To 4
3060                              If z = 1 Then
3070                                  If Trim(Comments(z) & "") <> "" Then
3080                                      .SelFontName = "Courier New"
3090                                      .SelFontSize = Fontz1
3100                                      .SelText = Comments(z) & vbCrLf
3110                                      CrCnt = CrCnt + 1
3120                                  End If
3130                              Else
3140                                  If Trim(Comments(z) & "") <> "" Then
3150                                      .SelFontName = "Courier New"
3160                                      .SelFontSize = Fontz1
3170                                      .SelText = "                            " & Comments(z) & vbCrLf
3180                                      CrCnt = CrCnt + 1
3190                                  End If
3200                              End If
3210                          Next
3220                      Else
3230                          If (InStr(UCase(lp(n)), "CHLAMYDIA") And InStr(UCase(lp(n)), "DETECTED") _
                                  And Not InStr(UCase(lp(n)), "NOT")) Then
3240                              .SelText = Left(lp(n), InStr(1, UCase(lp(n)), "DETECTED") - 1)
3250                              .SelBold = True
3260                              .SelText = Mid(lp(n), InStr(1, UCase(lp(n)), "DETECTED")) & vbCrLf
3270                          Else
3280                              .SelFontName = "Courier New"
3290                              .SelFontSize = Fontz1
3300                              .SelText = lp(n) & vbCrLf
3310                          End If
3320                          CrCnt = CrCnt + 1
3330                          .SelBold = False
3340                      End If
3350                  End If
                      'Comments
3360                  If lc(n) <> "" Then
3370                      .SelFontName = "Courier New"
3380                      .SelFontSize = Fontz1
3390                      .SelItalic = True
3400                      .SelBold = True
3410                      FillCommentLines lc(n), 4, Comments(), 85
3420                      For z = 1 To 4
3430                          If z = 1 Then
3440                              If Trim(Comments(z) & "") <> "" Then
3450                                  .SelFontName = "Courier New"
3460                                  .SelFontSize = Fontz1
3470                                  .SelText = " -> " & Comments(z) & vbCrLf
3480                                  CrCnt = CrCnt + 1
3490                              End If
3500                          Else
3510                              If Trim(Comments(z) & "") <> "" Then
3520                                  .SelBold = True
3530                                  .SelText = "   " & Comments(z) & vbCrLf
3540                                  CrCnt = CrCnt + 1
3550                              End If
3560                          End If
3570                      Next
3580                      .SelItalic = False
3590                      .SelBold = False
3600                  End If
3610              End If
3620          Next

3630          .SelFontName = "Courier New"
3640          .SelFontSize = Fontz1

3650          Do While CrCnt < 26
3660              .SelText = vbCrLf
3670              CrCnt = CrCnt + 1
3680          Loop

3690          If TestCount <= Val(frmMain.lblImmMoreThan) Then
                  '    Set Cx = Cxs.Load(RP.SampleID)
3700              Set OBS = OBS.Load(RP.SampleID, "Immunology", "Demographic")
3710              If Not OBS Is Nothing Then
3720                  For Each OB In OBS
3730                      Select Case UCase$(OB.Discipline)
                          Case "IMMUNOLOGY"
3740                          FillCommentLines OB.Comment, 8, Comments(), 80
3750                          For n = 1 To 8
3760                              If Trim(Comments(n) & "") <> "" Then
3770                                  .SelFontName = "Courier New"
3780                                  .SelFontSize = Fontz1
3790                                  .SelText = Comments(n) & vbCrLf
3800                                  CrCnt = CrCnt + 1
3810                              End If
3820                          Next
3830                      Case "DEMOGRAPHIC"
3840                          FillCommentLines OB.Comment, 2, Comments(), 80
3850                          For n = 1 To 2
3860                              If Trim(Comments(n) & "") <> "" Then
3870                                  .SelText = Comments(n) & vbCrLf
3880                                  CrCnt = CrCnt + 1
3890                              End If
3900                          Next
3910                      End Select
3920                  Next
3930              End If
3940          End If

3950          If Not IsDate(DoB) Or Trim(udtHeading.Sex) = "" Then
3960              .SelColor = vbBlack
3970              .SelText = "**** No Sex/DoB given. No reference range applied ****" & vbCrLf
3980              .SelText = vbCrLf
3990              CrCnt = CrCnt + 1
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
4000          End If
4010          .SelColor = vbBlack

4020          If IsDate(tb!SampleDate) Then
4030              SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
4040          Else
4050              SampleDate = ""
4060          End If
4070          If IsDate(RunTime) Then
4080              Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
4090          Else
4100              If IsDate(tb!Rundate) Then
4110                  Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
4120              Else
4130                  Rundate = ""
4140              End If
4150          End If

4160          If RP.FaxNumber <> "" Then
4170              PrintFooterRTBFax RP.Initiator, SampleDate, Rundate
4180          Else
                  '3980          PrintFooterRTB AuthorisedBy, SampleDate, Rundate, ExternalTestingNote
4190              If UCase(GetOptionSetting("GetLatestAuthorisedBy", "")) = UCase("True") Then
4200                  PrintFooterRTB GetLatestAuthorisedBy("Imm", RP.SampleID), SampleDate, GetLatestRunDateTime("Imm", RP.SampleID, Rundate), ExternalTestingNote
4210              Else
4220                  PrintFooterRTB AuthorisedBy, SampleDate, GetLatestRunDateTime("Imm", RP.SampleID, Rundate), ExternalTestingNote
4230              End If

4240          End If

4250          .SelStart = 0
4260          If RP.FaxNumber <> "" Then
4270              f = FreeFile
4280              Open SysOptFax(0) & RP.SampleID & "IMM1.doc" For Output As f
4290              Print #f, .TextRTF
4300              Close f
4310              SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "IMM1.doc"
4320          Else
                  'Do not print if Doctor is disabled in DisablePrinting
                  '*******************************************************************
4330              If CheckDisablePrinting(RP.Ward, "Immunology") Then

4340              ElseIf CheckDisablePrinting(RP.GP, "Immunology") Then
4350              Else
4360                  .SelPrint Printer.hdc
4370              End If
                  '*******************************************************************
                  '.SelPrint Printer.hDC
4380          End If
4390          sql = "SELECT * FROM Reports WHERE 0 = 1"
4400          Set tb = New Recordset
4410          RecOpenServer 0, tb, sql
4420          tb.AddNew
4430          tb!SampleID = RP.SampleID
4440          tb!Name = udtHeading.Name
4450          tb!Dept = "I"
4460          tb!Initiator = RP.Initiator
4470          tb!PrintTime = PrintTime
4480          tb!RepNo = "0I" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
4490          tb!PageNumber = 0
4500          tb!Report = .TextRTF
4510          tb!Printer = Printer.DeviceName
4520          tb.Update
4530      End With

          '##########################
          'Print Second Page
4540      CrCnt = 0
4550      If TestCount > Val(frmMain.lblImmMoreThan) Then

4560          If RP.FaxNumber <> "" Then
4570              PrintHeadingRTBFax
4580          Else
4590              PrintHeadingRTB ("Page 2 of 2")
4600          End If
4610          With frmRichText.rtb
4620              .SelFontName = "Courier New"
4630              .SelFontSize = Fontz1

                  '4260          .SelText = Space(39) & "Page 2 of 2" & vbCrLf
                  '4270          CrCnt = CrCnt + 1

                  '    n = Row_Count(lpc - Val(frmMain.lblimmmorethan))

                  '    For C = 1 To n
                  '      .SelText = vbCrLf
                  '      CrCnt = CrCnt + 1
                  '    Next

4640              For n = Val(frmMain.lblImmMoreThan) + 1 To Val(lpc)
4650                  .SelFontName = "Courier New"
4660                  .SelFontSize = Fontz1
4670                  If InStr(lp(n), " L ") Or InStr(lp(n), " H ") Then
4680                      If InStr(lp(n), " L ") Then
4690                          .SelColor = vbBlue
4700                      ElseIf InStr(lp(n), " H ") Then
4710                          .SelColor = vbBlack
4720                      End If
4730                      .SelText = Left(lp(n), 34)
4740                      .SelBold = True
4750                      .SelText = Mid(lp(n), 35, 3)
4760                      .SelBold = False
4770                      .SelText = Mid(lp(n), 38) & vbCrLf
4780                      CrCnt = CrCnt + 1
4790                  Else
4800                      .SelColor = vbBlack
4810                      If Len(lp(n)) > 70 Then
4820                          .SelText = Left(lp(n), 28)
4830                          FillCommentLines Mid(lp(n), 29, Len(lp(n)) - 28), 4, Comments(), 65
4840                          For z = 1 To 4
4850                              If z = 1 Then
4860                                  If Trim(Comments(z) & "") <> "" Then
4870                                      .SelFontName = "Courier New"
4880                                      .SelFontSize = Fontz1
4890                                      .SelText = Comments(z) & vbCrLf
4900                                      CrCnt = CrCnt + 1
4910                                  End If
4920                              Else
4930                                  If Trim(Comments(z) & "") <> "" Then
4940                                      .SelFontName = "Courier New"
4950                                      .SelFontSize = Fontz1
4960                                      .SelText = "                               " & Comments(z) & vbCrLf
4970                                      CrCnt = CrCnt + 1
4980                                  End If
4990                              End If
5000                          Next
5010                      Else
5020                          If (InStr(UCase(lp(n)), "CHLAMYDIA") And InStr(UCase(lp(n)), "DETECTED") _
                                  And Not InStr(UCase(lp(n)), "NOT")) Then
5030                              .SelText = Left(lp(n), InStr(1, UCase(lp(n)), "DETECTED") - 1)
5040                              .SelBold = True
5050                              .SelText = Mid(lp(n), InStr(1, UCase(lp(n)), "DETECTED")) & vbCrLf
5060                          Else
5070                              .SelText = lp(n) & vbCrLf
5080                          End If
5090                          CrCnt = CrCnt + 1
5100                      End If
5110                  End If
5120                  If lc(n) <> "" Then
5130                      .SelItalic = True
5140                      .SelBold = True
5150                      FillCommentLines lc(n), 4, Comments(), 80
5160                      For z = 1 To 4
5170                          If z = 1 Then
5180                              If Trim(Comments(z) & "") <> "" Then
5190                                  .SelFontName = "Courier New"
5200                                  .SelFontSize = Fontz1
5210                                  .SelText = " -> " & Comments(z) & vbCrLf
5220                                  CrCnt = CrCnt + 1
5230                              End If
5240                          Else
5250                              If Trim(Comments(z) & "") <> "" Then
5260                                  .SelFontName = "Courier New"
5270                                  .SelFontSize = Fontz1
5280                                  .SelText = "    " & Comments(z) & vbCrLf
5290                                  CrCnt = CrCnt + 1
5300                              End If
5310                          End If
5320                      Next
5330                      .SelItalic = False
5340                      .SelBold = False
5350                  End If
5360              Next

5370              Do While CrCnt < 26
5380                  .SelText = vbCrLf
5390                  CrCnt = CrCnt + 1
5400              Loop
                  '    Set Cx = Cxs.Load(RP.SampleID)
5410              Set OBS = OBS.Load(RP.SampleID, "Immunology", "Demographic")

5420              If Not OBS Is Nothing Then
5430                  For Each OB In OBS
5440                      Select Case UCase$(OB.Discipline)
                          Case "IMMUNOLOGY"
5450                          FillCommentLines OB.Comment, 8, Comments(), 80
5460                          For n = 1 To 8
5470                              If Trim(Comments(n) & "") <> "" Then
5480                                  .SelFontName = "Courier New"
5490                                  .SelFontSize = Fontz1
5500                                  .SelText = Comments(n) & vbCrLf
5510                                  CrCnt = CrCnt + 1
5520                              End If
5530                          Next
5540                      Case "DEMOGRAPHIC"
5550                          FillCommentLines OB.Comment, 2, Comments(), 80
5560                          For n = 1 To 2
5570                              If Trim(Comments(n) & "") <> "" Then
5580                                  .SelFontName = "Courier New"
5590                                  .SelFontSize = Fontz1
5600                                  .SelText = Comments(n) & vbCrLf
5610                                  CrCnt = CrCnt + 1
5620                              End If
5630                          Next
5640                      End Select
5650                  Next
5660              End If

5670              .SelFontName = "Courier New"
5680              .SelFontSize = Fontz1

5690              If Not IsDate(DoB) Or Trim(udtHeading.Sex) = "" Then
5700                  .SelColor = vbBlack
5710                  .SelText = "**** No Sex/DoB given. No reference range applied ****" & vbCrLf
5720                  .SelText = vbCrLf
5730                  CrCnt = CrCnt + 1
                      '        ElseIf Not IsDate(DoB) Then
                      '            .SelColor = vbBlack
                      '            .SelText = "*** No Dob. Adult Age 25 used for Normal Ranges! ***" & vbCrLf
                      '            .SelText = vbCrLf
                      '            CrCnt = CrCnt + 1
                      '        ElseIf Trim(udtHeading.Sex) = "" Then
                      '            .SelColor = vbBlack
                      '            .SelText = "No Sex given. No reference range applied" & vbCrLf
                      '            .SelText = vbCrLf
                      '            CrCnt = CrCnt + 1
5740              End If

                  '5160      If IsDate(tb!SampleDate) Then
                  '5170        SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
                  '5180      Else
                  '5190        SampleDate = ""
                  '5200      End If
                  '5210      If IsDate(RunTime) Then
                  '5220        Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
                  '5230      Else
                  '5240        If IsDate(tb!Rundate) Then
                  '5250          Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
                  '5260        Else
                  '5270          Rundate = ""
                  '5280        End If
                  '5290      End If

5750              If RP.FaxNumber <> "" Then
5760                  PrintFooterRTBFax RP.Initiator, SampleDate, Rundate
5770              Else
                      '5510              PrintFooterRTB AuthorisedBy, SampleDate, Rundate
5780                  If UCase(GetOptionSetting("GetLatestAuthorisedBy", "")) = UCase("True") Then
5790                      PrintFooterRTB GetLatestAuthorisedBy("IMM", RP.SampleID), SampleDate, GetLatestRunDateTime("IMM", RP.SampleID, Rundate)
5800                  Else
5810                      PrintFooterRTB AuthorisedBy, SampleDate, GetLatestRunDateTime("Imm", RP.SampleID, Rundate)
5820                  End If
5830                  If ExternalTestingNote <> "" Then
5840                      PrintTextRTB frmRichText.rtb, vbNewLine & " " & ExternalTestingNote
5850                  End If
5860              End If

5870              .SelStart = 0
5880              If RP.FaxNumber <> "" Then
5890                  f = FreeFile
5900                  Open SysOptFax(0) & RP.SampleID & "IMM2.doc" For Output As f
5910                  Print #f, .TextRTF
5920                  Close f
5930                  SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "IMM2.doc"
5940              Else
                      'Do not print if Doctor is disabled in DisablePrinting
                      '*******************************************************************
5950                  If CheckDisablePrinting(RP.Ward, "Immunology") Then

5960                  ElseIf CheckDisablePrinting(RP.GP, "Immunology") Then
5970                  Else
5980                      .SelPrint Printer.hdc
5990                  End If
                      '*******************************************************************
                      '.SelPrint Printer.hDC
6000              End If
6010              sql = "SELECT * FROM Reports WHERE 0 = 1"
6020              Set tb = New Recordset
6030              RecOpenServer 0, tb, sql
6040              tb.AddNew
6050              tb!SampleID = RP.SampleID
6060              tb!Name = udtHeading.Name
6070              tb!Dept = "I"
6080              tb!Initiator = RP.Initiator
6090              tb!PrintTime = PrintTime
6100              tb!RepNo = "1I" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
6110              tb!PageNumber = 1
6120              tb!Report = .TextRTF
6130              tb!Printer = Printer.DeviceName
6140              tb.Update
6150          End With
6160      End If

6170      Exit Sub

PrintResultImmWin_Error:

          Dim strES As String
          Dim intEL As Integer

6180      intEL = Erl
6190      strES = Err.Description
6200      LogError "modImmunology", "PrintResultImmWin", intEL, strES

End Sub

Private Function AllergyResultFor(ByVal BRs As BIEResults, _
                                  ByVal AllergyName As String) _
                                  As String

      Dim br As BIEResult
      Dim AName As String

10    On Error GoTo AllergyResultFor_Error

20    AName = UCase$(AllergyName)

30    AllergyResultFor = ""
40    For Each br In BRs
50        If UCase$(br.LongName) = AName Then
60            AllergyResultFor = br.Result
70            Exit For
80        End If
90    Next

100   Exit Function

AllergyResultFor_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modImmunology", "AllergyResultFor", intEL, strES

End Function

Private Function AllergyFlagFor(ByVal BRs As BIEResults, _
                                ByVal AllergyName As String) _
                                As String

    Dim br As BIEResult
    Dim AName As String
    Dim v As String

10  On Error GoTo AllergyFlagFor_Error

20  AName = UCase$(AllergyName)
30  AllergyFlagFor = ""
40  For Each br In BRs
50      If UCase$(br.LongName) = AName Then
60          v = br.Result
            'If IsAllergyTest(Br.Code) And Br.Result > 0.35 Then
70          If IsAllergyTest(br.Code) And Left(v, 1) = ">" And Trim(Mid(v, 2)) >= 0.35 Then
80              AllergyFlagFor = " H"
90              Exit For
100         End If
110         If IsNumeric(v) And udtHeading.Sex <> "" And IsDate(udtHeading.DoB) Then
120             If Val(v) > br.PlausibleHigh Then
130                 AllergyFlagFor = " X"
140             ElseIf Val(v) < br.PlausibleLow Then
150                 AllergyFlagFor = " X"
160             ElseIf Val(v) > br.High And br.High <> 0 Then
170                 AllergyFlagFor = " H"
180             ElseIf Val(v) < br.Low Then
190                 AllergyFlagFor = " L"
200             Else
210                 AllergyFlagFor = "  "
220             End If
230         ElseIf IsNumeric(v) = False And udtHeading.Sex <> "" And IsDate(udtHeading.DoB) Then
240             If Left(v, 1) = "<" And Trim(Mid(v, 2)) <= br.Low Then
250                 AllergyFlagFor = " L"
260             ElseIf Left(v, 1) = ">" And Trim(Mid(v, 2)) >= br.High Then
270                 AllergyFlagFor = " H"
280             End If
290         Else
300             AllergyFlagFor = "  "
310         End If
320         Exit For
330     End If
340 Next

350 Exit Function

AllergyFlagFor_Error:

    Dim strES As String
    Dim intEL As Integer

360 intEL = Erl
370 strES = Err.Description
380 LogError "modImmunology", "AllergyFlagFor", intEL, strES

End Function


Private Function AllergyUnitsFor(ByVal BRs As BIEResults, _
                                 ByVal AllergyName As String) _
                                 As String

      Dim br As BIEResult
      Dim AName As String

10    On Error GoTo AllergyUnitsFor_Error

20    AName = UCase$(AllergyName)

30    AllergyUnitsFor = ""
40    For Each br In BRs
50        If UCase$(br.LongName) = AName Then
60            AllergyUnitsFor = br.Units
70            Exit For
80        End If
90    Next

100   Exit Function

AllergyUnitsFor_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modImmunology", "AllergyUnitsFor", intEL, strES

End Function

Private Function IsAllergyTest(Code As String) As Boolean

    Dim sql As String
    Dim tb As Recordset

10  On Error GoTo IsAllergyTest_Error

20  sql = "SELECT Count(*) AS Cnt FROM ImmTestDefinitions WHERE IsAllergy = 1 AND Code = '" & Code & "'"
30  Set tb = New Recordset
40  RecOpenServer 0, tb, sql
50  IsAllergyTest = tb!Cnt > 0

60  Exit Function

IsAllergyTest_Error:

    Dim strES As String
    Dim intEL As Integer

70  intEL = Erl
80  strES = Err.Description
90  LogError "modImmunology", "IsAllergyTest", intEL, strES, sql

End Function

Private Sub GetContents(ByVal PanelName As String, _
                        ByRef ContentLine As Collection)

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer

10    On Error GoTo GetContents_Error

20    sql = "SELECT Content FROM IPanels WHERE " & _
            "PanelName = '" & PanelName & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    n = -1
60    Do While Not tb.EOF
70        ContentLine.Add tb!Content & "", tb!Content & ""
80        tb.MoveNext
90    Loop

100   Exit Sub

GetContents_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modImmunology", "GetContents", intEL, strES, sql

End Sub

Private Function IsHemochromatosis(ByVal BRs As BIEResults) As Boolean

      Dim br As BIEResult
      Dim IsHemo As Boolean

10    On Error GoTo IsHemochromatosis_Error

20    IsHemochromatosis = False
30    For Each br In BRs
40        If br.Code = "C28" Or br.Code = "H63" Then
50            IsHemochromatosis = True
60            Exit For
70        End If
80    Next

90    Exit Function

IsHemochromatosis_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modImmunology", "IsHemochromatosis", intEL, strES

End Function

Private Function IsAllergy(ByVal BRs As BIEResults) As Boolean

      Dim br As BIEResult
      Dim tb As Recordset
      Dim sql As String
      Dim RetVal As Boolean
      Dim Content As String

10    On Error GoTo IsAllergy_Error

20    RetVal = False
30    Content = ""

40    For Each br In BRs
50        If UCase$(Mid$(br.Code, 2, 1)) = "X" Then
60            RetVal = True
70            Exit For
80        Else
90            Content = Content & "LongName = '" & br.LongName & "' OR "
100       End If
110   Next

120   If Not RetVal Then
130       Content = Left$(Content, Len(Content) - 3)
140       sql = "SELECT COUNT(*) AS Tot FROM ImmTestDefinitions WHERE " & _
                "IsAllergy = 1 " & _
                "AND (" & Content & ")"
150       Set tb = New Recordset
160       RecOpenServer 0, tb, sql
170       RetVal = tb!Tot > 0
180   End If

190   IsAllergy = RetVal

200   Exit Function

IsAllergy_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "modImmunology", "IsAllergy", intEL, strES

End Function

'Private Sub PrintResultAllergy(ByVal BRs As BIEResults)
'
'      Dim tb As Recordset
'      Dim sql As String
'      Dim Br As BIEResult
'      'Dim Cx As Comment
'      'Dim Cxs As New Comments
'      Dim OB As Observation
'      Dim OBS As New Observations
'      Dim SampleDate As String
'      Dim Rundate As String
'      Dim Clin As String
'      Dim f As Integer
'      Dim PrintTime As String
'      Dim ContentLine As New Collection
'      Dim strContent As String
'      Dim s As Variant
'      Dim AuthorisedBy As String
'
'10    On Error GoTo PrintResultAllergy_Error
'
'20    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
'
'30    sql = "SELECT * FROM Demographics WHERE " & _
'            "SampleID = '" & RP.SampleID & "'"
'40    Set tb = New Recordset
'50    RecOpenClient 0, tb, sql
'
'60    ClearUdtHeading
'70    With udtHeading
'80        .SampleID = RP.SampleID
'90        .Dept = "Immunology"
'100       .Name = tb!PatName & ""
'110       .Ward = RP.Ward
'120       .Chart = tb!Chart & ""
'130       .Sex = tb!Sex & ""
'140       .DoB = tb!DoB & ""
'150       .Clinician = RP.Clinician
'160       .Address0 = tb!Addr0 & ""
'170       .Address1 = tb!Addr1 & ""
'180       .GP = RP.GP
'190       .Hospital = tb!Hospital & ""
'200       .SampleDate = tb!SampleDate & ""
'210       .RecDate = tb!RecDate & ""
'220       .Rundate = tb!Rundate & ""
'230       .GpClin = Clin
'240       .DocumentNo = GetOptionSetting("ImmAllergyDocumentNo", "")
'250   End With
'
'260   frmRichText.rtb.Text = ""
'
'270   If RP.FaxNumber <> "" Then
'280       PrintHeadingRTBFax
'290   Else
'300       PrintHeadingRTB
'310   End If
'
'320   With frmRichText.rtb
'330       .SelBold = False
'340       .SelFontSize = 10
'350       .SelBold = False
'360       .SelText = Space(35) & "Page 1 of 1" & vbCrLf
'
'370       .SelFontSize = 10
'380       .SelBold = True
'390       .SelUnderline = True
'400       .SelText = "Allergy Test"
'410       .SelUnderline = False
'420       .SelText = Space$(28)
'430       .SelUnderline = True
'440       .SelText = "Result"
'450       .SelUnderline = False
'460       .SelText = "    "
'470       .SelUnderline = True
'480       .SelText = "Units" & vbCrLf & vbCrLf
'490       If Not BRs Is Nothing Then
'500           AuthorisedBy = GetAuthorisedBy(GetLastValidatedBy(BRs))
'510       End If
'520       For Each Br In BRs
'530           If AuthorisedBy = "" Then AuthorisedBy = GetAuthorisedBy(Br.Operator)
'540           If UCase$(Mid$(Br.Code, 2, 1)) = "X" Then
'550               .SelBold = True
'560               .SelFontSize = 10
'570               .SelText = Left$(Br.LongName & Space$(40), 40)
'580               .SelText = Br.Result & vbCrLf
'
'590               Set ContentLine = New Collection
'600               GetContents Br.LongName, ContentLine
'                  'GetContents BR.LongName, AlreadyPrinted
'
'610               .SelFontSize = 8
'620               .SelBold = False
'630               .SelText = "("
'640               strContent = ""
'650               For Each s In ContentLine
'660                   strContent = strContent & s & ", "
'670               Next
'680               strContent = Left$(strContent, Len(strContent) - 2)
'690               .SelText = strContent
'700               .SelText = ")" & vbCrLf
'710           Else
'720               .SelFontSize = 10
'730               .SelBold = False
'740               .SelText = Left$(Br.LongName & Space$(40), 40)
'750               .SelText = Left$(AllergyResultFor(BRs, Br.LongName) & Space$(10), 10)
'760               .SelText = Left$(AllergyUnitsFor(BRs, Br.LongName) & Space$(10), 10)
'770               If udtHeading.Sex <> "" And IsDate(udtHeading.DoB) Then
'780                   If Br.Code = "a-IgE" Then .SelText = Br.Low & " - " & Br.High
'790               End If
'800               .SelText = vbCrLf
'                  '          If UCase$(BR.Result) = "POSITIVE" Then
'                  '            For Each s In ContentLine
'                  '              .SelFontSize = 10
'                  '              .SelBold = False
'                  '              .SelText = Left$(s & Space$(40), 40)
'                  '              .SelText = Left$(AllergyResultFor(BRs, s) & Space$(10), 10)
'                  '              .SelText = AllergyUnitsFor(BRs, s)
'                  '              .SelText = vbCrLf
'                  '            Next
'                  '          End If
'
'810           End If
'820           LogTestAsPrinted "Imm", Br.SampleID, Br.Code
'830       Next
'
'          '    For x = BRs.Count To 1 Step -1
'          '        Set BR = BRs(x)
'          '        If UCase$(Mid$(BR.Code, 2, 1)) = "X" Then
'          '            Set BR = Nothing
'          '            BRs.RemoveItem x
'          '        End If
'          '    Next
'          '
'          '  For Each BR In BRs
'          '    .SelFontSize = 10
'          '    .SelBold = False
'          '    .SelText = Left$(BR.LongName & Space$(40), 40)
'          '    .SelText = Left$(AllergyResultFor(BRs, BR.LongName) & Space$(10), 10)
'          '    .SelText = Left$(AllergyUnitsFor(BRs, BR.LongName) & Space$(10), 10)
'          '  If BR.Code = "a-IgE" Then .SelText = BR.Low & " - " & BR.High
'          '
'          '    .SelText = vbCrLf
'          '  Next
'840       .SelText = vbCrLf
'850       .SelFontSize = 9
'860       .SelBold = True
'870       .SelText = Space(5) & "Interpretive Comment for Specific Allergens (Not Total IgE):" & vbCrLf
'880       .SelBold = True
'890       .SelFontSize = 2
'900       .SelText = Space(20) & String(400, "-") & vbCrLf
'910       .SelFontSize = 9
'920       .SelBold = True
'930       .SelText = Left(Space(5) & "Ref Range kUA/L" & Space(20), 25) & "Clinical Implications" & vbCrLf
'940       .SelFontSize = 2
'950       .SelBold = True
'960       .SelText = Space(20) & String(400, "-") & vbCrLf
'
'970       .SelBold = False
'980       .SelFontSize = 9
'990       .SelText = Left(Space(5) & "< 0.1" & Space(20), 25) & "Negative/Absent/Undetectable." & vbCrLf
'1000      .SelFontSize = 2
'1010      .SelText = Space(20) & String(400, "-") & vbCrLf
'1020      .SelFontSize = 9
'1030      .SelText = Left(Space(5) & "0.10 - 0.35" & Space(20), 25) & "For Specialist use only; Clinical relevance undetermined." & vbCrLf
'1040      .SelFontSize = 2
'1050      .SelText = Space(20) & String(400, "-") & vbCrLf
'1060      .SelFontSize = 9
'1070      .SelText = Left(Space(5) & "0.35 - 0.70" & Space(20), 25) & "Low level of allergy; Indicative of ongoing sensitization." & vbCrLf
'1080      .SelFontSize = 2
'1090      .SelText = Space(20) & String(400, "-") & vbCrLf
'1100      .SelFontSize = 9
'1110      .SelText = Left(Space(5) & "0.70 - 3.50" & Space(20), 25) & "Moderate level of allergy; Indicative of ongoing sensitization." & vbCrLf
'1120      .SelFontSize = 2
'1130      .SelText = Space(20) & String(400, "-") & vbCrLf
'1140      .SelFontSize = 9
'1150      .SelText = Left(Space(5) & "3.50 - 17.5" & Space(20), 25) & "High level of allergy; Indicative of high level sensitization." & vbCrLf
'1160      .SelFontSize = 2
'1170      .SelText = Space(20) & String(400, "-") & vbCrLf
'1180      .SelFontSize = 9
'1190      .SelText = Left(Space(5) & "> 17.5 " & Space(20), 25) & "Very High level of allergy; Indicative of very high level sensitization." & vbCrLf
'1200      .SelFontSize = 2
'1210      .SelText = Space(20) & String(400, "-") & vbCrLf
'1220      .SelText = vbCrLf
'1230      .SelFontSize = 9
'1240      .SelText = Space(5) & Chr(149) & " For all positive specific IgE results, please interpret in context of the" & vbCrLf
'1250      .SelFontSize = 9
'1260      .SelText = Space(5) & "  clinical history." & vbCrLf
'1270      .SelText = Space(5) & Chr(149) & " High Total IgE levels can result in low level positivity (up to 3.5 kAU/l)" & vbCrLf
'1280      .SelFontSize = 9
'1290      .SelText = Space(5) & "  in specific IgE tests. This is particularly the case when the Total IgE is more" & vbCrLf
'1300      .SelFontSize = 9
'1310      .SelText = Space(5) & "  than 1000 kU/l. Please interpret the results of the specific IgE tests in context" & vbCrLf
'1320      .SelFontSize = 9
'1330      .SelText = Space(5) & "  of the clinical history." & vbCrLf
'
'          '    .SelText = Space(5) & Chr(149) & " Panels are reported as Positive or Negative only." & vbCrLf
'          '    .SelFontSize = 9
'          '    .SelText = Space(5) & Chr(149) & " Very low levels of specific IgE should be interpreted with caution" & vbCrLf
'          '    .SelFontSize = 9
'          '    .SelText = Space(5) & "  when Total IgE values are above 1000 kU/L." & vbCrLf
'          '    .SelFontSize = 9
'          '    .SelText = Space(5) & Chr(149) & " In food allergy specific IgE may remain undetectable even with a " & vbCrLf
'          '    .SelFontSize = 9
'          '    .SelText = Space(5) & "  convincing clinical history as the antibodies may be directed " & vbCrLf
'          '    .SelFontSize = 9
'          '    .SelText = Space(5) & "  towards allergens that are not revealed or may be altered during " & vbCrLf
'          '    .SelFontSize = 9
'          '    .SelText = Space(5) & "  industrial processing, cooking or digestion." & vbCrLf
'          '  .SelBold = False
'          '  .SelFontSize = 10
'          '  .SelText = Space$(20) & "1. Negative result is defined as <0.10 kIU/L" & vbCrLf
'          '  .SelBold = False
'          '  .SelFontSize = 10
'          '  .SelText = Space$(20) & "2. Panels are reported as Positive or Negative only." & vbCrLf & vbCrLf
'
'
'          '  Set Cx = Cxs.Load(RP.SampleID)
'1340      Set OBS = OBS.Load(RP.SampleID, "Immunology", "Demographic")
'1350      If Not OBS Is Nothing Then
'1360          .SelBold = True
'1370          .SelFontSize = 10
'1380          .SelText = "Comment:-"
'1390          For Each OB In OBS
'1400              .SelBold = False
'1410              .SelText = OB.Comment & vbCrLf
'1420          Next
'1430      End If
'
'1440      .SelFontSize = 10
'1450      .SelBold = False
'
'1460      If RP.FaxNumber <> "" Then
'1470          PrintFooterRTBFax RP.Initiator, SampleDate, Rundate
'1480          .SelStart = 0
'1490          f = FreeFile
'1500          Open SysOptFax(0) & RP.SampleID & "Imm2.doc" For Output As f
'1510          Print #f, .TextRTF
'1520          Close f
'1530          SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "Imm2.doc"
'1540      Else
'1550          PrintFooterA4RTB AuthorisedBy, udtHeading.SampleDate, udtHeading.Rundate
'1560          .SelStart = 0
'
'1570          .SelPrint Printer.hDC
'1580      End If
'
'1590      sql = "SELECT * FROM Reports WHERE 0 = 1"
'1600      Set tb = New Recordset
'1610      RecOpenServer 0, tb, sql
'1620      tb.AddNew
'1630      tb!SampleID = RP.SampleID
'1640      tb!Name = udtHeading.Name
'1650      tb!Dept = "I"
'1660      tb!Initiator = RP.Initiator
'1670      tb!PrintTime = PrintTime
'1680      tb!RepNo = "1A" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
'1690      tb!PageNumber = 0
'1700      tb!Report = .TextRTF
'1710      tb!Printer = Printer.DeviceName
'1720      tb.Update
'1730  End With
'
'1740  ResetPrinter
'
'1750  Exit Sub
'
'PrintResultAllergy_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'1760  intEL = Erl
'1770  strES = Err.Description
'1780  LogError "modImmunology", "PrintResultAllergy", intEL, strES, sql
'
'End Sub

Private Sub PrintResultAllergy(ByVal BRs As BIEResults)

          Dim tb As Recordset
          Dim sql As String
          Dim br As BIEResult
          Dim OB As Observation
          Dim OBS As New Observations
          Dim SampleDate As String
          Dim Rundate As String
          Dim Clin As String
          Dim f As Integer
          Dim PrintTime As String
          Dim ContentLine As New Collection
          Dim strContent As String
          Dim s As Variant
          Dim AuthorisedBy As String
10        ReDim Comments(1 To 8) As String
          Dim i As Integer
          Dim CommentTitle As String
          Dim Flag As String
          Dim BoldLine As Boolean

20        On Error GoTo PrintResultAllergy_Error

30        PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

40        sql = "SELECT * FROM Demographics WHERE " & _
                "SampleID = '" & RP.SampleID & "'"
50        Set tb = New Recordset
60        RecOpenClient 0, tb, sql

70        ClearUdtHeading
80        With udtHeading
90            .SampleID = RP.SampleID
100           .Dept = "Immunology"
110           .Name = tb!PatName & ""
120           .Ward = RP.Ward
130           .Chart = tb!Chart & ""
140           .Sex = tb!Sex & ""
150           .DoB = tb!DoB & ""
160           .Clinician = RP.Clinician
170           .Address0 = tb!Addr0 & ""
180           .Address1 = tb!Addr1 & ""
190           .GP = RP.GP
200           .Hospital = tb!Hospital & ""
210           .SampleDate = tb!SampleDate & ""
220           .RecDate = tb!RecDate & ""
230           .Rundate = tb!Rundate & ""
240           .GpClin = Clin
250           .DocumentNo = GetOptionSetting("ImmAllergyDocumentNo", "")
260       End With

270       frmRichText.rtb.Text = ""

280       If RP.FaxNumber <> "" Then
290           PrintHeadingRTBFax
300       Else
310           PrintHeadingRTB
320       End If

330       With frmRichText.rtb
              'heading
340           PrintTextRTB frmRichText.rtb, FormatString("Allergy Test", 36, " ", AlignLeft), 10, True, , True
350           PrintTextRTB frmRichText.rtb, FormatString("Result", 16, " ", AlignLeft), 10, True, , True
360           PrintTextRTB frmRichText.rtb, FormatString(" ", 3, " ", AlignLeft), 10, True, , True
370           PrintTextRTB frmRichText.rtb, FormatString("Units", 12, " ", AlignLeft), 10, True, , True
              'PrintTextRTB frmRichText.rtb, FormatString("", 12, " ", AlignLeft), 10, True, , True
380           PrintTextRTB frmRichText.rtb, FormatString("Reference Range", 15, " ", AlignLeft), 10, True, , True
              'PrintTextRTB frmRichText.rtb, FormatString("", 12, " ", AlignLeft), 10, True, , True
390           PrintTextRTB frmRichText.rtb, vbCrLf
              'Results
400           If Not BRs Is Nothing Then
410               AuthorisedBy = GetAuthorisedBy(GetLastValidatedBy(BRs))
420           End If
430           For Each br In BRs
440               If AuthorisedBy = "" Then AuthorisedBy = GetAuthorisedBy(br.Operator)
450               If UCase$(Mid$(br.Code, 2, 1)) = "X" Then

                      'PrintTextRTB frmRichText.rtb, FormatString(Br.LongName, Len(Br.LongName), , AlignLeft), 10, , , True
                      'PrintTextRTB frmRichText.rtb, FormatString(" ", 36 - Len(Br.LongName), " "), 10
460                   If IsNumeric(br.Result) Then
470                       If IsAllergyTest(br.Code) And br.Result > 0.35 Then
480                           PrintTextRTB frmRichText.rtb, FormatString(br.LongName, Len(br.LongName), , AlignLeft), 10, True, , True
490                           PrintTextRTB frmRichText.rtb, FormatString(" ", 36 - Len(br.LongName), " "), 10, True
500                           PrintTextRTB frmRichText.rtb, FormatString(br.Result, 16, " " + "H   ", AlignLeft), 10, True
510                           PrintTextRTB frmRichText.rtb, FormatString(br.Units, 12, " ", AlignLeft), 10, True
520                           If br.PrnRR = True Then
530                               PrintTextRTB frmRichText.rtb, FormatString(br.Low & " - " & br.High, 10, " ", AlignLeft), 10, True
540                           End If
550                       Else
560                           PrintTextRTB frmRichText.rtb, FormatString(br.LongName, Len(br.LongName), , AlignLeft), 10, , , True
570                           PrintTextRTB frmRichText.rtb, FormatString(" ", 36 - Len(br.LongName), " "), 10
580                           PrintTextRTB frmRichText.rtb, FormatString(br.Result, 20, " ", AlignLeft), 10
590                           PrintTextRTB frmRichText.rtb, FormatString(br.Units, 12, " ", AlignLeft), 10
600                           If br.PrnRR = True Then
610                               PrintTextRTB frmRichText.rtb, FormatString(br.Low & " - " & br.High, 10, " ", AlignLeft), 10
620                           End If
630                       End If
640                   Else
650                       PrintTextRTB frmRichText.rtb, FormatString(br.LongName, Len(br.LongName), , AlignLeft), 10, , , True
660                       PrintTextRTB frmRichText.rtb, FormatString(" ", 36 - Len(br.LongName), " "), 10
670                       PrintTextRTB frmRichText.rtb, FormatString(br.Result, 20, " ", AlignLeft), 10, FlagAllergyResultHigh(br.Result)
680                       PrintTextRTB frmRichText.rtb, FormatString(br.Units, 12, " ", AlignLeft), 10
690                       If br.PrnRR = True Then
700                           PrintTextRTB frmRichText.rtb, FormatString(br.Low & " - " & br.High, 10, " ", AlignLeft), 10
710                       End If
                          'PrintTextRTB frmRichText.rtb, FormatString(AllergyUnitsFor(BRs, Br.LongName), 12, " ", AlignLeft), 10
720                   End If

730                   PrintTextRTB frmRichText.rtb, vbCrLf
740                   Set ContentLine = New Collection
750                   GetContents br.LongName, ContentLine
760                   strContent = ""
770                   For Each s In ContentLine
780                       strContent = strContent & s & ", "
790                   Next
800                   strContent = Left$(strContent, Len(strContent) - 2)
810                   strContent = "(" & strContent & ")"
820                   PrintTextRTB frmRichText.rtb, FormatString(strContent, 100, " ", AlignLeft), 8
830                   PrintTextRTB frmRichText.rtb, vbCrLf
840               Else
850                   BoldLine = False
860                   If UCase(br.Result) = "POSITIVE" Then
870                       BoldLine = True
880                   End If
890                   Flag = Trim(AllergyFlagFor(BRs, br.LongName))
900                   If (Flag = "H") Then BoldLine = True
                      'BoldLine = IIf(Flag <> "" Or Left$(AllergyResultFor(BRs, Br.LongName), 1) = ">" Or Left$(AllergyResultFor(BRs, Br.LongName), 1) = "<", True, False)
910                   PrintTextRTB frmRichText.rtb, FormatString(br.LongName, 36, " ", AlignLeft), 10, BoldLine
920                   PrintTextRTB frmRichText.rtb, FormatString(AllergyResultFor(BRs, br.LongName), 16, " ", AlignLeft), 10, BoldLine
                      '            If IsNumeric(Br.Result) Then
                      '                    If IsAllergyTest(Br.Code) And Br.Result > 0.35 Then
                      '                        PrintTextRTB frmRichText.rtb, FormatString(AllergyResultFor(BRs, Br.LongName), 16, " ", AlignLeft), 10, True
                      '                    Else
                      '                        PrintTextRTB frmRichText.rtb, FormatString(AllergyResultFor(BRs, Br.LongName), 16, " ", AlignLeft), 10
                      '                    End If
                      '            Else
                      '
                      '                PrintTextRTB frmRichText.rtb, FormatString(AllergyResultFor(BRs, Br.LongName), 16, " ", AlignLeft), 10, FlagAllergyResultHigh(Br.Result)
                      '            End If

930                   PrintTextRTB frmRichText.rtb, FormatString(Flag, 3, " ", AlignLeft), 10, IIf(Flag = "", False, True)
940                   PrintTextRTB frmRichText.rtb, FormatString(AllergyUnitsFor(BRs, br.LongName), 12, " ", AlignLeft), 10, BoldLine
950                   If udtHeading.Sex <> "" And IsDate(udtHeading.DoB) And br.PrnRR = True Then
                          'If Br.Code = "a-IgE" Then
960                       PrintTextRTB frmRichText.rtb, FormatString(br.Low & " - " & br.High, 12, " ", AlignLeft), 10, BoldLine
                          'End If
970                   End If
980                   PrintTextRTB frmRichText.rtb, vbCrLf

990               End If
1000              LogTestAsPrinted "Imm", br.SampleID, br.Code
1010          Next


1020          .SelText = vbCrLf
1030          .SelFontSize = 9
1040          .SelBold = True
1050          .SelText = Space(5) & "Interpretive Comment for Specific Allergens (Not Total IgE):" & vbCrLf
1060          .SelBold = True
1070          .SelFontSize = 2
1080          .SelText = Space(20) & String(400, "-") & vbCrLf
1090          .SelFontSize = 9
1100          .SelBold = True
1110          .SelText = Left(Space(5) & "Ref Range kUA/L" & Space(20), 25) & "Clinical Implications" & vbCrLf
1120          .SelFontSize = 2
1130          .SelBold = True
1140          .SelText = Space(20) & String(400, "-") & vbCrLf

1150          .SelBold = False
1160          .SelFontSize = 9
1170          .SelText = Left(Space(5) & "< 0.1" & Space(20), 25) & "Negative/Absent/Undetectable." & vbCrLf
1180          .SelFontSize = 2
1190          .SelText = Space(20) & String(400, "-") & vbCrLf
1200          .SelFontSize = 9
1210          .SelText = Left(Space(5) & "0.10 - 0.35" & Space(20), 25) & "For Specialist use only; Clinical relevance undetermined." & vbCrLf
1220          .SelFontSize = 2
1230          .SelText = Space(20) & String(400, "-") & vbCrLf
1240          .SelFontSize = 9
1250          .SelText = Left(Space(5) & "0.36 - 0.70" & Space(20), 25) & "Low level of allergy; Indicative of ongoing sensitization." & vbCrLf
1260          .SelFontSize = 2
1270          .SelText = Space(20) & String(400, "-") & vbCrLf
1280          .SelFontSize = 9
1290          .SelText = Left(Space(5) & "0.71 - 3.50" & Space(20), 25) & "Moderate level of allergy; Indicative of ongoing sensitization." & vbCrLf
1300          .SelFontSize = 2
1310          .SelText = Space(20) & String(400, "-") & vbCrLf
1320          .SelFontSize = 9
1330          .SelText = Left(Space(5) & "3.51 - 17.5" & Space(20), 25) & "High level of allergy; Indicative of high level sensitization." & vbCrLf
1340          .SelFontSize = 2
1350          .SelText = Space(20) & String(400, "-") & vbCrLf
1360          .SelFontSize = 9
1370          .SelText = Left(Space(5) & "> 17.5 " & Space(20), 25) & "Very High level of allergy; Indicative of very high level sensitization." & vbCrLf
1380          .SelFontSize = 2
1390          .SelText = Space(20) & String(400, "-") & vbCrLf
1400          .SelText = vbCrLf
1410          .SelFontSize = 9
1420          .SelText = Space(5) & Chr(149) & " For all positive specific IgE results, please interpret in context of the" & vbCrLf
1430          .SelFontSize = 9
1440          .SelText = Space(5) & "  clinical history." & vbCrLf
1450          .SelText = Space(5) & Chr(149) & " High Total IgE levels can result in low level positivity (up to 3.5 kAU/l)" & vbCrLf
1460          .SelFontSize = 9
1470          .SelText = Space(5) & "  in specific IgE tests. This is particularly the case when the Total IgE is more" & vbCrLf
1480          .SelFontSize = 9
1490          .SelText = Space(5) & "  than 1000 kU/l. Please interpret the results of the specific IgE tests in context" & vbCrLf
1500          .SelFontSize = 9
1510          .SelText = Space(5) & "  of the clinical history." & vbCrLf


1520          Set OBS = OBS.Load(RP.SampleID, "Immunology", "Demographic")
1530          If Not OBS Is Nothing Then
1540              CommentTitle = "Comment:- "
1550              For Each OB In OBS
1560                  FillCommentLines CommentTitle & OB.Comment, 8, Comments(), 80
1570                  For i = 1 To 8
1580                      PrintTextRTB frmRichText.rtb, Comments(i), 10
1590                      PrintTextRTB frmRichText.rtb, vbCrLf
1600                  Next i
1610                  CommentTitle = ""
1620              Next

1630          End If

1640          .SelFontSize = 10
1650          .SelBold = False

1660          If RP.FaxNumber <> "" Then
1670              PrintFooterRTBFax RP.Initiator, SampleDate, Rundate
1680              .SelStart = 0
1690              f = FreeFile
1700              Open SysOptFax(0) & RP.SampleID & "Imm2.doc" For Output As f
1710              Print #f, .TextRTF
1720              Close f
1730              SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "Imm2.doc"
1740          Else
                  '1640          PrintFooterA4RTB AuthorisedBy, udtHeading.SampleDate, udtHeading.Rundate
1750              If UCase(GetOptionSetting("GetLatestAuthorisedBy", "")) = UCase("True") Then
1760                  PrintFooterA4RTB GetLatestAuthorisedBy("Imm", RP.SampleID), udtHeading.SampleDate, GetLatestRunDateTime("Imm", RP.SampleID, udtHeading.Rundate)
1770              Else
1780                  PrintFooterA4RTB AuthorisedBy, udtHeading.SampleDate, GetLatestRunDateTime("Imm", RP.SampleID, udtHeading.Rundate)
1790              End If
1800              .SelStart = 0

1810              .SelPrint Printer.hdc
1820          End If

1830          sql = "SELECT * FROM Reports WHERE 0 = 1"
1840          Set tb = New Recordset
1850          RecOpenServer 0, tb, sql
1860          tb.AddNew
1870          tb!SampleID = RP.SampleID
1880          tb!Name = udtHeading.Name
1890          tb!Dept = "I"
1900          tb!Initiator = RP.Initiator
1910          tb!PrintTime = PrintTime
1920          tb!RepNo = "1A" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
1930          tb!PageNumber = 0
1940          tb!Report = .TextRTF
1950          tb!Printer = Printer.DeviceName
1960          tb.Update
1970      End With

1980      ResetPrinter

1990      Exit Sub

PrintResultAllergy_Error:

          Dim strES As String
          Dim intEL As Integer

2000      intEL = Erl
2010      strES = Err.Description
2020      LogError "modImmunology", "PrintResultAllergy", intEL, strES, sql

End Sub

Function ImmTestAffected(ByVal br As BIEResult) As Boolean

      Dim TestName As String
      Dim tb As Recordset
      Dim sql As String
      Dim sn As Recordset

10    On Error GoTo ImmTestAffected_Error

20    ImmTestAffected = False
30    TestName = Trim(br.LongName)

40    sql = "SELECT * FROM ImmMasks WHERE SampleID = '" & br.SampleID & "'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql

70    If tb.EOF Then Exit Function

80    sql = "SELECT * FROM ImmTestDefinitions WHERE code = '" & br.Code & "' and shortname = '" & br.ShortName & "'"

90    Set sn = New Recordset
100   RecOpenServer 0, sn, sql
110   Do While Not sn.EOF
120       If sn!h And tb!h Then
130           ImmTestAffected = True
140           Exit Do
150       End If
160       If sn!s And tb!s Then
170           ImmTestAffected = True
180           Exit Do
190       End If
200       If sn!l And tb!l Then
210           ImmTestAffected = True
220           Exit Do
230       End If
240       If sn!o And tb!o Then
250           ImmTestAffected = True
260           Exit Do
270       End If
280       If sn!g And tb!g Then
290           ImmTestAffected = True
300           Exit Do
310       End If
320       If sn!J And tb!J Then
330           ImmTestAffected = True
340           Exit Do
350       End If
360       sn.MoveNext
370   Loop

380   Exit Function

ImmTestAffected_Error:

      Dim strES As String
      Dim intEL As Integer

390   intEL = Erl
400   strES = Err.Description
410   LogError "modImmunology", "ImmTestAffected", intEL, strES, sql

End Function

Function ImmReasonAffect(ByVal br As BIEResult) As String

      Dim TestName As String
      Dim sql As String
      Dim tb As Recordset
      Dim sn As Recordset

10    On Error GoTo ImmReasonAffect_Error

20    sql = "SELECT * FROM Immtestdefinitions WHERE code = '" & br.Code & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    ImmReasonAffect = ""
60    TestName = Trim(br.LongName)

70    If tb.EOF Then Exit Function

80    sql = "SELECT * FROM ImmMasks WHERE SampleID = '" & br.SampleID & "'"
90    Set sn = New Recordset
100   RecOpenServer 0, sn, sql
110   Do While Not sn.EOF
120       If tb!LongName = TestName And tb!Code = br.Code Then
130           If sn!g And tb!g Then
140               ImmReasonAffect = "Grossly Haemolysed"
150               Exit Do
160           End If
170           If sn!h And tb!h Then
180               ImmReasonAffect = "Haemolysed"
190               Exit Do
200           End If
210           If sn!s And tb!s Then
220               ImmReasonAffect = "Slightly Haemolysed"
230               Exit Do
240           End If
250           If sn!l And tb!l Then
260               ImmReasonAffect = "Lipaemic"
270               Exit Do
280           End If
290           If sn!J And tb!J Then
300               ImmReasonAffect = "Icteric"
310               Exit Do
320           End If
330           If sn!o And tb!o Then
340               ImmReasonAffect = "Old Sample"
350               Exit Do
360           End If
370       End If
380       sn.MoveNext
390   Loop

400   Exit Function

ImmReasonAffect_Error:

      Dim strES As String
      Dim intEL As Integer

410   intEL = Erl
420   strES = Err.Description
430   LogError "modImmunology", "ImmReasonAffect", intEL, strES, sql

End Function

Public Sub PrintResultHemochromatosis(ByVal BRs As BIEResults)

      Dim tb As Recordset
      Dim tbUN As Recordset
      Dim sql As String
      Dim Sex As String
10    ReDim lp(0 To 35) As String
20    ReDim lc(0 To 35) As String
      Dim lpc As Integer
      Dim cUnits As String
      Dim TempUnits As String
      Dim Flag As String
      Dim n As Integer
      Dim v As String
      Dim Low As Single
      Dim High As Single
      Dim strLow As String * 4
      Dim z As Long
      Dim s As String
      Dim strHigh As String * 4
      Dim br As BIEResult
      Dim TestCount As Integer
      Dim SampleType As String
      Dim ResultsPresent As Boolean
      'Dim Cx As Comment
      'Dim Cxs As New Comments
      Dim OB As Observation
      Dim OBS As New Observations
30    ReDim Comments(1 To 8) As String
      Dim SampleDate As String
      Dim Rundate As String
      Dim DoB As String
      Dim RunTime As String
      Dim RunDateTime As String
      Dim Fasting As String
      Dim Fx As Fasting
      Dim CodeGLU As String
      Dim CodeCHO As String
      Dim CodeTRI As String
      Dim udtPrintLine(0 To 35) As PrintLine
      Dim strFormat As String
      Dim SerumPrn As Boolean
      Dim UrinePrn As Boolean
      Dim C As Integer
      Dim f As Integer
      Dim Fontz1 As Integer
      Dim Fontz2 As Integer
      Dim PrintTime As String
      Dim AuthorisedBy As String
      Dim GPAddress0 As String
      Dim GPAddress1 As String
      Dim Margin As String


40    On Error GoTo PrintResultHemochromatosis_Error

50    AuthorisedBy = ""
60    RunDateTime = ""
70    Margin = Space(4)

80    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")


90    sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
100   Set tb = New Recordset
110   RecOpenClient 0, tb, sql

120   If tb.EOF Then
130       Exit Sub
140   End If

150   If IsDate(tb!DoB) Then
160       DoB = Format(tb!DoB, "dd/mmm/yyyy")
170   Else
180       DoB = ""
190   End If

200   If IsDate(tb!SampleDate) Then
210       SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
220   Else
230       SampleDate = ""
240   End If
250   If IsDate(RunTime) Then
260       Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
270   Else
280       If IsDate(tb!Rundate) Then
290           Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
300       Else
310           Rundate = ""
320       End If
330   End If


340   ResultsPresent = False
350   TestCount = 0

360   If Not BRs Is Nothing Then

370       AuthorisedBy = GetAuthorisedBy(GetLastValidatedBy(BRs))

380       For Each br In BRs
390           If br.Code <> "C28" And br.Code <> "H63" Then
400               BRs.RemoveItem br
410           End If
420       Next
430       TestCount = BRs.Count
440       If TestCount <> 0 Then
450           ResultsPresent = True
460           SampleType = BRs(1).SampleType
470           If Trim(SampleType) = "" Then SampleType = "B"

480       End If
490   End If

500   lpc = 0

510   If ResultsPresent Then




520       With frmRichText
              'LEAVE space for HSE logo

530           PrintTextRTB .rtb, String(10, vbCrLf)


              'REPORT HEADING
540           PrintTextRTB .rtb, FormatString("Page 1 of 1", 99, , AlignRight) & vbCrLf, 9, False
550           PrintTextRTB .rtb, FormatString("Regional Molecular Diagnostics Laboratory", 99, , AlignCenter) & vbCrLf, 10, True
560           PrintTextRTB .rtb, FormatString("Midlands Regional Hosptial " & HospName(0), 99, , AlignCenter) & vbCrLf, 10, True
570           PrintTextRTB .rtb, vbCrLf
580           PrintTextRTB .rtb, FormatString("Haemochromatosis Genetic Testing", 99, , AlignCenter) & vbCrLf, 10, True
590           PrintTextRTB .rtb, vbCrLf
600           PrintTextRTB .rtb, FormatString(GetOptionSetting("HaemochromatosisAccreditationLine1", ""), 99, , AlignCenter) & vbCrLf, 10, True
610           PrintTextRTB .rtb, FormatString(GetOptionSetting("HaemochromatosisAccreditationLine2", ""), 99, , AlignCenter) & vbCrLf, 10, True
620           PrintTextRTB .rtb, FormatString("", 106, , AlignCenter) & vbCrLf, 9, True
630           PrintTextRTB .rtb, vbCrLf
640           PrintTextRTB .rtb, vbCrLf

              'PATIENT DETAILS
650           PrintTextRTB .rtb, Margin, 10
660           PrintTextRTB .rtb, "PATIENT DETAILS" & vbCrLf, 10, True, , True
670           PrintTextRTB .rtb, vbCrLf
680           PrintTextRTB .rtb, Margin & FormatString("NAME:", 20), 10, True
690           PrintTextRTB .rtb, Margin & FormatString(udtHeading.Name, 71) & vbCrLf, 10
700           PrintTextRTB .rtb, Margin & FormatString("DOB:", 20), 10, True
710           PrintTextRTB .rtb, Margin & FormatString(udtHeading.DoB, 71) & vbCrLf, 10
720           PrintTextRTB .rtb, Margin & FormatString("ADDRESS:", 20), 10, True
730           PrintTextRTB .rtb, Margin & FormatString(udtHeading.Address0 & " " & udtHeading.Address1, 71) & vbCrLf, 10
740           PrintTextRTB .rtb, Margin & FormatString("LAB NO:", 20), 10, True
750           PrintTextRTB .rtb, Margin & FormatString(udtHeading.SampleID, 24), 10
760           PrintTextRTB .rtb, Space(3), 10
770           PrintTextRTB .rtb, FormatString("CHART NO:", 20), 10, True
780           PrintTextRTB .rtb, FormatString(udtHeading.Chart, 24) & vbCrLf, 10
790           PrintTextRTB .rtb, Margin & FormatString("CLINICAL DETAILS:", 20), 10, True
800           PrintTextRTB .rtb, Margin & FormatString(tb!ClDetails & "", 71) & vbCrLf, 10
810           PrintTextRTB .rtb, vbCrLf
820           PrintTextRTB .rtb, vbCrLf

              'CONSULTANT DETAILS
830           PrintTextRTB .rtb, Margin & "REQUESTING CONSULTANT / GENERAL PRACTICIONER" & vbCrLf, 10, True
840           PrintTextRTB .rtb, vbCrLf
850           PrintTextRTB .rtb, Margin & FormatString("REFFERRING CLICICIAN NAME:", 30), 10
860           If udtHeading.Clinician <> "" Then
870               PrintTextRTB .rtb, FormatString(udtHeading.Clinician, 61) & vbCrLf, 10
880               GPAddress0 = udtHeading.Ward
890           Else
900               PrintTextRTB .rtb, FormatString(udtHeading.GP, 61) & vbCrLf, 10
910               GPAddress0 = GetGPAddress(udtHeading.GP, 1)
920               GPAddress1 = GetGPAddress(udtHeading.GP, 2)
930           End If
940           PrintTextRTB .rtb, Margin & FormatString("ADDRESS FOR REPORT:", 30), 10
950           PrintTextRTB .rtb, FormatString(GPAddress0, 61) & vbCrLf, 10
960           PrintTextRTB .rtb, Margin & FormatString(" ", 30), 10
970           PrintTextRTB .rtb, FormatString(GPAddress1, 61) & vbCrLf, 10
980           PrintTextRTB .rtb, vbCrLf
990           PrintTextRTB .rtb, vbCrLf

1000          PrintTextRTB .rtb, Margin, 10
1010          PrintTextRTB .rtb, "TEST RESULT" & vbCrLf, 10, True, , True
1020          If Not BRs Is Nothing Then
1030              AuthorisedBy = GetAuthorisedBy(GetLastValidatedBy(BRs))
1040          End If
1050          For Each br In BRs
1060              If AuthorisedBy = "" Then AuthorisedBy = GetAuthorisedBy(br.Operator)
1070              If RunDateTime = "" Then RunDateTime = br.RunTime
1080              If br.Printable = True Then
1090                  Rundate = br.Rundate
                      'TEST RESULTS

1100                  PrintTextRTB .rtb, Margin & FormatString(br.LongName & ":", 30, , AlignRight), 10, True
1110                  PrintTextRTB .rtb, FormatString(br.Result, 61) & vbCrLf, 10, True

                      'check if phoned status required on Haemochromatosis report
                      '            If Br.Pc = "P" Then
                      '                lc(lpc) = Br.LongName & " - Phoned. "
                      '            ElseIf Br.Pc = "C" Then
                      '                lc(lpc) = Br.LongName & " - Checked. "
                      '            ElseIf Br.Pc = "PC" Then
                      '                lc(lpc) = Br.LongName & " - Phoned & Checked. "
                      '            End If

                      '            lc(lpc) = lc(lpc) & Br.Comment
1120                  RunTime = br.RunTime
1130                  LogTestAsPrinted "Imm", br.SampleID, br.Code
1140              End If

1150          Next
1160          PrintTextRTB .rtb, vbCrLf




              'INTERPRETATION
1170          PrintTextRTB .rtb, Margin & FormatString("INTERPRETATION:", 30)
1180          PrintTextRTB .rtb, FormatString("Homozygous = patient has 2 copies of the mutation", 70) & vbCrLf
1190          PrintTextRTB .rtb, Margin & FormatString(" ", 30)
1200          PrintTextRTB .rtb, FormatString("Heterozygous = patient has 1 copy of the mutation (carrier)", 70) & vbCrLf
1210          PrintTextRTB .rtb, Margin & FormatString(" ", 30)
1220          PrintTextRTB .rtb, FormatString("Normal = patient does not have either mutation", 70) & vbCrLf
1230          PrintTextRTB .rtb, vbCrLf
1240          PrintTextRTB .rtb, vbCrLf

              'COMMENTS
1250          Set OBS = OBS.Load(RP.SampleID, "Immunology")
1260          If Not OBS Is Nothing Then
1270              PrintTextRTB .rtb, Margin & FormatString("COMMENTS: ", 30, , AlignLeft)
1280              For Each OB In OBS
1290                  Select Case UCase$(OB.Discipline)
                      Case "IMMUNOLOGY"
1300                      FillCommentLines OB.Comment, 8, Comments(), 70
1310                      For n = 1 To 8
1320                          If Trim(Comments(n) & "") <> "" Then
1330                              If n > 1 Then
1340                                  PrintTextRTB .rtb, Margin & FormatString(" ", 30, , AlignLeft)
1350                              End If
1360                              PrintTextRTB .rtb, FormatString(Comments(n), 70) & vbCrLf
1370                          End If
1380                      Next
1390                  End Select
1400              Next
1410              PrintTextRTB .rtb, vbCrLf
1420              PrintTextRTB .rtb, vbCrLf
1430              PrintTextRTB .rtb, vbCrLf
1440              PrintTextRTB .rtb, vbCrLf
1450              PrintTextRTB .rtb, vbCrLf
1460          End If


              'FOOTER
1470          PrintTextRTB .rtb, Margin & FormatString("Signed:", 7, , AlignLeft)
1480          PrintTextRTB .rtb, FormatString("_________________________", 25, , AlignLeft)
1490          PrintTextRTB .rtb, FormatString(" ", 5, , AlignLeft)
1500          PrintTextRTB .rtb, FormatString("_________________________", 25, , AlignLeft) & vbCrLf
              'PrintTextRTB .rtb, FormatString(" ", 5, , AlignLeft)
              'PrintTextRTB .rtb, FormatString("_________________________", 25, , AlignLeft) & vbCrLf

1510          PrintTextRTB .rtb, Margin & FormatString(" ", 7, , AlignLeft)
1520          PrintTextRTB .rtb, FormatString(AuthorisedBy, 25, , AlignLeft)
1530          PrintTextRTB .rtb, FormatString(" ", 5, , AlignLeft)
1540          PrintTextRTB .rtb, FormatString("Helen Corrigan", 25, , AlignLeft) & vbCrLf
              'PrintTextRTB .rtb, FormatString(" ", 5, , AlignLeft)
              'PrintTextRTB .rtb, FormatString("Dr. K. Perera", 25, , AlignLeft) & vbCrLf

1550          PrintTextRTB .rtb, Margin & FormatString(" ", 7, , AlignLeft)
1560          PrintTextRTB .rtb, FormatString("Medical Scientist", 25, , AlignLeft)
1570          PrintTextRTB .rtb, FormatString(" ", 5, , AlignLeft)
1580          PrintTextRTB .rtb, FormatString("Chief Medical Scientist", 25, , AlignLeft) & vbCrLf
              'PrintTextRTB .rtb, FormatString(" ", 5, , AlignLeft)
              'PrintTextRTB .rtb, FormatString("Consultant Haematologist", 25, , AlignLeft) & vbCrLf


1590          PrintTextRTB .rtb, vbCrLf
1600          PrintTextRTB .rtb, vbCrLf
1610          PrintTextRTB .rtb, vbCrLf
1620          PrintTextRTB .rtb, vbCrLf
              'NOTES
1630          PrintTextRTB .rtb, Margin & "Notes:" & vbCrLf, 8, True
1640          PrintTextRTB .rtb, Margin & "1)   C282Y homozygotes are at risk of iron overload." & vbCrLf, 8
1650          PrintTextRTB .rtb, Margin & "2)   C282Y/H63D compound heterozygotes (carriers of both mutations) are at risk of iron overload." & vbCrLf, 8
1660          PrintTextRTB .rtb, Margin & "3)   C282Y heterozygotes and H63D heterozygotes do not develop significant iron overload." & vbCrLf, 8
1670          PrintTextRTB .rtb, Margin & "4)   H63D Homozygote's may in some rare cases develop significant iron overload." & vbCrLf, 8
1680          PrintTextRTB .rtb, Margin & "5)   93% of Irish Haemochromatosis patients are homozygous for the C282Y mutation." & vbCrLf, 8
1690          PrintTextRTB .rtb, Margin & "6)   Approximately 20% of the Irish population are carriers for C282Y and 25% are carriers for H63D." & vbCrLf, 8
1700          PrintTextRTB .rtb, Margin & "7)   As the carrier frequencies are so high in Ireland, spouses of homozygotes should be screened to " & vbCrLf, 8
1710          PrintTextRTB .rtb, Margin & "     establish potential status of offspring." & vbCrLf, 8
1720          PrintTextRTB .rtb, Margin & "8)   While homozygosity for C282Y is observed in the majority of Caucasian HH patients; the diagnostic " & vbCrLf, 8
1730          PrintTextRTB .rtb, Margin & "     value of other HFE mutations is still under investigation. More recently, mutations in the " & vbCrLf, 8
1740          PrintTextRTB .rtb, Margin & "     transferring receptor-2 (TFR-2) and ferroportin (FPN-1) genes have been reported." & vbCrLf, 8
1750          PrintTextRTB .rtb, Margin & "9)   Type 2 HH, also known as Juvenile Haemochromatosis results from mutations of an undefined locus on 1q." & vbCrLf, 8
1760          PrintTextRTB .rtb, Margin & "10)  Haemochromatosis genetic testing is conducted under the direction of Dr. K.Perera, Consultant Haematologist." & vbCrLf, 8


1770          PrintTextRTB .rtb, vbCrLf
1780          PrintTextRTB .rtb, FormatString("The Molecular Diagnostics laboratory is a participant in the ""UK NEQAS for HFE Genetic Testing Scheme""", 106, , AlignCenter) & vbCrLf, 9

1790          PrintTextRTB .rtb, vbCrLf
1800          PrintTextRTB .rtb, vbCrLf
1810          PrintTextRTB .rtb, Margin & String(140, "-") & vbCrLf, 6
1820          PrintTextRTB .rtb, Margin & FormatString("Date of test result:", 30)
1830          PrintTextRTB .rtb, FormatString(RunDateTime, 17, , AlignLeft)
1840          PrintTextRTB .rtb, Margin & FormatString("Sample Type: EDTA", 18)
1850          PrintTextRTB .rtb, Margin & FormatString("Sample Date:", 13)
1860          PrintTextRTB .rtb, FormatString(SampleDate, 17, , AlignLeft)


1870          Sex = tb!Sex & ""




              '    If TestCount <= Val(frmMain.lblImmMoreThan) Then
              '        '    Set Cx = Cxs.Load(RP.SampleID)
              '        Set OBS = OBS.Load(RP.SampleID, "Immunology", "Demographic")
              '        If Not OBS Is Nothing Then
              '            For Each OB In OBS
              '                Select Case UCase$(OB.Discipline)
              '                    Case "IMMUNOLOGY"
              '                        FillCommentLines OB.Comment, 8, Comments(), 80
              '                        For n = 1 To 8
              '                            If Trim(Comments(n) & "") <> "" Then
              '                                .SelFontName = "Courier New"
              '                                .SelFontSize = Fontz1
              '                                .SelText = Comments(n) & vbCrLf
              '                                CrCnt = CrCnt + 1
              '                            End If
              '                        Next
              '                    Case "DEMOGRAPHIC"
              '                        FillCommentLines OB.Comment, 2, Comments(), 80
              '                        For n = 1 To 2
              '                            If Trim(Comments(n) & "") <> "" Then
              '                                .SelText = Comments(n) & vbCrLf
              '                                CrCnt = CrCnt + 1
              '                            End If
              '                        Next
              '                End Select
              '            Next
              '        End If
              '    End If


1880          .rtb.SelStart = 0
1890          If RP.FaxNumber <> "" Then
1900              f = FreeFile
1910              Open SysOptFax(0) & RP.SampleID & "Hemo1.doc" For Output As f
1920              Print #f, .rtb.TextRTF
1930              Close f
1940              SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "Hemo1.doc"
1950          Else
                  'Do not print if Doctor is disabled in DisablePrinting
                  '*******************************************************************
1960              If CheckDisablePrinting(RP.Ward, "Immunology") Then

1970              ElseIf CheckDisablePrinting(RP.GP, "Immunology") Then
1980              Else
1990                  .rtb.SelPrint Printer.hdc
2000              End If
                  '*******************************************************************

                  '.rtb.SelPrint Printer.hDC
2010          End If
2020          sql = "SELECT * FROM Reports WHERE 0 = 1"
2030          Set tb = New Recordset
2040          RecOpenServer 0, tb, sql
2050          tb.AddNew
2060          tb!SampleID = RP.SampleID
2070          tb!Name = Trim(udtHeading.Name)
2080          tb!Dept = "L"
2090          tb!Initiator = RP.Initiator
2100          tb!PrintTime = PrintTime
2110          tb!RepNo = "0I" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
2120          tb!PageNumber = 0
2130          tb!Report = .rtb.TextRTF
2140          tb!Printer = Printer.DeviceName
2150          tb.Update

2160      End With
2170  End If


2180  Exit Sub

PrintResultHemochromatosis_Error:

      Dim strES As String
      Dim intEL As Integer

2190  intEL = Erl
2200  strES = Err.Description
2210  LogError "modImmunology", "PrintResultHemochromatosis", intEL, strES, sql

End Sub

'---------------------------------------------------------------------------------------
' Procedure : FlagAllergyResultHigh
' Author    : Babar Shahzad
' Date      : 03/10/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Function FlagAllergyResultHigh(ByVal Res As String) As Boolean


10    On Error GoTo FlagAllergyResultHigh_Error

20    Res = LTrim(RTrim(Res))
30    If Left(Res, 1) = ">" Then
40        Res = Mid(Res, 2, Len(Res))
50        Res = LTrim(RTrim(Res))
60        If IsNumeric(Res) Then
70            If Val(Res) + 0.01 > 0.35 Then
80                FlagAllergyResultHigh = True
90            End If
100       Else
110           FlagAllergyResultHigh = False
120       End If
130   End If

140   Exit Function

FlagAllergyResultHigh_Error:

       Dim strES As String
       Dim intEL As Integer

150    intEL = Erl
160    strES = Err.Description
170    LogError "modImmunology", "FlagAllergyResultHigh", intEL, strES
          
End Function


'********************************Unused code
'Public Sub PrintResultImmWinPortlaoise()
'
'Dim tb As Recordset
'Dim tbUN As Recordset
'Dim sql As String
'Dim Sex As String
'ReDim lp(0 To 35) As String
'Dim lpc As Integer
'Dim upc As Integer
'Dim cUnits As String
'Dim TempUnits As String
'Dim Flag As String
'Dim n As Integer
'Dim v As String
'Dim Low As Single
'Dim High As Single
'Dim strLow As String * 4
'Dim strHigh As String * 4
'Dim BRs As New BIEResults
'Dim Br As BIEResult
'Dim f As Integer
'Dim TestCount As Integer
'Dim sTestCount As Integer
'Dim SampleType As String
'Dim ResultsPresent As Boolean
''Dim Cx As Comment
''Dim Cxs As New Comments
'Dim OB As Observation
'Dim OBS As New Observations
'ReDim Comments(1 To 4) As String
'Dim SampleDate As String
'Dim Rundate As String
'Dim DoB As String
'Dim RunTime As String
'Dim sdtPrintLine(0 To 35) As PrintLine
'Dim udtPrintLine(0 To 35) As PrintLine
'Dim strFormat As String
'Dim C As Integer
'Dim PrintTime As String
'
'On Error GoTo PrintResultImmWinPortlaoise_Error
'
'PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")
'
'For n = 0 To 35
'    udtPrintLine(n).Analyte = ""
'    udtPrintLine(n).Result = ""
'    udtPrintLine(n).Flag = ""
'    udtPrintLine(n).Units = ""
'    udtPrintLine(n).NormalRange = ""
'    udtPrintLine(n).Fasting = ""
'    sdtPrintLine(n).Analyte = ""
'    sdtPrintLine(n).Result = ""
'    sdtPrintLine(n).Flag = ""
'    sdtPrintLine(n).Units = ""
'    sdtPrintLine(n).NormalRange = ""
'    sdtPrintLine(n).Fasting = ""
'Next
'
'sql = "SELECT * FROM Demographics WHERE " & _
 '      "SampleID = '" & RP.SampleID & "'"
'Set tb = New Recordset
'RecOpenClient 0, tb, sql
'
'If tb.EOF Then
'    Exit Sub
'End If
'
'If IsDate(tb!DoB) Then
'    DoB = Format(tb!DoB, "dd/mmm/yyyy")
'Else
'    DoB = ""
'End If
'
'ResultsPresent = False
'Set BRs = BRs.Load("Imm", RP.SampleID, "Results", 0, "Default", "")
'If Not BRs Is Nothing Then
'    If IsAllergy(BRs) Then
'        PrintResultAllergy BRs
'        Exit Sub
'    Else
'        ResultsPresent = True
'    End If
'End If
'
'lpc = 0
'If ResultsPresent Then
'    For Each Br In BRs
'        If Trim(Br.Operator) <> "" Then RP.Initiator = Br.Operator
'        sTestCount = sTestCount + 1
'        RunTime = Br.RunTime
'        v = Br.Result
'
'        High = Val(Br.High)
'        Low = Val(Br.Low)
'
'        If Low < 10 Then
'            strLow = Format(Low, "0.00")
'        ElseIf Low < 100 Then
'            strLow = Format(Low, "##.0")
'        Else
'            strLow = Format(Low, " ###")
'        End If
'        If High < 10 Then
'            strHigh = Format(High, "0.00")
'        ElseIf High < 100 Then
'            strHigh = Format(High, "##.0")
'        Else
'            strHigh = Format(High, "### ")
'        End If
'
'        If IsNumeric(v) And Sex <> "" Then
'            If Val(v) > Br.PlausibleHigh Then
'                sdtPrintLine(lpc).Flag = " X "
'                lp(lpc) = "  "
'                Flag = " X"
'            ElseIf Val(v) < Br.PlausibleLow Then
'                sdtPrintLine(lpc).Flag = " X "
'                lp(lpc) = "  "
'                Flag = " X"
'            ElseIf Val(v) > High Then
'                sdtPrintLine(lpc).Flag = " I "
'                lp(lpc) = "  "    'bold
'                Flag = " I"
'            ElseIf Val(v) < Low Then
'                sdtPrintLine(lpc).Flag = " N "
'                lp(lpc) = "  "    'bold
'                Flag = " N"
'            Else
'                sdtPrintLine(lpc).Flag = " E "
'                lp(lpc) = "  "    'unbold
'                Flag = " E"
'            End If
'        Else
'            sdtPrintLine(lpc).Flag = " C "
'            lp(lpc) = "  "    'unbold
'            Flag = " C"
'        End If
'        lp(lpc) = lp(lpc) & Left(Br.LongName & Space(20), 20)
'        sdtPrintLine(lpc).Analyte = Left(Br.LongName & Space(16), 16)
'        lp(lpc) = lp(lpc) & " " & Flag & " "
'
'        If IsNumeric(v) Then
'            Select Case Br.Printformat
'                Case 0: strFormat = "######"
'                Case 1: strFormat = "###0.0"
'                Case 2: strFormat = "##0.00"
'                Case 3: strFormat = "#0.000"
'            End Select
'            If Br.Code = "RUB" And Val(v) > 50 Then v = "> 50"
'            lp(lpc) = lp(lpc) & " " & Right(Space(6) & Format(v, strFormat), 6)
'            sdtPrintLine(lpc).Result = Format(v, strFormat)
'        Else
'            lp(lpc) = lp(lpc) & " " & Right(Space(6) & Left(v, 70), 70)
'            sdtPrintLine(lpc).Result = v
'        End If
'
'        sql = "SELECT * FROM Lists WHERE " & _
         '              "ListType = 'UN' and Code = '" & Br.Units & "'"
'        Set tbUN = Cnxn(0).Execute(sql)
'        If Not tbUN.EOF Then
'            cUnits = Left(tbUN!Text & Space(6), 6)
'        Else
'            cUnits = Left(Br.Units & Space(6), 6)
'        End If
'        sdtPrintLine(lpc).Units = cUnits
'        lp(lpc) = lp(lpc) & " " & cUnits
'
'        If Br.PrnRR = True And Sex <> "" Then
'            lp(lpc) = lp(lpc) & " (N/I 0 - " & strLow & ", E "
'            lp(lpc) = lp(lpc) & strLow & "-"
'            lp(lpc) = lp(lpc) & strHigh & ", I >" & strHigh & ")"
'            sdtPrintLine(lpc).NormalRange = "(" & strLow & "-" & strHigh & ")"
'        End If
'        sdtPrintLine(lpc).Fasting = ""
'        'LogImmAsPrinted br.Code
'        lpc = lpc + 1
'        upc = upc + 1
'    Next
'End If
'
'ClearUdtHeading
'With udtHeading
'    .SampleID = RP.SampleID
'    .Dept = "Immunology"
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
'    .SampleType = SampleType
'End With
'
'PrintHeadingRTB
'
'Sex = tb!Sex & ""
'
'With frmRichText.rtb
'    .SelFontSize = 10
'
'    '2490    If SysOptPrintMiddle(0) Then
'    '2500       n = Row_Count(lpc)
'    '2510       For C = 1 To n
'    '2520         .SelFontSize = 10
'    '2530         .SelText = vbCrLf
'    '2540         CrCnt = CrCnt + 1
'    '2550       Next
'    '2560    End If
'    '
'
'    For n = 0 To 19
'        If Trim(lp(n)) <> "" Then
'            Debug.Print lp(n)
'            If InStr(Mid(lp(n), 34, 3), "Neg") Or InStr(Mid(lp(n), 34, 3), "Pos") Then
'                If InStr(lp(n), "Pos") And InStr(UCase(lp(n)), "TITRE") Then
'                    .SelBold = True
'                    .SelText = Left(lp(n), 24)
'                    .SelText = Mid(lp(n), 30)
'                    .SelBold = False
'                ElseIf InStr(lp(n), "Neg") And InStr(UCase(lp(n)), "TITRE") Then
'                    .SelText = Left(lp(n), 24)
'                    .SelText = Mid(lp(n), 30)
'                ElseIf InStr(lp(n), "Pos") Then
'                    .SelBold = True
'                    .SelText = Left(lp(n), 24)
'                    .SelText = "Positive"
'                    .SelBold = False
'                Else
'                    .SelText = Left(lp(n), 24)
'                    .SelText = "Negative"
'                End If
'                .SelText = vbCrLf
'                CrCnt = CrCnt + 1
'            ElseIf InStr(Mid(lp(n), 24, 3), " N ") Or InStr(Mid(lp(n), 24, 3), " E ") Then
'                .SelBold = True
'                If InStr(lp(n), " N ") Then .SelColor = vbBlue Else .SelColor = vbRed
'                .SelText = Left(lp(n), 24) & " "
'                .SelText = Mid(lp(n), 29)
'                If InStr(lp(n), " N ") Then
'                    .SelText = " Non Immune"
'                ElseIf InStr(lp(n), " E ") Then
'                    .SelText = " Equivocal"
'                End If
'                .SelBold = False
'                .SelText = vbCrLf
'                CrCnt = CrCnt + 1
'            ElseIf InStr(Mid(lp(n), 24, 3), " C ") Then
'                .SelColor = vbBlack
'                .SelText = Left(lp(n), 24)
'                .SelText = Trim(Mid(lp(n), 30)) & vbCrLf
'                CrCnt = CrCnt + 1
'            Else
'                .SelColor = vbBlack
'                .SelText = Left(lp(n), 24)
'                .SelText = Mid(lp(n), 30)
'                .SelText = " Immune" & vbCrLf
'                CrCnt = CrCnt + 1
'            End If
'        End If
'
'    Next
'
'    .SelColor = vbBlack
'
'    Do While CrCnt < 28
'        .SelText = vbCrLf
'        CrCnt = CrCnt + 1
'    Loop
'    '  Set Cx = Cxs.Load(RP.SampleID)
'    Set OBS = OBS.Load(RP.SampleID, "Immunology", "Demographic")
'
'    If Not OBS Is Nothing Then
'        For Each OB In OBS
'            Select Case UCase$(OB.Discipline)
'                Case "IMMUNOLOGY"
'                    FillCommentLines OB.Comment, 4, Comments(), 97
'                    For n = 1 To 4
'                        If Trim(Comments(n)) <> "" Then
'                            .SelText = Comments(n) & vbCrLf
'                            CrCnt = CrCnt + 1
'                        End If
'                    Next
'                Case "DEMOGRAPHIC"
'                    FillCommentLines OB.Comment, 2, Comments(), 97
'                    For n = 1 To 2
'                        If Trim(Comments(n)) <> "" Then
'                            .SelText = Comments(n) & vbCrLf
'                            CrCnt = CrCnt + 1
'                        End If
'                    Next
'            End Select
'        Next
'    End If
'
'    If Not IsDate(tb!DoB) And Trim(Sex) = "" Then
'        .SelColor = vbBlack
'        .SelText = "No Sex/DoB given. Normal ranges may not be relevant" & vbCrLf
'        .SelText = vbCrLf
'        CrCnt = CrCnt + 1
'    ElseIf Not IsDate(tb!DoB) Then
'        .SelColor = vbBlack
'        .SelText = "*** No Dob. Adult Age 25 used for Normal Ranges! ***" & vbCrLf
'        .SelText = vbCrLf
'        CrCnt = CrCnt + 1
'    ElseIf Trim(Sex) = "" Then
'        .SelColor = vbBlack
'        .SelText = "No Sex given. No reference range applied" & vbCrLf
'        .SelText = vbCrLf
'        CrCnt = CrCnt + 1
'    End If
'    .SelColor = vbBlack
'
'    If IsDate(tb!SampleDate) Then
'        SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
'    Else
'        SampleDate = ""
'    End If
'    If IsDate(RunTime) Then
'        Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
'    Else
'        If IsDate(tb!Rundate) Then
'            Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
'        Else
'            Rundate = ""
'        End If
'    End If
'
'    If RP.FaxNumber <> "" Then
'        PrintFooterRTBFax RP.Initiator, SampleDate, Rundate
'        f = FreeFile
'        Open SysOptFax(0) & RP.SampleID & "IMM1.doc" For Output As f
'        Print #f, .TextRTF
'        Close f
'        SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "IMM1.doc"
'    Else
'        PrintFooterRTB RP.Initiator, SampleDate, Rundate
'        .SelStart = 0
'        .SelPrint Printer.hDC
'    End If
'    sql = "SELECT * FROM Reports WHERE 0 = 1"
'    Set tb = New Recordset
'    RecOpenServer 0, tb, sql
'    tb.AddNew
'    tb!SampleID = RP.SampleID
'    tb!Name = udtHeading.Name
'    tb!Dept = "I"
'    tb!Initiator = RP.Initiator
'    tb!PrintTime = PrintTime
'    tb!RepNo = "0I" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
'    tb!PageNumber = 0
'    tb!Report = .TextRTF
'    tb!Printer = Printer.DeviceName
'    tb.Update
'End With
'
''##########################
''Print Second Page
'If sTestCount > 20 Then
'    PrintHeadingRTB
'
'    With frmRichText.rtb
'        .SelFontSize = 10
'
'        For n = 20 To 35
'            If Trim(lp(n)) <> "" Then
'                Debug.Print lp(n)
'                If InStr(Mid(lp(n), 34, 3), "Neg") Or InStr(Mid(lp(n), 34, 3), "Pos") Then
'                    If InStr(lp(n), "Pos") Then
'                        .SelBold = True
'                        .SelText = Left(lp(n), 24)
'                        .SelText = "Positive"
'                        .SelBold = False
'                    Else
'                        .SelText = Left(lp(n), 24)
'                        .SelText = "Negative"
'                    End If
'                    .SelText = vbCrLf
'                    CrCnt = CrCnt + 1
'                ElseIf InStr(Mid(lp(n), 24, 3), " N ") Or InStr(Mid(lp(n), 24, 3), " E ") Then
'                    .SelBold = True
'                    If InStr(lp(n), " N ") Then .SelColor = vbBlue Else .SelColor = vbRed
'                    .SelText = Left(lp(n), 24) & " "
'                    .SelText = Mid(lp(n), 29)
'                    If InStr(lp(n), " N ") Then
'                        .SelText = " Non Immune"
'                    ElseIf InStr(lp(n), " E ") Then
'                        .SelText = " Equivocal"
'                    End If
'                    .SelBold = False
'                    .SelText = vbCrLf
'                    CrCnt = CrCnt + 1
'                ElseIf InStr(Mid(lp(n), 24, 3), " C ") Then
'                    .SelColor = vbBlack
'                    .SelText = Left(lp(n), 24)
'                    .SelText = Trim(Mid(lp(n), 30)) & vbCrLf
'                    CrCnt = CrCnt + 1
'                Else
'                    .SelColor = vbBlack
'                    .SelText = Left(lp(n), 24)
'                    .SelText = Mid(lp(n), 30)
'                    .SelText = " Immune" & vbCrLf
'                    CrCnt = CrCnt + 1
'                End If
'            End If
'        Next
'
'        Do While CrCnt < 28
'            .SelText = vbCrLf
'            CrCnt = CrCnt + 1
'        Loop
'        '    Set Cx = Cxs.Load(RP.SampleID)
'        Set OBS = OBS.Load(RP.SampleID, "Immunology", "Demographic")
'
'        If Not OBS Is Nothing Then
'            For Each OB In OBS
'                Select Case UCase$(OB.Discipline)
'                    Case "IMMUNOLOGY"
'                        FillCommentLines OB.Comment, 4, Comments(), 97
'                        For n = 1 To 4
'                            If Trim(Comments(n)) <> "" Then
'                                .SelText = Comments(n) & vbCrLf
'                                CrCnt = CrCnt + 1
'                            End If
'                        Next
'                    Case "DEMOGRAPHIC"
'                        FillCommentLines OB.Comment, 2, Comments(), 97
'                        For n = 1 To 2
'                            If Trim(Comments(n)) <> "" Then
'                                .SelText = Comments(n) & vbCrLf
'                                CrCnt = CrCnt + 1
'                            End If
'                        Next
'                End Select
'            Next
'        End If
'
'        If Not IsDate(tb!DoB) And Trim(Sex) = "" Then
'            .SelColor = vbBlack
'            .SelText = "No Sex/DoB given. Normal ranges may not be relevant" & vbCrLf
'            .SelText = vbCrLf
'            CrCnt = CrCnt + 1
'        ElseIf Not IsDate(tb!DoB) Then
'            .SelColor = vbBlack
'            .SelText = "*** No Dob. Adult Age 25 used for Normal Ranges! ***" & vbCrLf
'            .SelText = vbCrLf
'            CrCnt = CrCnt + 1
'        ElseIf Trim(Sex) = "" Then
'            .SelColor = vbBlack
'            .SelText = "No Sex given. No reference range applied" & vbCrLf
'            .SelText = vbCrLf
'            CrCnt = CrCnt + 1
'        End If
'        .SelColor = vbBlack
'
'        If IsDate(tb!SampleDate) Then
'            SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
'        Else
'            SampleDate = ""
'        End If
'        If IsDate(RunTime) Then
'            Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
'        Else
'            If IsDate(tb!Rundate) Then
'                Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
'            Else
'                Rundate = ""
'            End If
'        End If
'
'        If RP.FaxNumber <> "" Then
'            PrintFooterRTBFax RP.Initiator, SampleDate, Rundate
'            f = FreeFile
'            Open SysOptFax(0) & RP.SampleID & "IMM2.doc" For Output As f
'            Print #f, .TextRTF
'            Close f
'            SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "IMM2.doc"
'        Else
'            PrintFooterRTB RP.Initiator, SampleDate, Rundate
'            .SelStart = 0
'            .SelPrint Printer.hDC
'        End If
'        sql = "SELECT * FROM Reports WHERE 0 = 1"
'        Set tb = New Recordset
'        RecOpenServer 0, tb, sql
'        tb.AddNew
'        tb!SampleID = RP.SampleID
'        tb!Name = udtHeading.Name
'        tb!Dept = "I"
'        tb!Initiator = RP.Initiator
'        tb!PrintTime = PrintTime
'        tb!RepNo = "1I" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
'        tb!PageNumber = 1
'        tb!Report = frmRichText.rtb.TextRTF
'        tb!Printer = Printer.DeviceName
'        tb.Update
'    End With
'End If
'
'ResetPrinter
'
'sql = "UPDATE ImmResults SET Printed = '1' WHERE " & _
 '      "SampleID = '" & RP.SampleID & "'"
'Cnxn(0).Execute sql
'
'Exit Sub
'
'PrintResultImmWinPortlaoise_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'intEL = Erl
'strES = Err.Description
'
'LogError "modImmunology", "PrintResultImmWinPortlaoise", intEL, strES, sql
'
'sql = "Delete FROM printpending WHERE SampleID = '" & RP.SampleID & "' and department = '" & RP.Department & "'"
'Cnxn(0).Execute sql
'
'End Sub

