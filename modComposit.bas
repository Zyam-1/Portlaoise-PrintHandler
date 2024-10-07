Attribute VB_Name = "modComposit"
Option Explicit

Public Sub PrintComposit()

      Dim tbd As Recordset
      Dim fbc As Integer
      Dim TotalRetics As Long
      Dim Flag As String
      Dim lym As String
      Dim neut As String
      Dim mono As String
      Dim eos As String
      Dim bas As String
      Dim luc As String
      Dim DiffFound As Boolean
      Dim p As String
      Dim a As String
      Dim W As String
      Dim tb As Recordset
      Dim tbH As Recordset
      Dim n As Integer
      Dim Sex As String
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
      Dim DaysOld As String
      Dim tbUN As Recordset
      Dim lpc As Integer
      Dim cUnits As String
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
      Dim sn As Recordset
      Dim f As Integer
      Dim PrintTime As String

20    On Error GoTo PrintComposit_Error

30    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

40    For n = 0 To 60
50        udtPrintLine(n).Analyte = ""
60        udtPrintLine(n).Result = ""
70        udtPrintLine(n).Flag = ""
80        udtPrintLine(n).Units = ""
90        udtPrintLine(n).NormalRange = ""
100       udtPrintLine(n).Fasting = ""
110       udtPrintLine(n).Reason = ""
120   Next

130   sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
140   Set tb = New Recordset
150   RecOpenClient 0, tb, sql
160   If tb.EOF Then Exit Sub

170   If IsDate(tb!SampleDate) Then
180       SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
190   Else
200       SampleDate = ""
210   End If
220   If IsDate(tb!Rundate) Then
230       Rundate = Format(tb!Rundate, "dd/mmm/yyyy hh:mm")
240   Else
250       Rundate = ""
260   End If
270   If IsDate(tb!DoB) Then
280       DoB = Format(tb!DoB & "", "Short Date")
290   End If
300   sql = "SELECT * FROM HaemResults WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
310   Set tbH = New Recordset
320   RecOpenClient 0, tbH, sql

330   DoB = tb!DoB & ""

340   Select Case Left(UCase(tb!Sex & ""), 1)
          Case "M": Sex = "M"
350       Case "F": Sex = "F"
360       Case Else: Sex = ""
370   End Select

380   ClearUdtHeading
390   With udtHeading
400       .SampleID = RP.SampleID
410       .Dept = "Pathology"
420       .Name = tb!PatName & ""
430       .Ward = RP.Ward
440       .DoB = DoB
450       .Chart = tb!Chart & ""
460       .Clinician = RP.Clinician
470       .Address0 = tb!Addr0 & ""
480       .Address1 = tb!Addr1 & ""
490       .GP = RP.GP
500       .Sex = tb!Sex & ""
510       .Hospital = tb!Hospital & ""
520       .SampleDate = tb!SampleDate & ""
530       .RecDate = tb!RecDate & ""
540       .Rundate = tb!Rundate & ""
550       .GpClin = ""
560       .SampleType = SampleType
570       .DocumentNo = GetOptionSetting("CompositDocumentNo", "")
580       .AandE = tb!AandE & ""
590   End With

600   PrintHeadingRTBFax

610   With frmRichText.rtb
620       .SelFontName = "Courier New"
630       .SelFontSize = 12
640       .SelBold = True


      '******************************Start of Coagulation Printing
650       TestCount = 0
660       Set CRs = CRs.Load(RP.SampleID, Trim(SysOptExp(0)))

670       If Not CRs Is Nothing Then
680           For Each CR In CRs
690               If CR.Valid <> 0 Then TestCount = TestCount + 1
700           Next
710       End If
720       If TestCount > 0 Then
730           .SelText = "                      "
740           .SelUnderline = True
750           .SelText = "COAGULATION"
760           .SelText = vbCrLf
770           .SelUnderline = False
780           .SelText = vbCrLf

790           For Each CR In CRs
800               Rundate = Format(CR.RunTime, "dd/MMM/yyyy hh:mm")
810               If DoB <> "" And Len(DoB) > 9 Then DaysOld = Abs(DateDiff("d", SampleDate, DoB)) Else DaysOld = 12783
820               If DaysOld = 0 Then DaysOld = 1
830               sql = "SELECT * FROM coagtestdefinitions WHERE code = '" & Trim(CR.Code) & "' " & _
                        "and agefromdays <= '" & DaysOld & "' and agetodays >= '" & DaysOld & "'"
840               Set sn = New Recordset
850               RecOpenServer 0, sn, sql
860               If Not tb.EOF Then
870                   If sn!Printable Then
880                       If InterpCoag(0, Sex, CR.Code, CR.Result, DaysOld) <> "" And Trim(UCase(CR.Units)) <> "INR" Then
890                           .SelBold = True
900                       Else
910                           .SelBold = False
920                       End If
930                       .SelText = Left(" " & Space(5), 5)
940                       .SelText = Left(sn!TestName & Space(11), 11)
950                       Select Case sn!DP
                              Case 1: .SelText = Left(Format(CR.Result, "0.0") & Space(7), 7)
960                           Case 2: .SelText = Left(Format(CR.Result, "0.00") & Space(7), 7)
970                           Case 3: .SelText = Left(Format(CR.Result, "0.000") & Space(7), 7)
980                           Case Else: .SelText = Left(CR.Result & Space(7), 7)
990                       End Select
1000                      If Trim(CR.Units) = "ÆG/ML" Then .SelText = Left("ug/ML" & Space(6), 6) Else .SelText = Left(CR.Units & Space(6), 6)
1010                      If Trim(CR.Units) <> "INR" Then .SelText = Left(InterpCoag(0, Sex, CR.Code, CR.Result, DaysOld) & Space(10), 10)
1020                      If Trim(CR.Units) <> "INR" Then .SelText = nrCoag(CR.Code, Sex, DoB, SampleDate)
1030                      .SelText = vbCrLf
1040                  End If
1050              End If
1060          Next
          
1070          .SelText = vbCrLf
1080          .SelBold = False
          
              '  Set Cx = Cxs.Load(RP.SampleID)
          
1090          Set OBS = OBS.Load(RP.SampleID, "Coagulation")
1100          If Not OBS Is Nothing Then
1110              For Each OB In OBS
1120                  FillCommentLines OB.Comment, 4, Comments(), 80
1130                  For n = 1 To 4
1140                      If Comments(n) <> "" Then .SelText = "   " & Comments(n)
1150                  Next
1160              Next
1170          End If
1180      End If
      '******************************End of Coagulation Printing



      '******************************Start of Haematology Printing
1190      sql = "SELECT * FROM HaemResults WHERE " & _
                "SampleID = '" & RP.SampleID & "' and VALID = '1'"
1200      Set tbH = New Recordset
1210      RecOpenClient 0, tbH, sql
1220      If Not tbH.EOF Then

1230          .SelText = "                      "
1240          .SelUnderline = True
1250          .SelText = "HAEMATOLOGY"
1260          .SelText = vbCrLf
1270          .SelUnderline = False

1280          DoB = tb!DoB & ""

1290          fbc = Trim(tbH!WBC & "") <> ""

1300          sql = "SELECT * FROM differentials WHERE " & _
                    "runnumber = '" & RP.SampleID & "' and prndiff = 1 "
1310          Set tbd = New Recordset
1320          RecOpenClient 0, tbd, sql
1330          If Not tbd.EOF Then
1340              DiffFound = True
1350              For n = 0 To 14
1360                  W = Trim(tbd("Wording" & Format(n)) & "")
1370                  p = IIf(Val(tbd("P" & Format(n)) & "") = 0, "", tbd("P" & Format(n)))
1380                  a = IIf(Val(tbd("A" & Format(n)) & "") = 0, "", tbd("A" & Format(n)))
1390                  frmMain.gDiff.AddItem W & vbTab & p & vbTab & a
1400              Next
1410          End If

1420          .SelText = vbCrLf
1430          .SelFontName = "Courier New"
1440          .SelFontSize = 9

1450          If fbc Then
1460              Flag = InterpH(tbH!WBC & "", "WBC", Sex, DoB, SampleDate)
1470              .SelText = Left(" " & Space(3), 3)
1480              .SelText = "WBC   "
1490              If Flag <> "X" Then
1500                  .SelText = Right("     " & Trim(tbH!WBC & ""), 6)
1510              Else
1520                  .SelText = "XXXXX"
1530              End If
1540              .SelText = " x10"
1550              .SelFontSize = 4
1560              .SelCharOffset = 40
1570              .SelText = "9 "
1580              .SelCharOffset = 0
1590              .SelFontSize = 9
1600              .SelText = "/l "
1610              .SelBold = True
1620              .SelText = " " & Flag & " "
1630              .SelBold = False
1640              .SelText = Left(HaemNormalRange("WBC", Sex, DoB, SampleDate) & Space(15), 15)

1650              If DiffFound = True Then
1660                  For n = 1 To frmMain.gDiff.Rows - 1
1670                      If InStr(UCase(frmMain.gDiff.TextMatrix(n, 0)), "NEUT") > 0 Then
1680                          neut = frmMain.gDiff.TextMatrix(n, 2)
1690                          Exit For
1700                      End If
1710                  Next
1720              Else
1730                  neut = Trim$(tbH!neuta & "")
1740              End If

1750              If Trim(neut & "") <> "" Then
1760                  .SelText = "Neut  "
1770                  Flag = InterpH(neut & "", "NEUTA", Sex, DoB, SampleDate)
1780                  If Flag <> "X" Then
1790                      .SelText = Right("     " & neut & "", 6)
1800                  Else
1810                      .SelText = "XXXXX"
1820                  End If
1830                  .SelText = " x10"
1840                  .SelFontSize = 4
1850                  .SelCharOffset = 40
1860                  .SelText = "9 "
1870                  .SelCharOffset = 0
1880                  .SelFontSize = 9
1890                  .SelText = "/l "
1900                  .SelBold = True
1910                  .SelText = " " & Flag & " "
1920                  .SelBold = False
1930                  .SelText = HaemNormalRange("NEUTA", Sex, DoB, SampleDate) & vbCrLf
1940              Else
1950                  .SelText = vbCrLf
1960              End If

1970              If DiffFound = True Then
1980                  For n = 1 To frmMain.gDiff.Rows - 1
1990                      If InStr(UCase(frmMain.gDiff.TextMatrix(n, 0)), "LYM") > 0 Then
2000                          lym = frmMain.gDiff.TextMatrix(n, 2)
2010                          Exit For
2020                      End If
2030                  Next
2040              Else
2050                  lym = Trim$(tbH!lyma & "")
2060              End If

2070              If Trim(lym & "") <> "" Then
2080                  .SelFontSize = 4
2090                  .SelCharOffset = 40
2100                  .SelText = "  "
2110                  .SelCharOffset = 0
2120                  .SelFontSize = 9
2130                  .SelText = Left(" " & Space(40), 40)
2140                  .SelText = "Lymph "
2150                  Flag = InterpH(lym & "", "LYMA", Sex, DoB, SampleDate)
2160                  If Flag <> "X" Then
2170                      .SelText = Right("     " & lym & "", 6)
2180                  Else
2190                      .SelText = "XXXXXX"
2200                  End If
2210                  .SelText = " x10"
2220                  .SelFontSize = 4
2230                  .SelCharOffset = 40
2240                  .SelText = "9 "
2250                  .SelCharOffset = 0
2260                  .SelFontSize = 9
2270                  .SelText = "/l "
2280                  .SelBold = True
2290                  .SelText = " " & Flag & " "
2300                  .SelBold = False
2310                  .SelText = HaemNormalRange("LYMA", Sex, DoB, SampleDate) & vbCrLf
2320              Else
2330                  .SelText = vbCrLf
2340              End If

2350              .SelFontSize = 9
2360              Flag = InterpH(tbH!RBC & "", "RBC", Sex, DoB, SampleDate)
2370              .SelText = Left(" " & Space(3), 3) & "RBC   "
2380              If Flag <> "X" Then
2390                  .SelText = Right("     " & tbH!RBC & "", 6)
2400              Else
2410                  .SelText = "XXXXXX"
2420              End If
2430              .SelText = " x10"
2440              .SelFontSize = 4
2450              .SelCharOffset = 40
2460              .SelText = "12"
2470              .SelCharOffset = 0
2480              .SelFontSize = 9
2490              .SelText = "/l "
2500              .SelBold = True
2510              .SelText = " " & Flag & " "
2520              .SelBold = False
2530              .SelText = Left(HaemNormalRange("RBC", Sex, DoB, SampleDate) & Space(15), 15)

2540              If DiffFound = True Then
2550                  For n = 1 To frmMain.gDiff.Rows - 1
2560                      If InStr(UCase(frmMain.gDiff.TextMatrix(n, 0)), "MON") > 0 Then
2570                          mono = frmMain.gDiff.TextMatrix(n, 2)
2580                          Exit For
2590                      End If
2600                  Next
2610              Else
2620                  mono = Trim$(tbH!MonoA & "")
2630              End If

2640              If Trim(mono & "") <> "" Then
2650                  .SelText = "Mono  "
2660                  Flag = InterpH(mono & "", "MONOA", Sex, DoB, SampleDate)
2670                  If Flag <> "X" Then
2680                      .SelText = Right("     " & mono & "", 6)
2690                  Else
2700                      .SelText = "XXXXXX"
2710                  End If
2720                  .SelText = " x10"
2730                  .SelFontSize = 4
2740                  .SelCharOffset = 40
2750                  .SelText = "9 "
2760                  .SelCharOffset = 0
2770                  .SelFontSize = 9
2780                  .SelText = "/l "
2790                  .SelBold = True
2800                  .SelText = " " & Flag & " "
2810                  .SelBold = False
2820                  .SelText = HaemNormalRange("MONOA", Sex, DoB, SampleDate) & vbCrLf
2830              Else
2840                  .SelText = vbCrLf
2850              End If

2860              If DiffFound = True Then
2870                  For n = 1 To frmMain.gDiff.Rows - 1
2880                      If InStr(UCase(frmMain.gDiff.TextMatrix(n, 0)), "EOS") > 0 Then
2890                          eos = frmMain.gDiff.TextMatrix(n, 2)
2900                          Exit For
2910                      End If
2920                  Next
2930              Else
2940                  eos = Trim$(tbH!eosa & "")
2950              End If

2960              If Trim(eos & "") <> "" Then
2970                  .SelFontSize = 4
2980                  .SelCharOffset = 40
2990                  .SelText = "  "
3000                  .SelCharOffset = 0
3010                  .SelFontSize = 9
3020                  .SelText = Left(" " & Space(40), 40)
3030                  .SelText = "Eos   "
3040                  Flag = InterpH(eos & "", "EOSA", Sex, DoB, SampleDate)
3050                  If Flag <> "X" Then
3060                      .SelText = Right("     " & eos & "", 6)
3070                  Else
3080                      .SelText = "XXXXXX"
3090                  End If
3100                  .SelText = " x10"
3110                  .SelFontSize = 4
3120                  .SelCharOffset = 40
3130                  .SelText = "9 "
3140                  .SelCharOffset = 0
3150                  .SelFontSize = 9
3160                  .SelText = "/l "
3170                  .SelBold = True
3180                  .SelText = " " & Flag & " "
3190                  .SelBold = False
3200                  .SelText = HaemNormalRange("EOSA", Sex, DoB, SampleDate) & vbCrLf
3210              Else
3220                  .SelText = vbCrLf
3230              End If

3240              .SelFontSize = 9

3250              Flag = InterpH(tbH!Hgb & "", "Hgb", Sex, DoB, SampleDate)
3260              .SelText = Left(" " & Space(3), 3) & "Hgb   "
3270              If Flag <> "X" Then
3280                  .SelText = Right("     " & tbH!Hgb & "", 6)
3290              Else
3300                  .SelText = "XXXXX"
3310              End If
3320              .SelFontSize = 4
3330              .SelCharOffset = 40
3340              .SelText = "  "
3350              .SelCharOffset = 0
3360              .SelFontSize = 9
3370              .SelText = "  "
3380              .SelText = "g/dl "
3390              .SelBold = True
3400              .SelText = " " & Flag & " "
3410              .SelBold = False
3420              .SelText = Left(HaemNormalRange("Hgb", Sex, DoB, SampleDate) & Space(15), 15)

3430              If DiffFound = True Then
3440                  For n = 1 To frmMain.gDiff.Rows - 1
3450                      If InStr(UCase(frmMain.gDiff.TextMatrix(n, 0)), "BAS") > 0 Then
3460                          bas = frmMain.gDiff.TextMatrix(n, 2)
3470                          Exit For
3480                      End If
3490                  Next
3500              Else
3510                  bas = Trim$(tbH!basa & "")
3520              End If

3530              If Trim(bas & "") <> "" Then
3540                  .SelText = "Bas   "
3550                  Flag = InterpH(bas & "", "BASA", Sex, DoB, SampleDate)
3560                  If Flag <> "X" Then
3570                      .SelText = Right("     " & bas & "", 6)
3580                  Else
3590                      .SelText = "XXXXXX"
3600                  End If
3610                  .SelText = " x10"
3620                  .SelFontSize = 4
3630                  .SelCharOffset = 40
3640                  .SelText = "9 "
3650                  .SelCharOffset = 0
3660                  .SelFontSize = 9
3670                  .SelText = "/l "
3680                  .SelBold = True
3690                  .SelText = " " & Flag & " "
3700                  .SelBold = False
3710                  .SelText = HaemNormalRange("BASA", Sex, DoB, SampleDate) & vbCrLf
3720              Else
3730                  .SelText = vbCrLf
3740              End If

3750              If DiffFound = True Then
3760                  For n = 1 To frmMain.gDiff.Rows - 1
3770                      If InStr(UCase(frmMain.gDiff.TextMatrix(n, 0)), "LUC") > 0 Then
3780                          luc = frmMain.gDiff.TextMatrix(n, 2)
3790                          Exit For
3800                      End If
3810                  Next
3820              Else
3830                  luc = Trim$(tbH!luca & "")
3840              End If

3850              If DiffFound = True Then
3860                  luc = Format(frmMain.gDiff.TextMatrix(7, 2), "##0.0##")
3870              Else
3880                  luc = Trim$(tbH!luca & "")
3890              End If

3900              If Trim(luc & "") <> "" Then
3910                  If DiffFound = True Then
3920                      .SelFontSize = 4
3930                      .SelCharOffset = 40
3940                      .SelText = "  "
3950                      .SelCharOffset = 0
3960                      .SelFontSize = 9
3970                      .SelText = Left(" " & Space(40), 40)
3980                      .SelText = Left(Initial2Upper(frmMain.gDiff.TextMatrix(7, 0)) & Space(16), 16)
3990                      .SelText = Right(" " & Trim$(luc & ""), 6)
4000                  Else
4010                      .SelFontSize = 4
4020                      .SelCharOffset = 40
4030                      .SelText = "  "
4040                      .SelCharOffset = 0
4050                      .SelFontSize = 9
4060                      .SelText = Left(" " & Space(40), 40)
4070                      .SelText = "Luc   "
4080                      Flag = InterpH(luc & "", "LUCA", Sex, DoB, SampleDate)
4090                      If Flag <> "X" Then
4100                          .SelText = Right("     " & Trim$(luc & ""), 6)
4110                      Else
4120                          .SelText = " XXXXX"
4130                      End If
4140                      .SelText = " x10"
4150                      .SelFontSize = 4
4160                      .SelCharOffset = 40
4170                      .SelText = "9 "
4180                      .SelCharOffset = 0
4190                      .SelFontName = "Courier New"
4200                      .SelFontSize = 9
4210                      .SelText = "/l  "
4220                      .SelBold = True
4230                      .SelText = Flag & " "
4240                      .SelBold = False
4250                  End If
4260                  If DiffFound = True Then
4270                      .SelText = vbCrLf
4280                      CrCnt = CrCnt + 1
4290                  Else
4300                      .SelText = HaemNormalRange("LUCA", Sex, DoB, SampleDate) & vbCrLf
4310                      CrCnt = CrCnt + 1
4320                  End If
4330              Else
4340                  .SelText = vbCrLf
4350                  CrCnt = CrCnt + 1
4360              End If

4370              .SelFontSize = 9

4380              Flag = InterpH(tbH!Hct & "", "Hct", Sex, DoB, SampleDate)
4390              .SelText = Left(" " & Space(3), 3) & "Hct   "
4400              If Flag <> "X" Then
4410                  .SelText = Right("     " & tbH!Hct & "", 6)
4420              Else
4430                  .SelText = "XXXXXX"
4440              End If
4450              .SelFontSize = 4
4460              .SelCharOffset = 40
4470              .SelText = "  "
4480              .SelCharOffset = 0
4490              .SelFontSize = 9
4500              .SelText = "   "
4510              .SelText = "l/l "
4520              .SelBold = True
4530              .SelText = " " & Flag & " "
4540              .SelBold = False
4550              .SelText = Left(HaemNormalRange("Hct", Sex, DoB, SampleDate) & Space(15), 15)

4560              If DiffFound = True Then
4570                  luc = ""
4580                  If frmMain.gDiff.TextMatrix(8, 1) <> "" Then
4590                      luc = Format(frmMain.gDiff.TextMatrix(8, 2), "##0.0##")
4600                      If Trim(luc & "") <> "" Then
4610                          .SelText = Left(Initial2Upper(Left(frmMain.gDiff.TextMatrix(8, 0), 17)) & Space(16), 16)
4620                          .SelText = " " & luc
4630                          .SelText = vbCrLf
4640                          CrCnt = CrCnt + 1
4650                      Else
4660                          .SelText = vbCrLf
4670                          CrCnt = CrCnt + 1
4680                      End If
4690                  Else
4700                      .SelText = vbCrLf
4710                      CrCnt = CrCnt + 1
4720                  End If
4730              Else
4740                  .SelText = vbCrLf
4750                  CrCnt = CrCnt + 1
4760              End If

4770              If DiffFound = True Then
4780                  luc = ""
4790                  If frmMain.gDiff.TextMatrix(9, 1) <> "" Then
4800                      luc = Format(frmMain.gDiff.TextMatrix(9, 2), "##0.0##")
4810                      If Trim(luc & "") <> "" Then
4820                          .SelFontSize = 4
4830                          .SelCharOffset = 40
4840                          .SelText = "  "
4850                          .SelCharOffset = 0
4860                          .SelFontSize = 9
4870                          .SelText = Left(" " & Space(40), 40)
4880                          .SelText = Left(Initial2Upper(frmMain.gDiff.TextMatrix(9, 0)) & Space(16), 16)
4890                          .SelText = " " & luc
4900                          .SelText = vbCrLf
4910                          CrCnt = CrCnt + 1
4920                      Else
4930                          .SelText = vbCrLf
4940                          CrCnt = CrCnt + 1
4950                      End If
4960                  Else
4970                      .SelText = vbCrLf
4980                      CrCnt = CrCnt + 1
4990                  End If
5000              Else
5010                  .SelText = vbCrLf
5020                  CrCnt = CrCnt + 1
5030              End If

5040              .SelFontSize = 9

5050              Flag = InterpH(tbH!MCV & "", "MCV", Sex, DoB, SampleDate)
5060              .SelText = Left(" " & Space(3), 3) & "MCV   "
5070              If Flag <> "X" Then
5080                  .SelText = Right("     " & tbH!MCV & "", 6)
5090              Else
5100                  .SelText = "XXXXX"
5110              End If
5120              .SelText = "    "
5130              .SelFontSize = 4
5140              .SelCharOffset = 40
5150              .SelText = "  "
5160              .SelCharOffset = 0
5170              .SelFontSize = 9
5180              .SelText = "fl "
5190              .SelBold = True
5200              .SelText = " " & Flag & " "
5210              .SelBold = False
5220              .SelText = Left(HaemNormalRange("MCV", Sex, DoB, SampleDate) & Space(15), 15)

5230          End If

5240          If DiffFound = True Then
5250              luc = ""
5260              If frmMain.gDiff.TextMatrix(10, 1) <> "" Then
5270                  luc = Format(frmMain.gDiff.TextMatrix(10, 2), "##0.0##")
5280                  If Trim(luc & "") <> "" Then
5290                      .SelText = Left(Initial2Upper(Left(frmMain.gDiff.TextMatrix(10, 0), 17)) & Space(16), 16)
5300                      .SelText = " " & luc
5310                      .SelText = vbCrLf
5320                      CrCnt = CrCnt + 1
5330                  Else
5340                      .SelText = vbCrLf
5350                      CrCnt = CrCnt + 1
5360                  End If
5370              Else
5380                  .SelText = vbCrLf
5390                  CrCnt = CrCnt + 1
5400              End If
5410          Else
5420              .SelText = vbCrLf
5430              CrCnt = CrCnt + 1
5440          End If

5450          Flag = ""
5460          If DiffFound = True Then
5470              luc = ""
5480              If frmMain.gDiff.TextMatrix(11, 1) <> "" Then
5490                  luc = Format(frmMain.gDiff.TextMatrix(11, 2), "##0.0##")
5500                  If Trim(luc & "") <> "" Then
5510                      .SelFontSize = 4
5520                      .SelCharOffset = 40
5530                      .SelText = "  "
5540                      .SelCharOffset = 0
5550                      .SelFontSize = 9
5560                      .SelText = Left(" " & Space(40), 40)
5570                      .SelText = Left$(Initial2Upper(frmMain.gDiff.TextMatrix(11, 0)) & Space(16), 16)
5580                      .SelText = " " & luc
5590                      .SelText = vbCrLf
5600                      CrCnt = CrCnt + 1
5610                  Else
5620                      .SelText = vbCrLf
5630                      CrCnt = CrCnt + 1
5640                  End If
5650              Else
5660                  .SelText = vbCrLf
5670                  CrCnt = CrCnt + 1
5680              End If
5690          Else
5700              .SelText = vbCrLf
5710              CrCnt = CrCnt + 1
5720          End If

5730          .SelFontSize = 9

5740          If fbc Then
5750              Flag = InterpH(tbH!MCH & "", "MCH", Sex, DoB, SampleDate)
5760              .SelText = Left(" " & Space(3), 3) & "MCH   "
5770              If Flag <> "X" Then
5780                  .SelText = Right("     " & tbH!MCH & "", 6)
5790              Else
5800                  .SelText = "XXXXXX"
5810              End If
5820              .SelText = "    "
5830              .SelFontSize = 4
5840              .SelCharOffset = 40
5850              .SelText = "  "
5860              .SelCharOffset = 0
5870              .SelFontSize = 9
5880              .SelText = "pg "
5890              .SelBold = True
5900              .SelText = " " & Flag & " "
5910              .SelBold = False
5920              .SelText = HaemNormalRange("MCH", Sex, DoB, SampleDate) & vbCrLf
5930          End If

5940          .SelText = vbCrLf

5950          .SelFontSize = 9

5960          If fbc Then
5970              Flag = InterpH(tbH!MCHC & "", "MCHC", Sex, DoB, SampleDate)
5980              .SelText = Left(" " & Space(3), 3) & "MCHC  "
5990              If Flag <> "X" Then
6000                  .SelText = Right("     " & tbH!MCHC & "", 6)
6010              Else
6020                  .SelText = "XXXXXX"
6030              End If
6040              .SelText = "  "
6050              .SelFontSize = 4
6060              .SelCharOffset = 40
6070              .SelText = "  "
6080              .SelCharOffset = 0
6090              .SelFontSize = 9
6100              .SelText = "g/dl "
6110              .SelBold = True
6120              .SelText = " " & Flag & " "
6130              .SelBold = False
6140              .SelText = Left(HaemNormalRange("MCHC", Sex, DoB, SampleDate) & Space(15), 15)

6150          End If

6160          If Trim(tbH!ESR & "") <> "" Then
6170              Flag = InterpH(tbH!ESR & "", "ESR", Sex, DoB, SampleDate)
6180              .SelText = "ESR      "
6190              If Flag <> "X" Then
6200                  .SelText = Right("      " & tbH!ESR & "", 6)
6210              Else
6220                  .SelText = "XXXXXX"
6230              End If
6240              .SelFontSize = 4
6250              .SelCharOffset = 40
6260              .SelText = " "
6270              .SelCharOffset = 0
6280              .SelFontSize = 9
6290              .SelText = "mm/hr"
6300              .SelBold = True
6310              .SelText = " " & Flag & " "
6320              .SelBold = False
6330              .SelText = HaemNormalRange("ESR", Sex, DoB, SampleDate) & vbCrLf
6340          Else
6350              .SelText = vbCrLf
6360          End If

6370          .SelText = vbCrLf
6380          .SelFontSize = 9

6390          If fbc Then
6400              Flag = InterpH(tbH!rdwcv & "", "RDW", Sex, DoB, SampleDate)
6410              .SelText = Left(" " & Space(3), 3) & "RDW   "
6420              If Flag <> "X" Then
6430                  .SelText = Right("     " & tbH!rdwcv & "", 6)
6440              Else
6450                  .SelText = "XXXXXX"
6460              End If
6470              .SelText = " "
6480              .SelFontSize = 4
6490              .SelCharOffset = 40
6500              .SelText = "  "
6510              .SelCharOffset = 0
6520              .SelFontSize = 9
6530              .SelText = "    % "
6540              .SelBold = True
6550              .SelText = " " & Flag & " "
6560              .SelBold = False
6570              .SelText = Left(HaemNormalRange("RDW", Sex, DoB, SampleDate) & Space(15), 15)
6580              If Trim(tbH!retp & "") <> "" Then
6590                  .SelText = "Total Retics "
6600                  If Trim(tbH!reta) & "" <> "" And Trim(tbH!reta) <> "?" Then TotalRetics = Val(tbH!reta & "") Else Flag = "X"   'Val((tbH!retics) * Val(tbH!RBC) * 10)
6610                  If Flag <> "X" Then
6620                      .SelText = Format(TotalRetics, "###0")
6630                  Else
6640                      .SelText = "XXXX"
6650                  End If
6660                  .SelText = " x10"
6670                  .SelFontSize = 4
6680                  .SelCharOffset = 40
6690                  .SelText = "9 "
6700                  .SelCharOffset = 0
6710                  .SelFontSize = 9
6720                  .SelText = "/l "
6730                  .SelBold = True
6740                  .SelText = Flag & " "
6750                  .SelBold = False
6760                  .SelText = Trim(HaemNormalRange("RETA", Sex, DoB, SampleDate)) & vbCrLf
6770              Else
6780                  .SelText = vbCrLf
6790              End If
6800          End If

6810          .SelFontSize = 9
6820          If Trim(tbH!Monospot & "") <> "" Then
6830              .SelFontSize = 4
6840              .SelCharOffset = 40
6850              .SelText = "  "
6860              .SelCharOffset = 0
6870              .SelFontSize = 9
6880              .SelText = Left(" " & Space(40), 40)
6890              .SelText = "Monospot "
6900              If tbH!Monospot = "N" Then
6910                  .SelText = "Negative"
6920              ElseIf tbH!Monospot = "P" Then
6930                  .SelText = "Positive"
6940              Else
6950                  .SelText = "?"
6960              End If
6970              .SelText = vbCrLf
6980          Else
6990              .SelText = vbCrLf
7000          End If

7010          .SelFontSize = 9
7020          If fbc Then
7030              Flag = InterpH(tbH!Plt & "", "Plt", Sex, DoB, SampleDate)
7040              .SelText = Left(" " & Space(3), 3) & "Plt   "
7050              If Flag <> "X" Then
7060                  .SelText = Right("     " & tbH!Plt & "", 6)
7070              Else
7080                  .SelText = "XXXXX"
7090              End If
7100              .SelText = " x10"
7110              .SelFontSize = 4
7120              .SelCharOffset = 40
7130              .SelText = "9 "
7140              .SelCharOffset = 0
7150              .SelFontSize = 9
7160              .SelText = "/l "
7170              .SelBold = True
7180              .SelText = " " & Flag & " "
7190              .SelBold = False
7200              .SelText = Left(HaemNormalRange("Plt", Sex, DoB, SampleDate) & Space(15), 15)
7210          End If
              '**********************************
              '  Sickle Cell Screen Result
              '**********************************
7220          If Trim(tbH!sickledex & "") <> "" And Trim(tbH!sickledex & "") <> "?" Then
7230              .SelText = Left("Sickle Cell Screen  " & Trim(tbH!sickledex) & Space(40), 40)
                  'CrCnt = CrCnt + 1
7240          Else
7250              .SelText = vbCrLf
                  'CrCnt = CrCnt + 1
7260          End If
              '*******************************
              
7270          If Trim(tbH!tasot & "") <> "" And Trim(tbH!tasot & "") <> "?" Then
7280              .SelText = "Asot     " & Trim(tbH!tasot) & vbCrLf
7290          Else
7300              .SelText = vbCrLf
7310          End If

7320          .SelFontSize = 9
7330          If Trim(tbH!tra & "") <> "" And Trim(tbH!tra & "") <> "?" Then
7340              .SelFontSize = 4
7350              .SelCharOffset = 40
7360              .SelText = "  "
7370              .SelCharOffset = 0
7380              .SelFontSize = 9
7390              .SelText = Left(" " & Space(40), 40)
7400              .SelText = "Ra       " & Trim(tbH!tra) & vbCrLf
7410          Else
7420              .SelText = vbCrLf
7430          End If
              
7440          .SelText = vbCrLf

7450          .SelFontSize = 9
7460          If DiffFound = True Then
7470              .SelText = "          Manual Differential Reported" & vbCrLf
7480          End If

7490          .SelFontSize = 9
              '  Set Cx = Cxs.Load(RP.SampleID)
7500          Set OBS = OBS.Load(RP.SampleID, "Haematology")
7510          If Not OBS Is Nothing Then
7520              For Each OB In OBS
7530                  FillCommentLines OB.Comment, 4, Comments(), 80
7540                  For n = 1 To 4
7550                      If Comments(n) <> "" Then .SelText = "  " & Comments(n) & vbCrLf
7560                  Next
7570              Next
7580          End If
7590      End If

7600      .SelText = vbCrLf
      '******************************End of Haematology Printing

7610      If Not IsNull(tb!Fasting) Then
7620          Fasting = tb!Fasting
7630      Else
7640          Fasting = False
7650      End If

7660      Select Case UCase(HospName(0))
              Case "PORTLAOISE"
7670              CodeGLU = SysOptBioCodeForGlucose(0)
7680              CodeCHO = SysOptBioCodeForChol(0)
7690              CodeTRI = SysOptBioCodeForTrig(0)
7700              CodeGLUP = SysOptBioCodeForGlucoseP(0)
7710              CodeCHOP = SysOptBioCodeForCholP(0)
7720              CodeTRIP = SysOptBioCodeForTrigP(0)
7730      End Select
          
          
      '******************************Start of Biochemistry Printing
7740      ResultsPresent = False
7750      TestCount = 0
7760      Set BRs = BRs.Load("bio", RP.SampleID, "Results", 0, "", "")
7770      If Not BRs Is Nothing Then
7780          For Each br In BRs
7790              If br.Valid <> 0 Then TestCount = TestCount + 1
7800          Next
              'TestCount = BRs.Count
7810          If TestCount <> 0 Then
7820              ResultsPresent = True
7830              SampleType = BRs(1).SampleType
7840              If Trim(SampleType) = "" Then SampleType = "S"
7850          End If
7860      End If

7870      lpc = 0
7880      If ResultsPresent Then
7890          .SelFontName = "Courier New"
7900          .SelFontSize = 12
7910          .SelText = "                      "
7920          .SelUnderline = True
7930          .SelText = "BIOCHEMISTRY"
7940          .SelText = vbCrLf
7950          .SelUnderline = False
7960          .SelFontName = "Courier New"
7970          .SelFontSize = 10
7980          .SelText = vbCrLf
7990          For Each br In BRs
8000              If br.Printable = True Then
8010                  RunTime = br.RunTime
8020                  v = br.Result

8030                  If br.Code = CodeGLU Or br.Code = CodeCHO Or br.Code = CodeTRI Or _
                          br.Code = CodeGLUP Or br.Code = CodeCHOP Or br.Code = CodeTRIP Then
8040                      If Fasting Then
8050                          Set Fx = Nothing
8060                          If br.Code = CodeGLU Or br.Code = CodeGLUP Then
8070                              Set Fx = colFastings("GLU")
8080                          ElseIf br.Code = CodeCHO Or br.Code = CodeCHOP Then
8090                              Set Fx = colFastings("CHO")
8100                          ElseIf br.Code = CodeTRI Or br.Code = CodeTRIP Then
8110                              Set Fx = colFastings("TRI")
8120                          End If
8130                          If Not Fx Is Nothing Then
8140                              High = Fx.FastingHigh
8150                              Low = Fx.FastingLow
8160                          Else
8170                              High = Val(br.High)
8180                              Low = Val(br.Low)
8190                          End If
8200                      Else
8210                          High = Val(br.High)
8220                          Low = Val(br.Low)
8230                      End If
8240                  Else
8250                      High = Val(br.High)
8260                      Low = Val(br.Low)
8270                  End If

8280                  If Low < 10 Then
8290                      strLow = Format(Low, "0.00")
8300                  ElseIf Low < 100 Then
8310                      strLow = Format(Low, "##.0")
8320                  Else
8330                      strLow = Format(Low, " ###")
8340                  End If
8350                  If High < 10 Then
8360                      strHigh = Format(High, "0.00")
8370                  ElseIf High < 100 Then
8380                      strHigh = Format(High, "##.0")
8390                  Else
8400                      strHigh = Format(High, "### ")
8410                  End If

8420                  If IsNumeric(v) Then
8430                      If Val(v) > br.PlausibleHigh Then
8440                          udtPrintLine(lpc).Flag = " X "
8450                          udtPrintLine(lpc).Result = "***"
8460                          Flag = " X"
8470                      ElseIf Val(v) < br.PlausibleLow Then
8480                          udtPrintLine(lpc).Flag = " X "
8490                          udtPrintLine(lpc).Result = "***"
8500                          Flag = " X"
8510                      ElseIf Val(v) > High And High <> 0 Then
8520                          udtPrintLine(lpc).Flag = " H "
8530                          Flag = " H"
8540                      ElseIf Val(v) < Low Then
8550                          udtPrintLine(lpc).Flag = " L "
8560                          Flag = " L"
8570                      Else
8580                          udtPrintLine(lpc).Flag = "   "
8590                          Flag = "  "
8600                      End If
8610                  Else
8620                      udtPrintLine(lpc).Flag = "   "
8630                      Flag = "  "
8640                  End If

8650                  udtPrintLine(lpc).Analyte = Left(br.LongName & Space(25), 25) & " "

8660                  If TestAffected(br) = False Then
8670                      If IsNumeric(v) Then
8680                          Select Case br.Printformat
                                  Case 0: strFormat = "######"
8690                              Case 1: strFormat = "###0.0"
8700                              Case 2: strFormat = "##0.00"
8710                              Case 3: strFormat = "#0.000"
8720                          End Select
8730                          If Trim(udtPrintLine(lpc).Result) <> "***" Then
8740                              udtPrintLine(lpc).Result = Format(v, strFormat)
8750                          End If
8760                      Else
8770                          If Trim(udtPrintLine(lpc).Result) <> "***" Then
8780                              udtPrintLine(lpc).Result = Format(v, strFormat)
8790                          End If
8800                          If Trim(udtPrintLine(lpc).Result) <> "***" Then udtPrintLine(lpc).Result = Format(v, strFormat)
8810                      End If
8820                  Else
8830                      udtPrintLine(lpc).Result = "XXXXXX"
8840                  End If

8850                  sql = "SELECT * FROM Lists WHERE " & _
                            "ListType = 'UN' and Code = '" & br.Units & "'"
8860                  Set tbUN = Cnxn(0).Execute(sql)
8870                  If Not tbUN.EOF Then
8880                      cUnits = Left(tbUN!Text & Space(6), 6)
8890                  Else
8900                      cUnits = Left(br.Units & Space(6), 6)
8910                  End If
8920                  udtPrintLine(lpc).Units = cUnits
8930                  '11-15-23 Zyam
                      If Val(strLow) = 0 And Val(strHigh) = 0 Then
                         udtPrintLine(lpc).NormalRange = " "
                      Else
                         udtPrintLine(lpc).NormalRange = "(" & strLow & "-" & strHigh & ")"
                      End If
                      '11-15-23 Zyam

8940                  udtPrintLine(lpc).Fasting = ""
8950                  If Fasting Then
8960                      If br.Code = CodeGLU Or br.Code = CodeCHO Or br.Code = CodeTRI Or _
                              br.Code = CodeGLUP Or br.Code = CodeCHOP Or br.Code = CodeTRIP Then
8970                          udtPrintLine(lpc).Fasting = "(Fasting)"
8980                      End If
8990                  End If

9000                  If TestAffected(br) = True Then
9010                      udtPrintLine(lpc).Reason = ReasonAffect(br)
9020                  Else
9030                      udtPrintLine(lpc).Reason = ""
9040                  End If

9050                  lpc = lpc + 1
9060              End If
9070          Next
9080      End If

9090      .SelFontSize = 10

9100      If MultiColumn Then
9110          For n = 0 To Val(frmMain.txtMoreThan) - 1
9120              .SelColor = vbBlack
9130              If Trim(udtPrintLine(n).Flag) = "L" Then .SelColor = vbBlue
9140              If Trim(udtPrintLine(n).Flag) = "H" Then .SelColor = vbRed
9150              .SelBold = False
9160              .SelText = udtPrintLine(n).Analyte

9170              If udtPrintLine(n).Flag <> "   " Then
9180                  .SelBold = True
9190              End If
9200              .SelText = udtPrintLine(n).Flag
9210              .SelBold = False
9220              .SelFontSize = 8
9230              .SelText = udtPrintLine(n).Units
9240              .SelText = udtPrintLine(n).NormalRange
9250              .SelFontSize = 10
                  'Now Right Hand Column
9260              .SelText = "     "
9270              .SelText = udtPrintLine(n + Val(frmMain.txtMoreThan)).Analyte
9280              If udtPrintLine(n + Val(frmMain.txtMoreThan)).Flag <> "   " Then
9290                  .SelBold = True
9300              End If
9310              .SelText = udtPrintLine(n + Val(frmMain.txtMoreThan)).Result
9320              .SelText = udtPrintLine(n + Val(frmMain.txtMoreThan)).Flag
9330              .SelBold = False
9340              .SelFontSize = 8
9350              .SelText = udtPrintLine(n + Val(frmMain.txtMoreThan)).Units
9360              .SelText = udtPrintLine(n + Val(frmMain.txtMoreThan)).NormalRange & vbCrLf
9370              .SelFontSize = 10
9380          Next
9390          If Fasting Then
9400              .SelText = "(All above relate to Normal Fasting Ranges.)"
9410          End If
9420      Else
9430          For n = 0 To 35
9440              If Trim(udtPrintLine(n).Analyte) <> "" Then
9450                  .SelColor = vbBlack
9460                  If Trim(udtPrintLine(n).Flag) = "L" Then .SelColor = vbBlue
9470                  If Trim(udtPrintLine(n).Flag) = "H" Then .SelColor = vbRed
9480                  .SelText = Left(" " & Space(10), 10)
9490                  .SelBold = False
9500                  .SelText = udtPrintLine(n).Analyte
9510                  If udtPrintLine(n).Flag <> "   " Then
9520                      .SelBold = True
9530                  End If
9540                  .SelText = udtPrintLine(n).Result
9550                  .SelText = udtPrintLine(n).Flag
9560                  .SelBold = False
9570                  .SelText = udtPrintLine(n).Units
9580                  .SelText = udtPrintLine(n).NormalRange
9590                  .SelText = udtPrintLine(n).Fasting & vbCrLf
9600              End If
9610          Next
9620      End If
          '  Set Cx = Cxs.Load(RP.SampleID)
9630      Set OBS = OBS.Load(RP.SampleID, "Biochemistry", "Demographic")
9640      If Not OBS Is Nothing Then
9650          For Each OB In OBS
9660              Select Case UCase$(OB.Discipline)
                      Case "BIOCHEMISTRY"
9670                      FillCommentLines OB.Comment, 4, Comments(), 80
9680                      For n = 1 To 4
9690                          If Comments(n) <> "" Then .SelText = "  " & Comments(n) & vbCrLf
9700                      Next
9710                  Case "DEMOGRAPHIC"
9720                      FillCommentLines OB.Comment, 2, Comments(), 80
9730                      For n = 1 To 4
9740                          If Comments(n) <> "" Then .SelText = "  " & Comments(n) & vbCrLf
9750                      Next
9760              End Select
9770          Next
9780      End If

9790      If IsDate(tb!SampleDate) Then
9800          SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
9810      Else
9820          SampleDate = ""
9830      End If
9840      If IsDate(RunTime) Then
9850          Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
9860      Else
9870          If IsDate(tb!Rundate) Then
9880              Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
9890          Else
9900              Rundate = ""
9910          End If
9920      End If

9930      If Not IsDate(tb!DoB) And Trim(Sex) = "" Then
9940          .SelColor = vbBlue
9950          .SelText = "          " & "No Sex/DoB given. Normal ranges may not be relevant"
9960      ElseIf Not IsDate(tb!DoB) Then
9970          .SelColor = vbBlue
9980          .SelText = "          " & "No DoB given. Normal ranges may not be relevant"
9990      ElseIf Trim(Sex) = "" Then
10000         .SelColor = vbBlue
10010         .SelText = "          " & "No Sex given. Normal ranges may not be relevant"
10020     End If

10030     .SelColor = vbBlack

10040     CrCnt = 50
10050     PrintFooterRTBFax RP.Initiator, SampleDate, Rundate

10060     .SelStart = 0
10070     .SelLength = Len(.Text)

10080     f = FreeFile
10090     Open SysOptFax(0) & RP.SampleID & "COMP.doc" For Output As f
10100     Print #f, .TextRTF
10110     Close f

10120     SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "COMP.doc"

10130     sql = "SELECT * FROM Reports WHERE 0 = 1"
10140     Set tb = New Recordset
10150     RecOpenServer 0, tb, sql
10160     tb.AddNew
10170     tb!SampleID = RP.SampleID
10180     tb!Name = udtHeading.Name
10190     tb!Dept = "M"
10200     tb!Initiator = RP.Initiator
10210     tb!PrintTime = PrintTime
10220     tb!RepNo = "M" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
10230     tb!PageNumber = 0
10240     tb!Report = .TextRTF
10250     tb!Printer = "FAX - " & RP.FaxNumber
10260     tb.Update
10270 End With

10280 ResetPrinter

10290 Exit Sub

PrintComposit_Error:

      Dim strES As String
      Dim intEL As Integer

10300 intEL = Erl
10310 strES = Err.Description
10320 LogError "modComposit", "PrintComposit", intEL, strES, sql

End Sub

