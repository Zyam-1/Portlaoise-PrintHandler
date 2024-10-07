Attribute VB_Name = "modHaematology"
Option Explicit

Public Function HaemNormalRange(ByVal Analyte As String, _
                                ByVal Sex As String, _
                                ByVal DoB As String, _
                                ByVal SampleDate As String) _
                                As String

      Dim sql As String
      Dim tb As Recordset
      Dim DaysOld As Long
      Dim SexSQL As String
      Dim strFormat As String
      Dim strL As String * 4
      Dim strH As String * 4
      Dim strRange As String


10    On Error GoTo HaemNormalRange_Error



20    Select Case Left(UCase(Sex), 1)
          Case "M"
30            SexSQL = "MaleLow as Low, MaleHigh as High "
40        Case "F"
50            SexSQL = "FemaleLow as Low, FemaleHigh as High "
60        Case ""    'Sex not given
70            HaemNormalRange = "           "
80            Exit Function
90        Case Else
100           SexSQL = "FemaleLow as Low, MaleHigh as High "
110   End Select

120   If Not IsDate(DoB) Then
130       HaemNormalRange = "           "
140       Exit Function
150   End If

160   If IsDate(DoB) Then

170       DaysOld = Abs(DateDiff("d", SampleDate, DoB))

180       sql = "SELECT top 1 PrintFormat, " & _
                SexSQL & _
                "FROM HaemTestDefinitions WHERE " & _
                "AnalyteName = '" & Analyte & "' and AgeFromDays <= '" & DaysOld & "' " & _
                "and AgeToDays >= '" & DaysOld & "' " & _
                "order by AgeFromDays desc, AgeToDays asc"
190   Else
200       sql = "SELECT top 1 PrintFormat, " & _
                SexSQL & _
                "FROM HaemTestDefinitions WHERE " & _
                "AnalyteName = '" & Analyte & "' and AgeFromDays <='9125' " & _
                "and AgeToDays >= '9125'"
210   End If

220   Set tb = New Recordset
230   RecOpenClient 0, tb, sql

240   strRange = "(    -    )"

250   If Not tb.EOF Then
260       Select Case tb!Printformat
              Case 0: strFormat = "0"
270           Case 1: strFormat = "0.0"
280           Case 2: strFormat = "0.00"
290           Case 3: strFormat = "0.000"
300       End Select

310       If tb!High <> 999 Then
320           RSet strL = Format(tb!Low, strFormat)
330           Mid(strRange, 2, 4) = strL
340           LSet strH = Format(tb!High, strFormat)
350           Mid(strRange, 7, 4) = strH
360       End If
370   Else
380       strRange = "(    -    )"
390   End If

400   HaemNormalRange = strRange

410   Exit Function

HaemNormalRange_Error:

      Dim strES As String
      Dim intEL As Integer

420   intEL = Erl
430   strES = Err.Description
440   LogError "modHaematology", "HaemNormalRange", intEL, strES, sql

End Function




Public Function PrintResultHaemAdvia() As Boolean

      Dim tb As Recordset
      Dim tbH As Recordset
      Dim tbd As Recordset
      Dim n As Integer
      Dim Sex As String
      Dim fbc As Integer
      Dim TotalRetics As Single
      Dim DoB As String
      Dim Flag As String
      Dim sql As String
      'Dim Cx As Comment
      'Dim Cxs As New Comments
      Dim OB As Observation
      Dim OBS As New Observations
10    ReDim Comments(1 To 16) As String
20    ReDim commentstemp(1 To 16) As String
      Dim SampleDate As String
      Dim Rundate As String
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
      Dim copies As Long
      Dim d As Long
      Dim Clin As String
      Dim f As Integer
      Dim Fontz1 As Integer
      Dim Fontz2 As Integer
      Dim Fontz3 As Integer
      Dim SNP As Integer
      Dim PrintTime As String
      Dim InTheMiddle As Boolean
      Dim AuthorisedBy As String
      Const MaxResultLines As Integer = 27
      Dim CommentLines As Integer
      Dim cIndex As Integer
      Dim TotalPages As Integer
      Dim FooterStartLine As Integer
      Dim HosName As String

30    On Error GoTo PrintResultHaemAdvia_Error

40    TotalPages = 1
50    InTheMiddle = False
60    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

70    SNP = 4
80    PrintResultHaemAdvia = True
90    frmMain.gDiff.Rows = 2
100   frmMain.gDiff.AddItem ""
110   frmMain.gDiff.RemoveItem 1

120   sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
130   Set tb = New Recordset
140   RecOpenClient 0, tb, sql
150   If tb.EOF Then Exit Function
160   HosName = tb!Hospital
170   If IsDate(tb!SampleDate) Then
180       SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
190   Else
200       SampleDate = ""
210   End If

220   sql = "SELECT * FROM HaemResults WHERE " & _
            "SampleID = '" & RP.SampleID & "' "

230   If RP.Department = "H" Then
240       sql = sql & " and valid = 1"
250   End If

260   Set tbH = New Recordset
270   RecOpenServer 0, tbH, sql
280   If tbH.EOF Then Exit Function



290   AuthorisedBy = GetAuthorisedBy(tbH!Operator & "")

300   DoB = tb!DoB & ""

310   fbc = Trim(tbH!WBC & "") <> ""

320   If IsDate(tbH!RunDateTime) Then
330       Rundate = Format(tbH!RunDateTime, "dd/mmm/yyyy hh:mm")
340   Else
350       Rundate = ""
360   End If

370   sql = "SELECT * FROM Differentials WHERE " & _
            "RunNumber = '" & RP.SampleID & "' AND PrnDiff = 1 "
380   Set tbd = New Recordset
390   RecOpenClient 0, tbd, sql
400   If Not tbd.EOF Then
410       DiffFound = True
420       For n = 0 To 14
430           W = Trim(tbd("Wording" & Format(n)) & "")
440           p = IIf(Val(tbd("P" & Format(n)) & "") = 0, "", tbd("P" & Format(n)))
450           a = IIf(Val(tbd("A" & Format(n)) & "") = 0, "", tbd("A" & Format(n)))
460           frmMain.gDiff.AddItem W & vbTab & p & vbTab & a
470       Next
480   End If

490   Select Case Left(UCase(tb!Sex & ""), 1)
      Case "M": Sex = "M"
500   Case "F": Sex = "F"
510   Case Else: Sex = ""
520   End Select


530   ClearUdtHeading

540   With udtHeading
550       .SampleID = RP.SampleID
560       .Dept = "Haematology"
570       If RP.Department = "K" Then .Dept = "Draft Haem"
580       .Name = tb!PatName & ""
590       .Ward = RP.Ward
600       .DoB = DoB
610       .Chart = tb!Chart & ""
620       .Clinician = RP.Clinician
630       .Address0 = tb!Addr0 & ""
640       .Address1 = tb!Addr1 & ""
650       .GP = RP.GP
660       .Sex = tb!Sex & ""
670       .Hospital = tb!Hospital & ""
680       .SampleDate = tb!SampleDate & ""
690       .RecDate = tb!RecDate & ""
700       .Rundate = tbH!Rundate & ""
710       .GpClin = Clin
720       .SampleType = ""
730       .Notes = "*** - DRAFT REPORT. DO NOT RELEASE !!! - ***"
740       .DocumentNo = GetOptionSetting("HaemMainDocumentNo", "")
750       .AandE = tb!AandE & ""
760   End With

      'Load comment here so that total pages could be calculated
770   Set OBS = OBS.Load(RP.SampleID, "Haematology", "Demographic")
780   If Not OBS Is Nothing Then
790       For Each OB In OBS

800           Select Case UCase$(OB.Discipline)
              Case "HAEMATOLOGY"
810               FillCommentLines OB.Comment, 16, Comments(), 87
820               For n = 1 To 16

830                   If Trim(Comments(n) & "") <> "" Then
840                       CommentLines = CommentLines + 1
850                       commentstemp(CommentLines) = Comments(n)
860                   End If
870               Next
880           Case "DEMOGRAPHIC"
890               FillCommentLines OB.Comment, 8, Comments(), 87
900               For n = 1 To 8
910                   If Trim(Comments(n) & "") <> "" Then
920                       CommentLines = CommentLines + 1
930                       commentstemp(CommentLines) = Comments(n)
940                   End If
950               Next
960           End Select
970       Next
980   End If
990   FooterStartLine = GetOptionSetting("FooterStartLine", 33)
      'Total pages are going to be  (HeadingLines+BodyLines+CommentLines) / footerline
1000  If fbc Then
1010      TotalPages = IIf((FooterStartLine - 15 - 18 - CommentLines) >= 0, 1, 2)

1020  Else
1030      TotalPages = IIf((FooterStartLine - 15 - 14 - CommentLines) >= 0, 1, 2)
1040  End If


1050  If RP.FaxNumber <> "" Then
1060      PrintHeadingRTBFax
1070  Else
1080      PrintHeadingRTB "Page 1 Of " & TotalPages
1090  End If

1100  If RP.FaxNumber <> "" Then
1110      Fontz1 = 9
1120      Fontz2 = 5
1130      SNP = 0
1140  Else
1150      Fontz1 = 10
1160      Fontz2 = 6
1170      SNP = 5
1180  End If

1190  With frmRichText.rtb
1200      If fbc Then
              '*****************Start WBC + Nuet + Lymph*******************
1210          Flag = InterpH(tbH!WBC & "", "WBC", Sex, DoB, SampleDate)
1220          PrintTest "WBC", IIf(Flag = "X", "XXXXX", tbH!WBC & ""), GetUnitsHaem("WBC", "x10^9/l"), Flag, HaemNormalRange("WBC", Sex, DoB, SampleDate), , 10, , Trim(Flag) <> "", Trim(Flag) <> "", , Trim(Flag) <> ""

1230          If DiffFound = True Then
1240              For n = 1 To frmMain.gDiff.Rows - 1
1250                  If InStr(UCase(frmMain.gDiff.TextMatrix(n, 0)), "NEUT") > 0 Then
1260                      neut = Format(frmMain.gDiff.TextMatrix(n, 2), "##0.0##")
1270                      Exit For
1280                  End If
1290              Next
1300          Else
1310              neut = tbH!neuta & ""
1320          End If
1330          If Trim(neut) <> "" Then
1340              Flag = InterpH(neut & "", "NEUTA", Sex, DoB, SampleDate)
1350              PrintTest "Neut", IIf(Flag = "X", "XXXXX", neut), GetUnitsHaem("NEUTA", "x10^9/l"), Flag, HaemNormalRange("NEUTA", Sex, DoB, SampleDate), True, 10, , Trim(Flag) <> "", Trim(Flag) <> "", , Trim(Flag) <> ""
1360              CrCnt = CrCnt + 1
1370          Else
1380              .SelText = vbCrLf
1390              CrCnt = CrCnt + 1
1400          End If
              '*****************End WBC + Nuet   *******************


              '*****************Start Lymph*******************
1410          PrintTextRTB frmRichText.rtb, String(39, " ") + String(SNP, " "), 10

1420          If DiffFound = True Then
1430              For n = 1 To frmMain.gDiff.Rows - 1
1440                  If InStr(UCase(frmMain.gDiff.TextMatrix(n, 0)), "LYM") > 0 Then
1450                      lym = Format(frmMain.gDiff.TextMatrix(n, 2), "##0.0##")
1460                      Exit For
1470                  End If
1480              Next
1490          Else
1500              lym = tbH!lyma & ""
1510          End If

1520          If Trim(lym) <> "" Then
1530              Flag = InterpH(lym & "", "LYMA", Sex, DoB, SampleDate)
1540              PrintTest "Lymph", IIf(Flag = "X", "XXXXX", lym), GetUnitsHaem("LYMA", "x10^9/l"), Flag, HaemNormalRange("LYMA", Sex, DoB, SampleDate), True, 10, , Trim(Flag) <> "", Trim(Flag) <> "", , Trim(Flag) <> ""
1550              CrCnt = CrCnt + 1
1560          Else
1570              .SelText = vbCrLf
1580              CrCnt = CrCnt + 1
1590          End If

              '*****************End Lymph *******************


              '*****************Start RBC + Mono *******************

1600          Flag = InterpH(tbH!RBC & "", "RBC", Sex, DoB, SampleDate)
1610          PrintTest "RBC", IIf(Flag = "X", "XXXXX", tbH!RBC & ""), GetUnitsHaem("RBC", "x10^12/l"), Flag, HaemNormalRange("RBC", Sex, DoB, SampleDate), , 10, , Trim(Flag) <> "", Trim(Flag) <> "", , Trim(Flag) <> ""

1620          If DiffFound = True Then
1630              For n = 1 To frmMain.gDiff.Rows - 1
1640                  If InStr(UCase(frmMain.gDiff.TextMatrix(n, 0)), "MON") > 0 Then
1650                      mono = Format(frmMain.gDiff.TextMatrix(n, 2), "##0.0##")
1660                      Exit For
1670                  End If
1680              Next
1690          Else
1700              mono = tbH!MonoA & ""
1710          End If

1720          If Trim(mono) <> "" Then
1730              Flag = InterpH(mono & "", "MONOA", Sex, DoB, SampleDate)
1740              PrintTest "Mono", IIf(Flag = "X", "XXXXX", mono), GetUnitsHaem("MONOA", "x10^9/l"), Flag, HaemNormalRange("MONOA", Sex, DoB, SampleDate), True, 10, , Trim(Flag) <> "", Trim(Flag) <> "", , Trim(Flag) <> ""
1750              CrCnt = CrCnt + 1
1760          Else
1770              .SelText = vbCrLf
1780              CrCnt = CrCnt + 1
1790          End If

              '*****************End RBC + Mono *******************

              '*****************Start EOS *******************

1800          PrintTextRTB frmRichText.rtb, String(39, " ") + String(SNP, " "), 10

1810          If DiffFound = True Then
1820              For n = 1 To frmMain.gDiff.Rows - 1
1830                  If InStr(UCase(frmMain.gDiff.TextMatrix(n, 0)), "EOS") > 0 Then
1840                      eos = Format(frmMain.gDiff.TextMatrix(n, 2), "##0.0##")
1850                      Exit For
1860                  End If
1870              Next
1880          Else
1890              eos = tbH!eosa
1900          End If

1910          If Trim(eos & "") <> "" Then
1920              Flag = InterpH(eos & "", "EOSA", Sex, DoB, SampleDate)
1930              PrintTest "Eos", IIf(Flag = "X", "XXXXX", eos), GetUnitsHaem("EOSA", "x10^9/l"), Flag, HaemNormalRange("EOSA", Sex, DoB, SampleDate), True, 10, , Trim(Flag) <> "", Trim(Flag) <> "", , Trim(Flag) <> ""
1940              CrCnt = CrCnt + 1
1950          Else
1960              .SelText = vbCrLf
1970              CrCnt = CrCnt + 1
1980          End If

              '*****************Start EOS *******************

1990      End If


          '*****************Start HGB + Bas *******************

2000      If Trim(tbH!Hgb & "") <> "" Then
2010          Flag = InterpH(tbH!Hgb & "", "Hgb", Sex, DoB, SampleDate)
2020          PrintTest "Hgb", IIf(Flag = "X", "XXXXX", tbH!Hgb & ""), GetUnitsHaem("Hgb", "g/dl"), Flag, HaemNormalRange("Hgb", Sex, DoB, SampleDate), , 10, , Trim(Flag) <> "", Trim(Flag) <> "", , Trim(Flag) <> ""
2030      Else
2040          .SelText = String(40, " ")
2050      End If

2060      If fbc Then
2070          If DiffFound = True Then
2080              For n = 1 To frmMain.gDiff.Rows - 1
2090                  If InStr(UCase(frmMain.gDiff.TextMatrix(n, 0)), "BAS") > 0 Then
2100                      bas = Format(frmMain.gDiff.TextMatrix(n, 2), "##0.0##")
2110                      Exit For
2120                  End If
2130              Next
2140          Else
2150              bas = tbH!basa
2160          End If

2170          If Trim(bas & "") <> "" Then
2180              Flag = InterpH(bas & "", "BASA", Sex, DoB, SampleDate)
2190              PrintTest "Bas", IIf(Flag = "X", "XXXXX", bas), GetUnitsHaem("BASA", "x10^9/l"), Flag, HaemNormalRange("BASA", Sex, DoB, SampleDate), True, 10, , Trim(Flag) <> "", Trim(Flag) <> "", , Trim(Flag) <> ""
2200              CrCnt = CrCnt + 1
2210          Else
2220              .SelText = vbCrLf
2230              CrCnt = CrCnt + 1
2240          End If
2250      End If


          '*****************End HGB + Bas *******************

          '*****************Start LUCA *******************

2260      PrintTextRTB frmRichText.rtb, String(39, " ") + String(SNP, " "), 10

2270      luc = ""
2280      Flag = ""
2290      If DiffFound = True Then
2300          luc = Format(frmMain.gDiff.TextMatrix(7, 2), "##0.0##")
2310      Else
2320          luc = tbH!luca & ""
2330      End If
2340      If Trim(luc & "") <> "" Then
2350          If DiffFound = True Then
2360              PrintTest Initial2Upper(frmMain.gDiff.TextMatrix(7, 0)), luc, "", "", "", True, 10
2370          Else
2380              Flag = InterpH(luc & "", "LUCA", Sex, DoB, SampleDate)
2390              PrintTest "Luc", IIf(Flag = "X", "XXXXX", luc), GetUnitsHaem("LUCA", "x10^9/l"), Flag, IIf(DiffFound, "", HaemNormalRange("LUCA", Sex, DoB, SampleDate)), True, 10, , Trim(Flag) <> "", Trim(Flag) <> "", , Trim(Flag) <> ""
2400          End If
2410          CrCnt = CrCnt + 1
2420      Else
2430          .SelText = vbCrLf
2440          CrCnt = CrCnt + 1
2450      End If
          '*****************End LUCA *******************


          '*****************Start HCT *******************

2460      If fbc Then
2470          Flag = InterpH(tbH!Hct & "", "Hct", Sex, DoB, SampleDate)
2480          PrintTest "Hct", IIf(Flag = "X", "XXXXX", tbH!Hct & ""), GetUnitsHaem("Hct", "l/l"), Flag, HaemNormalRange("Hct", Sex, DoB, SampleDate), , 10, , Trim(Flag) <> "", Trim(Flag) <> "", , Trim(Flag) <> ""
2490      End If

2500      If DiffFound = True Then
2510          luc = ""
2520          If frmMain.gDiff.TextMatrix(8, 1) <> "" Then
2530              luc = Format(frmMain.gDiff.TextMatrix(8, 2), "##0.0##")
2540              If Trim(luc & "") <> "" Then
2550                  PrintTest Initial2Upper(frmMain.gDiff.TextMatrix(8, 0)), luc, "", "", "", True, 10
2560                  CrCnt = CrCnt + 1
2570              Else
2580                  .SelText = vbCrLf
2590                  CrCnt = CrCnt + 1
2600              End If
2610          Else
2620              .SelText = vbCrLf
2630              CrCnt = CrCnt + 1
2640          End If
2650      Else
2660          .SelText = vbCrLf
2670          CrCnt = CrCnt + 1
2680      End If

          '*****************End HCT + Unknown *******************

          '*****************Start  Unknown *******************
2690      PrintTextRTB frmRichText.rtb, String(39, " ") + String(SNP, " "), 10

2700      If DiffFound = True Then
2710          luc = ""
2720          If frmMain.gDiff.TextMatrix(9, 1) <> "" Then
2730              luc = Format(frmMain.gDiff.TextMatrix(9, 2), "##0.0##")
2740              If Trim(luc & "") <> "" Then
2750                  PrintTest Initial2Upper(frmMain.gDiff.TextMatrix(9, 0)), luc, "", "", "", True, 10
2760                  CrCnt = CrCnt + 1
2770              Else
2780                  .SelText = vbCrLf
2790                  CrCnt = CrCnt + 1
2800              End If
2810          Else
2820              .SelText = vbCrLf
2830              CrCnt = CrCnt + 1
2840          End If
2850      Else
2860          .SelText = vbCrLf
2870          CrCnt = CrCnt + 1
2880      End If

          '*****************End  Unknown *******************


          '*****************Start MCV + Unknown *******************

2890      If fbc Then
2900          Flag = InterpH(tbH!MCV & "", "MCV", Sex, DoB, SampleDate)
2910          PrintTest "MCV", IIf(Flag = "X", "XXXXX", tbH!MCV & ""), GetUnitsHaem("MCV", "fl"), Flag, HaemNormalRange("MCV", Sex, DoB, SampleDate), , 10, , Trim(Flag) <> "", Trim(Flag) <> "", , Trim(Flag) <> ""
2920      End If
2930      Flag = ""
2940      If DiffFound = True Then
2950          luc = ""
2960          If frmMain.gDiff.TextMatrix(10, 1) <> "" Then
2970              luc = Format(frmMain.gDiff.TextMatrix(10, 2), "##0.0##")
2980              If Trim(luc & "") <> "" Then
2990                  PrintTest Initial2Upper(frmMain.gDiff.TextMatrix(10, 0)), luc, "", "", "", True, 10
3000                  CrCnt = CrCnt + 1
3010              Else
3020                  .SelText = vbCrLf
3030                  CrCnt = CrCnt + 1
3040              End If
3050          Else
3060              .SelText = vbCrLf
3070              CrCnt = CrCnt + 1
3080          End If
3090      Else
3100          .SelText = vbCrLf
3110          CrCnt = CrCnt + 1
3120      End If

          '*****************End MCV + Unknown *******************


          '*****************Start Unknown *******************

3130      PrintTextRTB frmRichText.rtb, String(39, " ") + String(SNP, " "), 10

3140      Flag = ""
3150      If DiffFound = True Then
3160          luc = ""
3170          If frmMain.gDiff.TextMatrix(11, 1) <> "" Then
3180              luc = Format(frmMain.gDiff.TextMatrix(11, 2), "##0.0##")
3190              If Trim(luc & "") <> "" Then
3200                  PrintTest Initial2Upper(frmMain.gDiff.TextMatrix(11, 0)), luc, "", "", "", True, 10
3210                  CrCnt = CrCnt + 1
3220              Else
3230                  .SelText = vbCrLf
3240                  CrCnt = CrCnt + 1
3250              End If
3260          Else
3270              .SelText = vbCrLf
3280              CrCnt = CrCnt + 1
3290          End If
3300      Else
3310          .SelText = vbCrLf
3320          CrCnt = CrCnt + 1
3330      End If

          '*****************End  Unknown *******************


          '*****************Start Comment Lines + Rest of tests on right side *******************





3340      If fbc Then
3350          Flag = InterpH(tbH!MCH & "", "MCH", Sex, DoB, SampleDate)
3360          PrintTest "MCH", IIf(Flag = "X", "XXXXX", tbH!MCH & ""), GetUnitsHaem("MCH", "pg"), Flag, HaemNormalRange("MCH", Sex, DoB, SampleDate), , 10, , Trim(Flag) <> "", Trim(Flag) <> "", , Trim(Flag) <> ""
3370      End If
3380      Flag = ""
3390      If DiffFound = True Then
3400          luc = ""
3410          If frmMain.gDiff.TextMatrix(12, 1) <> "" Then
3420              luc = Format(frmMain.gDiff.TextMatrix(12, 2), "##0.0##")
3430              If Trim(luc & "") <> "" Then
3440                  PrintTest Initial2Upper(Format(frmMain.gDiff.TextMatrix(12, 0))), luc, "", "", "", True, 10
3450                  CrCnt = CrCnt + 1
3460              Else
3470                  .SelText = vbCrLf
3480                  CrCnt = CrCnt + 1
3490              End If
3500          Else
3510              .SelText = vbCrLf
3520              CrCnt = CrCnt + 1
3530          End If
3540      Else
3550          .SelText = vbCrLf
3560          CrCnt = CrCnt + 1
3570      End If


          'ESR*******************
3580      PrintTextRTB frmRichText.rtb, String(39, " ") + String(SNP, " "), 10

3590      If Trim(tbH!ESR & "") <> "" Then
3600          Flag = InterpH(tbH!ESR & "", "ESR", Sex, DoB, SampleDate)
3610          PrintTest "ESR", IIf(Flag = "X", "XXXXX", Trim(tbH!ESR & "")), GetUnitsHaem("ESR", "mm/hr"), Flag, HaemNormalRange("ESR", Sex, DoB, SampleDate), True, 10, , Trim(Flag) <> "", Trim(Flag) <> "", , Trim(Flag) <> ""
3620          CrCnt = CrCnt + 1
3630      Else
3640          .SelText = vbCrLf
3650          CrCnt = CrCnt + 1
3660      End If
          'RETA******************
3670      If fbc Then
3680          Flag = InterpH(tbH!MCHC & "", "MCHC", Sex, DoB, SampleDate)
3690          PrintTest "MCHC", IIf(Flag = "X", "XXXXX", tbH!MCHC & ""), GetUnitsHaem("MCHC", "g/dl"), Flag, HaemNormalRange("MCHC", Sex, DoB, SampleDate), , 10, , Trim(Flag) <> "", Trim(Flag) <> "", , Trim(Flag) <> ""
3700      End If
3710      If Trim(tbH!reta & "") <> "" Then
3720          Flag = InterpH(tbH!reta & "", "RETA", Sex, DoB, SampleDate)
3730          If Trim(tbH!reta) & "" <> "" And Trim(tbH!reta) <> "?" Then TotalRetics = Val(Trim(tbH!reta & "")) Else Flag = "X"
3740          PrintTest "Retics", IIf(Flag = "X", "XXXXX", CStr(TotalRetics)), GetUnitsHaem("RETA", "x10^9/l"), Flag, HaemNormalRange("RETA", Sex, DoB, SampleDate), True, 10, , Trim(Flag) <> "", Trim(Flag) <> "", , Trim(Flag) <> ""
3750          CrCnt = CrCnt + 1
3760      Else
3770          .SelText = vbCrLf
3780          CrCnt = CrCnt + 1
3790      End If
          'Monospot***************
3800      PrintTextRTB frmRichText.rtb, String(39, " ") + String(SNP, " "), 10
3810      If Trim(tbH!Monospot & "") <> "" And Trim(tbH!Monospot & "") <> "?" Then
3820          If tbH!Monospot = "N" Then luc = "Negative"
3830          If tbH!Monospot = "P" Then luc = "Positive"
3840          If tbH!Monospot = "I" Then luc = "Inconclusive"
3850          PrintTest "Infectious Mononucleosis Screen", luc, "", "", "", True, 8, True
3860          CrCnt = CrCnt + 1
3870      Else
3880          .SelText = vbCrLf
3890          CrCnt = CrCnt + 1
3900      End If
          'Asot*******************
3910      If fbc Then
3920          Flag = InterpH(tbH!rdwcv & "", "RDW", Sex, DoB, SampleDate)
3930          PrintTest "RDW", IIf(Flag = "X", "XXXXX", tbH!rdwcv & ""), GetUnitsHaem("RDW", "%"), Flag, HaemNormalRange("RDW", Sex, DoB, SampleDate), , 10, , Trim(Flag) <> "", Trim(Flag) <> "", , Trim(Flag) <> ""
3940      Else
3950          PrintTextRTB frmRichText.rtb, String(39, " ") + String(SNP, " "), 10
3960      End If

3970      If Trim(tbH!tasot & "") <> "" And Trim(tbH!tasot & "") <> "?" Then
3980          PrintTest "Asot", Trim(tbH!tasot) & "", "", "", "", True, 10, True
3990          CrCnt = CrCnt + 1
4000      Else
4010          .SelText = vbCrLf
4020          CrCnt = CrCnt + 1
4030      End If

          'Malaria****************
4040      PrintTextRTB frmRichText.rtb, String(39, " ") + String(SNP, " "), 10
4050      If Trim(tbH!malaria & "") <> "" And Trim(tbH!malaria & "") <> "?" Then
4060          PrintTest "Malaria Antigen Screen", Trim(tbH!malaria) & "", "", "", "", True, 10, True
4070      Else
4080          .SelText = vbCrLf
4090      End If
4100      CrCnt = CrCnt + 1
          'Sickledex**************
4110      If Trim(tbH!Plt) & "" <> "" Then
4120          Flag = InterpH(tbH!Plt & "", "Plt", Sex, DoB, SampleDate)
4130          PrintTest "Plt", IIf(Flag = "X", "XXXXX", tbH!Plt & ""), GetUnitsHaem("Plt", "x10^9/l"), Flag, HaemNormalRange("Plt", Sex, DoB, SampleDate), , 10, , Trim(Flag) <> "", Trim(Flag) <> "", , Trim(Flag) <> ""
4140      Else
4150          PrintTextRTB frmRichText.rtb, String(39, " ") + String(SNP, " "), 10
4160      End If
4170      If Trim(tbH!sickledex & "") <> "" And Trim(tbH!sickledex & "") <> "?" Then
4180          PrintTest "Sickle Cell Screen", Trim(tbH!sickledex) & "", "", "", "", True, 10, True
4190          CrCnt = CrCnt + 1
4200      Else
4210          .SelText = vbCrLf
4220          CrCnt = CrCnt + 1
4230      End If

          'Rheumatoid*************
4240      If GetOptionSetting("PRINTNRBC", 0) = 1 Then
4250          PrintTest "NRBC", Trim(tbH!nrbcp & ""), "", "", "", , 10
4260      Else
4270          PrintTest "", "", "", "", "", , 10
4280      End If
          '    If PrintCommentOnFirstPage Then
          '        PrintTextRTB frmRichText.rtb, FormatString("Comments:", 39) + String(SNP, " "), 10, True
          '    Else
          '        PrintTextRTB frmRichText.rtb, FormatString(" ", 39) + String(SNP, " "), 10, True
          '    End If
4290      If Trim(tbH!tra & "") <> "" And Trim(tbH!tra & "") <> "?" Then
4300          PrintTest "Rheumatoid Factor", Trim(tbH!tra) & "", "", "", "", True, 10, True
4310          CrCnt = CrCnt + 1
4320      Else
4330          .SelText = vbCrLf
4340          CrCnt = CrCnt + 1
4350      End If


          'NRBC*******************



          '*****************End Comment Lines + Unknown *******************


          'Not sure about md4 and md5 results. if required can be added later.
          '.SelText = tbH!md4 & "" & vbCrLf
          '    CrCnt = CrCnt + 1
          '
          '.SelText = tbH!md5 & "" & vbCrLf
          'CrCnt = CrCnt + 1






4360      If DiffFound Then
4370          PrintTextRTB frmRichText.rtb, FormatString("Manual Differential Reported", 39) + String(SNP, " ") & vbCrLf, 10, True
4380          CrCnt = CrCnt + 1
4390      End If

4400      If Not IsDate(tb!DoB) Or Trim(Sex) = "" Then
4410          PrintTextRTB frmRichText.rtb, FormatString("No Sex/DoB given. No ref range applied ", 39) + String(SNP, " ") & vbCrLf, 10, True
4420          CrCnt = CrCnt + 1
4430      End If
4440      cIndex = 0
4450      If CrCnt < GetOptionSetting("FooterStartLine", 33) And CommentLines > 0 Then
              '        PrintTextRTB frmRichText.rtb, FormatString("Comments ", 82) & vbCrLf, 10, True
              '        CrCnt = CrCnt + 1
4460          For n = 1 To (GetOptionSetting("FooterStartLine", 33) - CrCnt)
4470              PrintTextRTB frmRichText.rtb, FormatString(commentstemp(n), 87) & vbCrLf, 9
4480              CrCnt = CrCnt + 1
4490              cIndex = cIndex + 1
4500          Next n
4510      End If

          'Print Footer and Save Report
4520      If RP.FaxNumber <> "" Then
4530          PrintFooterRTBFax RP.Initiator, SampleDate, Rundate
4540          f = FreeFile
4550          Open SysOptFax(0) & RP.SampleID & "HAEM.doc" For Output As f
4560          Print #f, .TextRTF
4570          Close f
4580          SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "HAEM.doc"
4590      Else
4600          PrintFooterRTB AuthorisedBy, SampleDate, Rundate, , "Haem", HosName
4610          .SelStart = 0
              'Do not print if Doctor is disabled in DisablePrinting
              '*******************************************************************
4620          If CheckDisablePrinting(RP.Ward, "Haematology") Then

4630          ElseIf CheckDisablePrinting(RP.GP, "Haematology") Then
4640          Else
4650              .SelPrint Printer.hdc
4660          End If
              '*******************************************************************
              '.SelPrint Printer.hDC
4670      End If

4680      sql = "SELECT * FROM Reports WHERE 0 = 1"
4690      Set tb = New Recordset
4700      RecOpenServer 0, tb, sql
4710      tb.AddNew
4720      tb!SampleID = RP.SampleID
4730      tb!Name = udtHeading.Name
4740      tb!Dept = RP.Department
4750      tb!Initiator = RP.Initiator
4760      tb!PrintTime = PrintTime
4770      tb!RepNo = "0" & RP.Department & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
4780      tb!PageNumber = 0
4790      tb!Report = .TextRTF
4800      tb!Printer = Printer.DeviceName
4810      tb.Update


4820      If cIndex < CommentLines Then
4830          If RP.FaxNumber <> "" Then
4840              PrintHeadingRTBFax
4850          Else
4860              PrintHeadingRTB "Page 2 Of " & TotalPages
4870          End If
4880          For n = cIndex + 1 To UBound(commentstemp)
4890              If commentstemp(n) <> "" Then
4900                  PrintTextRTB frmRichText.rtb, FormatString(commentstemp(n), 87) & vbCrLf, 9
4910                  CrCnt = CrCnt + 1
4920              End If
4930          Next n

4940          If RP.Department = "K" Then
4950              .SelBold = True
4960              .SelText = "                     *** - DRAFT REPORT. DO NOT RELEASE !!! - ***" & vbCrLf
4970              .SelBold = False
4980              CrCnt = CrCnt + 1
4990          End If
5000          If RP.FaxNumber <> "" Then
5010              PrintFooterRTBFax RP.Initiator, SampleDate, Rundate
5020              f = FreeFile
5030              Open SysOptFax(0) & RP.SampleID & "HAEM.doc" For Output As f
5040              Print #f, .TextRTF
5050              Close f
5060              SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "HAEM.doc"
5070          Else
5080              PrintFooterRTB AuthorisedBy, SampleDate, Rundate, , "Haem", HosName
5090              .SelStart = 0
                  'Do not print if Doctor is disabled in DisablePrinting
                  '*******************************************************************
5100              If CheckDisablePrinting(RP.Ward, "Haematology") Then

5110              ElseIf CheckDisablePrinting(RP.GP, "Haematology") Then
5120              Else
5130                  .SelPrint Printer.hdc
5140              End If
                  '*******************************************************************
                  '.SelPrint Printer.hDC
5150          End If

5160          sql = "SELECT * FROM Reports WHERE 0 = 1"
5170          Set tb = New Recordset
5180          RecOpenServer 0, tb, sql
5190          tb.AddNew
5200          tb!SampleID = RP.SampleID
5210          tb!Name = udtHeading.Name
5220          tb!Dept = RP.Department
5230          tb!Initiator = RP.Initiator
5240          tb!PrintTime = PrintTime
5250          tb!RepNo = "0" & RP.Department & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
5260          tb!PageNumber = 1
5270          tb!Report = .TextRTF
5280          tb!Printer = Printer.DeviceName
5290          tb.Update
5300      End If


5310  End With

5320  ResetPrinter

5330  Exit Function

PrintResultHaemAdvia_Error:

      Dim strES As String
      Dim intEL As Integer

5340  intEL = Erl
5350  strES = Err.Description

5360  sql = "Delete FROM printpending WHERE SampleID = '" & RP.SampleID & "' and department = '" & RP.Department & "'"
5370  Cnxn(0).Execute sql

5380  LogError "modHaematology", "PrintResultHaemAdvia", intEL, strES, sql

End Function



Public Function PrintResultHaem(Optional ByVal PrintA4 As Boolean = True) As Boolean

      Dim tb         As Recordset
      Dim tbH        As Recordset
      Dim tbd        As Recordset
      Dim n          As Integer
      Dim Sex        As String
      Dim fbc        As Integer
      Dim TotalRetics As Single
      Dim DoB        As String
      Dim Flag       As String
      Dim sql        As String

      Dim OB         As Observation
      Dim OBS        As New Observations
10    ReDim Comments(1 To 16) As String
20    ReDim commentstemp(1 To 16) As String
      Dim SampleDate As String
      Dim Rundate    As String
      Dim lym        As String
      Dim neut       As String
      Dim mono       As String
      Dim eos        As String
      Dim bas        As String
      Dim luc        As String
      Dim dv         As String
      Dim DiffFound  As Boolean
      Dim p          As String
      Dim a          As String
      Dim W          As String
      Dim Clin       As String
      Dim f          As Integer
      Dim Fontz1     As Integer
      Dim Fontz2     As Integer
      Dim SNP        As Integer
      Dim PrintTime  As String
      Dim InTheMiddle As Boolean
      Dim AuthorisedBy As String
      Dim HosName    As String

      Dim udtPrintLine() As ResultLine
      Dim lpc        As Integer
      Dim TotalLines As Integer
      Dim CommentLines As Integer
      Dim PerPageLines As Integer
      Dim BodyLines  As Integer
      Dim FooterLines As Integer
      Dim LineNoStartComment As Integer
      Dim TotalPages As Integer
      Dim i          As Integer
      Dim PageNumber As Integer
      Dim FontBold   As Boolean


30    On Error GoTo PrintResultHaemAdvia_Error


40    If PrintA4 Then
50        TotalLines = 100
60        CommentLines = 10
70        PerPageLines = 77
80        FooterLines = 3
90    Else
100       TotalLines = 100
110       CommentLines = 4
120       PerPageLines = 35
130       FooterLines = 3
140   End If


150   TotalPages = 1
160   InTheMiddle = False
170   PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

180   ReDim lp(0 To TotalLines) As String
190   ReDim udtPrintLine(0 To TotalLines) As ResultLine
200   ReDim Comments(1 To CommentLines) As String



      'Clear All
210   For n = 0 To TotalLines
220       udtPrintLine(n).Analyte = ""
230       udtPrintLine(n).Result = ""
240       udtPrintLine(n).Flag = ""
250       udtPrintLine(n).Units = ""
260       udtPrintLine(n).NormalRange = ""
270       udtPrintLine(n).Fasting = ""
280       udtPrintLine(n).Reason = ""
290   Next


300   SNP = 4
310   PrintResultHaem = True
320   frmMain.gDiff.Rows = 2
330   frmMain.gDiff.AddItem ""
340   frmMain.gDiff.RemoveItem 1

350   sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
360   Set tb = New Recordset
370   RecOpenClient 0, tb, sql
380   If tb.EOF Then Exit Function

390   HosName = tb!Hospital
400   If IsDate(tb!SampleDate) Then
410       SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
420   Else
430       SampleDate = ""
440   End If

450   sql = "SELECT * FROM HaemResults WHERE " & _
            "SampleID = '" & RP.SampleID & "' "

460   If RP.Department = "H" Then
470       sql = sql & " and valid = 1"
480   End If

490   Set tbH = New Recordset
500   RecOpenServer 0, tbH, sql
510   If tbH.EOF Then Exit Function

520   AuthorisedBy = GetAuthorisedBy(tbH!Operator & "")

530   DoB = tb!DoB & ""

540   fbc = Trim(tbH!WBC & "") <> ""

550   If IsDate(tbH!RunDateTime) Then
560       Rundate = Format(tbH!RunDateTime, "dd/mmm/yyyy hh:mm")
570   Else
580       Rundate = ""
590   End If

600   sql = "SELECT * FROM Differentials WHERE " & _
            "RunNumber = '" & RP.SampleID & "' AND PrnDiff = 1 "
610   Set tbd = New Recordset
620   RecOpenClient 0, tbd, sql
630   If Not tbd.EOF Then
640       DiffFound = True
650       For n = 0 To 14
660           W = Trim(tbd("Wording" & Format(n)) & "")
670           p = IIf(Val(tbd("P" & Format(n)) & "") = 0, "", tbd("P" & Format(n)))
680           a = IIf(Val(tbd("A" & Format(n)) & "") = 0, "", tbd("A" & Format(n)))
690           frmMain.gDiff.AddItem W & vbTab & p & vbTab & a
700       Next
710   End If

720   Select Case Left(UCase(tb!Sex & ""), 1)
          Case "M": Sex = "M"
730       Case "F": Sex = "F"
740       Case Else: Sex = ""
750   End Select


760   ClearUdtHeading

770   With udtHeading
780       .SampleID = RP.SampleID
790       .Dept = "Haematology"
800       If RP.Department = "K" Then .Dept = "Draft Haem"
810       .Name = tb!PatName & ""
820       .Ward = RP.Ward
830       .DoB = DoB
840       .Chart = tb!Chart & ""
850       .Clinician = RP.Clinician
860       .Address0 = tb!Addr0 & ""
870       .Address1 = tb!Addr1 & ""
880       .GP = RP.GP
890       .Sex = tb!Sex & ""
900       .Hospital = tb!Hospital & ""
910       .SampleDate = tb!SampleDate & ""
920       .RecDate = tb!RecDate & ""
930       .Rundate = tbH!Rundate & ""
940       .GpClin = Clin
950       .SampleType = ""
960       .Notes = "*** - DRAFT REPORT. DO NOT RELEASE !!! - ***"
970       .DocumentNo = GetOptionSetting("HaemMainDocumentNo", "")
980       .AandE = tb!AandE & ""
990   End With



1000  If RP.FaxNumber <> "" Then
1010      PrintHeadingRTBFax
1020  Else
1030      PrintHeadingRTB "Page 1 Of " & TotalPages
1040  End If

1050  AddResultToLP udtPrintLine, lp, lpc, "Test", "Result", "Unit", "Ref. Range", "Flag", , , , True, , True
1060  CrCnt = CrCnt + 1

1070  If fbc Then

1080      Flag = InterpH(tbH!WBC & "", "WBC", Sex, DoB, SampleDate)
1090      AddResultToLP udtPrintLine, lp, lpc, "WBC", IIf(Flag = "X", "XXXXX", tbH!WBC & ""), GetUnitsHaem("WBC", "x10^9/l"), HaemNormalRange("WBC", Sex, DoB, SampleDate), Flag, "", "", ""
1100      CrCnt = CrCnt + 1


          'Differentials (Manual and Analyser)
1110      neut = IIf(DiffFound, GetDifferentail("NEUT"), tbH!neuta & "")
1120      If Trim(neut) <> "" Then
1130          Flag = InterpH(neut & "", "NEUTA", Sex, DoB, SampleDate)
1140          AddResultToLP udtPrintLine, lp, lpc, "   Neutrophils", IIf(Flag = "X", "XXXXX", neut), GetUnitsHaem("NEUTA", "x10^9/l"), HaemNormalRange("NEUTA", Sex, DoB, SampleDate), Flag, "", "", ""
1150          CrCnt = CrCnt + 1
1160      End If

1170      lym = IIf(DiffFound, GetDifferentail("LYM"), tbH!lyma & "")
1180      If Trim(lym) <> "" Then
1190          Flag = InterpH(lym & "", "LYMA", Sex, DoB, SampleDate)
1200          AddResultToLP udtPrintLine, lp, lpc, "   Lymphocytes", IIf(Flag = "X", "XXXXX", lym), GetUnitsHaem("LYMA", "x10^9/l"), HaemNormalRange("LYMA", Sex, DoB, SampleDate), Flag, "", "", ""
1210          CrCnt = CrCnt + 1
1220      End If

1230      mono = IIf(DiffFound, GetDifferentail("MON"), tbH!MonoA & "")
1240      If Trim(mono) <> "" Then
1250          Flag = InterpH(mono & "", "MONOA", Sex, DoB, SampleDate)
1260          AddResultToLP udtPrintLine, lp, lpc, "   Monocytes", IIf(Flag = "X", "XXXXX", mono), GetUnitsHaem("MONOA", "x10^9/l"), HaemNormalRange("MONOA", Sex, DoB, SampleDate), Flag, "", "", ""
1270          CrCnt = CrCnt + 1
1280      End If

1290      eos = IIf(DiffFound, GetDifferentail("EOS"), tbH!eosa & "")
1300      If Trim(eos & "") <> "" Then
1310          Flag = InterpH(eos & "", "EOSA", Sex, DoB, SampleDate)
1320          AddResultToLP udtPrintLine, lp, lpc, "   Eosinophils", IIf(Flag = "X", "XXXXX", eos), GetUnitsHaem("EOSA", "x10^9/l"), HaemNormalRange("EOSA", Sex, DoB, SampleDate), Flag, "", "", ""
1330          CrCnt = CrCnt + 1
1340      End If

1350      bas = IIf(DiffFound, GetDifferentail("BAS"), tbH!basa & "")
1360      If Trim(bas & "") <> "" Then
1370          Flag = InterpH(bas & "", "BASA", Sex, DoB, SampleDate)
1380          AddResultToLP udtPrintLine, lp, lpc, "   Basophils", bas, GetUnitsHaem("BASA", "g/dl"), HaemNormalRange("BASA", Sex, DoB, SampleDate), Flag, "", "", ""
1390          CrCnt = CrCnt + 1
1400      End If

1410      If DiffFound Then
1420          luc = Format(frmMain.gDiff.TextMatrix(7, 2), "##0.0##")
1430      Else
1440          luc = tbH!luca & ""
1450      End If
1460      If Trim(luc & "") <> "" Then
1470          If DiffFound = True Then
1480              AddResultToLP udtPrintLine, lp, lpc, "   " & Initial2Upper(frmMain.gDiff.TextMatrix(7, 0)), luc, "", "", "", "", "", ""
1490          Else
1500              Flag = InterpH(luc & "", "LUCA", Sex, DoB, SampleDate)
1510              AddResultToLP udtPrintLine, lp, lpc, "   Lucocytes", luc, GetUnitsHaem("LUCA", "x10^9/l"), IIf(DiffFound, "", HaemNormalRange("LUCA", Sex, DoB, SampleDate)), Flag, "", "", ""
1520          End If
1530          CrCnt = CrCnt + 1
1540      End If

1550      If DiffFound = True Then
1560          dv = ""
1570          If frmMain.gDiff.TextMatrix(8, 1) <> "" Then
1580              dv = Format(frmMain.gDiff.TextMatrix(8, 2), "##0.0##")
1590              AddResultToLP udtPrintLine, lp, lpc, "   " & Initial2Upper(frmMain.gDiff.TextMatrix(8, 0)), dv
1600              CrCnt = CrCnt + 1
1610          End If
1620          dv = ""
1630          If frmMain.gDiff.TextMatrix(9, 1) <> "" Then
1640              dv = Format(frmMain.gDiff.TextMatrix(9, 2), "##0.0##")
1650              AddResultToLP udtPrintLine, lp, lpc, "   " & Initial2Upper(frmMain.gDiff.TextMatrix(9, 0)), dv
1660              CrCnt = CrCnt + 1
1670          End If
1680          dv = ""
1690          If frmMain.gDiff.TextMatrix(10, 1) <> "" Then
1700              dv = Format(frmMain.gDiff.TextMatrix(10, 2), "##0.0##")
1710              AddResultToLP udtPrintLine, lp, lpc, "   " & Initial2Upper(frmMain.gDiff.TextMatrix(10, 0)), dv
1720              CrCnt = CrCnt + 1
1730          End If
1740          dv = ""
1750          If frmMain.gDiff.TextMatrix(11, 1) <> "" Then
1760              dv = Format(frmMain.gDiff.TextMatrix(11, 2), "##0.0##")
1770              AddResultToLP udtPrintLine, lp, lpc, "   " & Initial2Upper(frmMain.gDiff.TextMatrix(11, 0)), dv
1780              CrCnt = CrCnt + 1
1790          End If
1800          dv = ""
1810          If frmMain.gDiff.TextMatrix(12, 1) <> "" Then
1820              dv = Format(frmMain.gDiff.TextMatrix(12, 2), "##0.0##")
1830              AddResultToLP udtPrintLine, lp, lpc, "   " & Initial2Upper(frmMain.gDiff.TextMatrix(12, 0)), dv
1840              CrCnt = CrCnt + 1
1850          End If
1860      End If

1870      Flag = InterpH(tbH!RBC & "", "RBC", Sex, DoB, SampleDate)
1880      AddResultToLP udtPrintLine, lp, lpc, "RBC", IIf(Flag = "X", "XXXXX", tbH!RBC), GetUnitsHaem("RBC", "x10^12/l"), HaemNormalRange("RBC", Sex, DoB, SampleDate), Flag, "", "", ""
1890      CrCnt = CrCnt + 1

1900  End If

1910  If Trim(tbH!Hgb & "") <> "" Then
1920      Flag = InterpH(tbH!Hgb & "", "Hgb", Sex, DoB, SampleDate)
1930      AddResultToLP udtPrintLine, lp, lpc, "Hgb", IIf(Flag = "X", "XXXXX", tbH!Hgb & ""), GetUnitsHaem("Hgb", "g/dl"), HaemNormalRange("Hgb", Sex, DoB, SampleDate), Flag, "", "", ""
1940      CrCnt = CrCnt + 1
1950  End If

1960  If fbc Then
1970      Flag = InterpH(tbH!Hct & "", "Hct", Sex, DoB, SampleDate)
1980      AddResultToLP udtPrintLine, lp, lpc, "Hct", tbH!Hct & "", GetUnitsHaem("Hct", "l/l"), HaemNormalRange("Hct", Sex, DoB, SampleDate), Flag, "", "", ""
1990      CrCnt = CrCnt + 1

2000      Flag = InterpH(tbH!MCV & "", "MCV", Sex, DoB, SampleDate)
2010      AddResultToLP udtPrintLine, lp, lpc, "MCV", tbH!MCV & "", GetUnitsHaem("MCV", "fl"), HaemNormalRange("MCV", Sex, DoB, SampleDate), Flag, "", "", ""
2020      CrCnt = CrCnt + 1

2030      Flag = InterpH(tbH!MCH & "", "MCH", Sex, DoB, SampleDate)
2040      AddResultToLP udtPrintLine, lp, lpc, "MCH", tbH!MCH & "", GetUnitsHaem("MCH", "pg"), HaemNormalRange("MCH", Sex, DoB, SampleDate), Flag, "", "", ""
2050      CrCnt = CrCnt + 1

2060      Flag = InterpH(tbH!MCHC & "", "MCHC", Sex, DoB, SampleDate)
2070      AddResultToLP udtPrintLine, lp, lpc, "MCHC", Trim(tbH!MCHC & ""), GetUnitsHaem("MCHC", "g/dl"), HaemNormalRange("MCHC", Sex, DoB, SampleDate), Flag
2080      CrCnt = CrCnt + 1

2090      Flag = InterpH(tbH!rdwcv & "", "RDW", Sex, DoB, SampleDate)
2100      AddResultToLP udtPrintLine, lp, lpc, "RDW", tbH!rdwcv & "", GetUnitsHaem("RDW", "%"), HaemNormalRange("RDW", Sex, DoB, SampleDate), Flag
2110      CrCnt = CrCnt + 1

2120  End If

2130  If Trim(tbH!Plt) & "" <> "" Then
2140      Flag = InterpH(tbH!Plt & "", "Plt", Sex, DoB, SampleDate)
2150      AddResultToLP udtPrintLine, lp, lpc, "Plt", tbH!Plt & "", GetUnitsHaem("Plt", "x10^9/l"), HaemNormalRange("Plt", Sex, DoB, SampleDate), Flag
2160      CrCnt = CrCnt + 1
2170  End If


      'Extra Tests
2180  If Trim(tbH!reta & "") <> "" Then
2190      Flag = InterpH(tbH!reta & "", "RETA", Sex, DoB, SampleDate)
2200      If Trim(tbH!reta) & "" <> "" And Trim(tbH!reta) <> "?" Then TotalRetics = Val(Trim(tbH!reta & "")) Else Flag = "X"
2210      AddResultToLP udtPrintLine, lp, lpc, "Retics", CStr(TotalRetics), GetUnitsHaem("RETA", "x10^9/l"), HaemNormalRange("RETA", Sex, DoB, SampleDate), Flag
2220      CrCnt = CrCnt + 1
2230  End If

2240  If Trim(tbH!ESR & "") <> "" Then
2250      Flag = InterpH(tbH!ESR & "", "ESR", Sex, DoB, SampleDate)
2260      AddResultToLP udtPrintLine, lp, lpc, "ESR", Trim(tbH!ESR & ""), GetUnitsHaem("ESR", "mm/hr"), HaemNormalRange("ESR", Sex, DoB, SampleDate), Flag
2270      CrCnt = CrCnt + 1
2280  End If

2290  If Trim(tbH!Monospot & "") <> "" And Trim(tbH!Monospot & "") <> "?" Then
2300      If tbH!Monospot = "N" Then luc = "Negative"
2310      If tbH!Monospot = "P" Then luc = "Positive"
2320      If tbH!Monospot = "I" Then luc = "Inconclusive"
2330      AddResultToLP udtPrintLine, lp, lpc, "Infectious Mononucleosis Screen", luc
2340      CrCnt = CrCnt + 1
2350  End If

2360  If Trim(tbH!tra & "") <> "" And Trim(tbH!tra & "") <> "?" Then
2370      AddResultToLP udtPrintLine, lp, lpc, "Rheumatoid Factor", Trim(tbH!tra) & ""
2380      CrCnt = CrCnt + 1
2390  End If

2400  If Trim(tbH!tasot & "") <> "" And Trim(tbH!tasot & "") <> "?" Then
2410      AddResultToLP udtPrintLine, lp, lpc, "Asot", Trim(tbH!tasot) & ""
2420      CrCnt = CrCnt + 1
2430  End If

2440  If Trim(tbH!malaria & "") <> "" And Trim(tbH!malaria & "") <> "?" Then
2450      AddResultToLP udtPrintLine, lp, lpc, "Malaria Antigen Screen", Trim(tbH!malaria) & ""
2460      CrCnt = CrCnt + 1
2470  End If

2480  If Trim(tbH!sickledex & "") <> "" And Trim(tbH!sickledex & "") <> "?" Then
2490      AddResultToLP udtPrintLine, lp, lpc, "Sickle Cell Screen", Trim(tbH!sickledex) & ""
2500      CrCnt = CrCnt + 1
2510  End If

2520  If GetOptionSetting("PRINTNRBC", 0) = 1 Then
2530      If Trim(tbH!nrbcp & "") <> "" Then
2540          AddResultToLP udtPrintLine, lp, lpc, "NRBC", Trim(tbH!nrbcp & "")
2550          CrCnt = CrCnt + 1
2560      End If
2570  End If

      'add blank line before comment

2580  AddResultToLP udtPrintLine, lp, lpc, "", "", "", "", ""

      'Comments

2590  Set OBS = OBS.Load(RP.SampleID, "HAEMATOLOGY", "Demographic")
2600  If Not OBS Is Nothing Then
2610      For Each OB In OBS
2620          Select Case UCase$(OB.Discipline)
                  Case "HAEMATOLOGY"
2630                  AddCommentToLP udtPrintLine, lp, lpc, OB.Comment, ""
2640                  CrCnt = CrCnt + 1
2650              Case "DEMOGRAPHIC"
2660                  AddCommentToLP udtPrintLine, lp, lpc, OB.Comment, ""
2670                  CrCnt = CrCnt + 1
2680          End Select
2690      Next
2700  End If

2710  If DiffFound Then
2720      AddCommentToLP udtPrintLine, lp, lpc, "Manual Differential Reported"
2730      CrCnt = CrCnt + 1
2740  End If

2750  If Not IsDate(tb!DoB) Or Trim(Sex) = "" Then
2760      AddCommentToLP udtPrintLine, lp, lpc, "No Sex/DoB given. No ref range applied "
2770      CrCnt = CrCnt + 1
2780  End If


2790  PrintReport udtPrintLine, lp, lpc, "Haem", PrintA4, SampleDate, Rundate, AuthorisedBy, PrintTime, "", ""
2800  Exit Function

PrintResultHaemAdvia_Error:

      Dim strES      As String
      Dim intEL      As Integer

2810  intEL = Erl
2820  strES = Err.Description

2830  sql = "Delete FROM printpending WHERE SampleID = '" & RP.SampleID & "' and department = '" & RP.Department & "'"
2840  Cnxn(0).Execute sql

2850  LogError "modHaematology", "PrintResultHaem", intEL, strES, sql

End Function


Private Function GetDifferentail(ByVal AnalyteName As String) As String

      Dim n          As Integer
      Dim v As String

10    On Error GoTo GetDifferentail_Error

20    v = ""
30    For n = 1 To frmMain.gDiff.Rows - 1
40        If InStr(UCase(frmMain.gDiff.TextMatrix(n, 0)), AnalyteName) > 0 Then
50            v = Format(frmMain.gDiff.TextMatrix(n, 2), "##0.0##")
60            Exit For
70        End If
80    Next
90    GetDifferentail = v

100   Exit Function

GetDifferentail_Error:

110   LogError "modHaematology", "GetDifferentail", Erl, Err.Description


End Function



Private Sub PrintUnits(ByVal Units As String, _
                       ByVal SmallFontSize As Single, _
                       ByVal NormalFontSize As Single)


10    With frmRichText.rtb

20        Select Case Units
              Case "10^9/l"
30                .SelText = " x10"
40                .SelFontSize = SmallFontSize
50                .SelCharOffset = 40
60                .SelText = "9 "
70                .SelCharOffset = 0
80                .SelFontName = "Courier New"
90                .SelFontSize = NormalFontSize
100               .SelText = "/l  "
110           Case "10^12/l"
120               .SelText = " x10"
130               .SelFontSize = SmallFontSize
140               .SelCharOffset = 40
150               .SelText = "12"
160               .SelCharOffset = 0
170               .SelFontName = "Courier New"
180               .SelFontSize = NormalFontSize
190               .SelText = "/l  "
200       End Select
210   End With

End Sub



Private Sub PrintTest(TestName As String, Result As String, Unit As String, Flag As String, _
                      RefRange As String, Optional LineFeed As Boolean = False, Optional PrintFont As Integer = 10, _
                      Optional QualativeResult As Boolean = False, Optional TestNameBold As Boolean = False, _
                      Optional ResultBold As Boolean = False, Optional UnitBold As Boolean = False, _
                      Optional FlagBold As Boolean = False, Optional RefRangeBold As Boolean = False)

          Dim UnitP1 As String
          Dim UnitP2 As String
          Dim UnitPart As String
          Dim ChrCnt As Integer
          Dim TestNameLength As Integer
          Dim TestResultLength As Integer


10        If PrintFont = 8 Then
20            If QualativeResult Then
30                TestNameLength = 31
40                TestResultLength = 15
50            Else
60                TestNameLength = 6
70                TestResultLength = 6
80            End If
90        ElseIf PrintFont = 10 Then
100           If QualativeResult Then
110               TestNameLength = 22
120               TestResultLength = 15
130           Else
140               TestNameLength = 6
150               TestResultLength = 6
160           End If
170       End If


180       TestName = FormatString(TestName, TestNameLength, , IIf(LineFeed, AlignLeft, AlignCenter))
190       PrintTextRTB frmRichText.rtb, TestName, PrintFont, TestNameBold
200       PrintTextRTB frmRichText.rtb, " ", PrintFont

210       Result = FormatString(Trim(Result), TestResultLength, , AlignRight)
220       PrintTextRTB frmRichText.rtb, Result, PrintFont, ResultBold
230       PrintTextRTB frmRichText.rtb, " ", PrintFont

240       If QualativeResult = False Then
250           PrintTextRTB frmRichText.rtb, " ", PrintFont
260           Unit = Trim(FormatString(Unit, 8))
270           If InStr(1, Unit, "^") > 0 And InStr(1, Unit, "/") > 0 And (InStr(1, Unit, "/") - InStr(1, Unit, "^")) > 0 Then
280               ChrCnt = 0
290               UnitPart = Left$(Unit, InStr(1, Unit, "^") - 1)
300               ChrCnt = ChrCnt + Len(UnitPart)
310               PrintTextRTB frmRichText.rtb, UnitPart, PrintFont, UnitBold

320               UnitPart = FormatString(Mid$(Unit, InStr(1, Unit, "^") + 1, InStr(1, Unit, "/") - InStr(1, Unit, "^") - 1), 2, , AlignLeft)
330               ChrCnt = ChrCnt + Len(UnitPart)
340               PrintTextRTB frmRichText.rtb, UnitPart, PrintFont - 3, UnitBold, , , , True
350               PrintTextRTB frmRichText.rtb, " ", 7

360               UnitPart = Mid$(Unit, InStr(1, Unit, "/"), Len(Unit))
370               ChrCnt = ChrCnt + Len(UnitPart)
380               PrintTextRTB frmRichText.rtb, UnitPart, PrintFont, UnitBold

390               If ChrCnt < 8 Then
400                   PrintTextRTB frmRichText.rtb, String(8 - ChrCnt, " "), 10
410               End If
420           Else
430               PrintTextRTB frmRichText.rtb, FormatString(Unit, 8), PrintFont, UnitBold
440           End If
450           PrintTextRTB frmRichText.rtb, " ", PrintFont

460           Flag = FormatString(Flag, 1, , AlignCenter)
470           PrintTextRTB frmRichText.rtb, Flag, PrintFont, FlagBold
480           PrintTextRTB frmRichText.rtb, " ", PrintFont

490           RefRange = FormatString(RefRange, 14, , AlignLeft)
500           PrintTextRTB frmRichText.rtb, RefRange, PrintFont, RefRangeBold

510       End If

520       If LineFeed = True Then
530           PrintTextRTB frmRichText.rtb, vbCrLf
540       Else
550           PrintTextRTB frmRichText.rtb, String(4, " "), PrintFont
560       End If

End Sub

Private Function GetUnitsHaem(ByVal Analyte As String, ByVal CurrentUnit As String) As String

      Dim sql        As String
      Dim tb         As Recordset
      Dim s          As String

10    On Error GoTo GetUnitsHaem_Error

20    sql = "Select distinct AnalyteName,Units from haemtestdefinitions where analytename = '" & Analyte & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    If tb.EOF Then
60        GetUnitsHaem = CurrentUnit
70    Else
80        GetUnitsHaem = tb!Units & ""
90    End If

100   Exit Function

GetUnitsHaem_Error:

110   MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetUnitsHaem of Form frmViewResults"

End Function


