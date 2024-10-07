Attribute VB_Name = "modMicro"
Option Explicit

Public Function PrintUrine()

      Dim tb As Recordset
      Dim tu As Recordset
      Dim ds As Recordset
10    ReDim sn(0 To 1) As Recordset
      Dim sql As String
20    ReDim organism(0 To 4) As String
30    ReDim Entries(0 To 1) As Integer
      Dim so As Recordset
      Dim maxentries As Integer
      Dim n As Integer
      Dim ABDetails(1 To 18, 0 To 3) As String
      Dim Found As Boolean
      Dim Suppressed(0 To 2) As Boolean
      Dim MicroPresent As Boolean
      Dim BioPresent As Boolean
      Dim MiscPresent As Boolean
      Dim CulturesPresent As Integer
      Dim ABsPresent As Boolean
      Dim Dept As String
      Dim Rundate As String
      Dim SampleDate As String
      'Dim Cx As Comment
      'Dim Cxs As New Comments
      Dim OB As Observation
      Dim OBS As New Observations
40    ReDim Comments(1 To 4) As String
      Dim lngU As Long
      Dim rs As Recordset
      Dim xT As Long
      Dim CT As Long
      Dim Fontz1 As Integer
      Dim Fontz2 As Integer
      Dim Fontz3 As Integer
      Dim Fontz4 As Integer
      Dim xFound As Boolean
      Dim DoB As String
      Dim Clin As String
      Dim f As Integer
      Dim Fx As Integer
      Dim PrintTime As String

50    On Error GoTo PrintUrine_Error

60    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

70    xT = 0

80    If InStr(UCase(pForcePrintTo), "FAX") > 0 Then
90      xT = 20
100     sql = "SELECT * FROM options WHERE description = 'FX'"
110     Set tb = New Recordset
120     RecOpenServer 0, tb, sql
130     If tb.EOF Then
140       Fx = 0
150     Else
160       Fx = Val(tb!Contents)
170     End If
180   End If


190   ABsPresent = False

200   sql = "SELECT * FROM micrositedetails WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
210   Set tu = New Recordset
220   RecOpenServer 0, tu, sql

230   If Not tu.EOF Then
240     Dept = Trim(tu!SiteDetails & "")
250   End If

260   sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
270   Set tb = New Recordset
280   RecOpenServer 0, tb, sql

290   DoB = Format(tb!DoB, "dd/MMM/yyyy")

300   ClearUdtHeading
310   With udtHeading
320     .SampleID = RP.SampleID - SysOptMicroOffset(0)
330     .Dept = "Microbiology"
340     .Name = tb!PatName & ""
350     .Ward = RP.Ward
360     .DoB = DoB
370     .Chart = tb!Chart & ""
380     .Clinician = RP.Clinician
390     .Address0 = tb!Addr0 & ""
400     .Address1 = tb!Addr1 & ""
410     .GP = RP.GP
420     .Sex = tb!Sex & ""
430     .Hospital = tb!Hospital & ""
440     .SampleDate = tb!SampleDate & ""
450     .RecDate = tb!RecDate & ""
460     .Rundate = tb!Rundate & ""
470     .GpClin = Clin
480     .SampleType = ""
490   End With
500   If RP.FaxNumber <> "" Then
510     PrintHeadingRTBFax
520   Else
530     PrintHeadingRTB
540   End If

550   If RP.FaxNumber <> "" Then
560     Fontz1 = 9
570     Fontz2 = 2
580     Fontz3 = 8
590     Fontz4 = 12
600   Else
610     Fontz1 = 10
620     Fontz2 = 2
630     Fontz3 = 9
640     Fontz4 = 14
650   End If

660   With frmRichText.rtb
670     .SelFontSize = Fontz1
  
680     sql = "SELECT coalesce(valid,0), * FROM Urine WHERE " & _
              "SampleID = '" & RP.SampleID & "'"
690     Set tu = New Recordset
700     RecOpenServer 0, tu, sql
  
710     If tu.EOF Then
720       .SelText = vbCrLf
730       CrCnt = CrCnt + 1
740     Else
750       If Trim$(tu!UserName & "") <> "" Then
760         RP.Initiator = tu!UserName
770       End If

780     .SelFontSize = Fontz1
790     If tu!Valid = 0 Then
800       .SelText = Left$(" " & Space(40), 40) & "DRAFT REPORT" & vbCrLf
810       CrCnt = CrCnt + 1
820     End If
830       MicroPresent = IsMicroPresent(tu)
840       BioPresent = IsBioPresent(tu)
850       MiscPresent = IsMiscPresent(tu)
860       sql = "SELECT * FROM micrositedetails WHERE SampleID = " & RP.SampleID & ""
870       Set rs = New Recordset
880       RecOpenServer 0, rs, sql
890       If Not rs.EOF Then
900         .SelBold = True
910         .SelText = "  SAMPLE TYPE : " & UCase(Trim(rs!SiteDetails & "")) & vbCrLf
920         CrCnt = CrCnt + 1
930         .SelBold = False
940       End If
950       If MicroPresent Or BioPresent Then
960         If MicroPresent Then
970           .SelText = "    "
980           .SelUnderline = True
990           .SelText = "MICROSCOPY"
1000          .SelUnderline = False
1010          .SelText = Left(" " & Space(37), 37)
1020        Else
1030          .SelText = Left(" " & Space(47), 47)
1040        End If
1050        If BioPresent Then
1060          .SelUnderline = True
1070          .SelText = "BIOCHEMISTRY" & vbCrLf
1080          .SelUnderline = False
1090        Else
1100          .SelText = vbCrLf
1110        End If
1120        CrCnt = CrCnt + 1

1130        If MicroPresent Then
1140          .SelText = Left("    Leucocytes: " & Trim(tu!WCC) & "" & " /cmm" & Space(51), 51)
1150        Else
1160          .SelText = Left(" " & Space(51), 51)
1170        End If
1180        If BioPresent Then
1190          .SelText = "          pH: "
1200          If tu!pH & "" = "Unsuitable" Then
1210            .SelBold = True
1220            .SelText = "  ( Sample unsuitable for )" & vbCrLf
1230            CrCnt = CrCnt + 1
1240            .SelBold = False
1250          Else
1260            .SelText = tu!pH & "" & vbCrLf
1270            CrCnt = CrCnt + 1
1280          End If
1290        Else
1300          .SelText = vbCrLf
1310          CrCnt = CrCnt + 1
1320        End If
  
1330        If MicroPresent Then
1340          .SelText = Left("  Erythrocytes: " & Trim(tu!RCC) & "" & " /cmm" & Space(47), 47)
1350        Else
1360          .SelText = Left(" " & Space(51), 51)
1370        End If
1380        If BioPresent Then
1390          .SelText = "     Protein: "
1400          If tu!pH & "" = "Unsuitable" Then
1410            .SelBold = True
1420            .SelText = "  ( Biochemistry Analysis.)" & vbCrLf
1430            .SelText = vbCrLf
1440            .SelBold = False
1450            CrCnt = CrCnt + 1
1460          Else
1470            .SelText = Trim(tu!Protein & "") & " " & "mg/dL" & vbCrLf
1480            CrCnt = CrCnt + 1
1490          End If
1500        Else
1510          .SelText = vbCrLf
1520          CrCnt = CrCnt + 1
1530        End If
  
1540        If MicroPresent Then
1550          .SelText = Left("         Casts: " & Trim(tu!Casts & "") & Space(51), 51)
1560        Else
1570          .SelText = Left(" " & Space(51), 51)
1580        End If
1590        If BioPresent Then
1600          .SelText = "     Glucose: " & Trim(tu!Glucose & "") & " mmol/l"
1610        End If
1620        .SelText = vbCrLf
1630        CrCnt = CrCnt + 1
  
1640        If MicroPresent Then
1650          .SelText = Left("      Crystals: " & Trim(tu!Crystals & "") & Space(51), 51)
1660        Else
1670          .SelText = Left(" " & Space(51), 51)
1680        End If
1690        If BioPresent Then
1700          .SelText = "     Ketones: " & Trim(tu!ketones & "") & vbCrLf
1710          CrCnt = CrCnt + 1
1720        Else
1730          .SelText = vbCrLf
1740          CrCnt = CrCnt + 1
1750        End If
  
1760        If MiscPresent Then
1770          .SelText = Left(" Miscellaneous: " & Trim(tu!Misc0 & "") & Space(51), 51)
1780        Else
1790          .SelText = Left(" " & Space(51), 51)
1800        End If
  
1810        If BioPresent Then
1820          .SelText = "Urobilinogen: " & Trim(tu!urobilinogen & "")
1830        End If
1840        .SelText = vbCrLf
1850        CrCnt = CrCnt + 1

1860        If MiscPresent Then
1870          .SelText = Left("              : " & Trim(tu!Misc1 & "") & Space(51), 51)
1880        Else
1890          .SelText = Left(" " & Space(51), 51)
1900        End If
  
1910        If BioPresent Then
1920          .SelText = "   Bilirubin: " & Trim(tu!bilirubin & "")
1930        End If
1940        .SelText = vbCrLf
1950        CrCnt = CrCnt + 1
  
1960        If MiscPresent Then
1970          .SelText = Left("              : " & Trim(tu!Misc2 & "") & Space(51), 51)
1980        Else
1990          .SelText = Left(" " & Space(51), 51)
2000        End If

2010        If BioPresent Then
2020          .SelText = "    Blood Hb: " & Trim(tu!BloodHb & "")
2030        End If
2040        .SelText = vbCrLf
2050        CrCnt = CrCnt + 1
  
2060      Else
2070        For n = 1 To 8
2080          .SelText = vbCrLf
2090          CrCnt = CrCnt + 1
2100        Next
2110      End If

2120      .SelFontSize = Fontz2
2130      If RP.FaxNumber <> "" Then .SelText = String(280, "-") Else .SelText = String(420, "-")
2140      .SelFontSize = Fontz3
2150      .SelText = vbCrLf
2160      CrCnt = CrCnt + 1

2170      CulturesPresent = 0

2180      sql = "SELECT * FROM isolates WHERE SampleID = " & RP.SampleID & " order by isolatenumber"
2190      Set ds = New Recordset
2200      RecOpenServer 0, ds, sql
2210      Do While Not ds.EOF
2220        CulturesPresent = ds!IsolateNumber
2230        organism(ds!IsolateNumber) = ds!OrganismName
2240        If Len(ds!OrganismName) > 32 Then CT = CT + 1 + (Len(ds!OrganismName) - 32)
2250        ds.MoveNext
2260      Loop

          '  Suppressed(0) = tu!SuppressSensitivities0
          '  Suppressed(1) = tu!SuppressSensitivities1
          '  Suppressed(2) = tu!SuppressSensitivities2
  
2270      If CulturesPresent > 0 Or Trim(tu!Count & "") <> "" Then
2280        .SelText = " "
2290        .SelBold = True
2300        .SelUnderline = True
2310        .SelText = "Colony count:"
2320        .SelUnderline = False
2330        sql = "SELECT * FROM sensitivities WHERE " & _
                  "SampleID = '" & RP.SampleID & "'  order by antibioticcode"
2340        Set so = New Recordset
2350        RecOpenServer 0, so, sql
2360        Do While Not so.EOF
2370          ABsPresent = True
2380          Found = False
2390          For n = 1 To 18
2400            If Left(Trim(AntiName(Right(Space(16) & so!AntibioticCode, 16))), 16) = Left(Trim(ABDetails(n, 0)), 16) Then
2410              If so!IsolateNumber = 1 Then
2420                If so!Report = True Then
2430                  ABDetails(n, 1) = so!RSI & ""
2440                Else
2450                  ABDetails(n, 1) = "   "
2460                End If
2470              ElseIf so!IsolateNumber = 2 Then
2480                If so!Report = True Then
2490                  ABDetails(n, 2) = so!RSI & ""
2500                Else
2510                  ABDetails(n, 2) = "   "
2520                End If
2530              ElseIf so!IsolateNumber = 3 Then
2540                If so!Report = True Then
2550                  ABDetails(n, 3) = so!RSI & ""
2560                Else
2570                  ABDetails(n, 3) = "   "
2580                End If
2590              End If
2600              Found = True
2610            End If
2620          Next
2630          If Not Found Then
2640            For n = 1 To 18
2650              If Trim(ABDetails(n, 0)) = "" Then
2660                ABDetails(n, 0) = Left(AntiName(Right(Space(16) & so!AntibioticCode, 16)), 16)
2670                If so!IsolateNumber = 1 Then
2680                  If so!Report = True Then
2690                    ABDetails(n, 1) = so!RSI & ""
2700                  Else
2710                    ABDetails(n, 1) = "   "
2720                  End If
2730                ElseIf so!IsolateNumber = 2 Then
2740                  If so!Report = True Then
2750                    ABDetails(n, 2) = so!RSI & ""
2760                  Else
2770                    ABDetails(n, 2) = "   "
2780                  End If
2790                ElseIf so!IsolateNumber = 3 Then
2800                  If so!Report = True Then
2810                    ABDetails(n, 3) = so!RSI & ""
2820                  Else
2830                    ABDetails(n, 3) = "   "
2840                  End If
2850                End If
2860                Exit For
2870              End If
2880            Next
2890          End If
2900          so.MoveNext
2910        Loop
2920      Else
2930        For n = 1 To 18
2940          ABDetails(n, 0) = " "
2950          ABDetails(n, 1) = " "
2960          ABDetails(n, 2) = " "
2970          ABDetails(n, 3) = " "
2980        Next
2990      End If
  
3000      If CulturesPresent = 0 Then
3010        .SelFontSize = Fontz3
3020        .SelText = vbCrLf
3030        CrCnt = CrCnt + 1
3040      Else
3050        If ABsPresent Then
3060          If CulturesPresent = 1 Then
3070            .SelFontSize = Fontz3
3080            .SelText = Left(" " & Space(33 - Fx), 33 - Fx)
3090            .SelUnderline = True
3100            .SelText = "Culture           1      "
3110            .SelUnderline = False
3120          ElseIf CulturesPresent = 2 Then
3130            .SelFontSize = Fontz3
3140            .SelText = Left(" " & Space(33 - Fx), 33 - Fx)
3150            .SelUnderline = True
3160            .SelText = "Culture           1  2   "
3170            .SelUnderline = False
3180          Else
3190            .SelFontSize = Fontz3
3200            .SelText = Left(" " & Space(33 - Fx), 33 - Fx)
3210            .SelUnderline = True
3220            .SelText = "Culture           1  2  3"
3230            .SelUnderline = False
3240          End If
3250        End If
3260        Found = False
3270        For n = 10 To 18
3280          If Trim(ABDetails(n, 0)) <> "" Then
3290            Found = True
3300            Exit For
3310          End If
3320        Next
3330        If Found Then
3340          .SelFontSize = Fontz3
3350          .SelText = "  "
3360          If CulturesPresent = 1 Then
3370            .SelFontSize = Fontz3
3380            .SelUnderline = True
3390            .SelText = "Culture           1"
3400            .SelUnderline = False
3410          ElseIf CulturesPresent = 2 Then
3420            .SelFontSize = Fontz3
3430            .SelUnderline = True
3440            .SelText = "Culture           1  2"
3450            .SelUnderline = False
3460          Else
3470            .SelFontSize = Fontz3
3480            .SelUnderline = True
3490            .SelText = "Culture           1  2  3"
3500            .SelUnderline = False
3510          End If
3520        End If
3530        .SelText = vbCrLf
3540        CrCnt = CrCnt + 1
3550      End If
  
3560      .SelBold = False
3570      .SelFontSize = Fontz3
  
3580      Found = False
3590      For n = 1 To 18
3600        If Trim(ABDetails(n, 0)) <> "" Then
3610          Found = True
3620          Exit For
3630        End If
3640      Next
3650      If Not Found Then
3660        .SelText = vbCrLf
3670        CrCnt = CrCnt + 1
3680      Else
3690        .SelFontSize = Fontz3
3700        .SelText = Left(" " & Space(47 - Fx), 47 - Fx)
3710        PrintABD 1, ABDetails()
3720      End If
  
3730      .SelFontSize = Fontz3
3740      .SelText = Left("   " & tu!Count & "" & Space(12), 12)
3750      .SelBold = False
3760      If UCase(tu!Count & "") <> "NORMAL" And Trim(tu!Count & "") <> "" Then
3770        .SelText = Left(" Cfu's per mL" & Space(37 - Fx), 35 - Fx)
3780      Else
3790        .SelText = Left(" " & Space(37 - Fx), 35 - Fx)
3800      End If
  
3810      If CulturesPresent > 0 Then
3820        PrintABD 2, ABDetails()
3830      End If
    
3840      If CulturesPresent > 0 Then
3850        .SelFontSize = Fontz3
3860        .SelText = Left(" " & Space(47 - Fx), 47 - Fx)
3870        PrintABD 3, ABDetails()
3880        .SelText = " "
3890        .SelBold = True
3900        .SelUnderline = True
3910        .SelText = "Culture:"
3920        .SelBold = False
3930        .SelUnderline = False
3940        .SelText = Left(" " & Space(38 - Fx), 38 - Fx)
3950        PrintABD 4, ABDetails()

3960        If InStr(organism(0) & "", "?") Then
3970          .SelFontSize = Fontz3
3980            .SelText = Left(" Query Significance." & Space(47 - Fx), 47 - Fx)
3990          Else
4000            .SelFontSize = Fontz3
4010            .SelText = Left(" " & Space(47 - Fx), 47 - Fx)
4020          End If
4030          PrintABD 5, ABDetails()
4040          If CulturesPresent > 0 Then
4050            .SelFontSize = Fontz3
4060            .SelText = Left(" 1:" & organism(1) & Space(47 - Fx), 47 - Fx)
4070          End If
4080          PrintABD 6, ABDetails()
4090          sql = "SELECT * FROM sensitivities WHERE SampleID = " & RP.SampleID & " and isolatenumber = 1"
4100          Set rs = New Recordset
4110          RecOpenServer 0, rs, sql
4120          If Not rs.EOF Then
4130            sql = "SELECT count(report) as tot FROM sensitivities WHERE " & _
                      "SampleID = " & RP.SampleID & " and isolatenumber = 1 and report = 1"
4140            Set rs = New Recordset
4150            RecOpenServer 0, rs, sql
4160            If Not tb.EOF Then
4170              If rs!Tot = 0 Then
4180                .SelFontSize = Fontz3
4190                .SelText = Left(" " & Space(47 - Fx), 47 - Fx)
4200              Else
4210                .SelFontSize = Fontz3
4220                .SelText = Left(" " & Space(47 - Fx), 47 - Fx)
4230              End If
4240            Else
4250              .SelText = Left(" " & Space(47 - Fx), 47 - Fx)
4260            End If
4270          Else
4280            .SelFontSize = Fontz3
4290            .SelText = Left(" " & Space(47 - Fx), 47 - Fx)
4300          End If
4310          PrintABD 7, ABDetails()
4320          If CulturesPresent > 1 Then
4330            .SelFontSize = Fontz3
4340            .SelText = Left(" 2:" & organism(2) & Space(47 - Fx), 47 - Fx)
4350          Else
4360            .SelFontSize = Fontz3
4370            .SelText = Left(" " & Space(47 - Fx), 47 - Fx)
4380          End If
4390          PrintABD 8, ABDetails()
4400          If CulturesPresent > 1 Then
4410            sql = "SELECT count(report) as tot FROM sensitivities WHERE " & _
                      "SampleID = " & RP.SampleID & " and isolatenumber = 2 and report = 1"
4420            Set rs = New Recordset
4430            RecOpenServer 0, rs, sql
4440            If Not tb.EOF Then
4450              If rs!Tot = 0 Then
4460                .SelFontSize = Fontz3
4470                .SelText = Left(" " & Space(47 - Fx), 47 - Fx)
4480              Else
4490                .SelFontSize = Fontz3
4500                .SelText = Left(" " & Space(47 - Fx), 47 - Fx)
4510              End If
4520            Else
4530              .SelFontSize = Fontz3
4540              .SelText = Left(" " & Space(47 - Fx), 47 - Fx)
4550            End If
4560          Else
4570            .SelFontSize = Fontz3
4580            .SelText = Left(" " & Space(47 - Fx), 47 - Fx)
4590          End If
4600          PrintABD 9, ABDetails()
4610          If CulturesPresent > 2 Then
4620            .SelText = Left(" 3:" & organism(3) & Space(47 - Fx), 47 - Fx)
4630          Else
4640            .SelText = Left(" " & Space(47 - Fx), 47 - Fx)
4650          End If
4660        End If
4670      End If
    
4680    Do While CrCnt < 29
4690      .SelText = vbCrLf
4700      CrCnt = CrCnt + 1
4710    Loop
    
4720    .SelFontSize = Fontz2
4730    .SelText = vbCrLf
4740    If RP.FaxNumber <> "" Then .SelText = String(280, "-") Else .SelText = String(420, "-")
4750    .SelFontSize = Fontz1
4760    .SelText = vbCrLf
4770    CrCnt = CrCnt + 1

4780    If Trim(tu!Pregnancy & "") <> "" Then
4790      .SelBold = True
4800      If Left(tu!Pregnancy, 1) = "I" Then
4810        .SelText = "  Pregnancy Test: Inconclusive - Please repeat with early morning sample in 3 - 5 days" & vbCrLf
4820        CrCnt = CrCnt + 1
4830      Else
4840        .SelFontSize = Fontz4
4850        If tu!Pregnancy = "N" Then
4860          .SelText = "  Pregnancy Test: Negative" & vbCrLf
4870          CrCnt = CrCnt + 1
4880        ElseIf tu!Pregnancy = "P" Then
4890          .SelText = "  Pregnancy Test: Positive" & vbCrLf
4900          CrCnt = CrCnt + 1
4910        Else
4920          .SelText = "  Pregnancy Test: " & tu!Pregnancy & vbCrLf
4930          CrCnt = CrCnt + 1
4940        End If
4950        .SelFontSize = Fontz1
4960      End If
4970      .SelBold = False
4980    End If
4990    If Trim(tu!HCGLevel & "") <> "" Then
5000      .SelText = "   "
5010      .SelBold = True
5020      .SelText = "hCG:"
5030      .SelBold = False
5040      .SelText = Trim(tu!HCGLevel) & " mIU/mL"
5050    End If
5060    If Trim(tu!BenceJones & "") <> "" Then
5070      .SelText = " "
5080      .SelBold = True
5090      .SelText = "Bence Jones Protein:"
5100      .SelBold = False
5110      .SelText = Trim(tu!BenceJones)
5120    End If
5130    If Trim(tu!FatGlobules & "") <> "" Then
5140      .SelText = "   "
5150      .SelBold = True
5160      .SelText = "Fat Globules:"
5170      .SelBold = False
5180      .SelText = Trim(tu!FatGlobules)
5190    End If
5200    If Trim(tu!SG & "") <> "" Then
5210      .SelText = "   "
5220      .SelBold = True
5230      .SelText = "SG:"
5240      .SelBold = False
5250      .SelText = Trim(tu!SG)
5260    End If
5270    .SelText = vbCrLf
5280    CrCnt = CrCnt + 1
  
5290    If tu!Valid = 1 Then
5300      sql = "UPDATE Urine " & _
                "SET Printed = 1 WHERE " & _
                "SampleID = '" & RP.SampleID & "'"
5310      Cnxn(0).Execute sql
5320    End If
  
5330    Do While CrCnt < 31
5340      .SelText = vbCrLf
5350      CrCnt = CrCnt + 1
5360    Loop
    
      '  Set Cx = Cxs.Load(RP.SampleID)
5370    Set OBS = OBS.Load(RP.SampleID, "MicroCS", "Demographic", "MicroGeneral", "MicroConsultant")
5380    If Not OBS Is Nothing Then
5390      .SelBold = True
5400      .SelText = "Comment : "
5410      .SelBold = False
5420      For Each OB In OBS
5430          Select Case UCase$(OB.Discipline)
                  Case "DEMOGRAPHIC"
5440                  FillCommentLines OB.Comment, 2, Comments(), 87
5450                  For n = 1 To 4
5460                    If Trim(Comments(n)) <> "" Then .SelText = "     " & Comments(n) & vbCrLf
5470                    CrCnt = CrCnt + 1
5480                  Next
5490              Case Else
5500                  FillCommentLines OB.Comment, 4, Comments(), 87
5510                  For n = 1 To 4
5520                    If Trim(Comments(n)) <> "" Then .SelText = "     " & Comments(n) & vbCrLf
5530                    CrCnt = CrCnt + 1
5540                  Next
5550          End Select
5560      Next
5570    End If
    
5580    If IsDate(tb!SampleDate) Then
5590      SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
5600    Else
5610      SampleDate = ""
5620    End If
5630    If IsDate(tb!Rundate) Then
5640      Rundate = Format(tb!Rundate, "dd/mmm/yyyy")
5650    Else
5660      Rundate = ""
5670    End If

5680    If RP.FaxNumber <> "" Then
5690      PrintFooterRTBFax RP.Initiator, SampleDate, Rundate
5700      f = FreeFile
5710      Open SysOptFax(0) & RP.SampleID & "URN.doc" For Output As f
5720      .SelStart = 0
5730      Print #f, .TextRTF
5740      Close f
5750      SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "URN.doc"
5760    Else
5770      PrintFooterRTB RP.Initiator, SampleDate, Rundate
5780      .SelStart = 0
5790      .SelPrint Printer.hDC
5800    End If
5810    sql = "SELECT * FROM Reports WHERE 0 = 1"
5820    Set tb = New Recordset
5830    RecOpenServer 0, tb, sql
5840    tb.AddNew
5850    tb!SampleID = RP.SampleID
5860    tb!Name = udtHeading.Name
5870    tb!Dept = "U"
5880    tb!Initiator = RP.Initiator
5890    tb!PrintTime = PrintTime
5900    tb!RepNo = "0U" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
5910    tb!PageNumber = 0
5920    tb!Report = .TextRTF
5930    tb!Printer = Printer.DeviceName
5940    tb.Update
5950  End With

5960  Exit Function

PrintUrine_Error:

      Dim strES As String
      Dim intEL As Integer

5970  intEL = Erl
5980  strES = Err.Description
5990  LogError "modMicro", "PrintUrine", intEL, strES, sql

End Function


Public Function IsBioPresent(ByVal tu As Recordset) As Boolean
      'Present
10    On Error GoTo IsBioPresent_Error

20    If (Trim(tu!pH & "") <> "") Or _
         (Trim(tu!Protein & "") <> "") Or _
         (Trim(tu!Glucose & "") <> "") Or _
         (Trim(tu!ketones & "") <> "") Or _
         (Trim(tu!urobilinogen & "") <> "") Or _
         (Trim(tu!bilirubin & "") <> "") Or _
         (Trim(tu!BloodHb & "") <> "") Then
   
30      IsBioPresent = True

40    Else
50      IsBioPresent = False
60    End If

70    Exit Function

IsBioPresent_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "modMicro", "IsBioPresent", intEL, strES

End Function


Public Function IsMicroPresent(ByVal tu As Recordset) As Boolean
        'Present
10    On Error GoTo IsMicroPresent_Error

20    If (Trim(tu!WCC & "") <> "") Or _
         (Trim(tu!RCC & "") <> "") Or _
         (Trim(tu!Casts & "") <> "") Or _
         (Trim(tu!Crystals & "") <> "") Or _
         (Trim(tu!Misc0 & "") <> "") Or _
         (Trim(tu!Misc1 & "") <> "") Or _
         (Trim(tu!Misc2 & "") <> "") Then
  
30      IsMicroPresent = True

40    Else
50      IsMicroPresent = False
60    End If

70    Exit Function

IsMicroPresent_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "modMicro", "IsMicroPresent", intEL, strES

End Function


Public Function IsMiscPresent(ByVal tu As Recordset) As Boolean

10    On Error GoTo IsMiscPresent_Error

20    If (Trim(tu!Misc0 & "") <> "") Or _
         (Trim(tu!Misc1 & "") <> "") Or _
         (Trim(tu!Misc2 & "") <> "") Then
   
30      IsMiscPresent = True

40    Else
50      IsMiscPresent = False
60    End If

70    Exit Function

IsMiscPresent_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "modMicro", "IsMiscPresent", intEL, strES

End Function

Public Sub PrintABD(ByVal Start As Integer, _
                    ByRef ABDetails() As String)
  
10    On Error GoTo PrintABD_Error

20    With frmRichText.rtb
30      If RP.FaxNumber <> "" Then
40        .SelFontSize = 9
50      End If
60      .SelText = Left(ABDetails(Start, 0) & Space(18), 18)
70      .SelText = Left(ABDetails(Start, 1) & Space(3), 3)
80      .SelText = Left(ABDetails(Start, 2) & Space(3), 3)
90      .SelText = Left(ABDetails(Start, 3) & Space(3), 3)
100     .SelText = "" & Left(ABDetails(Start + 9, 0) & Space(18), 18)
110     .SelText = Left(ABDetails(Start + 9, 1) & Space(3), 3)
120     .SelText = Left(ABDetails(Start + 9, 2) & Space(3), 3)
130     .SelText = Left(ABDetails(Start + 9, 3) & Space(3), 3)
140     .SelText = vbCrLf
150   End With

160   CrCnt = CrCnt + 1

170   Exit Sub

PrintABD_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "modMicro", "PrintABD", intEL, strES

End Sub

Private Function AntiName(ByVal Name As String) As String

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo AntiName_Error

20    sql = "SELECT * FROM antibiotics WHERE code = '" & Trim(Name) & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If Not tb.EOF Then
60        AntiName = tb!AntibioticName
70    Else
80        AntiName = Name
90    End If

100   Exit Function

AntiName_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modMicro", "AntiName", intEL, strES, sql

End Function



Public Function PrintFaeces()

      Dim sql As String
      Dim DoB As String
      Dim tb As Recordset
      Dim tbO As Recordset
      Dim tbR As Recordset
      Dim tbsalm As Recordset
      Dim n As Integer
      'Dim Cx As Comment
      'Dim Cxs As New Comments
      Dim OB As Observation
      Dim OBS As New Observations
10    ReDim Comments(1 To 4) As String
      Dim CulFound As Boolean
      Dim SampleDate  As String
      Dim Rundate  As String
      Dim RunTime  As String
      Dim PrintTime As String

20    On Error GoTo PrintFaeces_Error

30    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")


40    sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql

70    DoB = Format(tb!DoB, "dd/MMM/yyyy")
80    ClearUdtHeading
90    With udtHeading
100     .SampleID = RP.SampleID - SysOptMicroOffset(0)
110     .Dept = "Microbiology"
120     .Name = tb!PatName & ""
130     .Ward = RP.Ward
140     .DoB = DoB
150     .Chart = tb!Chart & ""
160     .Clinician = RP.Clinician
170     .Address0 = tb!Addr0 & ""
180     .Address1 = tb!Addr1 & ""
190     .GP = RP.GP
200     .Sex = tb!Sex & ""
210     .Hospital = tb!Hospital & ""
220     .SampleDate = tb!SampleDate & ""
230     .RecDate = tb!RecDate & ""
240     .Rundate = tb!Rundate & ""
250     .GpClin = ""
260     .SampleType = ""
270   End With

280   PrintHeadingRTB

290   sql = "SELECT * FROM microrequests WHERE SampleID = " & RP.SampleID & ""
300   Set tbR = New Recordset
310   RecOpenServer 0, tbR, sql

320   sql = "SELECT * FROM faeces WHERE SampleID = " & RP.SampleID & ""
330   Set tbO = New Recordset
340   RecOpenServer 0, tbO, sql

350   sql = "SELECT * FROM salmshig WHERE SampleID = " & RP.SampleID & ""
360   Set tbsalm = New Recordset
370   RecOpenServer 0, tbsalm, sql

380   With frmRichText.rtb
390     .SelFontSize = 10
  
400     If tbR.EOF Then   'no requests
410       For n = 1 To 14
420         .SelText = vbCrLf
430         CrCnt = CrCnt + 1
440       Next
450     Else
460       If tbR!faecal And 2 ^ 2 Then  '?ova and parasite
470         If tbO!OP0 & tbO!OP1 & tbO!OP2 & "" = "" Then
480           .SelText = Left(" " & Space(23), 23)
490           .SelText = "Ova/Parasite Report:-  Not Ready" & vbCrLf
500           .SelText = vbCrLf
510           CrCnt = CrCnt + 2
520         Else
530           .SelText = Left(" " & Space(23), 23)
540           .SelText = "Ova/Parasite Report:-  " & tbO!OP0 & "" & vbCrLf
550           If Trim(tbO!OP1) & "" <> "" Then
560             .SelText = Left(" " & Space(46), 46)
570             .SelText = tbO!OP1 & "" & vbCrLf
580             CrCnt = CrCnt + 1
590           End If
600           If Trim(tbO!OP2) & "" <> "" Then
610             .SelText = Left(" " & Space(46), 46)
620             .SelText = tbO!OP2 & "" & vbCrLf
630             CrCnt = CrCnt + 1
640           End If
650         End If
660       End If
  
670       .SelFontSize = 10
    
680       If tbR!faecal And 2 ^ 6 Then  'rota virus
690         If Trim(tbO!Rota & "") = "" Then
700           .SelText = Space$(23)
710           .SelText = "Rota Virus         :-  Report Not Ready" & vbCrLf
720           CrCnt = CrCnt + 1
730         Else
740           .SelText = Space$(23)
750           .SelText = "Rota Virus         :- "
760           If tbO!Rota = "P" Then
770             .SelBold = True
780             .SelText = "*** Positive ***" & vbCrLf
790             .SelBold = False
800             CrCnt = CrCnt + 1
810           ElseIf tbO!Rota = "N" Then
820             .SelText = "Negative" & vbCrLf
830             CrCnt = CrCnt + 1
840           End If
850         End If
860       End If
    
870       If tbR!faecal And 2 ^ 6 Then  'adeno virus
880         If Trim(tbO!Adeno & "") = "" Then
890           .SelText = Left(" " & Space(23), 23)
900           .SelText = "Adeno Virus        :- Report Not Ready" & vbCrLf
910           CrCnt = CrCnt + 1
920         Else
930           .SelText = Space$(23)
940           .SelText = "Adeno Virus        :- "
950           If tbO!Adeno = "P" Then
960             .SelBold = True
970             .SelText = "*** Positive ***" & vbCrLf
980             CrCnt = CrCnt + 1
990             .SelBold = False
1000          ElseIf tbO!Adeno = "N" Then
1010            .SelText = "Negative" & vbCrLf
1020            CrCnt = CrCnt + 1
1030          End If
1040        End If
1050      End If
    
1060      If tbR!faecal And 2 ^ 7 Then  'Toxin A
1070        If Trim(tbO!ToxinAL & "") = "" And Trim(tbO!toxinata) & "" = "" Then
1080          .SelText = Space(23)
1090          .SelText = "C Diff Toxin A/B   :- Report Not Ready" & vbCrLf
1100          CrCnt = CrCnt + 1
1110        ElseIf Trim(tbO!ToxinAL & "") = "N" Or Trim(tbO!toxinata & "") = "N" Then
1120          .SelText = Space(23)
1130          .SelText = "C Diff Toxin A/B   :- Negative." & vbCrLf
1140          CrCnt = CrCnt + 1
1150        ElseIf Trim(tbO!ToxinAL & "") = "P" Or Trim(tbO!toxinata & "") = "P" Then
1160          .SelText = Space(23)
1170          .SelBold = True
1180          .SelText = "C Diff Toxin A/B   :- *** Positive ***" & vbCrLf
1190          CrCnt = CrCnt + 1
1200          .SelBold = False
1210        End If
1220      End If

1230      If tbR!faecal And 2 ^ 3 Or tbR!faecal And 2 ^ 4 Or tbR!faecal And 2 ^ 5 Then   'occult blood
1240        If Trim(tbO!Occult & "") = "" Then
1250          .SelText = "Occult Blood       :- Report Not Ready" & vbCrLf
1260          CrCnt = CrCnt + 1
1270        Else
1280          .SelText = Space(23)
1290          .SelText = "Occult Blood Report:- " & vbCrLf
1300          CrCnt = CrCnt + 1
1310          If Left(tbO!Occult, 1) = "N" Then
1320            .SelText = Space(45)
1330            .SelText = "1:" & "Negative" & vbCrLf
1340            CrCnt = CrCnt + 1
1350          ElseIf Left(tbO!Occult, 1) = "P" Then
1360            .SelBold = True
1370            .SelText = Space(45)
1380            .SelText = "1:" & "Positive" & vbCrLf
1390            CrCnt = CrCnt + 1
1400            .SelBold = False
1410          End If
1420          If Mid(tbO!Occult, 2, 1) = "N" Then
1430            .SelText = Space(45)
1440            .SelText = "2:" & "Negative" & vbCrLf
1450            CrCnt = CrCnt + 1
1460          ElseIf Mid(tbO!Occult, 2, 1) = "P" Then
1470            .SelBold = True
1480            .SelText = Space(45)
1490            .SelText = "2:" & "Positive" & vbCrLf
1500            CrCnt = CrCnt + 1
1510            .SelBold = False
1520          End If
1530          If Right(tbO!Occult, 1) = "N" Then
1540            .SelText = Space(45)
1550            .SelText = "3:" & "Negative" & vbCrLf
1560            CrCnt = CrCnt + 1
1570          ElseIf Right(tbO!Occult, 1) = "P" Then
1580            .SelBold = True
1590            .SelText = Space(45)
1600            .SelText = "3:" & "Positive" & vbCrLf
1610            CrCnt = CrCnt + 1
1620            .SelBold = False
1630          End If
1640        End If
1650      Else
1660        .SelText = vbCrLf
1670        CrCnt = CrCnt + 1
1680      End If

1690      If Not tbO.EOF Then
1700        If Trim(tbO!crp & "") = "N" Then
1710          .SelText = Space(23)
1720          .SelText = "Cryptosporidium Screen Negative" & vbCrLf
1730          CrCnt = CrCnt + 1
1740        ElseIf Trim(tbO!crp & "") = "P" Then
1750          .SelText = Space(23)
1760          .SelText = "Cryptosporidium Screen Positive" & vbCrLf
1770          CrCnt = CrCnt + 1
1780        End If
1790      End If

1800      If tbR!faecal And 2 ^ 8 Or tbR!faecal And 2 ^ 8 Or tbR!faecal And 2 ^ 0 Or tbR!faecal And 2 ^ 10 Then
1810        .SelText = vbCrLf
1820        CrCnt = CrCnt + 1
1830        .SelBold = True
1840        .SelText = Space(15)
1850        .SelText = "CULTURE : -" & vbCrLf
1860        CrCnt = CrCnt + 1
1870        CulFound = False
1880        .SelBold = False
1890      End If

1900      If tbR!faecal And 2 ^ 0 Then 'culture
1910        If Trim(tbO!CampCulture & "") <> "" Then
1920          .SelText = Space(23)
1930          CulFound = True
1940          If InStr(tbO!CampCulture, "No") > 0 Then
1950            .SelText = "No Campylobacter Isolated" & vbCrLf
1960            CrCnt = CrCnt + 1
1970          Else
1980            .SelText = "Campylobacter      :- "
1990            .SelBold = True
2000            .SelText = tbO!CampCulture & vbCrLf
2010            CrCnt = CrCnt + 1
2020            .SelBold = False
2030          End If
2040        End If
2050      End If
    
2060      If tbR!faecal And 2 ^ 10 Then  'salmonella and shigella
2070        .SelText = Space(23)
2080        If tbsalm.EOF Then
2090          CulFound = True
2100          .SelText = "Salmonella         :- Report Not Ready" & vbCrLf
2110          CrCnt = CrCnt + 1
2120        Else
2130          CulFound = True
2140          If Not tbsalm.EOF Then
2150            If Trim(tbsalm!salmident & "") = "" Then
2160              .SelText = "Salmonella         :- Report Not Ready" & vbCrLf
2170              CrCnt = CrCnt + 1
2180            Else
2190              If InStr(tbsalm!salmident, "No") > 0 Then
2200                .SelText = tbsalm!salmident & "" & vbCrLf
2210                CrCnt = CrCnt + 1
2220              Else
2230                .SelBold = True
2240                .SelText = tbsalm!salmident & "" & vbCrLf
2250                CrCnt = CrCnt + 1
2260                .SelBold = False
2270              End If
2280            End If
2290          Else
2300            .SelText = "Salmonella         :- Report Not Ready" & vbCrLf
2310            CrCnt = CrCnt + 1
2320          End If
2330        End If
2340      End If
    
2350      If tbR!faecal And 2 ^ 10 Then 'salmonella and shigella
2360        CulFound = True
2370        .SelText = Space(23)
2380        If tbsalm.EOF Then
2390          .SelText = "Shigella           :- Report Not Ready" & vbCrLf
2400          CrCnt = CrCnt + 1
2410        Else
2420          If Not tbsalm.EOF Then
2430            If Trim(tbsalm!shigType & "") = "" Then
2440              .SelText = "Shigella           :- Report Not Ready" & vbCrLf
2450              CrCnt = CrCnt + 1
2460            Else
2470              If InStr(tbsalm!shigType, "No") > 0 Then
2480                .SelText = tbsalm!shigType & "" & vbCrLf
2490                CrCnt = CrCnt + 1
2500              Else
2510                .SelBold = True
2520                .SelText = tbsalm!shigType & "" & " Isolated" & vbCrLf
2530                CrCnt = CrCnt + 1
2540                .SelBold = False
2550              End If
2560            End If
2570          Else
2580            .SelText = "Shigella           :- Report Not Ready" & vbCrLf
2590            CrCnt = CrCnt + 1
2600          End If
2610        End If
2620      End If
2630    End If
    
2640    If Not tbR.EOF Then
2650      If tbR!faecal And 2 ^ 8 Or tbR!faecal And 2 ^ 8 Or tbR!faecal And 2 ^ 0 Or tbR!faecal And 2 ^ 10 Then
2660        If CulFound = False Then
2670          .SelText = Space(23)
2680          .SelText = "Culture Not Ready" & vbCrLf
2690          CrCnt = CrCnt + 1
2700        End If
2710      End If
2720    End If
  
2730    If Not tbR.EOF Then
2740      If tbR!faecal And 2 ^ 9 Then 'e/p coli
2750        .SelText = Space(23)
2760        If InStr(tbO!epc, "P") > 0 Then
2770          .SelBold = True
2780          .SelText = "*** Enteropathogenic E Coli Isolated ***" & vbCrLf
2790          CrCnt = CrCnt + 1
2800          .SelBold = False
2810        ElseIf Len(Trim(tbO!epc)) < 4 Then
2820          .SelText = "Enteropathogenic E Coli Report not ready." & vbCrLf
2830          CrCnt = CrCnt + 1
2840        Else
2850          .SelText = "No Enteropathogenic E Coli isolated." & vbCrLf
2860          CrCnt = CrCnt + 1
2870        End If
2880      End If
2890    End If
  
2900    If Not tbO.EOF Then
2910      If Trim(tbO!pc0157report & "") <> "" Then   'coli 0157
2920        .SelText = Space(23)
2930        If InStr(tbO!pc0157report, "No") > 0 Then
2940          .SelText = "No E Coli 0157 Isolated" & vbCrLf
2950        Else
2960          .SelText = tbO!pc0157report & vbCrLf
2970        End If
2980        CrCnt = CrCnt + 1
2990      End If
3000    End If
  
3010    Do While CrCnt < 31
3020      .SelText = vbCrLf
3030      CrCnt = CrCnt + 1
3040    Loop
  
      '  Set Cx = Cxs.Load(RP.SampleID)
3050    Set OBS = OBS.Load(RP.SampleID, "MicroCS", "Demographic", "MicroGeneral")

3060    If Not OBS Is Nothing Then
3070      .SelBold = True
3080      .SelText = "Comment : "
3090      .SelBold = False
3100      For Each OB In OBS
3110          Select Case UCase$(OB.Discipline)
                  Case "DEMOGRAPHIC"
3120                  FillCommentLines OB.Comment, 2, Comments(), 87
3130                  For n = 1 To 4
3140                    If Trim(Comments(n)) <> "" Then .SelText = " " & Comments(n) & vbCrLf
3150                    CrCnt = CrCnt + 1
3160                  Next
3170              Case Else
3180                  FillCommentLines OB.Comment, 4, Comments(), 87
3190                  For n = 1 To 4
3200                    If Trim(Comments(n)) <> "" Then .SelText = " " & Comments(n) & vbCrLf
3210                    CrCnt = CrCnt + 1
3220                  Next
3230          End Select
3240      Next
3250    End If
    
3260    If IsDate(udtHeading.SampleDate) Then
3270      SampleDate = Format(udtHeading.SampleDate, "dd/mmm/yyyy hh:mm")
3280    Else
3290      SampleDate = ""
3300    End If
3310    If IsDate(RunTime) Then
3320      Rundate = Format(RunTime, "dd/mmm/yyyy hh:mm")
3330    Else
3340      If IsDate(udtHeading.Rundate) Then
3350        Rundate = Format(udtHeading.Rundate, "dd/mmm/yyyy")
3360      Else
3370        Rundate = ""
3380      End If
3390    End If
    
3400    PrintFooterRTB RP.Initiator, SampleDate, Rundate
  
3410    .SelStart = 0
3420    .SelPrint Printer.hDC
  
3430    sql = "SELECT * FROM Reports WHERE 0 = 1"
3440    Set tb = New Recordset
3450    RecOpenServer 0, tb, sql
3460    tb.AddNew
3470    tb!SampleID = RP.SampleID
3480    tb!Name = udtHeading.Name
3490    tb!Dept = "F"
3500    tb!Initiator = RP.Initiator
3510    tb!PrintTime = PrintTime
3520    tb!RepNo = "0F" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
3530    tb!PageNumber = 0
3540    tb!Report = .TextRTF
3550    tb!Printer = Printer.DeviceName
3560    tb.Update
3570  End With

3580  Exit Function

PrintFaeces_Error:

      Dim strES As String
      Dim intEL As Integer

3590  intEL = Erl
3600  strES = Err.Description
3610  LogError "modMicro", "PrintFaeces", intEL, strES

End Function

