Attribute VB_Name = "modCytoHist"
Option Explicit

Public Sub PrintCytology(ByVal NoOfCopies As Integer)

      Dim tb As Recordset
      Dim n As Integer
      Dim Sex As String
      Dim DoB As String
      Dim sql As String
10    ReDim Comments(1 To 4) As String
      Dim SampleDate As String
      Dim Rundate As String
      Dim tbUN As Recordset
      Dim lpc As Integer
      Dim cUnits As String
      Dim v As String
      Dim Low As Single
      Dim High As Single
      Dim strLow As String * 4
      Dim strHigh As String * 4
      Dim TestCount As Integer
      Dim SampleType As String
      Dim ResultsPresent As Boolean
      Dim RunTime As String
      Dim Fasting As String
      Dim Fx As Fasting
      Dim strFormat As String
      Dim MultiColumn As Boolean
      Dim Samp As String
20    ReDim pl(1 To 1) As String
      Dim plCounter As Integer
      Dim crPos As Integer
      Dim HR As String
      Dim LinesAllowed As Integer
      Dim TotalPages As Integer
      Dim ThisPage As Integer
      Dim TopLine As Integer
      Dim BottomLine As Integer
      Dim crlfFound As Boolean
      Dim ds As Recordset
      Dim SID As Double
      Dim Yadd As Long
      'Dim Cx As Comment
      'Dim Cxs As New Comments
      Dim OB As Observation
      Dim OBS As New Observations
      Dim PrintTime As String
      Dim i As Integer

30    On Error GoTo PrintCytology_Error

40    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

50    sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "' and hyear = '" & RP.Year & "'"
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

190   DoB = tb!DoB & ""

200   Select Case Left(UCase(tb!Sex & ""), 1)
          Case "M": Sex = "M"
210       Case "F": Sex = "F"
220       Case Else: Sex = ""
230   End Select

240   Printer.Font.Name = "Courier New"
250   Printer.Font.Size = 10

260   Printer.Print

270   LinesAllowed = 58

280   sql = "SELECT * FROM cytoresults WHERE SampleID = '" & RP.SampleID & "' and hyear = '" & RP.Year & "'"
290   Set ds = New Recordset
300   RecOpenServer 0, ds, sql
310   If ds.EOF Then Exit Sub

320   If Trim(ds!UserName & "") <> "" Then RP.Initiator = Trim(ds!UserName)

330   HR = Trim(ds!cytoreport & "")
340   crlfFound = True
350   Do While crlfFound
360       HR = RTrim(HR)
370       crlfFound = False
380       If Right(HR, 1) = vbCr Or Right(HR, 1) = vbLf Then
390           HR = Left(HR, Len(HR) - 1)
400           crlfFound = True
410       End If
420   Loop

430   plCounter = 0
440   Do While Len(HR) > 0
450       crPos = InStr(HR, vbCr)
460       If crPos > 0 And crPos < 81 Then
470           plCounter = plCounter + 1
480           ReDim Preserve pl(1 To plCounter)
490           pl(plCounter) = Left(HR, crPos - 1)
500           HR = Mid(HR, crPos + 2)
510       Else
520           If Len(HR) > 81 Then
530               For n = 81 To 1 Step -1
540                   If Mid(HR, n, 1) = " " Then
550                       Exit For
560                   End If
570               Next
580               plCounter = plCounter + 1
590               ReDim Preserve pl(1 To plCounter)
600               pl(plCounter) = Left(HR, n)
610               HR = Mid(HR, n + 1)
620           Else
630               plCounter = plCounter + 1
640               ReDim Preserve pl(1 To plCounter)
650               pl(plCounter) = HR
660               Exit Do
670           End If
680       End If
690   Loop

700   Samp = RP.SampleID - (((Val(Swap_Year(Trim(tb!Hyear))) * 1000) + SysOptCytoOffset(0)))
710   Samp = Trim(tb!Hyear) & "/" & Samp & "C"

720   ClearUdtHeading
730   With udtHeading
740       .SampleID = Samp
750       .Dept = "Cytology"
760       .Name = tb!PatName & ""
770       .Ward = RP.Ward
780       .DoB = DoB
790       .Chart = tb!Chart & ""
800       .Clinician = RP.Clinician
810       .Address0 = tb!Addr0 & ""
820       .Address1 = tb!Addr1 & ""
830       .GP = RP.GP
840       .Sex = tb!Sex & ""
850       .Hospital = tb!Hospital & ""
860       .SampleDate = tb!SampleDate & ""
870       .RecDate = tb!RecDate & ""
880       .Rundate = tb!Rundate & ""
890       .GpClin = ""
900       .SampleType = ""
910       .AandE = tb!AandE & ""
920   End With

930   TotalPages = Int((plCounter - 1) / LinesAllowed) + 1
940   If TotalPages = 0 Then TotalPages = 1
950   For i = 1 To NoOfCopies
960       For ThisPage = 1 To TotalPages

970           With frmRichText.rtb
980               PrintHeadingRTB
990               .SelFontName = "Courier New"
1000              .SelFontSize = 10

1010              .SelText = vbCrLf
1020              CrCnt = CrCnt + 1

1030              .SelBold = False
1040              .SelText = "Nature of Specimen" & vbCrLf
1050              .SelText = "[A]: "
1060              .SelBold = True
1070              .SelText = Left(ds!natureofspecimen & "" & Space(50), 50)
1080              If ds!natureofspecimenb & "" <> "" Then
1090                  .SelBold = False
1100                  .SelText = "[B]: "
1110                  .SelBold = True
1120                  .SelText = ds!natureofspecimenb & "" & vbCrLf
1130                  CrCnt = CrCnt + 1
1140              Else
1150                  CrCnt = CrCnt + 1
1160                  .SelText = vbCrLf
1170              End If

1180              If ds!natureofspecimenC & "" <> "" Then
1190                  .SelBold = False
1200                  .SelText = "[C]: "
1210                  .SelBold = True
1220                  .SelText = Left(ds!natureofspecimenC & "" & Space(50), 50)
1230                  .SelBold = False
1240                  .SelText = "[D]: "
1250                  .SelBold = True
1260                  .SelText = ds!natureofspecimenD & "" & vbCrLf
1270                  CrCnt = CrCnt + 1
1280              Else
1290                  .SelText = vbCrLf
1300                  CrCnt = CrCnt + 1
1310              End If

1320              .SelFontSize = 2
1330              .SelText = String$(420, "-") & vbCrLf

1340              .SelFontSize = 10

1350              TopLine = (ThisPage - 1) * LinesAllowed + 1
1360              BottomLine = (ThisPage - 1) * LinesAllowed + LinesAllowed
1370              If BottomLine > plCounter Then
1380                  BottomLine = plCounter
1390              End If
1400              For n = TopLine To BottomLine
1410                  If UCase(Left(pl(n), 24)) = "MICROSCOPIC EXAMINATION:" Or _
                         UCase(Left(pl(n), 29)) = "BONE MARROW ASPIRATE & BIOPSY" Or _
                         UCase(Left(pl(n), 18)) = "GROSS EXAMINATION:" Or _
                         UCase(Left(pl(n), 21)) = "SUPPLEMENTARY REPORT:" Or _
                         UCase(Left(pl(n), 14)) = "DR. K. CUNNANE" Or _
                         UCase(Left(pl(n), 17)) = "DR. KEVIN CUNNANE" Or _
                         UCase(Left(pl(n), 11)) = "DR GERARD C" Or _
                         UCase(Left(pl(n), 11)) = "PATHOLOGIST" Or _
                         UCase(Left(pl(n), 18)) = "DR. J. D. GILSENAN" Or _
                         UCase(Left(pl(n), 14)) = "FURTHER REPORT" Or _
                         UCase(Left(pl(n), 10)) = "APPEARANCE" Or _
                         UCase(Left(pl(n), 23)) = "MICROSCOPIC EXAMINATION" Or _
                         UCase(Left(pl(n), 10)) = "CONSULTANT" Or _
                         UCase(Left(pl(n), 7)) = "COMMENT" Or _
                         UCase(Left(pl(n), 21)) = "SUPPLEMENTARY REPORT" Or _
                         UCase(Left(pl(n), 11)) = "APPEARANCE:" Then
1420                      .SelBold = True
1430                  Else
1440                      .SelBold = False
1450                  End If
1460                  .SelText = "   " & pl(n) & vbCrLf
1470                  CrCnt = CrCnt + 1
1480              Next

1490              .SelText = vbCrLf
1500              If Trim(ds!cytocomment & "") <> "" Then
1510                  .SelText = Trim(ds!cytocomment & "")
1520              End If
1530              .SelText = vbCrLf
                  '        Set Cx = Cxs.Load(RP.SampleID)
1540              Set OBS = OBS.Load(RP.SampleID, "Demographic")
1550              If Not OBS Is Nothing Then
1560                  For Each OB In OBS
1570                      FillCommentLines OB.Comment, 2, Comments(), 87
1580                      For n = 1 To 2
1590                          If Trim(Comments(n) & "") <> "" Then
1600                              .SelFontSize = 10
1610                              .SelText = Comments(n)
1620                              .SelText = vbCrLf
1630                              CrCnt = CrCnt + 1
1640                          End If
1650                      Next
1660                  Next
1670              End If

1680              Do While CrCnt < 70
1690                  .SelText = vbCrLf
1700                  CrCnt = CrCnt + 1
1710              Loop

1720              .SelText = Space(40) & "Page " & ThisPage & " of " & TotalPages & vbCrLf

1730              PrintFooterRTB RP.Initiator, SampleDate, Rundate
1740              .SelStart = 0
1750              .SelPrint Printer.hDC
1760              Yadd = Val(Swap_Year(RP.Year)) * 1000
1770              SID = RP.SampleID + SysOptCytoOffset(0) + Yadd

1780              sql = "SELECT * FROM Reports WHERE 0 = 1"
1790              Set tb = New Recordset
1800              RecOpenServer 0, tb, sql
1810              tb.AddNew
1820              tb!SampleID = RP.SampleID
1830              tb!Name = udtHeading.Name
1840              tb!Dept = "Y"
1850              tb!Initiator = RP.Initiator
1860              tb!PrintTime = PrintTime
1870              tb!RepNo = Format(ThisPage - 1) & "Y" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
1880              tb!PageNumber = ThisPage - 1
1890              tb!Report = frmRichText.rtb.TextRTF
1900              tb!Printer = Printer.DeviceName
1910              tb.Update
1920          End With
1930      Next
1940  Next i
1950  Exit Sub

PrintCytology_Error:

      Dim strES As String
      Dim intEL As Integer

1960  intEL = Erl
1970  strES = Err.Description
1980  LogError "modCytoHist", "PrintCytology", intEL, strES, sql

End Sub

Public Sub PrintHistology(NoOfCopies As Integer)

      'Dim Cx As Comment
      'Dim Cxs As New Comments
      Dim OB As Observation
      Dim OBS As New Observations
      Dim tb As Recordset
      Dim sn As Recordset
      Dim n As Integer
      Dim Sex As String
      Dim DoB As String
      Dim sql As String
10    ReDim Comments(1 To 4) As String
      Dim SampleDate As String
      Dim Rundate As String
      Dim tbUN As Recordset
      Dim lpc As Integer
      Dim cUnits As String
      Dim v As String
      Dim Low As Single
      Dim High As Single
      Dim strLow As String * 4
      Dim strHigh As String * 4
      Dim TestCount As Integer
      Dim SampleType As String
      Dim ResultsPresent As Boolean
      Dim RunTime As String
      Dim Fasting As String
      Dim Fx As Fasting
      Dim strFormat As String
      Dim MultiColumn As Boolean
20    ReDim pl(1 To 1) As String
      Dim plCounter As Integer
      Dim crPos As Integer
      Dim HR As String
      Dim LinesAllowed As Integer
      Dim TotalPages As Integer
      Dim ThisPage As Integer
      Dim TopLine As Integer
      Dim BottomLine As Integer
      Dim crlfFound As Boolean
      Dim Samp As String
      Dim SID As Double
      Dim Yadd As Long
      Dim Sampid As String
      Dim PrintTime As String
      Dim i As Integer

30    On Error GoTo PrintHistology_Error


40    Sampid = RP.SampleID

50    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

60    sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "' and hyear = '" & RP.Year & "'"
70    Set tb = New Recordset
80    RecOpenClient 0, tb, sql
90    If tb.EOF Then Exit Sub

100   If IsDate(tb!SampleDate) Then
110       SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
120   Else
130       SampleDate = ""
140   End If
150   If IsDate(tb!Rundate) Then
160       Rundate = Format(tb!Rundate, "dd/mmm/yyyy hh:mm")
170   Else
180       Rundate = ""
190   End If

200   DoB = tb!DoB & ""

210   Select Case Left(UCase(tb!Sex & ""), 1)
          Case "M": Sex = "M"
220       Case "F": Sex = "F"
230       Case Else: Sex = ""
240   End Select

250   Printer.Font.Name = "Courier New"
260   Printer.Font.Size = 10

270   Printer.Print

280   LinesAllowed = 58

290   sql = "SELECT * FROM historesults WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
300   Set sn = New Recordset
310   With sn
320       .CursorLocation = adUseServer
330       .CursorType = adOpenStatic
340       .LockType = adLockPessimistic
350       .ActiveConnection = Cnxn(0)
360       .Source = sql
370       .Open
380   End With
390   If tb.EOF Then Exit Sub

400   Rundate = Format(tb!Rundate, "dd/MMM/yyyy")

410   If sn.EOF Then Exit Sub

420   If Trim(sn!validdate & "") <> "" Then Rundate = Format(sn!validdate, "dd/MMM/yyyy")

430   If Trim(sn!UserName & "") <> "" Then RP.Initiator = Trim(sn!UserName) Else RP.Initiator = ""

440   HR = Trim(sn!historeport & "")
450   crlfFound = True
460   Do While crlfFound
470       HR = RTrim(HR)
480       crlfFound = False
490       If Right(HR, 1) = vbCr Or Right(HR, 1) = vbLf Then
500           HR = Left(HR, Len(HR) - 1)
510           crlfFound = True
520       End If
530   Loop

540   plCounter = 0
550   Do While Len(HR) > 0
560       crPos = InStr(HR, vbCr)
570       If crPos > 0 And crPos < 81 Then
580           plCounter = plCounter + 1
590           ReDim Preserve pl(1 To plCounter)
600           pl(plCounter) = Left(HR, crPos - 1)
610           HR = Mid(HR, crPos + 2)
620       Else
630           If Len(HR) > 81 Then
640               For n = 81 To 1 Step -1
650                   If Mid(HR, n, 1) = " " Then
660                       Exit For
670                   End If
680               Next
690               plCounter = plCounter + 1
700               ReDim Preserve pl(1 To plCounter)
710               pl(plCounter) = Left(HR, n)
720               HR = Mid(HR, n + 1)
730           Else
740               plCounter = plCounter + 1
750               ReDim Preserve pl(1 To plCounter)
760               pl(plCounter) = HR
770               Exit Do
780           End If
790       End If
800   Loop

810   If RP.SampleID > 30000000 Then Samp = RP.SampleID - (((Val(Swap_Year(Trim(tb!Hyear))) * 1000) + 30000000))
820   Samp = Trim(tb!Hyear) & "/" & Samp & "H"

830   ClearUdtHeading
840   With udtHeading
850       .SampleID = Samp
860       .Dept = "Histology"
870       .Name = tb!PatName & ""
880       .Ward = RP.Ward
890       .DoB = DoB
900       .Chart = tb!Chart & ""
910       .Clinician = RP.Clinician
920       .Address0 = tb!Addr0 & ""
930       .Address1 = tb!Addr1 & ""
940       .GP = RP.GP
950       .Sex = tb!Sex & ""
960       .Hospital = tb!Hospital & ""
970       .SampleDate = tb!SampleDate & ""
980       .RecDate = tb!RecDate & ""
990       .Rundate = Rundate
1000      .SampleType = ""
1010      .AandE = tb!AandE & ""
1020  End With

1030  TotalPages = Int((plCounter - 1) / LinesAllowed) + 1
1040  If TotalPages = 0 Then TotalPages = 1

1050  For i = 1 To NoOfCopies
1060      For ThisPage = 1 To TotalPages

1070          With frmRichText.rtb
1080              PrintHeadingRTB
1090              .SelFontName = "Courier New"
1100              .SelFontSize = 10
1110              .SelBold = False
1120              .SelText = "Nature of Specimen" & vbCrLf
1130              .SelText = "[A]: "
1140              .SelBold = True
1150              .SelText = Left(sn!natureofspecimen & "" & Space(40), 40)
1160              If sn!natureofspecimenC & "" <> "" Then
1170                  .SelBold = False
1180                  .SelText = " [C]: "
1190                  .SelBold = True
1200                  .SelText = sn!natureofspecimenC & "" & vbCrLf
1210                  CrCnt = CrCnt + 1
1220              Else
1230                  CrCnt = CrCnt + 1
1240                  .SelText = vbCrLf
1250              End If

1260              If sn!natureofspecimenb & "" <> "" Then
1270                  .SelBold = False
1280                  .SelText = "[B]: "
1290                  .SelBold = True
1300                  .SelText = Left(sn!natureofspecimenb & "" & Space(40), 40)
1310                  .SelBold = False
1320                  If sn!natureofspecimenC & "" <> "" Then
1330                      .SelText = " [D]: "
1340                      .SelBold = True
1350                      .SelText = sn!natureofspecimenD & "" & vbCrLf
1360                      CrCnt = CrCnt + 1
1370                  Else
1380                      .SelText = vbCrLf
1390                      CrCnt = CrCnt + 1
1400                  End If
1410              Else
1420                  .SelText = vbCrLf
1430                  CrCnt = CrCnt + 1
1440              End If

1450              If sn!NatureOfSpecimenE & "" <> "" Then
1460                  .SelBold = False
1470                  .SelText = Left(" " & Space(22), 22) & "[E]: "
1480                  .SelBold = True
1490                  .SelText = Left(sn!NatureOfSpecimenE & "" & Space(33), 33)
1500                  .SelBold = False
1510                  .SelText = "[F]: "
1520                  .SelBold = True
1530                  .SelText = sn!NatureOfSpecimenF & "" & vbCrLf
1540                  CrCnt = CrCnt + 1
1550              Else
1560                  .SelText = vbCrLf
1570                  CrCnt = CrCnt + 1
1580              End If

1590              .SelFontSize = 2
1600              .SelText = String$(420, "-") & vbCrLf

1610              .SelFontSize = 10

1620              TopLine = (ThisPage - 1) * LinesAllowed + 1
1630              BottomLine = (ThisPage - 1) * LinesAllowed + LinesAllowed
1640              If BottomLine > plCounter Then
1650                  BottomLine = plCounter
1660              End If
1670              For n = TopLine To BottomLine
1680                  If UCase(Left(pl(n), 24)) = "MICROSCOPIC EXAMINATION:" Or _
                         UCase(Left(pl(n), 29)) = "BONE MARROW ASPIRATE & BIOPSY" Or _
                         UCase(Left(pl(n), 18)) = "GROSS EXAMINATION:" Or _
                         UCase(Left(pl(n), 21)) = "SUPPLEMENTARY REPORT:" Or _
                         UCase(Left(pl(n), 14)) = "DR. K. CUNNANE" Or _
                         UCase(Left(pl(n), 17)) = "DR. KEVIN CUNNANE" Or _
                         UCase(Left(pl(n), 11)) = "DR GERARD C" Or _
                         UCase(Left(pl(n), 11)) = "PATHOLOGIST" Or _
                         UCase(Left(pl(n), 18)) = "DR. J. D. GILSENAN" Or _
                         UCase(Left(pl(n), 14)) = "FURTHER REPORT" Or _
                         UCase(Left(pl(n), 10)) = "APPEARANCE" Or _
                         UCase(Left(pl(n), 23)) = "MICROSCOPIC EXAMINATION" Or _
                         UCase(Left(pl(n), 10)) = "CONSULTANT" Or _
                         UCase(Left(pl(n), 7)) = "COMMENT" Or _
                         UCase(Left(pl(n), 21)) = "SUPPLEMENTARY REPORT" Then
1690                      .SelBold = True
1700                  Else
1710                      .SelBold = False
1720                  End If
1730                  .SelText = "   " & pl(n) & vbCrLf
1740                  CrCnt = CrCnt + 1
1750              Next

1760              Set OBS = OBS.Load(RP.SampleID, "Demographic")
1770              If Not OBS Is Nothing Then
1780                  For Each OB In OBS
1790                      FillCommentLines OBS.Item(1).Comment, 2, Comments(), 87
1800                      For n = 1 To 2
1810                          If Trim(Comments(n) & "") <> "" Then
1820                              .SelFontSize = 10
1830                              .SelText = Comments(n)
1840                              .SelText = vbCrLf
1850                              CrCnt = CrCnt + 1
1860                          End If
1870                      Next
1880                  Next
1890              End If

1900              Do While CrCnt < 70
1910                  .SelText = vbCrLf
1920                  CrCnt = CrCnt + 1
1930              Loop

1940              .SelText = Space(40) & "Page " & ThisPage & " of " & TotalPages & vbCrLf

1950              PrintFooterRTB RP.Initiator, SampleDate, Rundate
1960              .SelStart = 0
1970              .SelPrint Printer.hDC


1980              sql = "SELECT * FROM Reports WHERE RepNo = '" & Format(ThisPage - 1) & "P" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss") & "'"
1990              Set tb = New Recordset
2000              RecOpenServer 0, tb, sql
2010              If tb.EOF Then
2020                  tb.AddNew
2030                  tb!SampleID = RP.SampleID
2040                  tb!Name = udtHeading.Name
2050                  tb!Dept = "P"
2060                  tb!Initiator = RP.Initiator
2070                  tb!PrintTime = PrintTime
2080                  tb!RepNo = Format(ThisPage - 1) & "P" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
2090                  tb!PageNumber = ThisPage - 1
2100                  tb!Report = frmRichText.rtb.TextRTF
2110                  tb!Printer = Printer.DeviceName
2120                  tb.Update
2130              End If
2140          End With
2150      Next
2160  Next
2170  Exit Sub

PrintHistology_Error:

      Dim strES As String
      Dim intEL As Integer

2180  intEL = Erl
2190  strES = Err.Description
2200  LogError "modCytoHist", "PrintHistology", intEL, strES, sql

End Sub

Public Function Swap_Year(ByVal Hyear As String) As String

10    On Error GoTo Swap_Year_Error

20    Swap_Year = Right(Hyear, 1) & Mid(Hyear, 3, 1) & Mid(Hyear, 2, 1) & Left(Hyear, 1)

30    Exit Function

Swap_Year_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "modCytoHist", "Swap_Year", intEL, strES

End Function
