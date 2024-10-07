Attribute VB_Name = "modSemen"

Option Explicit

Public Sub PrintSAReport()

      Dim tb As Recordset
      Dim tu As Recordset
      Dim tm As Recordset
      Dim sql As String
      'Dim Cx As Comment
      'Dim Cxs As New Comments
      Dim OB As Observation
      Dim OBS As New Observations
10    ReDim Comments(1 To 4) As String
      Dim DoB As String
      Dim n As Integer
      Dim SampleDate As String
      Dim Rundate As String
      Dim f As Integer
      Dim PrintTime As String
      Dim SemenMorph As String
      Dim Morphology As String

20    On Error GoTo PrintSAReport_Error

30    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

40    sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql

70    DoB = Format(tb!DoB, "dd/MMM/yyyy")

80    ClearUdtHeading
90    With udtHeading
100       .SampleID = RP.SampleID - SysOptSemenOffset(0)
110       .Dept = "Microbiology"
120       .Name = tb!PatName & ""
130       .Ward = RP.Ward
140       .DoB = DoB
150       .Chart = tb!Chart & ""
160       .Clinician = RP.Clinician
170       .Address0 = tb!Addr0 & ""
180       .Address1 = tb!Addr1 & ""
190       .GP = RP.GP
200       .Sex = tb!Sex & ""
210       .Hospital = tb!Hospital & ""
220       .SampleDate = tb!SampleDate & ""
230       .RecDate = tb!RecDate & ""
240       .Rundate = tb!Rundate & ""
250       .GpClin = ""
260       .SampleType = ""
270       .AandE = tb!AandE & ""
280   End With

290   PrintHeadingRTB

300   sql = "SELECT * FROM SemenResults WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
310   Set tu = New Recordset
320   RecOpenServer 0, tu, sql

330   With frmRichText.rtb
340       .SelFontName = "Courier New"

350       .SelFontSize = 10

360       If Not tu.EOF Then
370           If Trim$(tu!UserName & "") <> "" Then
380               RP.Initiator = tu!UserName
390           End If
400           .SelText = vbCrLf
410           If Trim(tb!ClDetails) & "" <> "" Then
420               .SelText = "      Clinical Details: " & tb!ClDetails & vbCrLf
430           End If
440           .SelText = vbCrLf
450           .SelText = "       Semen Analysis :" & vbCrLf
460           .SelText = vbCrLf
470           .SelText = "                Volume: " & tu!Volume & "" & " mL" & vbCrLf
480           .SelText = vbCrLf
490           .SelText = "           Consistency: " & tu!Consistency & "" & vbCrLf
500           .SelText = vbCrLf
510           .SelText = "     Spermatozoa Count: "
520           If InStr(UCase(tu!semenCount & ""), "SEEN") Then
530               .SelText = tu!semenCount & "" & vbCrLf
540           Else
550               .SelText = tu!semenCount & "" & " Million per mL" & vbCrLf
560           End If
570           .SelText = vbCrLf
580           .SelText = vbCrLf

590           If Trim$(tu!Motility & tu!MotilityPro & tu!MotilityNonPro & tu!MotilityNonMotile & "") <> "" Then
600               .SelText = "             Motility :" & vbCrLf
610               If Trim$(tu!Motility & "") <> "" Then
620                   .SelText = "                   " & Right$("   " & Trim$(tu!Motility), 3) & " % Motile." & vbCrLf
630               End If
640               If Trim$(tu!MotilityPro & "") <> "" Then
650                   .SelText = "                   " & Right$("   " & Trim$(tu!MotilityPro), 3) & " % Motile Progressive." & vbCrLf
660               End If
670               If Trim$(tu!MotilitySlow & "") <> "" Then
680                   .SelText = "                   " & Right$("   " & Trim$(tu!MotilitySlow), 3) & " % Motile Slow Progressive." & vbCrLf
690               End If
700               If Trim$(tu!MotilityNonPro & "") <> "" Then
710                   .SelText = "                   " & Right$("   " & Trim$(tu!MotilityNonPro), 3) & " % Motile Non-Progressive." & vbCrLf
720               End If
730               If Trim$(tu!MotilityNonMotile & "") <> "" Then
740                   .SelText = "                   " & Right$("   " & Trim$(tu!MotilityNonMotile), 3) & " % Non Motile." & vbCrLf
750               End If
760           End If
770       End If

780       .SelText = vbCrLf
790       .SelText = vbCrLf

800       Morphology = ""
810       SemenMorph = ""
820       sql = "SELECT Result, UserName FROM GenericResults WHERE " & _
                "SampleID = '" & RP.SampleID & "' " & _
                "AND TestName = 'SemenMorphResult'"
830       Set tm = New Recordset
840       RecOpenServer 0, tm, sql
850       If Not tm.EOF Then
860           If Trim$(tm!UserName & "") <> "" Then
870               RP.Initiator = tm!UserName
880           End If
890           Morphology = tm!Result & ""
900       End If

910       sql = "SELECT * FROM GenericResults WHERE " & _
                "SampleID = '" & RP.SampleID & "' " & _
                "AND TestName = 'SemenMorphDescription'"
920       Set tm = New Recordset
930       RecOpenServer 0, tm, sql
940       If Not tm.EOF Then
950           If Trim$(tm!UserName & "") <> "" Then
960               RP.Initiator = tm!UserName
970           End If
980           SemenMorph = tm!Result & ""
990       End If

1000      If SemenMorph <> "" And Morphology <> "" Then
1010          .SelText = "           Morphology : "
1020          .SelText = Morphology & " "
1030          .SelText = SemenMorph & vbCrLf
1040      End If

          '  Set Cx = Cxs.Load(RP.SampleID)
1050      Set OBS = OBS.Load(RP.SampleID, "Semen", "Demographic")

1060      If Not OBS Is Nothing Then
1070          For Each OB In OBS
1080              Select Case UCase$(OB.Discipline)
                  Case "SEMEN"
1090                  FillCommentLines OB.Comment, 4, Comments(), 80
1100                  For n = 1 To 4
1110                      If Trim(Comments(n) & "") <> "" Then
1120                          .SelFontName = "Courier New"
1130                          .SelFontSize = 10
1140                          .SelText = Comments(n) & vbCrLf
1150                          CrCnt = CrCnt + 1
1160                      End If
1170                  Next
1180              Case "DEMOGRAPHIC"
1190                  FillCommentLines OB.Comment, 2, Comments(), 80
1200                  For n = 1 To 2
1210                      If Trim(Comments(n) & "") <> "" Then
1220                          .SelText = Comments(n) & vbCrLf
1230                          CrCnt = CrCnt + 1
1240                      End If
1250                  Next
1260              End Select
1270          Next
1280      End If

1290      If IsDate(tb!SampleDate) Then
1300          SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
1310      Else
1320          SampleDate = ""
1330      End If
1340      If IsDate(tb!Rundate) Then
1350          Rundate = Format(tb!Rundate, "dd/mmm/yyyy hh:mm")
1360      Else
1370          Rundate = ""
1380      End If

1390      CrCnt = 25

1400      If RP.FaxNumber <> "" Then
1410          PrintFooterRTBFax RP.Initiator, SampleDate, Rundate
1420          f = FreeFile
1430          Open SysOptFax(0) & RP.SampleID & "URN.doc" For Output As f
1440          Print #f, .TextRTF
1450          Close f
1460          SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "URN.doc"
1470      Else
1480          PrintFooterRTB RP.Initiator, SampleDate, Rundate
1490          .SelStart = 0
              'Do not print if Doctor is disabled in DisablePrinting
              '*******************************************************************
1500          If CheckDisablePrinting(RP.Ward, "Semen Analysis") Then

1510          ElseIf CheckDisablePrinting(RP.GP, "Semen Analysis") Then
1520          Else
1530              .SelPrint Printer.hDC
1540          End If
              '*******************************************************************
              '.SelPrint Printer.hDC
1550      End If

1560      sql = "SELECT * FROM Reports WHERE 0 = 1"
1570      Set tb = New Recordset
1580      RecOpenServer 0, tb, sql
1590      tb.AddNew
1600      tb!SampleID = RP.SampleID
1610      tb!Name = udtHeading.Name
1620      tb!Dept = "Z"
1630      tb!Initiator = RP.Initiator
1640      tb!PrintTime = PrintTime
1650      tb!RepNo = "0Z" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
1660      tb!PageNumber = 0
1670      tb!Report = .TextRTF
1680      tb!Printer = Printer.DeviceName
1690      tb.Update
1700  End With

1710  Exit Sub

PrintSAReport_Error:

      Dim strES As String
      Dim intEL As Integer

1720  intEL = Erl
1730  strES = Err.Description
1740  LogError "modSemen", "PrintSAReport", intEL, strES, sql

End Sub


