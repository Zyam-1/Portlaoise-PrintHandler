Attribute VB_Name = "External"
Option Explicit

Public Sub PrintExternal()

      Dim tb As Recordset
      Dim sql As String
      Dim SampleDate As String
      Dim Rundate As String
      Dim DoB As String
      Dim Clin As String
      Dim SampleType As String
      Dim PrintTime As String
      Dim ClDetails As String
      Dim HospitalName As String
      Dim OB As Observation
      Dim OBS As New Observations
10    ReDim Comments(1 To 4) As String
      Dim n As Integer
      Dim CommentETC As String

20    On Error GoTo PrintExternal_Error

30    HospitalName = GetOptionSetting("STJAMESHOSPITAL", "St James Hospital")

40    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

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

140   Rundate = tb!Rundate

150   If IsDate(tb!DoB) Then
160       DoB = Format(tb!DoB, "dd/mmm/yyyy")
170   Else
180       DoB = ""
190   End If

200   ClDetails = Trim(tb!ClDetails & "")

210   ClearUdtHeading
220   With udtHeading
230       .SampleID = RP.SampleID
240       .Dept = "External"
250       .Name = tb!PatName & ""
260       .Ward = RP.Ward
270       .DoB = DoB
280       .Chart = tb!Chart & ""
290       .Clinician = RP.Clinician
300       .Address0 = tb!Addr0 & ""
310       .Address1 = tb!Addr1 & ""
320       .GP = RP.GP
330       .Sex = tb!Sex & ""
340       .Hospital = tb!Hospital & ""
350       .SampleDate = tb!SampleDate & ""
360       .RecDate = tb!RecDate & ""
370       .Rundate = tb!Rundate & ""
380       .GpClin = Clin
390       .SampleType = SampleType
400       .AandE = tb!AandE & ""
410   End With

420   PrintHeadingRTB

430   With frmRichText.rtb
440       .SelFontSize = 10
450       .SelText = vbCrLf
460       .SelText = "Tests Requested : " & vbCrLf
470       CrCnt = CrCnt + 2

480       sql = "SELECT * FROM ExtResults WHERE SampleID = " & RP.SampleID
490       Set tb = New Recordset
500       RecOpenClient 0, tb, sql
510       If Not tb.EOF Then
520           While Not tb.EOF
530               If UCase(HospitalName) <> UCase(tb!SendTo & "") Then
540                   PrintTextRTB frmRichText.rtb, Space(10) & tb!Analyte & "" & vbCrLf
550               End If
560               tb.MoveNext
570           Wend
580       End If

590       Do While CrCnt < 31
600           .SelText = vbCrLf
610           CrCnt = CrCnt + 1
620       Loop

630       If ClDetails <> "" Then
640           .SelText = "Clinical Details : " & ClDetails & vbCrLf
650           CrCnt = CrCnt + 1
660       End If


670       Set OBS = OBS.Load(RP.SampleID, "Demographic")
680       If Not OBS Is Nothing Then
690           For Each OB In OBS
700               Select Case UCase$(OB.Discipline)

                  Case "DEMOGRAPHIC"
710                   FillCommentLines OB.Comment, 2, Comments(), 87
720                   For n = 1 To 2
730                       If Trim(Comments(n) & "") <> "" Then
740                           .SelFontSize = 8
750                           .SelBold = True
760                           .SelText = Comments(n)
770                           .SelBold = False
780                           .SelText = vbCrLf
790                           CrCnt = CrCnt + 1
800                       End If
810                   Next
820               End Select
830           Next
840       End If

850       sql = "SELECT * from etc WHERE sampleid = '" & RP.SampleID & "'"
860       Set tb = New Recordset
870       RecOpenServer 0, tb, sql
880       If Not tb.EOF Then
890           CommentETC = tb!etc0 & "" & tb!etc1 & "" & tb!etc2 & "" & tb!etc3 & "" & _
                           tb!etc4 & "" & tb!etc5 & "" & tb!etc6 & "" & tb!etc7 & ""
900           If Trim(CommentETC) <> "" Then
910               .SelText = "Comment:" & vbCrLf
920               CrCnt = CrCnt + 1
930               FillCommentLines CommentETC, 4, Comments(), 87
940               For n = 1 To 4
950                   If Trim(Comments(n) & "") <> "" Then
960                       .SelFontSize = 8
970                       .SelBold = True
980                       .SelText = Comments(n)
990                       .SelBold = False
1000                      .SelText = vbCrLf
1010                      CrCnt = CrCnt + 1
1020                  End If
1030              Next
1040          End If
1050      End If
1060      PrintFooterRTB RP.Initiator, SampleDate, Rundate
1070      .SelStart = 0
          'Do not print if Doctor is disabled in DisablePrinting
          '*******************************************************************
1080      If CheckDisablePrinting(RP.Ward, "Externals") Then

1090      ElseIf CheckDisablePrinting(RP.GP, "Externals") Then
1100      Else
1110          .SelPrint Printer.hdc
1120      End If
          '*******************************************************************
          '1080      .SelPrint Printer.hDC

1130      sql = "SELECT * FROM Reports WHERE 0 = 1"
1140      Set tb = New Recordset
1150      RecOpenServer 0, tb, sql
1160      tb.AddNew
1170      tb!SampleID = RP.SampleID
1180      tb!Name = udtHeading.Name
1190      tb!Dept = "X"
1200      tb!Initiator = RP.Initiator
1210      tb!PrintTime = PrintTime
1220      tb!RepNo = "0X" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
1230      tb!PageNumber = 0
1240      tb!Report = .TextRTF
1250      tb!Printer = Printer.DeviceName
1260      tb.Update
1270  End With

1280  Exit Sub

PrintExternal_Error:

      Dim strES As String
      Dim intEL As Integer

1290  intEL = Erl
1300  strES = Err.Description
1310  LogError "External", "PrintExternal", intEL, strES, sql

End Sub

Public Sub PrintExternalMicro()

      Dim tb As Recordset
      Dim sql As String
      Dim SampleDate As String
      Dim Rundate As String
      Dim DoB As String
      Dim Clin As String
      Dim SampleType As String
      Dim PrintTime As String
      Dim Site As String
      Dim SiteDetails As String

10    On Error GoTo PrintExternalMicro_Error

20    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

30    sql = "SELECT Site, SiteDetails FROM MicroSiteDetails WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    If Not tb.EOF Then
70        Site = tb!Site & ""
80        SiteDetails = tb!SiteDetails & ""
90    End If

100   sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
110   Set tb = New Recordset
120   RecOpenClient 0, tb, sql
130   If tb.EOF Then Exit Sub

140   If IsDate(tb!SampleDate) Then
150       SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
160   Else
170       SampleDate = ""
180   End If

190   Rundate = tb!Rundate

200   If IsDate(tb!DoB) Then
210       DoB = Format(tb!DoB, "dd/mmm/yyyy")
220   Else
230       DoB = ""
240   End If

250   ClearUdtHeading
260   With udtHeading
270       .SampleID = RP.SampleID
280       .Dept = "External"
290       .Name = tb!PatName & ""
300       .Ward = RP.Ward
310       .DoB = DoB
320       .Chart = tb!Chart & ""
330       .Clinician = RP.Clinician
340       .Address0 = tb!Addr0 & ""
350       .Address1 = tb!Addr1 & ""
360       .GP = RP.GP
370       .Sex = tb!Sex & ""
380       .Hospital = tb!Hospital & ""
390       .SampleDate = tb!SampleDate & ""
400       .RecDate = tb!RecDate & ""
410       .Rundate = tb!Rundate & ""
420       .GpClin = Clin
430       .SampleType = SampleType
440       .AandE = tb!AandE & ""
450   End With

460   PrintHeadingRTB

470   With frmRichText.rtb
480       .SelFontSize = 10
490       .SelText = vbCrLf
500       .SelText = vbCrLf
510       If Trim$(Site) <> "" Then
520           .SelText = "Site : " & Site & vbCrLf
530           .SelText = "       " & SiteDetails
540           .SelText = vbCrLf & vbCrLf
550       End If

560       .SelText = "Test Request : " & vbCrLf
570       CrCnt = CrCnt + 2

          '620     Do While CrCnt < 31
          '630       .SelText = vbCrLf
          '640       CrCnt = CrCnt + 1
          '650     Loop

580       If Trim(tb!ClDetails & "") <> "" Then
590           .SelText = "Clinical Details : " & Trim(tb!ClDetails & "") & vbCrLf
600           CrCnt = CrCnt + 1
610       End If

620       PrintFooterRTB RP.Initiator, SampleDate, Rundate
630       .SelStart = 0
640       .SelPrint Printer.hdc

650       sql = "SELECT * FROM Reports WHERE 0 = 1"
660       Set tb = New Recordset
670       RecOpenServer 0, tb, sql
680       tb.AddNew
690       tb!SampleID = RP.SampleID
700       tb!Name = udtHeading.Name
710       tb!Dept = "X"
720       tb!Initiator = RP.Initiator
730       tb!PrintTime = PrintTime
740       tb!RepNo = "0X" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
750       tb!PageNumber = 0
760       tb!Report = .TextRTF
770       tb!Printer = Printer.DeviceName
780       tb.Update
790   End With

800   Exit Sub

PrintExternalMicro_Error:

      Dim strES As String
      Dim intEL As Integer

810   intEL = Erl
820   strES = Err.Description
830   LogError "External", "PrintExternalMicro", intEL, strES, sql

End Sub


