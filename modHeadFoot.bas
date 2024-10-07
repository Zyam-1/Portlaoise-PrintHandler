Attribute VB_Name = "modHeadFoot"
Option Explicit
Public CrCnt As Long

Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const EM_GETLINECOUNT = &HBA
Public Sub PrintFooterRTB(ByVal Initiator As String, _
                          ByVal SampleDate As String, _
                          ByVal Rundate As String, _
                          Optional ByVal ExternalTestingNote As String = "", _
                          Optional ByVal Department As String = "", _
                          Optional ByVal HosName As String = "")

10        On Error GoTo PrintFooterRTB_Error

20        With frmRichText.rtb
30            .SelFontName = "Courier New"
40            .SelColor = vbBlack

50            If UCase(pForcePrintTo) = "FAX" Then
60                Do While CrCnt < GetOptionSetting("FooterStartLine", 33)
70                    .SelText = vbCrLf
80                    CrCnt = CrCnt + 1
90                Loop
100           Else
110               Do While CrCnt < GetOptionSetting("FooterStartLine", 33)
120                   .SelText = vbCrLf
130                   CrCnt = CrCnt + 1
140               Loop
150           End If


              '        If Department = "Haem" And HosName = "Portlaoise" Then
              '            .SelFontSize = 8
              '            .SelText = "FBC and Automated Differential are not under current scope of accreditation."
              '        End If
160           .SelFontSize = 4
170           .SelText = vbCrLf

180           .SelFontSize = 2
190           .SelText = String(420, "-") & vbCrLf

200           .SelFontSize = 9
210           .SelBold = False

220           If Len(SampleDate) > 10 Then
230               If Format$(SampleDate, "hh:mm") = "00:00" Then
240                   .SelText = Left$(" Sample Date : " & Format$(SampleDate, "dd/MM/yy") & Space(30), 30)
250               Else
260                   .SelText = Left$(" Sample Date : " & Format$(SampleDate, "dd/MM/yy hh:mm") & Space(30), 30)
270               End If
280           Else
290               .SelText = Left$(" Sample Date : " & Format$(SampleDate, "dd/MM/yy") & Space(30), 30)
300           End If
310           If RP.Department = "P" Or RP.Department = "Y" Then
320               .SelText = Left$(" " & Space(30), 30)
330           Else

340               Rundate = Format(Rundate, "dd/MM/yyyy hh:mm")
350               If Right(Rundate, 5) = "00:00" Then
360                   .SelText = Left$("Run Date : " & Format$(Rundate, "dd/mm/yy") & Space(30), 30)
370               Else
380                   .SelText = Left$("Run Date : " & Format$(Rundate, "dd/mm/yy hh:mm") & Space(30), 30)
390               End If

400           End If
              Dim qSampleID As String
              qSampleID = CStr(RP.SampleID)
410           If RP.Department = "K" Then
420               .SelText = " DRAFT REPORT"
                  '    ElseIf RP.Department = "P" Then
                  '        .SelText = " Printed by " & Initiator
                  '    ElseIf RP.Department = "Y" Then
                  '        .SelText = " Printed by " & Initiator
430           ElseIf RP.Department = "C" Then
                  'Zyam
                  Dim sqlAuthorised As String
                  Dim sqlusername As String
                  Dim tb As Recordset
                  Dim tb1 As Recordset
                  
                  sqlAuthorised = "SELECT username from Coagresults WHERE SampleID = '" & RP.SampleID & "' AND Username is not null"
                  Set tb = New Recordset
                  RecOpenClient 0, tb, sqlAuthorised
                  If tb!UserName = "HEM" Then
                    .SelText = " Authorised by: HemoHub AutoVal"
                  ElseIf tb!UserName = "CHA" Then
                    .SelText = " Authorised by: Charlotte Muldowney"
                  Else
                    
                    sqlusername = "SELECT Name from Users WHERE Code = '" & tb!UserName & "'"
                    Set tb1 = New Recordset
                    RecOpenClient 0, tb1, sqlusername
440                 .SelText = " Authorised by: " & tb1!Name
                     
                    
                  End If
              ElseIf RP.Department = "B" Then
                  .SelText = " Authorised by: " & getOperatorName("BioResults", qSampleID)
              ElseIf RP.Department = "H" Then
                  .SelText = " Authorised by: " & getOperatorName("HaemResults", qSampleID)
              ElseIf RP.Department = "I" Or RP.Department = "J" Then
                  .SelText = " Authorised by: " & getOperatorName("ImmResults", qSampleID)
              ElseIf RP.Department = "E" Then
                  .SelText = " Authorised by: " & getOperatorName("EndResults", qSampleID)
              Else
                  Dim sqlAuthorised1 As String
                  Dim tb2 As Recordset
                  sqlAuthorised1 = "SELECT username from ExtResults WHERE SampleID = '" & RP.SampleID & "' AND Username is not null"
                  Set tb2 = New Recordset
                  RecOpenClient 0, tb2, sqlAuthorised1
441                 .SelText = " Authorised by: " & tb2!UserName
                  'Zyam
450           End If
460           .SelText = vbCrLf
470           .SelText = ExternalTestingNote
480       End With

490       Exit Sub

PrintFooterRTB_Error:

          Dim strES As String
          Dim intEL As Integer

500       intEL = Erl
510       strES = Err.Description
520       LogError "modHeadFoot", "PrintFooterRTB", intEL, strES

End Sub


Public Sub PrintFooterA4RTB(ByVal Initiator As String, _
                            ByVal SampleDate As String, _
                            ByVal Rundate As String)

      Dim LineCount As Long
      Dim y As Long

10    On Error GoTo PrintFooterA4RTB_Error

20    y = Printer.Height

30    With frmRichText.rtb

40        LineCount = SendMessage(.hWnd, EM_GETLINECOUNT, 0&, 0&)

50        .SelFontName = "Courier New"
60        .SelColor = vbBlack

70        Do While LineCount < 100
80            .SelText = vbCrLf
90            LineCount = SendMessage(.hWnd, EM_GETLINECOUNT, 0&, 0&)
100       Loop

110       .SelFontSize = 2
120       .SelText = String(420, "-") & vbCrLf

130       .SelFontSize = 10
140       .SelBold = False

150       If Len(SampleDate) > 10 Then
160           .SelText = Left$(" Sample Date : " & Format$(SampleDate, "dd/MM/yy hh:mm") & Space(30), 30)
170       Else
180           .SelText = Left$(" Sample Date : " & Format$(SampleDate, "dd/MM/yy") & Space(30), 30)
190       End If
200       If RP.Department = "P" Or RP.Department = "Y" Then
210           .SelText = Left$(" " & Space(30), 30)
220       Else

230           Rundate = Format(Rundate, "dd/MM/yyyy hh:mm")
240           If Right(Rundate, 5) = "00:00" Then
250               .SelText = Left$("Run Date : " & Format$(Rundate, "dd/mm/yy") & Space(30), 30)
260           Else
270               .SelText = Left$("Run Date : " & Format$(Rundate, "dd/mm/yy hh:mm") & Space(30), 30)
280           End If

290       End If
          'Zyam
          Dim qSampleID  As String
          qSampleID = RP.SampleID
          If RP.Department = "C" Then
                  'Zyam
                  Dim sqlAuthorised As String
                  Dim sqlusername As String
                  Dim tb As Recordset
                  Dim tb1 As Recordset
                  
                  sqlAuthorised = "SELECT username from Coagresults WHERE SampleID = '" & RP.SampleID & "' AND Username is not null"
                  Set tb = New Recordset
                  RecOpenClient 0, tb, sqlAuthorised
                  If tb!UserName = "HEM" Then
                    .SelText = " Authorised by: HemoHub AutoVal"
                  ElseIf tb!UserName = "CHA" Then
                    .SelText = " Authorised by: Charlotte Muldowney"
                  Else
                    
                    sqlusername = "SELECT Name from Users WHERE Code = '" & tb!UserName & "'"
                    Set tb1 = New Recordset
                    RecOpenClient 0, tb1, sqlusername
440                 .SelText = " Authorised by: " & tb1!Name
                  End If
          ElseIf RP.Department = "B" Then
                  .SelText = " Authorised by: " & getOperatorName("BioResults", qSampleID)
          ElseIf RP.Department = "H" Then
                  .SelText = " Authorised by: " & getOperatorName("HaemResults", qSampleID)
          ElseIf RP.Department = "I" Or RP.Department = "J" Then
                  .SelText = " Authorised by: " & getOperatorName("ImmResults", qSampleID)
          ElseIf RP.Department = "E" Then
                  .SelText = " Authorised by: " & getOperatorName("EndResults", qSampleID)
450       End If
          'Zyam
          '    sql = "SELECT Code FROM Users WHERE " & _
               '          "Code = '" & Initiator & "' " & _
               '          "OR Name = '" & Initiator & "'"
          '    Set tb = New Recordset
          '    RecOpenServer 0, tb, sql
          '    If tb.EOF Then
          '        .SelText = " Authorised by " & Initiator
          '    Else
          '        .SelText = " Authorised by " & tb!Code
          '    End If
310   End With

320   Exit Sub

PrintFooterA4RTB_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "modHeadFoot", "PrintFooterA4RTB", intEL, strES

End Sub

Public Sub PrintFooterMicroRTB(ByVal SampleDate As String, _
                               ByVal Rundate As String)

      Dim LineCount As Long

10    On Error GoTo PrintFooterMicroRTB_Error
20    With frmRichText.rtb
30        LineCount = SendMessage(.hWnd, EM_GETLINECOUNT, 0&, 0&)
40        .SelFontName = "Courier New"
50        .SelColor = vbBlack

60        If UCase(pForcePrintTo) = "FAX" Then
70            Do While CrCnt < 34
80                .SelText = vbCrLf
90                CrCnt = CrCnt + 1
100           Loop
110       Else
120           Do While LineCount < 39
130               .SelText = vbCrLf
140               LineCount = SendMessage(.hWnd, EM_GETLINECOUNT, 0&, 0&)
150           Loop
160       End If

170       .SelFontSize = 4
180       .SelText = vbCrLf

190       .SelFontSize = 2
200       If RP.FaxNumber <> "" Then
210           .SelText = String(388, "-") & vbCrLf
220       Else
230           .SelText = String(420, "-") & vbCrLf
240       End If

250       .SelFontSize = 8
260       .SelBold = False

270       If Format$(SampleDate, "HH:mm") <> "00:00" Then
280           .SelText = Left$(" Sample Date : " & Format$(SampleDate, "dd/MM/yy hh:mm") & Space(30), 30)
290       Else
300           .SelText = Left$(" Sample Date : " & Format$(SampleDate, "dd/MM/yy") & Space(30), 30)
310       End If
320       If RP.Department = "P" Or RP.Department = "Y" Then
330           .SelText = Left$(" " & Space(30), 30)
340       Else

350           Rundate = Format(Rundate, "dd/MM/yyyy hh:mm")
360           If Right(Rundate, 5) = "00:00" Then
370               .SelText = Left$("Run Date : " & Format$(Rundate, "dd/mm/yy") & Space(30), 30)
380           Else
390               .SelText = Left$("Run Date : " & Format$(Rundate, "dd/mm/yy hh:mm") & Space(30), 30)
400           End If

410       End If

420       .SelText = " Validated by " & GetAuthorisedBy(MicroValidatedBy(RP.SampleID))
          'If RP.PrintAction = "PrintSaveFinal" Then
430       If GetAuthorisedStatus(RP.SampleID) = "1" Then
440           .SelText = vbCrLf
              'Zyam
              Dim sqlAuthorised As String
              Dim tb As Recordset
              sqlAuthorised = "SELECT username from Demographics WHERE SampleID = '" & RP.SampleID & "'"
              Set tb = New Recordset
              RecOpenClient 0, tb, sqlAuthorised
450           .SelText = " Authorised by: " & tb!UserName
              'Zyam
460       End If

470   End With

480   Exit Sub

PrintFooterMicroRTB_Error:

      Dim strES As String
      Dim intEL As Integer

490   intEL = Erl
500   strES = Err.Description
510   LogError "modHeadFoot", "PrintFooterMicroRTB", intEL, strES

End Sub
Public Sub PrintHeadingRTB(Optional ByVal PageNo As String = "")

          Dim SampleID As String
          Dim Dept As String
          Dim Name As String
          Dim Ward As String
          Dim DoB As String
          Dim Chart As String
          Dim Clinician As String
          Dim Address0 As String
          Dim Address1 As String
          Dim GP As String
          Dim Sex As String
          Dim Hospital As String
          Dim SampleDate As String
          Dim RecDate As String
          Dim Rundate As String
          Dim GpClin As String
          Dim SampleType As String
          Dim AccreditationText As String
          Dim AccreditationText2 As String
          Dim DocumentNumber As String
          Dim AandE As String
          Dim ReportTitle As String
          Dim IsHaem As Boolean
          Dim IsBio As Boolean

10        On Error GoTo PrintHeadingRTB_Error

20        CrCnt = 0

30        IsHaem = False
32        IsBio = False
40        With udtHeading
50            SampleID = .SampleID
60            Dept = .Dept
70            Name = .Name
80            Ward = .Ward
90            DoB = .DoB
100           Chart = .Chart
110           Clinician = .Clinician
120           Address0 = .Address0
130           Address1 = .Address1
140           GP = .GP
150           Sex = .Sex
160           Hospital = .Hospital
170           SampleDate = .SampleDate
180           RecDate = .RecDate
190           Rundate = .Rundate
200           GpClin = .GpClin
210           SampleType = .SampleType
220           DocumentNumber = .DocumentNo
230           AandE = .AandE
240       End With

250       If IsNumeric(SampleID) Then
260           If SampleID > SysOptSemenOffset(0) And SampleID < SysOptMicroOffset(0) Then    'Semen Analyses
270               SampleID = SampleID - SysOptSemenOffset(0)
280           ElseIf SampleID > SysOptMicroOffset(0) Then        'micro external
290               SampleID = SampleID - SysOptMicroOffset(0)
300           End If
310       End If

320       With frmRichText
330           .rtb.Text = ""
              '    .Font.Name = "Courier New"
              '    .Font.Size = 10

340           If DocumentNumber <> "" Then
350               PrintTextRTB .rtb, FormatString(DocumentNumber, 84, , AlignRight) & vbCrLf, 10, , , , vbRed
                  '.SelColor = vbRed
                  '.SelText = Right(Space(84) & DocumentNumber, 84) & vbCrLf
360               CrCnt = CrCnt + 1
370           End If



'                  .SelFontName = "Courier New"
'                  .SelFontSize = 12
              '    .SelBold = True
              '    .SelItalic = True
              '    .SelColor = vbBlack

              '.SelText = Left("     Regional Hospital " & StrConv(HospName(0), vbProperCase) & "." & Space(25), 25)
'+++ Junaid
'380           ReportTitle = Left("Regional Hospital " & StrConv(HospName(0), vbProperCase) & "." & Space(25), 25)
380           ReportTitle = "   MRH @ " & StrConv(HospName(0), vbProperCase) & "." & Space(1)
'--- Junaid
390           Select Case Left(Dept, 4)
              Case "Haem"
400               If SysOptHaemAddress(0) <> "" Then
410                   ReportTitle = ReportTitle & Left(SysOptHaemAddress(0) & " Phone " & SysOptHaemPhone(0) & Space(38), 38)
            Else:
420                   ReportTitle = ReportTitle & Left("Haematology Dept" & " Phone " & SysOptHaemPhone(0) & Space(38), 38)
430               End If
440           Case "Bioc"
450               If SysOptBioAddress(0) <> "" Then
460                   ReportTitle = ReportTitle & Left(SysOptBioAddress(0) & " Phone " & SysOptBioPhone(0) & Space(36), 36)
            Else:
470                   ReportTitle = ReportTitle & Left("Biochemistry Dept" & " Phone " & SysOptBioPhone(0) & Space(36), 36)
480               End If
490           Case "Path"
500               If SysOptBioPhone(0) <> "" Then
510                   ReportTitle = ReportTitle & Left("Pathology Lab Phone " & SysOptBioPhone(0) & Space(36), 36)
520               End If
530           Case "Bloo"
540               ReportTitle = ReportTitle & " Phone 38830"
550           Case "Endo"
560               If SysOptEndAddress(0) <> "" Then
570                   ReportTitle = ReportTitle & Left(SysOptEndAddress(0) & " Phone " & SysOptEndPhone(0) & Space(36), 36)
            Else:
580                   ReportTitle = ReportTitle & Left("Endocrinology Dept" & " Phone " & SysOptEndPhone(0) & Space(36), 36)
590               End If
600           Case "Immu"
610               If SysOptImmAddress(0) <> "" Then
620                   ReportTitle = ReportTitle & Left(SysOptImmAddress(0) & " Phone " & SysOptImmPhone(0) & Space(36), 36)
            Else:
630                   ReportTitle = ReportTitle & Left("Immunology Dept" & " Phone " & SysOptImmPhone(0) & Space(36), 36)
640               End If
650           Case "Coag"
660               If SysOptCoagAddress(0) <> "" Then
670                   ReportTitle = ReportTitle & Left(SysOptCoagAddress(0) & " Phone " & SysOptCoagPhone(0) & Space(36), 36)
            Else:
680                   ReportTitle = ReportTitle & Left("Coagulation Dept" & " Phone " & SysOptCoagPhone(0) & Space(36), 36)
690               End If
700           Case "Micr"
710               If SysOptMicroAddress(0) <> "" Then
720                   ReportTitle = ReportTitle & Left(SysOptMicroAddress(0) & " Phone " & SysOptMicroPhone(0) & Space(36), 36)
            Else:
730                   ReportTitle = ReportTitle & Left("Microbiology Dept" & " Phone " & SysOptMicroPhone(0) & Space(36), 36)
740               End If
750           Case "Exte"
760               If SysOptExtAddress(0) <> "" Then
770                   ReportTitle = ReportTitle & Left(SysOptExtAddress(0) & " Phone " & SysOptExtPhone(0) & Space(36), 36)
            Else:
780                   ReportTitle = ReportTitle & Left("External Requests " & " Phone " & SysOptExtPhone(0) & Space(36), 36)
790               End If
800           Case "Hist"
810               If SysOptHistoAddress(0) <> "" Then
820                   ReportTitle = ReportTitle & SysOptHistoAddress(0)
830               Else
840                   ReportTitle = ReportTitle & "Histology Dept"
850               End If
860               If SysOptHistoPhone(0) <> "" Then
870                   ReportTitle = ReportTitle & " Phone " & SysOptHistoPhone(0)
880               End If
890           Case "Cyto"
900               If SysOptCytoAddress(0) <> "" Then
910                   ReportTitle = ReportTitle & SysOptCytoAddress(0)
920               Else
930                   ReportTitle = ReportTitle & "Cytology Dept"
940               End If
950               If SysOptCytoPhone(0) <> "" Then
960                   ReportTitle = ReportTitle & " Phone " & SysOptCytoPhone(0)
970               End If
980           Case Else
990               ReportTitle = ReportTitle & "Laboratory Phone : " & SysOptBioPhone(0)
1000          End Select
1010          PrintTextRTB .rtb, ReportTitle & vbCrLf, 14, True, True
              '.SelText = vbCrLf
1020          CrCnt = CrCnt + 1

              '.SelBold = False
              '.SelItalic = False

              '.SelFontSize = 2
              '.SelText = String(420, "-") & vbCrLf
1030          PrintTextRTB .rtb, String(420, "-") & vbCrLf, 2
1040          CrCnt = CrCnt + 1


              'QMS Ref 818255 PRINT ACCREDITATION STATEMENT
1050          Select Case RP.Department
              Case "H":
1060              AccreditationText = GetOptionSetting("HaemAccreditation", "")
1070              AccreditationText2 = GetOptionSetting("HAEMAccreditation2", "")
1080              IsHaem = True
1090          Case "B":
1100              AccreditationText = GetOptionSetting("BioAccreditation", "")
1102              AccreditationText2 = GetOptionSetting("BioAccreditation2", "")
1104              IsBio = True
1110          Case "C":
1120              AccreditationText = GetOptionSetting("CoagAccreditation", "")
1130          Case "E":
1140              AccreditationText = GetOptionSetting("EndAccreditation", "")
1150          Case "Q":
1160              AccreditationText = GetOptionSetting("BgaAccreditation", "")
1170          Case "I", "J":
1180              AccreditationText = GetOptionSetting("ImmAccreditation", "")


1190          End Select

1200          If AccreditationText <> "" Then

1210              PrintTextRTB frmRichText.rtb, FormatString(AccreditationText, 108, , AlignCenter) & vbCrLf, 8, , , , vbRed
1220              CrCnt = CrCnt + 1
                  '        .SelFontSize = 2
                  '        .SelText = String(420, "-") & vbCrLf
                  '        CrCnt = CrCnt + 1
1230          End If



              '.SelFontName = "Courier New"
              '.SelFontSize = 11

              'line 1
              '    .SelText = "     NAME: "
              '    .SelBold = True
              '    .SelText = Left$(StrConv(Left(Name, 45), vbUpperCase) & Space(45), 45)
              '    .SelBold = False
1240          PrintTextRTB .rtb, FormatString("NAME:", 9, " ", AlignRight), 11
1250          PrintTextRTB .rtb, FormatString(StrConv(Left(Name, 45), vbUpperCase), 45, " "), 11, True

              '    .SelText = " LAB NO.: "
              '    .SelBold = True
              '    .SelText = SampleID
              '    .SelBold = False
1260          PrintTextRTB .rtb, FormatString("LAB NO:", 9, " ", AlignRight), 11
1270          PrintTextRTB .rtb, RP.SampleID, 11, True

1280          PrintTextRTB .rtb, vbCrLf
1290          CrCnt = CrCnt + 1
              '.SelText = vbCrLf


              '    .SelText = "  CHART #: "
              '    .SelBold = True
              '    'QMS Ref 818219
              '    If Left$(Trim(Chart), 1) = "T" Or Left$(Trim(Chart), 1) = "P" Or Left$(Trim(Chart), 1) = "M" Then
              '        Chart = Mid(Trim(Chart), 2, Len(Trim(Chart)))
              '    End If
              '    .SelText = Left$(Trim(Chart) & Space(23), 23)
              '    .SelBold = False

              'chart
1300          If Left$(Trim(Chart), 1) = "T" Or Left$(Trim(Chart), 1) = "P" Or Left$(Trim(Chart), 1) = "M" Then
1310              Chart = Mid(Trim(Chart), 2, Len(Trim(Chart)))
1320          End If
1330          PrintTextRTB .rtb, FormatString("CHART #:", 9, " ", AlignRight), 11
1340          PrintTextRTB .rtb, FormatString(Chart, 9, " "), 11, True

              'AndE
1350          If GetOptionSetting("PrintAandE", "0") = 1 Then
1360              PrintTextRTB .rtb, FormatString("AandE:", 6, " ", AlignRight), 11
1370              PrintTextRTB .rtb, FormatString(AandE, 8, " "), 11, True
1380          Else
1390              PrintTextRTB .rtb, FormatString(" ", 15, " ", AlignRight), 11
                  'PrintTextRTB .rtb, FormatString(Space(6), 8, " "), 11, True
1400          End If

              'dob
1410          PrintTextRTB .rtb, FormatString("DOB:", 4, " ", AlignRight), 11
1420          PrintTextRTB .rtb, FormatString(Format(DoB, "dd/mm/yyyy"), 19, " "), 11, True

              'sex
1430          PrintTextRTB .rtb, FormatString("SEX:", 4, " ", AlignRight), 11
1440          If Sex = "M" Then
1450              PrintTextRTB .rtb, FormatString("MALE", 6), 11, True

1460          ElseIf Sex = "F" Then
1470              PrintTextRTB .rtb, FormatString("FEMALE", 6), 11, True
1480          Else
1490              PrintTextRTB .rtb, FormatString("", 6), 11, True
1500          End If

              'PrintTextRTB .rtb, FormatString(IIf(Sex = "M", "MALE", "FEMALE"), 6), 11, True

1510          PrintTextRTB .rtb, vbCrLf
1520          CrCnt = CrCnt + 1



              '    .SelText = "      DOB: "
              '    .SelBold = True
              '    .SelText = Left(Format(DoB, "dd/mm/yyyy") & Space(10), 10)
              '    .SelBold = False
              '
              '    .SelText = "      SEX: "
              '    .SelFontSize = 11
              '    .SelBold = True
              '    If Sex = "F" Then Sex = "FEMALE"
              '    If Sex = "M" Then Sex = "MALE"
              '    .SelText = Sex
              '    .SelBold = False
              '
              '    .SelText = vbCrLf
              '
              '    .SelBold = False

1530          If GpClin = "CLIN" Then
                  '.SelBold = True
                  '.SelText = " Copy CON: "
1540              PrintTextRTB .rtb, FormatString("Copy CON:", 9, " ", AlignRight), 11, True
1550          Else
                  '.SelText = "     CONS: "
1560              PrintTextRTB .rtb, FormatString("CON:", 9, " ", AlignRight), 11
1570          End If
1580          PrintTextRTB .rtb, FormatString(Initial2Upper(Clinician), 24, " "), 11, True
              '    .SelBold = True
              '    .SelText = Left$(Initial2Upper(Clinician) & Space(22), 22)
              '    .SelBold = False
1590          PrintTextRTB .rtb, FormatString("WARD:", 5, " ", AlignRight), 11
1600          PrintTextRTB .rtb, FormatString(UCase(Trim(Ward)), 37), 11, True

1610          PrintTextRTB .rtb, vbCrLf
1620          CrCnt = CrCnt + 1
              '    .SelText = "      WARD: "
              '    .SelBold = True
              '    .SelText = UCase(Trim(Ward))
              '
              '    .SelBold = False
              '    .SelText = vbCrLf


              '    .SelText = "     HOSP: "
              '    .SelBold = True
              '    .SelText = Left(Initial2Upper(Hospital) & Space(22), 22)
              '    .SelBold = False

1630          PrintTextRTB .rtb, FormatString("HOSP:", 9, " ", AlignRight), 11
1640          PrintTextRTB .rtb, FormatString(Initial2Upper(Hospital), 22), 11, True




1650          If GpClin = "GP" Then
                  '.SelBold = False
                  '.SelText = "   Copy Gp: "
1660              PrintTextRTB .rtb, FormatString("COPY GP:", 8, " ", AlignRight), 11
1670          Else
                  '.SelText = "        GP: "
1680              PrintTextRTB .rtb, FormatString("GP:", 8, " ", AlignRight), 11
1690          End If
1700          PrintTextRTB .rtb, FormatString(Initial2Upper(GP), 37), 11, True

1710          PrintTextRTB .rtb, vbCrLf
1720          CrCnt = CrCnt + 1

              '    .SelBold = True
              '    .SelText = Initial2Upper(GP)
              '    .SelBold = False
              '    .SelText = vbCrLf


1730          PrintTextRTB .rtb, FormatString("PT ADDR:", 9, " ", AlignRight), 11
1740          PrintTextRTB .rtb, FormatString(Trim$(Address0) & " " & Trim$(Address1), 66), 11, True

1750          PrintTextRTB .rtb, vbCrLf
1760          CrCnt = CrCnt + 1

              '    .SelText = "  PT ADDR: "
              '    .SelText = Left(Trim$(Address0) & " " & Trim$(Address1) & Space(69), 69)
              '    .SelText = vbCrLf
              '    CrCnt = CrCnt + 1

1770          PrintTextRTB .rtb, FormatString("GP ADDR:", 9, " ", AlignRight), 11
1780          PrintTextRTB .rtb, FormatString(GetGPAddress(GP), 66), 11, True

1790          PrintTextRTB .rtb, vbCrLf
1800          CrCnt = CrCnt + 1

              '    .SelText = "  GP ADDR: "
              '    .SelText = Left(GetGPAddress(GP) & Space(69), 69)
              '    .SelText = vbCrLf
              '    CrCnt = CrCnt + 1

1810          PrintTextRTB .rtb, FormatString("SAMPLE TYPE:", 12, " ", AlignRight), 11
1820          PrintTextRTB .rtb, FormatString(GetSampleType(RP.Department, RP.SampleID), 30), 11, True

              '    .SelText = "  Sample Type: "
              '    .SelText = Left(GetSampleType(RP.Department, RP.SampleID) & Space(30), 30)


1830          If udtHeading.Dept = "Draft Haem" Then
                  '        .SelColor = vbRed
                  '        .SelFontSize = 8
                  '        .SelText = Left(udtHeading.Notes & Space(49), 49)
                  '        .SelColor = vbBlack
1840              PrintTextRTB .rtb, FormatString(udtHeading.Notes, 49), 8, , , , vbRed
1850          End If

1860          PrintTextRTB .rtb, vbCrLf
1870          CrCnt = CrCnt + 1

              '    .SelText = vbCrLf
              '    CrCnt = CrCnt + 1

1880          If RP.SendCopyTo <> "" Then
1890              PrintTextRTB .rtb, FormatString("This is a COPY Report for Attention of ", 40, " ", AlignRight), 11
1900              PrintTextRTB .rtb, RP.SendCopyTo, 11, True

1910              PrintTextRTB .rtb, vbCrLf
1920              CrCnt = CrCnt + 1
                  '        .SelBold = False
                  '        .SelText = "This is a COPY Report for Attention of "
                  '        .SelBold = True
                  '        .SelText = RP.SendCopyTo & vbCrLf
                  '        .SelBold = False

1930          End If



1940          PrintTextRTB .rtb, String(420, "-") & vbCrLf, 2
1950          CrCnt = CrCnt + 1

              '.SelFontSize = 2
              '.SelText = String(420, "-") & vbCrLf

              '.SelFontSize = 8
              '    If SampleType <> "" Then
              '        .SelText = Left("Sample Type: " & ListText("ST", SampleType) & Space(20), 20)
              '    Else
              '        .SelText = Left(" " & Space(20), 20)
              '    End If
              '.SelText = Space(20) & "Printed on " & Format$(Now, "dd/mm/yy") & " at  " & Format$(Now, "hh:mm")
1960          PrintTextRTB .rtb, FormatString("Printed on " & Format$(Now, "dd/mm/yy") & " at  " & Format$(Now, "hh:mm"), 40, "        ", AlignRight), 8

1970          RecDate = Format(RecDate, "dd/MM/yyyy hh:mm")
1980          If Right(RecDate, 5) = "00:00" Then
                  '.SelText = Left("Received : " & Format(RecDate, "dd/MM/yyyy") & Space(35), 35)
1990              PrintTextRTB .rtb, FormatString("Received : " & Format(RecDate, "dd/MM/yyyy"), 30, , AlignLeft), 8
2000          Else
                  '.SelText = Left("Received : " & Format(RecDate, "dd/MM/yyyy hh:mm") & Space(35), 35)
2010              PrintTextRTB .rtb, FormatString("Received : " & Format(RecDate, "dd/MM/yyyy hh:mm"), 30, , AlignLeft), 8
2020          End If
2030          If PageNo <> "" Then
2040              PrintTextRTB .rtb, FormatString(PageNo, 20, , AlignRight), 8
2050          End If
              '.SelText = Left(" " & Space(22), 22)
              '.SelText = vbCrLf
2060          PrintTextRTB .rtb, vbCrLf
2070          CrCnt = CrCnt + 1

2080          If IsHaem And UCase(HospName(0)) = "PORTLAOISE" Then
2090              PrintTextRTB .rtb, FormatString(AccreditationText2, 120, vbCrLf, AlignLeft), 8
2100              CrCnt = CrCnt + 1
2110          End If
2112          If IsBio And UCase(HospName(0)) = "PORTLAOISE" Then
2114             PrintTextRTB .rtb, FormatString(AccreditationText2, 120, vbCrLf, AlignLeft), 8
2116             CrCnt = CrCnt + 1
2118          End If
              '.SelFontSize = 2
              '.SelText = String(420, "-") & vbCrLf
2120          PrintTextRTB .rtb, String(420, "-") & vbCrLf, 2
2130          CrCnt = CrCnt + 1


2140      End With

2150      Exit Sub

PrintHeadingRTB_Error:

          Dim strES As String
          Dim intEL As Integer

2160      intEL = Erl
2170      strES = Err.Description
2180      LogError "modHeadFoot", "PrintHeadingRTB", intEL, strES

End Sub
Public Sub PrintHeadingNew(ByVal PageNumber As Integer, _
                           ByVal TotalPages As Integer)

      Dim sql As String
      Dim tb As Recordset
      Dim SampleID As String
      Dim Dept As String
      Dim PatName As String
      Dim Ward As String
      Dim DoB As String
      Dim Chart As String
      Dim Clinician As String
      Dim Address0 As String
      Dim Address1 As String
      Dim GP As String
      Dim Sex As String
      Dim Hospital As String
      Dim SampleDate As String
      Dim RecDate As String
      Dim Rundate As String
      Dim GpClin As String
      Dim SampleType As String
      Dim DocumentNumber As String
      Dim AandE As String
      Dim AccreditationText As String

10    On Error GoTo PrintHeadingNew_Error




20    sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If Not tb.EOF Then
60        PatName = tb!PatName & ""
70        If IsDate(tb!DoB) Then
80            DoB = Format(tb!DoB, "dd/mmm/yyyy")
90        End If
100       Chart = tb!Chart & ""
110       AandE = tb!AandE & ""
120       Address0 = tb!Addr0 & ""
130       Address1 = tb!Addr1 & ""
140       Sex = tb!Sex & ""
150       Hospital = tb!Hospital & ""
160       SampleDate = tb!SampleDate & ""
170       Rundate = tb!Rundate & ""
180       RecDate = tb!RecDate & ""
190       Ward = RP.Ward
          'Ward = tb!Ward & ""
200       Clinician = RP.Clinician
          'Clinician = tb!Clinician
210       GP = RP.GP
220       Dept = udtHeading.Dept
          'GP = tb!GP
230       DocumentNumber = udtHeading.DocumentNo
240   End If

250   If RP.Department = "Z" Then
260       SampleType = GetSemenSampleType(RP.SampleID)
270   Else
280       SampleType = GetMicroSiteDetails(RP.SampleID)
290   End If

300   With frmRichText.rtb
310       .Text = ""

320       .Font.Name = "Courier New"
330       .Font.Size = 10


340       If DocumentNumber <> "" Then
350           .SelColor = vbRed
360           .SelText = Right(Space(84) & DocumentNumber, 84) & vbCrLf
370           CrCnt = CrCnt + 1
380       End If


390       .SelFontName = "Courier New"
400       .SelFontSize = 12
410       .SelBold = True
420       .SelItalic = True
430       .SelColor = vbBlack
'+++ Junaid
'440       .SelText = Left("     Regional Hospital " & StrConv(HospName(0), vbProperCase) & "." & Space(25), 25)
440       .SelText = "   MRH @ " & StrConv(HospName(0), vbProperCase) & "." & Space(1)
'--- Junaid

450       Select Case Left(Dept, 4)
          Case "Haem"
460           If SysOptHaemAddress(0) <> "" Then
470               .SelText = Left(SysOptHaemAddress(0) & " Phone " & SysOptHaemPhone(0) & Space(38), 38)
480           Else
490               .SelText = Left("Haematology Dept" & " Phone " & SysOptHaemPhone(0) & Space(38), 38)
500           End If
510       Case "Bioc"
520           If SysOptBioAddress(0) <> "" Then
530               .SelText = Left(SysOptBioAddress(0) & " Phone " & SysOptBioPhone(0) & Space(36), 36)
  Else:
540               .SelText = Left("Biochemistry Dept" & " Phone " & SysOptBioPhone(0) & Space(36), 36)
550           End If
560       Case "Path"
570           If SysOptBioPhone(0) <> "" Then Printer.Print Left("Pathology Lab Phone " & SysOptBioPhone(0) & Space(36), 36)
580       Case "Bloo"
590           Printer.Print " Phone 38830";
600       Case "Endo"
610           If SysOptEndAddress(0) <> "" Then
620               .SelText = Left(SysOptEndAddress(0) & " Phone " & SysOptEndPhone(0) & Space(36), 36)
  Else:
630               .SelText = Left("Endocrinology Dept" & " Phone " & SysOptEndPhone(0) & Space(36), 36)
640           End If
650       Case "Immu"
660           If SysOptImmAddress(0) <> "" Then
670               .SelText = Left(SysOptImmAddress(0) & " Phone " & SysOptImmPhone(0) & Space(36), 36)
  Else:
680               .SelText = Left("Immunology Dept" & " Phone " & SysOptImmPhone(0) & Space(36), 36)
690           End If
700       Case "Coag"
710           If SysOptCoagAddress(0) <> "" Then
720               .SelText = Left(SysOptCoagAddress(0) & " Phone " & SysOptCoagPhone(0) & Space(36), 36)
  Else:
730               .SelText = Left("Coagulation Dept" & " Phone " & SysOptCoagPhone(0) & Space(36), 36)
740           End If
750       Case "Micr"
760           If SysOptMicroAddress(0) <> "" Then
770               .SelText = Left(SysOptMicroAddress(0) & " Phone " & SysOptMicroPhone(0) & Space(36), 36)
  Else:
780               .SelText = Left("Microbiology Dept" & " Phone " & SysOptMicroPhone(0) & Space(36), 36)
790           End If
800       Case "Exte"
810           If SysOptExtAddress(0) <> "" Then
820               .SelText = Left(SysOptExtAddress(0) & " Phone " & SysOptExtPhone(0) & Space(36), 36)
  Else:
830               .SelText = Left("External Requests " & " Phone " & SysOptExtPhone(0) & Space(36), 36)
840           End If
850       Case "Hist"
860           If SysOptHistoAddress(0) <> "" Then
870               .SelText = SysOptHistoAddress(0)
880           Else
890               .SelText = "Histology Dept"
900           End If
910           If SysOptHistoPhone(0) <> "" Then
920               .SelText = " Phone " & SysOptHistoPhone(0)
930           End If
940       Case "Cyto"
950           If SysOptCytoAddress(0) <> "" Then
960               .SelText = SysOptCytoAddress(0)
970           Else
980               .SelText = "Cytology Dept"
990           End If
1000          If SysOptCytoPhone(0) <> "" Then
1010              .SelText = " Phone " & SysOptCytoPhone(0)
1020          End If
1030      Case Else
1040          .SelText = "Laboratory Phone : " & SysOptBioPhone(0)
1050      End Select

1060      .SelText = vbCrLf

1070      .SelBold = False
1080      .SelItalic = False

1090      .SelFontSize = 2
1100      If RP.FaxNumber <> "" Then
1110          .SelText = String(388, "-") & vbCrLf
1120      Else
1130          .SelText = String(420, "-") & vbCrLf
1140      End If

          'QMS Ref 818255 PRINT ACCREDITATION STATEMENT
1150      AccreditationText = GetOptionSetting("MicroAccreditation", "")
1160      If AccreditationText <> "" Then
1170          PrintTextRTB frmRichText.rtb, FormatString(AccreditationText, 108, , AlignCenter) & vbCrLf, 8, , , , vbRed
1180          CrCnt = CrCnt + 1
1190      End If

1200      If RP.FaxNumber <> "" Then
1210          .SelFontName = "Courier New"
1220          .SelFontSize = 9
1230      Else
1240          .SelFontName = "Courier New"
1250          .SelFontSize = 11
1260      End If


          'line 1
1270      .SelText = "     NAME: "
1280      .SelBold = True
1290      .SelText = Left$(StrConv(Left(PatName, 45), vbUpperCase) & Space(45), 45)
1300      .SelBold = False

1310      .SelText = " LAB NO.: "
1320      .SelBold = True
1330      .SelText = DisplaySampleID
1340      .SelBold = False

1350      .SelText = vbCrLf
1360      CrCnt = CrCnt + 1

1370      .SelText = "  CHART #: "
1380      .SelBold = True
1390      .SelText = Left$(Trim(Chart) & Space(14), 14)
1400      .SelBold = False

1410      If GetOptionSetting("PrintAandE", "0") = 1 Then
1420          .SelText = FormatString("AandE: ", 7)
1430          .SelBold = True
1440          .SelText = FormatString(Trim(AandE), 8)
1450          .SelBold = False
1460      Else
1470          .SelText = FormatString(" ", 7)
1480          .SelBold = True
1490          .SelText = FormatString(" ", 8)
1500          .SelBold = False
1510      End If


1520      .SelText = "DOB: "
1530      .SelBold = True
1540      .SelText = Left(Format(DoB, "dd/mm/yyyy") & Space(16), 16)
1550      .SelBold = False

1560      .SelText = "SEX: "
1570      .SelFontSize = 11
1580      .SelBold = True
1590      If Sex = "F" Then Sex = "FEMALE"
1600      If Sex = "M" Then Sex = "MALE"
1610      .SelText = Sex
1620      .SelBold = False

1630      .SelText = vbCrLf
1640      CrCnt = CrCnt + 1

1650      .SelBold = False

1660      If GpClin = "CLIN" Then
1670          .SelBold = True
1680          .SelText = " Copy CON: "
1690      Else
1700          .SelText = "     CONS: "
1710      End If
1720      .SelBold = True
1730      .SelText = Left$(Initial2Upper(Clinician) & Space(22), 22)
1740      .SelBold = False

1750      .SelText = "      WARD: "
1760      .SelBold = True
1770      .SelText = UCase(Trim(Ward))

1780      .SelBold = False
1790      .SelText = vbCrLf
1800      CrCnt = CrCnt + 1

1810      .SelText = "     HOSP: "
1820      .SelBold = True
1830      .SelText = Left(Initial2Upper(Hospital) & Space(22), 22)
1840      .SelBold = False

1850      If GpClin = "GP" Then
1860          .SelBold = False
1870          .SelText = "   Copy Gp: "
1880      Else
1890          .SelText = "        GP: "
1900      End If
1910      .SelBold = True
1920      .SelText = Initial2Upper(GP)
1930      .SelBold = False
1940      .SelText = vbCrLf
1950      CrCnt = CrCnt + 1




1960      .SelText = "  PT ADDR: "
1970      .SelText = Left(Trim$(Address0) & " " & Trim$(Address1) & Space(69), 69)
1980      .SelText = vbCrLf
1990      CrCnt = CrCnt + 1

2000      .SelText = "  GP ADDR: "
2010      .SelText = Left(GetGPAddress(GP) & Space(69), 69)
2020      .SelText = vbCrLf

2030      .SelText = "  Sample Type: "
2040      .SelBold = True
2050      .SelText = Left(SampleType & Space(60), 60)
2060      .SelBold = False



2070      .SelText = vbCrLf

2080      If RP.SendCopyTo <> "" Then
2090          .SelBold = False
2100          .SelText = "  This is a COPY Report for Attention of "
2110          .SelBold = True
2120          .SelText = RP.SendCopyTo & vbCrLf
2130          .SelBold = False
2140      End If
2150      .SelFontSize = 2
2160      If RP.FaxNumber <> "" Then
2170          .SelText = String(388, "-") & vbCrLf
2180      Else
2190          .SelText = String(420, "-") & vbCrLf
2200      End If

2210      If RP.FaxNumber <> "" Then
2220          .SelFontSize = 7
2230      Else
2240          .SelFontSize = 8
2250      End If
2260      .SelText = Left(" " & Space(20), 20)

2270      .SelText = "     Printed on " & Format$(Now, "dd/mm/yy") & " at  " & Format$(Now, "hh:mm")
2280      RecDate = Format(RecDate, "dd/MM/yyyy hh:mm")
2290      If Right(RecDate, 5) = "00:00" Then
2300          .SelText = Left("       Received : " & Format(RecDate, "dd/MM/yyyy") & Space(35), 35)
2310      Else
2320          .SelText = Left("       Received : " & Format(RecDate, "dd/MM/yyyy hh:mm") & Space(35), 35)
2330      End If
2340      .SelText = Space(5)
2350      .SelFontSize = 8
2360      .SelText = "Page " & Format$(PageNumber) & " of " & Format$(TotalPages)
2370      .SelText = vbCrLf
2380      .SelFontSize = 2
2390      If RP.FaxNumber <> "" Then
2400          .SelText = String(388, "-") & vbCrLf
2410      Else
2420          .SelText = String(420, "-") & vbCrLf
2430      End If

2440  End With

2450  Exit Sub

PrintHeadingNew_Error:

      Dim strES As String
      Dim intEL As Integer

2460  intEL = Erl
2470  strES = Err.Description
2480  LogError "modHeadFoot", "PrintHeadingNew", intEL, strES, sql

End Sub

Public Sub PrintFooterRTBFax(ByVal Initiator As String, _
                             ByVal SampleDate As String, _
                             ByVal Rundate As String)

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo PrintFooterRTBFax_Error

20    With frmRichText.rtb
30        .SelFontName = "Courier New"
40        .SelColor = vbBlack

50        .SelFontSize = 4
60        .SelText = vbCrLf

70        .SelFontSize = 2
80        .SelText = String(280, "-") & vbCrLf

90        .SelFontSize = 7
100       .SelBold = False

110       .SelText = Left$("     Sample Date : " & Format$(SampleDate, "dd/MM/yy") & Space(30), 30)

120       Rundate = Format(Rundate, "dd/MM/yyyy hh:mm")
130       If Right(Rundate, 5) = "00:00" Then
140           .SelText = Left$("Run Date : " & Format$(Rundate, "dd/mm/yy") & Space(30), 30)
150       Else
160           .SelText = Left$("Run Date : " & Format$(Rundate, "dd/mm/yy hh:mm") & Space(30), 30)
170       End If

180       sql = "SELECT * FROM users WHERE code = '" & Initiator & "' or name = '" & Initiator & "'"
190       Set tb = New Recordset
200       RecOpenServer 0, tb, sql
210       If tb.EOF Then
220           .SelText = " Faxed by " & Initiator
230       Else
240           .SelText = " Faxed by " & tb!Code
250       End If
260   End With

270   Exit Sub

PrintFooterRTBFax_Error:

      Dim strES As String
      Dim intEL As Integer

280   intEL = Erl
290   strES = Err.Description
300   LogError "modHeadFoot", "PrintFooterRTBFax", intEL, strES, sql

End Sub

Public Sub PrintHeadingRTBFax()

      Dim SampleID As String
      Dim Dept As String
      Dim Name As String
      Dim Ward As String
      Dim DoB As String
      Dim Chart As String
      Dim Clinician As String
      Dim Address0 As String
      Dim Address1 As String
      Dim GP As String
      Dim Sex As String
      Dim Hospital As String
      Dim SampleDate As String
      Dim RecDate As String
      Dim Rundate As String
      Dim GpClin As String
      Dim SampleType As String

10    On Error GoTo PrintHeadingRTBFax_Error

20    CrCnt = 0

30    With udtHeading
40        SampleID = .SampleID
50        Dept = .Dept
60        Name = .Name
70        Ward = .Ward
80        DoB = .DoB
90        Chart = .Chart
100       Clinician = .Clinician
110       Address0 = .Address0
120       Address1 = .Address1
130       GP = .GP
140       Sex = .Sex
150       Hospital = .Hospital
160       SampleDate = .SampleDate
170       RecDate = .RecDate
180       Rundate = .Rundate
190       GpClin = .GpClin
200       SampleType = .SampleType
210   End With

220   With frmRichText.rtb
230       .Text = ""

240       .SelFontName = "Courier New"
250       .SelFontSize = 12
260       .SelBold = True
270       .SelItalic = True
280       .SelColor = vbBlack
'+++ Junaid
'290       .SelText = Left("Regional Hospital " & StrConv(HospName(0), vbProperCase) & "." & Space(20), 20)
290       .SelText = "MRH @ " & StrConv(HospName(0), vbProperCase) & "." & Space(10)
'--- Junaid
300       Select Case Left(Dept, 4)
          Case "Haem"
310           If SysOptHaemAddress(0) <> "" Then
320               .SelText = Left(SysOptHaemAddress(0) & " Phone " & SysOptHaemPhone(0) & Space(38), 38)
  Else:
330               .SelText = Left("Haematology Dept" & " Phone " & SysOptHaemPhone(0) & Space(38), 38)
340           End If
350       Case "Bioc"
360           If SysOptBioAddress(0) <> "" Then
370               .SelText = Left(SysOptBioAddress(0) & " Phone " & SysOptBioPhone(0) & Space(36), 36)
  Else:
380               .SelText = Left("Biochemistry Dept" & " Phone " & SysOptBioPhone(0) & Space(36), 36)
390           End If
400       Case "Path"
410           If SysOptBioPhone(0) <> "" Then .SelText = Left("Pathology Lab Phone " & SysOptBioPhone(0) & Space(36), 36)
420       Case "Bloo"
430           Printer.Print " Phone 38830";
440       Case "Endo"
450           If SysOptEndAddress(0) <> "" Then
460               .SelText = Left(SysOptEndAddress(0) & " Phone " & SysOptEndPhone(0) & Space(36), 36)
  Else:
470               .SelText = Left("Endocrinology Dept" & " Phone " & SysOptEndPhone(0) & Space(36), 36)
480           End If
490       Case "Immu"
500           If SysOptImmAddress(0) <> "" Then
510               .SelText = Left(SysOptImmAddress(0) & " Phone " & SysOptImmPhone(0) & Space(36), 36)
  Else:
520               .SelText = Left("Immunology Dept" & " Phone " & SysOptImmPhone(0) & Space(36), 36)
530           End If
540       Case "Coag"
550           If SysOptCoagAddress(0) <> "" Then
560               .SelText = Left(SysOptCoagAddress(0) & " Phone " & SysOptCoagPhone(0) & Space(36), 36)
  Else:
570               .SelText = Left("Coagulation Dept" & " Phone " & SysOptCoagPhone(0) & Space(36), 36)
580           End If
590       Case "Micr"
600           If SysOptMicroAddress(0) <> "" Then
610               .SelText = Left(SysOptMicroAddress(0) & " Phone " & SysOptMicroPhone(0) & Space(36), 36)
  Else:
620               .SelText = Left("Microbiology Dept" & " Phone " & SysOptMicroPhone(0) & Space(36), 36)
630           End If
640       Case "Exte"
650           If SysOptExtAddress(0) <> "" Then
660               .SelText = Left(SysOptExtAddress(0) & " Phone " & SysOptExtPhone(0) & Space(36), 36)
  Else:
670               .SelText = Left("External Requests " & " Phone " & SysOptExtPhone(0) & Space(36), 36)
680           End If
690       Case "Hist"
700           If SysOptHistoAddress(0) <> "" Then Printer.Print SysOptHistoAddress(0); Else Printer.Print "Histology Dept";
710           If SysOptHistoPhone(0) <> "" Then Printer.Print " Phone " & SysOptHistoPhone(0);
720       Case "Cyto"
730           If SysOptCytoAddress(0) <> "" Then Printer.Print SysOptCytoAddress(0); Else Printer.Print "Cytology Dept";
740           If SysOptCytoPhone(0) <> "" Then Printer.Print " Phone " & SysOptCytoPhone(0);
750       Case Else
760           Printer.Print "Laboratory Phone : " & SysOptBioPhone(0);
770       End Select

780       .SelText = vbCrLf
790       CrCnt = CrCnt + 1

800       .SelBold = False
810       .SelItalic = False

820       .SelFontSize = 2
830       .SelText = String(280, "-") & vbCrLf

840       .SelFontName = "Courier New"
850       .SelFontSize = 9

          'line 1
860       .SelText = "     NAME: "
870       .SelBold = True
880       .SelText = Left$(StrConv(Left(Name, 45), vbUpperCase) & Space(45), 45)
890       .SelBold = False

900       .SelText = " LAB NO.: "
910       .SelBold = True
920       .SelText = SampleID
930       .SelBold = False

940       .SelText = vbCrLf
950       CrCnt = CrCnt + 1

960       .SelFontSize = 9

970       .SelText = "  CHART #: "
980       .SelBold = True
990       .SelText = Left$(Trim(Chart) & Space(23), 23)
1000      .SelBold = False

1010      .SelText = "      DOB: "
1020      .SelBold = True
1030      .SelText = Left(Format(DoB, "dd/mm/yyyy") & Space(10), 10)
1040      .SelBold = False

1050      .SelText = "      SEX: "
1060      .SelFontSize = 11
1070      .SelBold = True
1080      If Sex = "F" Then Sex = "FEMALE"
1090      If Sex = "M" Then Sex = "MALE"
1100      .SelText = Sex
1110      .SelBold = False

1120      .SelText = vbCrLf
1130      CrCnt = CrCnt + 1
1140      .SelFontSize = 9

1150      .SelBold = False

1160      If GpClin = "CLIN" Then
1170          .SelBold = True
1180          .SelText = " Copy CON: "
1190      Else
1200          .SelText = "     CONS: "
1210      End If
1220      .SelBold = True
1230      .SelText = Left$(Initial2Upper(Clinician) & Space(22), 22)
1240      .SelBold = False

1250      .SelText = "      WARD: "
1260      .SelBold = True
1270      .SelText = UCase(Trim(Ward))

1280      .SelBold = False
1290      .SelText = vbCrLf
1300      CrCnt = CrCnt + 1
1310      .SelFontSize = 9

1320      .SelText = "     HOSP: "
1330      .SelBold = True
1340      .SelText = Left(Initial2Upper(Hospital) & Space(22), 22)
1350      .SelBold = False

1360      If GpClin = "GP" Then
1370          .SelBold = False
1380          .SelText = "   Copy Gp: "
1390      Else
1400          .SelText = "        GP: "
1410      End If
1420      .SelBold = True
1430      .SelText = Initial2Upper(GP)
1440      .SelBold = False
1450      .SelText = vbCrLf
1460      CrCnt = CrCnt + 1

1470      .SelFontSize = 9

          '.SelUnderline = True
1480      .SelText = "  ADDRESS: "
1490      .SelText = Left(Trim$(Address0) & " " & Trim$(Address1) & Space(69), 69)
1500      .SelText = vbCrLf
          '.SelUnderline = False
1510      CrCnt = CrCnt + 1

1520      .SelFontSize = 2
1530      .SelText = String(280, "-") & vbCrLf

1540      .SelFontSize = 7
1550      If SampleType <> "" Then
1560          .SelText = Left("Sample Type: " & ListText("ST", SampleType) & Space(20), 20)
1570      Else
1580          .SelText = Left(" " & Space(20), 20)
1590      End If
1600      .SelText = "     Printed on " & Format$(Now, "dd/mm/yy") & " at  " & Format$(Now, "hh:mm")

1610      RecDate = Format(RecDate, "dd/MM/yyyy hh:mm")
1620      If Right(RecDate, 5) = "00:00" Then
1630          .SelText = Left("       Received : " & Format(RecDate, "dd/MM/yyyy") & Space(35), 35)
1640      Else
1650          .SelText = Left("       Received : " & Format(RecDate, "dd/MM/yyyy hh:mm") & Space(35), 35)
1660      End If
1670      .SelText = Left(" " & Space(22), 22)
1680      .SelText = vbCrLf
1690      CrCnt = CrCnt + 2
1700      .SelFontSize = 2
1710      .SelText = String(280, "-") & vbCrLf
1720  End With

1730  Exit Sub

PrintHeadingRTBFax_Error:

      Dim strES As String
      Dim intEL As Integer

1740  intEL = Erl
1750  strES = Err.Description
1760  LogError "modHeadFoot", "PrintHeadingRTBFax", intEL, strES

End Sub

'To get the authorised Name for each department
'Zyam
'This work with EndResults, ImmResults, BioResults, HaemResults

Public Function getOperatorName(ByVal TableName As String, ByVal SampleID As String) As String
     Dim sqlAuthorised As String
     Dim tb As Recordset
     Dim tb1 As Recordset
     Dim sqlusername As String
     sqlAuthorised = "SELECT Operator from " & TableName & " WHERE SampleID = '" & SampleID & "' AND Operator is not null"
     Set tb = New Recordset
     RecOpenClient 0, tb, sqlAuthorised
     If tb!Operator = "HEM" Then
        getOperatorName = "HemoHub AutoVal"
     ElseIf tb!Operator = "CHA" Then
        getOperatorName = "Charlotte Muldowney"
     Else
     sqlusername = "SELECT Name from Users WHERE Code = '" & tb!Operator & "'"
     Set tb1 = New Recordset
     RecOpenClient 0, tb1, sqlusername
        If Not tb1.EOF Then
            getOperatorName = tb1!Name
        Else
            getOperatorName = "AutoVal"
        End If
     End If

End Function

'Zyam


