Attribute VB_Name = "modBloodgas"
Option Explicit

Public Sub PrintResultBloodGas()

      Dim f As Integer
      Dim tb As Recordset
      Dim tbH As Recordset
      Dim n As Integer
      Dim Sex As String
      Dim DoB As String
      Dim sql As String
      'Dim Cx As Comment
      'Dim Cxs As New Comments
      Dim OB As Observation
      Dim OBS As New Observations
10    ReDim Comments(1 To 4) As String
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
      Dim nrpH As String
      Dim nrPO2 As String
      Dim nrPCO2 As String
      Dim nrHCO3 As String
      Dim nrTotCO2 As String
      Dim nrBE As String
      Dim nrO2SAT As String
      Dim PrintTime As String
      Dim BGs As New BGAResults
      Dim BG As BGAResult
      Dim AuthorisedBy As String

20    On Error GoTo PrintResultBloodGas_Error

30    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

40    frmMain.gDiff.Rows = 2
50    frmMain.gDiff.AddItem ""
60    frmMain.gDiff.RemoveItem 1

70    sql = "SELECT * FROM Demographics WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
80    Set tb = New Recordset
90    RecOpenClient 0, tb, sql
100   If tb.EOF Then Exit Sub

110   If IsDate(tb!SampleDate) Then
120     SampleDate = Format(tb!SampleDate, "dd/mmm/yyyy hh:mm")
130   Else
140     SampleDate = ""
150   End If
160   If IsDate(tb!Rundate) Then
170     Rundate = Format(tb!Rundate, "dd/mmm/yyyy hh:mm")
180   Else
190     Rundate = ""
200   End If

210   sql = "SELECT * FROM bgaResults WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
220   Set tbH = New Recordset
230   RecOpenClient 0, tbH, sql
240   If tbH.EOF Then Exit Sub

250   AuthorisedBy = GetAuthorisedBy(tbH!Operator & "")

260   DoB = tb!DoB & ""

270   Select Case Left(UCase(tb!Sex & ""), 1)
        Case "M": Sex = "M"
280     Case "F": Sex = "F"
290     Case Else: Sex = ""
300   End Select

310   ClearUdtHeading
320   With udtHeading
330     .SampleID = RP.SampleID
340     .Dept = "Blood Gas"
350     .Name = tb!PatName & ""
360     .Ward = RP.Ward
370     .DoB = DoB
380     .Chart = tb!Chart & ""
390     .Clinician = RP.Clinician
400     .Address0 = tb!Addr0 & ""
410     .Address1 = tb!Addr1 & ""
420     .GP = RP.GP
430     .Sex = tb!Sex & ""
440     .Hospital = tb!Hospital & ""
450     .SampleDate = tb!SampleDate & ""
460     .RecDate = tb!RecDate & ""
470     .Rundate = tb!Rundate & ""
480     .GpClin = ""
490     .SampleType = ""
500     .AandE = tb!AandE & ""
510   End With

520   Set BG = BGs.LoadResults(RP.SampleID)
530   If BG Is Nothing Then Exit Sub

540   If RP.FaxNumber <> "" Then
550     PrintHeadingRTBFax
560   Else
570     PrintHeadingRTB
580   End If

590   sql = "SELECT * FROM BGDefinitions"
600   Set tb = New Recordset
610   RecOpenServer 0, tb, sql
620   If Not tb.EOF Then
630      nrpH = tb!pH & ""
640      nrPO2 = tb!PO2 & ""
650      nrPCO2 = tb!PCO2 & ""
660      nrHCO3 = tb!HCO3 & ""
670      nrTotCO2 = tb!TotCO2 & ""
680      nrBE = tb!BE & ""
690      nrO2SAT = tb!O2SAT & ""
700   End If

710   With frmRichText.rtb
720     .SelText = vbCrLf
730     .SelText = vbCrLf
740     .SelFontSize = 12
750     .SelText = Space(10) & "         pH : " & BG.pH & Space(35) & nrpH & vbCrLf
760     .SelText = Space(10) & "        PO2 : " & BG.PO2 & Space(35) & nrPO2 & vbCrLf
770     .SelText = Space(10) & "       PCO2 : " & BG.PCO2 & Space(35) & nrPCO2 & vbCrLf
780     .SelText = Space(10) & "       HCO3 : " & BG.HCO3 & Space(35) & nrHCO3 & vbCrLf
790     .SelText = Space(10) & "    Tot CO2 : " & BG.TotCO2 & Space(35) & nrTotCO2 & vbCrLf
800     .SelText = Space(10) & "         BE : " & BG.BE & Space(35) & nrBE & vbCrLf
810     .SelText = Space(10) & "      O2Sat : " & BG.O2SAT & Space(35) & nrO2SAT & vbCrLf
        
820     .SelText = vbCrLf

830     Set OBS = OBS.Load(RP.SampleID, "BloodGas", "Demographic")
840     If Not OBS Is Nothing Then
850       For Each OB In OBS
860           Select Case UCase$(OB.Discipline)
                  Case "BLOODGAS"
870                   FillCommentLines OB.Comment, 4, Comments(), 90
880                   For n = 1 To 4
890                     .SelText = Comments(n) & vbCrLf
900                   Next
910               Case "DEMOGRAPHIC"
920                   FillCommentLines OB.Comment, 2, Comments(), 97
930                   For n = 1 To 4
940                     .SelText = Comments(n) & vbCrLf
950                   Next
960           End Select
970       Next
980     End If
        
990     If Not IsDate(DoB) And Trim(Sex) = "" Then
1000      .SelColor = vbBlue
1010      .SelText = Space(24) & "No Sex/DoB given. Normal ranges may not be relevant"
1020    ElseIf Not IsDate(DoB) Then
1030      .SelColor = vbBlue
1040      .SelText = Space(24) & "No DoB given. Normal ranges may not be relevant"
1050    ElseIf Trim(Sex) = "" Then
1060      .SelColor = vbBlue
1070      .SelText = Space(24) & "No Sex given. Normal ranges may not be relevant"
1080    End If
1090    .SelText = vbCrLf
1100    .SelColor = vbBlack

1110    If RP.FaxNumber <> "" Then
1120      PrintFooterRTBFax RP.Initiator, SampleDate, Rundate
1130      f = FreeFile
1140      Open SysOptFax(0) & RP.SampleID & "BGA.doc" For Output As f
1150      .SelStart = 0
1160      Print #f, .TextRTF
1170      Close f
1180      SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "BGA.doc"
1190    Else
1200      PrintFooterRTB AuthorisedBy, SampleDate, Rundate
1210      .SelStart = 0
1220      .SelPrint Printer.hDC
1230    End If
        
1240    sql = "SELECT * FROM Reports WHERE 0 = 1"
1250    Set tb = New Recordset
1260    RecOpenServer 0, tb, sql
1270    tb.AddNew
1280    tb!SampleID = RP.SampleID
1290    tb!Name = udtHeading.Name
1300    tb!Dept = "G"
1310    tb!Initiator = RP.Initiator
1320    tb!PrintTime = PrintTime
1330    tb!RepNo = "0G" & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
1340    tb!PageNumber = 0
1350    tb!Report = .TextRTF
1360    tb!Printer = Printer.DeviceName
1370    tb.Update
1380  End With

1390  ResetPrinter

1400  sql = "Update bgaResults " & _
            "set Printed = 1, Valid = 1 " & _
            "WHERE SampleID = '" & RP.SampleID & "'"
1410  Cnxn(0).Execute sql

1420  Exit Sub

PrintResultBloodGas_Error:

      Dim strES As String
      Dim intEL As Integer

1430  intEL = Erl
1440  strES = Err.Description
1450  LogError "modBloodgas", "PrintResultBloodGas", intEL, strES, sql

End Sub
