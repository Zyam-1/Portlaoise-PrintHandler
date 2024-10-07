Attribute VB_Name = "modPrintMicroRTF"
Option Explicit

'A Rota/Adeno
'B Biochemistry
'C Coagulation
'D C and S
'E Endocrinology
'F FOB
'G C.diff
'H Haematology
'I Immunology
'M Micro
'O Ova/Parasites
'R Red Sub
'S ESR
'U Urine
'V RSV
'X External
'Y H.Pylori

Public Type OrgGroupRTF
    OrgGroup As String
    OrgName As String
    ShortName As String
    ReportName As String
    Qualifier As String

End Type

Public Type ABResultRTF
    AntibioticName As String
    AntibioticCode As String
    Result(1 To 8) As String
    Report(1 To 8) As Boolean
    RSI(1 To 8) As String
    CPO(1 To 8) As String
End Type

Public ValidatedByRTF As String

Public Type OrgGroup
    OrgGroup As String
    OrgName As String
    ShortName As String
    ReportName As String
    Qualifier As String
    NonReportable As Integer
End Type

Public Type ABResult
    AntibioticName As String
    AntibioticCode As String
    Result(1 To 8) As String
    Report(1 To 8) As Boolean
    RSI(1 To 8) As String
    CPO(1 To 8) As String
End Type

Public Type AntibioticPrintLine
    AntibioticName As String
    RSI(1 To 6) As String
End Type

Public ValidatedBy As String

Private Function CountLines(ByVal strIP As String) As Integer

10    ReDim Comments(1 To 5) As String
      Dim n As Integer

20    FillCommentLines strIP, 5, Comments()

30    For n = 5 To 1 Step -1
40        If Trim$(Comments(n)) <> "" Then
50            CountLines = n
60            Exit For
70        End If
80    Next

End Function

Private Function GetMicroscopyLineCount(ByVal SampleID As String) As Integer

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GetMicroscopyLineCount_Error

20    sql = "Select * from Urine where " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql

50    If tb.EOF Then
60        GetMicroscopyLineCount = 0
70    ElseIf Trim$(tb!Pregnancy & "") <> "" Then
80        GetMicroscopyLineCount = 1
90    ElseIf Trim$(tb!Bacteria & tb!WCC & tb!RCC & _
                   tb!Crystals & tb!Casts & _
                   tb!Misc0 & tb!Misc1 & tb!Misc2 & "") = "" Then
100       GetMicroscopyLineCount = 0
110   Else
120       GetMicroscopyLineCount = 3
130   End If

140   Exit Function

GetMicroscopyLineCount_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "modPrintMicro", "GetMicroscopyLineCount", intEL, strES, sql


End Function

Private Function GetIsolateCount(ByVal SampleID As String) As Integer

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo GetIsolateCount_Error

20    sql = "Select Count(DISTINCT IsolateNumber) as tot from Isolates where " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    GetIsolateCount = tb!Tot

60    Exit Function

GetIsolateCount_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "modPrintMicro", "GetIsolateCount", intEL, strES, sql


End Function

Private Function GetABCount(ByVal SampleID As String) _
        As Integer

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo GetABCount_Error

20    sql = "SELECT COUNT (DISTINCT AntibioticCode) AS Tot FROM Sensitivities WHERE " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "AND Report = 1"

30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    GetABCount = tb!Tot

60    Exit Function

GetABCount_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "modPrintMicro", "GetABCount", intEL, strES, sql

End Function


Private Function GetCommentLineCount(ByVal SampleID As String) As Integer

      Dim n As Integer
      Dim OB As Observation
      Dim OBS As New Observations

10    On Error GoTo GetCommentLineCount_Error

      'sql = "Select * from Comments where " & _
      '      "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "'"
      'Set tb = New Recordset
      'RecOpenClient 0, tb, sql
      'If tb.EOF Then
20    n = 0
30    Set OBS = OBS.Load(Val(SampleID) + SysOptMicroOffset(0), "MicroCS", "Demographic", "MicroConsultant", "MicroGeneral", "MicroCDiff")

40    If Not OBS Is Nothing Then
50        For Each OB In OBS
60            n = n + CountLines(OB.Comment)
70        Next
80    Else
90        GetCommentLineCount = 0
100   End If

110   GetCommentLineCount = n

120   Exit Function

GetCommentLineCount_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "modPrintMicro", "GetCommentLineCount", intEL, strES

End Function

Public Function GetMicroSite(ByVal SampleID As Double) As String

      Dim tb As Recordset
      Dim sql As String
      Dim RetVal As String

10    On Error GoTo GetMicroSite_Error

20    RetVal = ""

30    sql = "SELECT Site FROM MicroSiteDetails WHERE " & _
            "SampleID = '" & SampleID & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70        RetVal = UCase$(Trim$(tb!Site & ""))
80    End If

90    GetMicroSite = RetVal

100   Exit Function

GetMicroSite_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modPrintMicroRTF", "GetMicroSite", intEL, strES, sql

End Function

Public Function GetMicroSiteDetails(ByVal SampleID As Double) As String

      Dim tb As Recordset
      Dim sql As String
      Dim RetVal As String

10    On Error GoTo GetMicroSiteDetails_Error

20    RetVal = ""

30    sql = "SELECT Site, SiteDetails FROM MicroSiteDetails WHERE " & _
            "SampleID = '" & SampleID & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70        RetVal = UCase$(Trim$(tb!Site & "")) & " " & Trim$(tb!SiteDetails & "")
80    End If

90    GetMicroSiteDetails = RetVal

100   Exit Function

GetMicroSiteDetails_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modPrintMicroRTF", "GetMicroSiteDetails", intEL, strES, sql

End Function

Public Function GetSemenSampleType(ByVal SampleID As Double) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GetSemenSampleType_Error

20    sql = "Select SpecimenType From SemenResults Where SampleID = '%sampleid'"
30    sql = Replace(sql, "%sampleid", SampleID)

40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    If tb.EOF Then
70        GetSemenSampleType = ""
80    Else
90        GetSemenSampleType = tb!SpecimenType & ""
100   End If


110   Exit Function

GetSemenSampleType_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "modPrintMicroRTF", "GetSemenSampleType", intEL, strES, sql

End Function



Private Function GetMiscLineCount(ByVal SampleID As String) As Long

      'FOB+CDiff+Rota/Adeno+OP

      Dim sql As String
      Dim tb As Recordset
      Dim intCount As Integer

10    On Error GoTo GetMiscLineCount_Error

20    sql = "SELECT COUNT(OB0) + COUNT(OB1) + COUNT(OB2) AS FOB FROM Faeces WHERE " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "AND ((OB0 IS NOT NULL AND OB0 <> '') " & _
            "      OR (OB1 IS NOT NULL AND OB1 <> '') " & _
            "      OR (OB2 IS NOT NULL AND OB2 <> ''))"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    intCount = tb!FOB

60    sql = "Select ToxinAB, Rota, Adeno, Cryptosporidium, OP0, OP1, OP2, HPylori from Faeces where " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "'"
70    Set tb = New Recordset
80    RecOpenClient 0, tb, sql
90    If Not tb.EOF Then
100       If Trim$(tb!ToxinAB & "") <> "" Then
110           intCount = intCount + 1
120       End If
130       If Trim$(tb!Rota & "") <> "" Then
140           intCount = intCount + 1
150       End If
160       If Trim$(tb!Adeno & "") <> "" Then
170           intCount = intCount + 1
180       End If
190       If Trim$(tb!Cryptosporidium & "") <> "" Then
200           intCount = intCount + 1
210       End If
220       If Trim$(tb!OP0 & "") <> "" Then
230           intCount = intCount + 1
240       End If
250       If Trim$(tb!OP1 & "") <> "" Then
260           intCount = intCount + 1
270       End If
280       If Trim$(tb!OP2 & "") <> "" Then
290           intCount = intCount + 1
300       End If
310       If Trim$(tb!HPylori & "") <> "" Then
320           intCount = intCount + 1
330       End If
340   End If

350   sql = "SELECT SampleID FROM GenericResults WHERE " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' "
360   Set tb = New Recordset
370   RecOpenServer 0, tb, sql
380   If Not tb.EOF Then
390       intCount = intCount + 1
400   End If

410   GetMiscLineCount = intCount

420   Exit Function

GetMiscLineCount_Error:

      Dim strES As String
      Dim intEL As Integer

430   intEL = Erl
440   strES = Err.Description
450   LogError "modPrintMicro", "GetMiscLineCount", intEL, strES, sql


End Function

Private Function GetFluidCount(ByVal SampleID As String) As Long

      Dim sql As String
      Dim tb As Recordset
      Dim intCount As Integer

10    On Error GoTo GetFluidCount_Error

20    sql = "SELECT TOP 1 * FROM GenericResults WHERE " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "AND TestName LIKE 'CSFH%'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If Not tb.EOF Then
60        intCount = 6
70    End If

80    sql = "SELECT COUNT (*) AS Tot FROM GenericResults WHERE " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "AND ((TestName LIKE 'CSF%' " & _
            "AND TestName NOT LIKE 'CSFH%') OR TestName LIKE 'Fluid%')"
90    Set tb = New Recordset
100   RecOpenClient 0, tb, sql
110   If tb!Tot > 0 Then
120       intCount = intCount + tb!Tot + 2
130   End If

140   GetFluidCount = intCount

150   Exit Function

GetFluidCount_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "modPrintMicroRTF", "GetFluidCount", intEL, strES, sql

End Function


Private Function IsForcedTo(ByVal TrueOrFalse As String, _
                            ByVal ABName As String, _
                            ByVal SID As Long, _
                            ByVal Index As Integer) _
                            As Boolean

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo IsForcedTo_Error

20    sql = "Select * from ForcedABReport where " & _
            "SampleID = " & SID & " " & _
            "and ABName = '" & Trim$(ABName) & "' " & _
            "and Report = '" & IIf(TrueOrFalse = "Yes", "1", "0") & "' " & _
            "and [Index] = " & Index
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    IsForcedTo = Not tb.EOF

60    Exit Function

IsForcedTo_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "modPrintMicroRTF", "IsForcedTo", intEL, strES, sql

End Function


Public Sub LoadResultArray(ByVal SampleIDWithOffset As Double, _
                           ByRef ResultArray() As AntibioticPrintLine)

      Dim tb As Recordset
      Dim sql As String
      Dim Site As String
      Dim U As Integer
      Dim IsolateNumber As Integer

10    On Error GoTo LoadResultArray_Error

20    sql = "Select Site from MicroSiteDetails where " & _
            "SampleID = '" & SampleIDWithOffset & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60        Site = "Generic"
70    ElseIf tb!Site & "" = "" Then
80        Site = "Generic"
90    Else
100       Site = tb!Site
110   End If

120   sql = "SELECT  LTRIM(RTRIM(AntibioticName)) AntibioticName, " & _
            "CASE RSI WHEN 'R' THEN 'Resistant' " & _
            "         WHEN 'S' THEN 'Sensitive' " & _
            "         WHEN 'I' THEN 'Intermediate' " & _
            "         ELSE '' END RSI, " & _
            "IsolateNumber FROM Antibiotics A " & _
            "JOIN Sensitivities S ON " & _
            "A.AntibioticName = S.Antibiotic " & _
            "WHERE AntibioticName IN " & _
            "  (SELECT DISTINCT Antibiotic FROM Sensitivities WHERE " & _
            "   SampleID = '" & SampleIDWithOffset & "' " & _
            "   AND Report = 1 ) " & _
            "AND SampleID = '" & SampleIDWithOffset & "' " & _
            "ORDER BY A.ListOrder"
130   Set tb = New Recordset
140   RecOpenClient 0, tb, sql
150   U = UBound(ResultArray)
160   Do While Not tb.EOF
170       If ResultArray(U).AntibioticName <> tb!AntibioticName Then
180           U = UBound(ResultArray) + 1
190           ReDim Preserve ResultArray(0 To U)
200           ResultArray(U).AntibioticName = tb!AntibioticName
210       End If
220       IsolateNumber = tb!IsolateNumber
230       ResultArray(U).RSI(IsolateNumber) = tb!RSI & ""
240       tb.MoveNext
250   Loop

260   Exit Sub

LoadResultArray_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "modPrintMicroRTF", "LoadResultArray", intEL, strES, sql

End Sub


Private Function IsNegativeResults(ByVal SampleID As String) As Boolean

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo IsNegativeResults_Error

20    IsNegativeResults = False

30    sql = "Select OrganismGroup from Isolates where " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "'"
40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql
60    If Not tb.EOF Then

70        If UCase$(tb!OrganismGroup & "") = "NO GROWTH" Or _
             UCase$(tb!OrganismGroup & "") = "NEGATIVE RESULTS" Then

80            IsNegativeResults = True

90        End If

100   End If

110   Exit Function

IsNegativeResults_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "modPrintMicroRTF", "IsNegativeResults", intEL, strES, sql

End Function

Public Sub LoadResultArrayRTF(ByVal SampleIDWithOffset As Double, _
                              ByRef ResultArray() As ABResult)

      Dim tb As Recordset
      Dim tbR As Recordset
      Dim sql As String
      Dim Site As String
      Dim U As Integer
      Dim ReportThis As Boolean
      Dim NewABAdded As Boolean
      Dim IsolateNumber As Integer

10    On Error GoTo LoadResultArray_Error

20    sql = "Select Site from MicroSiteDetails where " & _
            "SampleID = '" & SampleIDWithOffset & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60        Site = "Generic"
70    ElseIf tb!Site & "" = "" Then
80        Site = "Generic"
90    Else
100       Site = tb!Site
110   End If

120   sql = "Select Code, AntibioticName, MAX(ListOrder) AS M from Antibiotics where " & _
            "Code in ( " & _
            "         Select distinct AntibioticCode from Sensitivities where " & _
            "         SampleID = '" & SampleIDWithOffset & "' and Report = 1 " & _
            "        ) " & _
            "GROUP BY Code, AntiBioticName Order by M"

130   Set tb = New Recordset
140   RecOpenClient 0, tb, sql
150   Do While Not tb.EOF
160       Debug.Print tb!AntibioticName
170       sql = "Select * from Sensitivities where " & _
                "AntibioticCode = '" & tb!Code & "' " & _
                "and SampleID = " & SampleIDWithOffset
180       Set tbR = New Recordset
190       RecOpenServer 0, tbR, sql
200       NewABAdded = False
210       Do While Not tbR.EOF
220           ReportThis = False
230           If Not IsForcedTo("No", tb!AntibioticName, SampleIDWithOffset, tbR!IsolateNumber) Then
240               ReportThis = True
250           End If
              '    Else
              '      If IsForcedTo("Yes", tb!AntibioticName, SampleIDWithOffset, tbR!IsolateNumber) Then
              '        ReportThis = True
              '      End If
              '    End If
260           If ReportThis Then
270               If Not NewABAdded Then
280                   U = UBound(ResultArray) + 1
290                   ReDim Preserve ResultArray(0 To U)
300                   ResultArray(U).AntibioticCode = tb!Code
310                   ResultArray(U).AntibioticName = Trim$(tb!AntibioticName)
320                   NewABAdded = True
330               End If
340               IsolateNumber = tbR!IsolateNumber
350               ResultArray(U).RSI(IsolateNumber) = tbR!RSI & ""
360           End If
370           tbR.MoveNext
380       Loop
390       tb.MoveNext
400   Loop

410   Exit Sub

LoadResultArray_Error:

      Dim strES As String
      Dim intEL As Integer

420   intEL = Erl
430   strES = Err.Description
440   LogError "modPrintMicro", "LoadResultArray", intEL, strES, sql

End Sub

Private Sub PrintMicroBloodCulture(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo PrintMicroBloodCulture_Error

20    If Val(SampleID) = 0 Then Exit Sub

30    sql = "SELECT * FROM BloodCultureResults WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "ORDER BY RunDateTime DESC"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    With frmRichText.rtb
70        If Not tb.EOF Then
80            .SelUnderline = True
90            .SelFontSize = 10
100           .SelText = "Blood Culture:-" & vbCrLf

110           Do While Not tb.EOF
120               .SelText = "Bottle " & tb!BottleNumber
130               Select Case tb!TypeOfTest & ""
                      Case GetOptionSetting("BcAerobicBottle", "BSA"): .SelText = " (Aerobic)"
140                   Case GetOptionSetting("BcAnarobicBottle", "BSN"): .SelText = " (Fan Aerobic)"
150                   Case GetOptionSetting("BcFanBottle", "BFA"): .SelText = " (Anaerobic)"
160               End Select
170               Select Case tb!Result & ""
                      Case "+": .SelText = " Positive"
180                   Case "-": .SelText = " Negative"
190                   Case "*": .SelText = " Negative to date. Still under Test."
200                   Case Else: .SelText = "Unknown"
210               End Select
220               .SelText = " After " & tb!TTD & " Hours."
230               If Not tb!Valid = 1 Then
240                   .SelText = " Not yet Validated."
250               End If
260               .SelText = vbCrLf
270               tb.MoveNext
280           Loop
290       End If
300   End With

310   Exit Sub

PrintMicroBloodCulture_Error:

      Dim strES As String
      Dim intEL As Integer

320   intEL = Erl
330   strES = Err.Description
340   LogError "modPrintMicroRTF", "PrintMicroBloodCulture", intEL, strES, sql

End Sub

Private Sub PrintMicroCDiff(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo PrintMicroCDiff_Error

20    sql = "SELECT " & _
            "COALESCE(LTRIM(RTRIM(LEFT(ToxinAB + '|', CHARINDEX('|', ToxinAB) - 1))), '') ToxAB, " & _
            "COALESCE(LTRIM(RTRIM(LEFT(CDiffCulture + '|', CHARINDEX('|', CDiffCulture) - 1))), '') ToxC " & _
            "FROM Faeces F, PrintValidLog P WHERE " & _
            "F.SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "AND F.SampleID = P.SampleID " & _
            "AND P.Department = 'G' " & _
            "AND P.Valid = 1"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then Exit Sub

60    With frmRichText.rtb
70        If Trim$(tb!ToxAB & "") <> "" Then
80            .SelText = vbCrLf: CrCnt = CrCnt + 1
90            .SelFontSize = 8
100           .SelBold = False
110           .SelText = String(10, " ") & "Clostridium difficile Toxin A/B : "
120           .SelBold = True
130           .SelText = tb!ToxAB & vbCrLf: CrCnt = CrCnt + 1
140       End If

150       If Trim$(tb!ToxC & "") <> "" Then
160           .SelFontSize = 8
170           .SelBold = False
180           .SelText = String(10, " ") & "Clostridium difficile Culture : "
190           .SelBold = True
200           .SelText = tb!ToxC & vbCrLf: CrCnt = CrCnt + 1
210       End If

          '220     .SelBold = False
          '230     .SelFontSize = 6
          '240     .SelText = "Note: "
          '250     .SelItalic = True
          '260     .SelText = "C. difficile"
          '270     .SelItalic = False
          '280     .SelText = " should be requested only when there is a high index "
          '290     .SelText = "of suspicion. The clinical details received with this test request" & vbCrLf: CrCnt = CrCnt + 1
          '300     .SelFontSize = 6
          '310     .SelText = "fail to meet the criteria for "
          '320     .SelText = "testing and as such has been deemed unsuitable for analysis. "
          '330     .SelText = "Please refer to the following guidelines for " & vbCrLf: CrCnt = CrCnt + 1
          '340     .SelFontSize = 6
          '350     .SelText = "requesting "
          '360     .SelItalic = True
          '370     .SelText = "C. difficile"
          '380     .SelItalic = False
          '390     .SelText = " toxin testing: - "
          '
          '400     .SelText = "Acute onset of loose stools (more than three within a "
          '410     .SelText = "24-hour period) for two days without" & vbCrLf: CrCnt = CrCnt + 1
          '420     .SelFontSize = 6
          '430     .SelText = "another aetiology, onset after >3 days in hospital,"
          '440     .SelText = " and a history of antibiotic use "
          '450     .SelText = "or chemotherapy; "
          '
          '460     .SelBold = True
          '470     .SelUnderline = True
          '480     .SelText = "or"
          '490     .SelBold = False
          '500     .SelUnderline = False
          '
          '510     .SelText = " Recurrence of diarrhoea within" & vbCrLf: CrCnt = CrCnt + 1
          '520     .SelFontSize = 6
          '530     .SelText = "eight weeks of the end of previous treatment of "
          '540     .SelItalic = True
          '550     .SelText = "C. difficile "
          '560     .SelItalic = False
          '570     .SelText = "infection." & vbCrLf: CrCnt = CrCnt + 1
220   End With

230   UpdatePrintValidLog Val(SampleID) + SysOptMicroOffset(0), "CDIFF"

240   Exit Sub

PrintMicroCDiff_Error:

      Dim strES As String
      Dim intEL As Integer

250   intEL = Erl
260   strES = Err.Description
270   LogError "modPrintMicro", "PrintMicroCDiff", intEL, strES, sql

End Sub



Private Sub PrintMicroUrineComment(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo PrintMicroUrineComment_Error

20    sql = "Select Site from MicroSiteDetails where " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "AND Site like 'Urine'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then Exit Sub

60    sql = "Select * from Isolates where " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "AND OrganismGroup <> 'Negative results' " & _
            "AND OrganismName <> ''"
70    Set tb = New Recordset
80    RecOpenServer 0, tb, sql
90    If tb.EOF Then Exit Sub
      '100   With frmRichText.rtb
      '110       .SelBold = False
      '120       .SelFontSize = 8
      '
      '130       .SelText = "Positive cultures "
      '140       .SelUnderline = True
      '150       .SelText = "must"
      '160       .SelUnderline = False
      '170       .SelText = " be correlated with signs and symptoms of UTI" & vbCrLf: CrCnt = CrCnt + 1
      '180       .SelFontSize = 8
      '190       .SelText = "Particularly with low colony counts" & vbCrLf: CrCnt = CrCnt + 1
      '
      '200   End With

100   Exit Sub

PrintMicroUrineComment_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modPrintMicro", "PrintMicroUrineComment", intEL, strES, sql

End Sub

Private Sub PrintMicroClinDetails(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim CLD As String

10    On Error GoTo PrintMicroClinDetails_Error

20    sql = "Select ClDetails from Demographics where " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    If Not tb.EOF Then
60        If Trim$(tb!ClDetails & "") <> "" Then
70            With frmRichText.rtb
80                .SelBold = False
90                .SelText = "Clinical Details:"
100               .SelBold = True
110               CLD = tb!ClDetails & ""
120               CLD = Replace(CLD, vbCr, " ")
130               CLD = Replace(CLD, vbLf, " ")
140               .SelText = CLD & vbCrLf: CrCnt = CrCnt + 1
150           End With
160       End If
170   End If

180   Exit Sub

PrintMicroClinDetails_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "modPrintMicro", "PrintMicroClinDetails", intEL, strES, sql

End Sub

Private Sub PrintMicroCurrentABs(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim s As String

10    On Error GoTo PrintMicroCurrentABs_Error


20    sql = "Select PCA0, PCA1, PCA2, PCA3 " & _
            "from MicroSiteDetails where " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql

50    s = ""
60    If Not tb.EOF Then
70        s = tb!PCA0 & " " & tb!PCA1 & " " & tb!PCA2 & " " & tb!PCA3 & ""
80    End If
90    If Trim$(s) <> "" Then
100       With frmRichText.rtb
110           .SelFontSize = 8
120           .SelBold = False
130           .SelText = "Current Antibiotics:"
140           .SelBold = True
150           .SelText = s & vbCrLf: CrCnt = CrCnt + 1
              '    .SelFontSize = 4
              '    .SelText = String(210, "-") & vbCrLf: CrCnt = CrCnt + 1
              '    .SelFontSize = 8
160       End With
170   End If

180   Exit Sub

PrintMicroCurrentABs_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "modPrintMicro", "PrintMicroCurrentABs", intEL, strES, sql


End Sub

Private Function GetPDefault(ByVal SampleID As String) As Integer

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GetPDefault_Error

20    sql = "Select L.[Default] " & _
            "from MicroSiteDetails as M, Lists as L where " & _
            "M.SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "and L.ListType = 'SI' " & _
            "and L.[Text] like M.Site "
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql

50    GetPDefault = 3
60    If Not tb.EOF Then
70        GetPDefault = Val(tb!Default)
80    End If

90    Exit Function

GetPDefault_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modPrintMicro", "GetPDefault", intEL, strES, sql


End Function

Private Sub PrintLine()

10    With frmRichText.rtb
20        .SelFontSize = 4
30        .SelText = String(210, "-") & vbCrLf: CrCnt = CrCnt + 1
40        .SelFontSize = 8
50    End With

End Sub

Private Sub PrintMicroComment(ByVal SampleID As String, _
                              ByVal Source As String)

      Dim pSource As String
      Dim OB As Observation
      Dim OBS As New Observations

10    On Error GoTo PrintMicroComment_Error

20    Select Case UCase$(Left$(Source, 1))
          Case "D": Source = "Demographic": pSource = "Demographics Comment:"
30        Case "C": Source = "MicroConsultant": pSource = "Consultant Comment:"
40        Case "M": Source = "MicroCS": pSource = "Medical Scientist Comment:"
50        Case "P": Source = "MicroGeneral": pSource = "Urine Specimen Comment:"
60        Case "F": Source = "MicroCDiff": pSource = "CDiff Comment:"
70    End Select

80    Set OBS = OBS.Load(Val(SampleID) + SysOptMicroOffset(0), Source)

90    With frmRichText.rtb
100       .SelFontSize = 8
110       If Not OBS Is Nothing Then
120           For Each OB In OBS
130               .SelBold = False
140               .SelText = pSource
150               .SelBold = True
160               .SelText = OB.Comment & vbCrLf: CrCnt = CrCnt + 1
170           Next
180       End If
190       .SelBold = False
200   End With

210   Exit Sub

PrintMicroComment_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "modPrintMicro", "PrintMicroComment", intEL, strES

End Sub

Private Sub PrintMicroAssIDBC(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim AssID As String

10    On Error GoTo PrintMicroAssIDBC_Error

20    sql = "SELECT AssID FROM Demographics WHERE " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    With frmRichText.rtb
60        .SelFontSize = 8
70        .SelBold = False

80        If Not tb.EOF Then

90            If Trim$(tb!AssID & "") <> "" Then
100               AssID = Format$(tb!AssID - SysOptMicroOffset(0))
110               .SelText = "Please refer to Lab number " & AssID & _
                             " for associated Lab Result." & vbCrLf: CrCnt = CrCnt + 1
120           End If


130       End If
140   End With
150   Exit Sub

PrintMicroAssIDBC_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "modPrintMicro", "PrintMicroAssIDBC", intEL, strES, sql


End Sub


Private Sub PrintMicroAssIDMRSA(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
10    On Error GoTo PrintMicroAssIDMRSA_Error

20    ReDim AssID(0 To 0) As Long
      Dim ThisID As String
      Dim n As Integer
      Dim s As String
      Dim X As Integer
      Dim Found As Boolean

30    sql = "SELECT AssID FROM AssociatedIDs WHERE " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    n = -1
70    Do While Not tb.EOF
80        n = n + 1
90        ReDim Preserve AssID(0 To n) As Long
100       AssID(n) = tb!AssID
110       tb.MoveNext
120   Loop
130   sql = "SELECT SampleID FROM AssociatedIDs WHERE " & _
            "AssID = '" & Val(SampleID) + SysOptMicroOffset(0) & "'"
140   Set tb = New Recordset
150   RecOpenServer 0, tb, sql
160   Do While Not tb.EOF
170       Found = False
180       For X = 0 To UBound(AssID)
190           If AssID(X) = tb!SampleID Then
200               Found = True
210               Exit For
220           End If
230       Next
240       If Not Found Then
250           n = n + 1
260           ReDim Preserve AssID(0 To n) As Long
270           AssID(n) = tb!SampleID
280       End If
290       tb.MoveNext
300   Loop

310   If n = -1 Then Exit Sub

320   With frmRichText.rtb
330       .SelFontSize = 8
340       .SelBold = False

350       .SelText = "This Result relates to the Site specified on this form only." & vbCrLf: CrCnt = CrCnt + 1
360       .SelText = "Please refer to Results for Lab numbers "
370       s = ""
380       For n = 0 To UBound(AssID)
390           ThisID = Format$(AssID(n) - SysOptMicroOffset(0))
400           s = s & ThisID & ", "
410       Next
420       s = Left$(s, Len(s) - 2)
430       .SelText = s & " as part of this series of screens." & vbCrLf: CrCnt = CrCnt + 1
440   End With
450   Exit Sub

PrintMicroAssIDMRSA_Error:

      Dim strES As String
      Dim intEL As Integer

460   intEL = Erl
470   strES = Err.Description
480   LogError "modPrintMicro", "PrintMicroAssIDMRSA", intEL, strES, sql

End Sub
Private Sub PrintMicroOvaParasites(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim blnRejectFound As Boolean
      Dim blnHeadingPrinted As Boolean

10    On Error GoTo PrintMicroOvaParasites_Error

20    sql = "SELECT " & _
            "COALESCE(LTRIM(RTRIM(LEFT(F.Cryptosporidium + '|', CHARINDEX('|', F.Cryptosporidium) - 1))), '') Crypto, " & _
            "OP0, OP1, OP2 FROM Faeces F, PrintValidLog P WHERE " & _
            "F.SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "AND F.SampleID = P.SampleID " & _
            "AND P.Department = 'O' " & _
            "AND P.Valid = 1"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then Exit Sub
60    If Trim$(tb!Crypto) = "" And Trim$(tb!OP0 & "") = "" _
         And Trim$(tb!OP1 & "") = "" And Trim$(tb!OP2 & "") = "" Then Exit Sub

70    With frmRichText.rtb
80        .SelText = vbCrLf: CrCnt = CrCnt + 1
90        .SelFontSize = 8

100       If Trim$(tb!Crypto) <> "" Then
110           .SelBold = False
120           .SelText = Space(10) & "Cryptosporidium : "
130           .SelBold = True
140           .SelText = tb!Crypto
150           .SelText = vbCrLf: CrCnt = CrCnt + 1
160           .SelBold = False
170       End If

180       blnRejectFound = False
190       For n = 0 To 2
200           If InStr(UCase$(tb("OP" & Format(n)) & ""), "REJECTED") <> 0 Then
210               .SelText = Space(10) & "Ova and Parasites : "
220               .SelBold = True
230               .SelText = "Sample Rejected" & vbCrLf: CrCnt = CrCnt + 1
240               .SelBold = False
250               .SelFontSize = 8
260               .SelText = "I wish to remind you that "
270               .SelItalic = True
280               .SelText = "Ova and Parasites"
290               .SelItalic.Font.Italic = False
300               .SelText = " should be requested only when there is a" & vbCrLf: CrCnt = CrCnt + 1
310               .SelText = "high index of suspicion. The clinical details received "
320               .SelText = "with this test request fail to meet the" & vbCrLf: CrCnt = CrCnt + 1
330               .SelText = "criteria for testing and as such has been deemed "
340               .SelText = "unsuitable for analysis. Please refer to" & vbCrLf: CrCnt = CrCnt + 1
350               .SelText = "the following guidelines for requesting "
360               .SelItalic = True
370               .SelText = "Ova and Parasites." & vbCrLf: CrCnt = CrCnt + 1
380               .SelItalic = False

390               .SelBold = True
400               .SelUnderline = True
410               .SelText = Space(36) & "Ova and Parasites" & vbCrLf: CrCnt = CrCnt + 1
420               .SelUnderline = False
430               .SelBold = False

440               .SelText = Space(33) & "Submit one stool sample if:" & vbCrLf: CrCnt = CrCnt + 1
450               .SelText = Space(32) & "Persistent diarrhoea > 7 days ;" & vbCrLf: CrCnt = CrCnt + 1

460               .SelBold = True
470               .SelUnderline = True
480               .SelText = Space(42) & "or" & vbCrLf: CrCnt = CrCnt + 1
490               .SelUnderline = False
500               .SelBold = False

510               .SelText = Space(31) & "Patient is immunocompromised;" & vbCrLf: CrCnt = CrCnt + 1

520               .SelBold = True
530               .SelUnderline = True
540               .SelText = Space(42) & "or" & vbCrLf: CrCnt = CrCnt + 1
550               .SelUnderline = False
560               .SelBold = False

570               .SelText = Space(27) & "Patient has visited a developing country" & vbCrLf: CrCnt = CrCnt + 1

580               blnRejectFound = True
590               Exit For
600           End If
610       Next

620       If Not blnRejectFound Then
630           blnHeadingPrinted = False
640           For n = 0 To 2
650               If Trim$(tb("OP" & Format(n)) & "") <> "" Then
660                   .SelText = vbCrLf: CrCnt = CrCnt + 1
670                   .SelBold = False
680                   If Not blnHeadingPrinted Then
690                       .SelText = Space(10) & "Ova and Parasites : "
700                       blnHeadingPrinted = True
710                   Else
720                       .SelText = Space(10) & "                    "
730                   End If
740                   .SelBold = True
750                   .SelText = Trim$(tb("OP" & Format(n)) & "")
760                   .SelBold = False
770               End If
780           Next
790           .SelText = vbCrLf: CrCnt = CrCnt + 1
800       End If

810       .SelFontSize = 8
820   End With
830   UpdatePrintValidLog Val(SampleID) + SysOptMicroOffset(0), "OP"

840   Exit Sub

PrintMicroOvaParasites_Error:

      Dim strES As String
      Dim intEL As Integer

850   intEL = Erl
860   strES = Err.Description
870   LogError "modPrintMicro", "PrintMicroOvaParasites", intEL, strES, sql

End Sub

Private Sub PrintMicroPage(ByVal PrintTime As String, _
                           ByVal HCLM As String, _
                           ByRef RP As ReportToPrint, _
                           ByVal PageNumber As String, _
                           ByVal PageCount As String, _
                           ByVal CommentsPresent As Boolean)

      'A Current Antibiotics
      'B Occult Blood
      'C Clin Details
      'D Demographic Comment
      'E RSV
      'F Footer
      'G Pregnancy
      'H Heading
      'I Specimen Type
      'J HPylori
      'K Reducing Substances
      'L Print a line
      'M Microscopy
      'N Negative Results
      'O Consultant Comments
      'P Specimen Comments
      'Q Micro Comment
      'R Rota/Adeno
      'S Sensitivities
      'T CDiff
      'U Fluids
      'V Ova/Parasites
      'W Blood Culture Associated SampleID
      'X MRSA / VRE Associated SampleIDs
      'Y Urine Comment
      'Z Blood Culture

      Dim n As Integer
      Dim sql As String
      Dim tb As Recordset
      Dim PatName As String
      Dim DoB As String
      Dim Chart As String
      Dim Address0 As String
      Dim Address1 As String
      Dim Sex As String
      Dim Hospital As String
      Dim SampleDate As Date
      Dim Rundate As Date
      Dim RecDate As String

      Dim f As Integer

10    On Error GoTo PrintMicroPage_Error

20    PatName = ""
30    DoB = ""
40    Chart = ""
50    Address0 = ""
60    Address1 = ""
70    Sex = ""
80    Hospital = ""
90    SampleDate = 0
100   Rundate = 0
110   RecDate = ""

120   sql = "Select * from Demographics where " & _
            "SampleID = '" & RP.SampleID + SysOptMicroOffset(0) & "'"
130   Set tb = New Recordset
140   RecOpenClient 0, tb, sql
150   If Not tb.EOF Then
160       PatName = tb!PatName & ""
170       If IsDate(tb!DoB) Then
180           DoB = Format(tb!DoB, "dd/mmm/yyyy")
190       End If
200       Chart = tb!Chart & ""
210       Address0 = tb!Addr0 & ""
220       Address1 = tb!Addr1 & ""
230       Sex = tb!Sex & ""
240       Hospital = tb!Hospital & ""
250       SampleDate = tb!SampleDate & ""
260       Rundate = tb!Rundate & ""
270       RecDate = tb!RecDate & ""
280   End If

290   ClearUdtHeading

300   With udtHeading
310       .SampleID = RP.SampleID
320       .Dept = "MicroBiology"
330       .Name = PatName
340       .Ward = RP.Ward & ""
350       .DoB = DoB
360       .Chart = Chart
370       .Clinician = RP.Clinician & ""
380       .Address0 = Address0
390       .Address1 = Address1
400       .GP = RP.GP & ""
410       .Sex = Sex
420       .Hospital = Hospital
430       .SampleDate = SampleDate
440       .RecDate = RecDate
450       .Rundate = Rundate
460       .GpClin = RP.Clinician & ""
470       .SampleType = ""
480       .AandE = tb!AandE & ""
490   End With


500   For n = 1 To Len(HCLM)
510       Select Case Mid$(HCLM, n, 1)
          Case "A": PrintMicroCurrentABs RP.SampleID
520       Case "B": PrintMicroOccultBlood RP.SampleID
530       Case "C": PrintMicroClinDetails RP.SampleID
540       Case "D": PrintMicroComment RP.SampleID, "Demographics"
550       Case "E": PrintMicroRSV RP.SampleID
560       Case "F":
570           If RP.FaxNumber <> "" Then
580               PrintFooterRTBFax RP.Initiator, SampleDate, Rundate
590               f = FreeFile
600               Open SysOptFax(0) & RP.SampleID & "MICRO.doc" For Output As f
610               Print #f, frmRichText.rtb.TextRTF
620               Close f
630               SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "MICRO.doc"
640           Else
650               PrintFooterMicroRTB SampleDate, Rundate
660               frmRichText.rtb.SelStart = 0
                  'Do not print if Doctor is disabled in DisablePrinting
                  '*******************************************************************
670               If CheckDisablePrinting(RP.Ward, "Microbiology") Then

680               ElseIf CheckDisablePrinting(RP.GP, "Microbiology") Then
690               Else
700                   frmRichText.rtb.SelPrint Printer.hDC
710               End If
                  '*******************************************************************
                  'frmRichText.rtb.SelPrint Printer.hDC
720           End If

730       Case "G": PrintMicroPregnancy RP.SampleID
740       Case "H":
750           If RP.FaxNumber <> "" Then
760               PrintHeadingRTBFax
770           Else
780               PrintHeadingRTB
790           End If
800       Case "I": PrintSpecTypeRTF RP, PageNumber, PageCount
810       Case "J": PrintMicroHPylori RP.SampleID
820       Case "K": PrintMicroRedSub RP.SampleID
830       Case "L": PrintLine
840       Case "M": PrintMicroscopyRTF RP.SampleID
850       Case "N": PrintNegativeResults RP.SampleID
860       Case "O": PrintMicroComment RP.SampleID, "Consultant"
870       Case "P": PrintMicroComment RP.SampleID, "P"
880       Case "Q": PrintMicroComment RP.SampleID, "M"
890       Case "R": PrintMicroRotaAdeno RP.SampleID
900       Case "S": PrintMicroSensitivities4Wide RP.SampleID, CommentsPresent
910       Case "T": PrintMicroCDiff RP.SampleID
920       Case "U": PrintMicroFluids RP.SampleID
930       Case "V": PrintMicroOvaParasites RP.SampleID
940       Case "W": PrintMicroAssIDBC RP.SampleID
950       Case "X": PrintMicroAssIDMRSA RP.SampleID
960       Case "Y": PrintMicroUrineComment RP.SampleID
970       Case "Z": PrintMicroBloodCulture RP.SampleID
980       End Select
990   Next


      'Save report data to database
1000  sql = "SELECT * FROM Reports WHERE 0 = 1"
1010  Set tb = New Recordset
1020  RecOpenServer 0, tb, sql
1030  tb.AddNew
1040  tb!SampleID = RP.SampleID + SysOptMicroOffset(0)
1050  tb!Name = udtHeading.Name
1060  tb!Dept = RP.Department
1070  tb!Initiator = RP.Initiator
1080  tb!PrintTime = PrintTime
1090  tb!RepNo = Format$(PageNumber - 1) & RP.Department & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
1100  tb!PageNumber = PageNumber - 1
1110  tb!Report = frmRichText.rtb.TextRTF
1120  tb!Printer = Printer.DeviceName
1130  tb.Update

1140  Exit Sub

PrintMicroPage_Error:

      Dim strES As String
      Dim intEL As Integer

1150  intEL = Erl
1160  strES = Err.Description
1170  LogError "modPrintMicroRTF", "PrintMicroPage", intEL, strES, sql

End Sub
Private Sub PrintMicroOccultBlood(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim Result As String

10    On Error GoTo PrintMicroOccultBlood_Error

20    sql = "SELECT " & _
            "COALESCE(LTRIM(RTRIM(LEFT(F.OB0 + '|', CHARINDEX('|', F.OB0) - 1))), '') B0, " & _
            "COALESCE(LTRIM(RTRIM(LEFT(F.OB1 + '|', CHARINDEX('|', F.OB1) - 1))), '') B1, " & _
            "COALESCE(LTRIM(RTRIM(LEFT(F.OB2 + '|', CHARINDEX('|', F.OB2) - 1))), '') B2 " & _
            "FROM Faeces F, PrintValidLog P WHERE " & _
            "F.SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "AND F.SampleID = P.SampleID " & _
            "AND P.Department = 'F' " & _
            "AND P.Valid = 1"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then Exit Sub

60    With frmRichText.rtb
70        .SelText = vbCrLf: CrCnt = CrCnt + 1
80        For n = 0 To 2
90            Result = tb("B" & Format$(n))
100           If Trim$(Result) <> "" Then
110               .SelBold = False
120               .SelFontSize = 8
130               .SelText = String(10, " ") & "Occult Blood (" & Format$(n + 1) & ") : "
140               .SelBold = True
150               .SelText = Result & vbCrLf: CrCnt = CrCnt + 1
160           End If
170       Next
180       .SelBold = False
190       .SelFontSize = 8
200   End With

210   UpdatePrintValidLog Val(SampleID) + SysOptMicroOffset(0), "FOB"

220   Exit Sub

PrintMicroOccultBlood_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "modPrintMicro", "PrintMicroOccultBlood", intEL, strES, sql

End Sub

Private Sub PrintMicroPregnancy(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim fSize As Long


10    On Error GoTo PrintMicroPregnancy_Error

20    sql = "SELECT Pregnancy, HCGLevel FROM Urine WHERE " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "AND COALESCE(Pregnancy, '') <> ''"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then Exit Sub

60    With frmRichText.rtb
70        fSize = .SelFontSize

80        .SelText = vbCrLf: CrCnt = CrCnt + 1
90        .SelText = vbCrLf: CrCnt = CrCnt + 1
100       .SelText = vbCrLf: CrCnt = CrCnt + 1
110       .SelText = vbCrLf: CrCnt = CrCnt + 1
120       .SelText = vbCrLf: CrCnt = CrCnt + 1
130       .SelText = vbCrLf: CrCnt = CrCnt + 1

140       .SelBold = False
150       .SelFontSize = 9
160       .SelText = Space(10) & "Pregnancy Test: "
170       .SelBold = True
180       If UCase$(tb!Pregnancy) = "N" Then
190           .SelText = "Negative" & vbCrLf: CrCnt = CrCnt + 1
200       ElseIf UCase$(tb!Pregnancy) = "P" Then
210           .SelText = "Positive" & vbCrLf: CrCnt = CrCnt + 1
220       ElseIf UCase$(tb!Pregnancy) = "E" Then
230           .SelText = "Equivocal" & vbCrLf: CrCnt = CrCnt + 1
240       ElseIf UCase$(tb!Pregnancy) = "I" Then
250           .SelText = "Inconclusive" & vbCrLf: CrCnt = CrCnt + 1
260       ElseIf UCase$(tb!Pregnancy) = "S" Then
270           .SelText = "Unsuitable" & vbCrLf: CrCnt = CrCnt + 1
280       Else
290           .SelText = tb!Pregnancy & "" & vbCrLf: CrCnt = CrCnt + 1
300       End If
310       .SelBold = False
320       .SelText = Space(10) & "   HCG Level: "
330       .SelBold = True
340       .SelText = Trim$(tb!HCGLevel & "") & " IU/L" & vbCrLf: CrCnt = CrCnt + 1

350       PrintMicroComment RP.SampleID, "Pregnancy Comment"

360       .SelBold = False
370       .SelFontSize = fSize
380   End With

390   UpdatePrintValidLog Val(SampleID) + SysOptMicroOffset(0), "URINE"

400   Exit Sub

PrintMicroPregnancy_Error:

      Dim strES As String
      Dim intEL As Integer

410   intEL = Erl
420   strES = Err.Description
430   LogError "modPrintMicro", "PrintMicroPregnancy", intEL, strES, sql


End Sub

Private Sub PrintMicroHPylori(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim fSize As Single

10    On Error GoTo PrintMicroHPylori_Error

20    sql = "SELECT " & _
            "COALESCE(LTRIM(RTRIM(LEFT(F.HPylori + '|', CHARINDEX('|', F.HPylori) - 1))), '') H " & _
            "FROM Faeces F, PrintValidLog P WHERE " & _
            "F.SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "AND F.SampleID = P.SampleID " & _
            "AND P.Department = 'Y' " & _
            "AND P.Valid = 1"

30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then Exit Sub
60    If Trim$(tb!h & "") = "" Then Exit Sub

70    With frmRichText.rtb

80        .SelText = vbCrLf: CrCnt = CrCnt + 1
90        fSize = .SelFontSize
100       .SelBold = False
110       .SelFontSize = 8

120       .SelText = Space(10) & "Helicobacter pylori Antigen Test: "
130       .SelBold = True
140       .SelText = tb!h
150       .SelText = vbCrLf: CrCnt = CrCnt + 1
160       .SelBold = False
170       .SelFontSize = fSize

180   End With

190   UpdatePrintValidLog Val(SampleID) + SysOptMicroOffset(0), "HPYLORI"

200   Exit Sub

PrintMicroHPylori_Error:

      Dim strES As String
      Dim intEL As Integer

210   intEL = Erl
220   strES = Err.Description
230   LogError "modPrintMicroRTF", "PrintMicroHPylori", intEL, strES, sql


End Sub

Private Sub PrintMicroRotaAdeno(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim fSize As Single

10    On Error GoTo PrintMicroRotaAdeno_Error

20    sql = "SELECT " & _
            "COALESCE(LTRIM(RTRIM(LEFT(F.Rota + '|', CHARINDEX('|', F.Rota) - 1))), '') R, " & _
            "COALESCE(LTRIM(RTRIM(LEFT(F.Adeno + '|', CHARINDEX('|', F.Adeno) - 1))), '') A " & _
            "FROM Faeces F, PrintValidLog P WHERE " & _
            "F.SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "AND F.SampleID = P.SampleID " & _
            "AND P.Department = 'A' " & _
            "AND P.Valid = 1"

30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then Exit Sub
60    If Trim$(tb!R) = "" And Trim$(tb!a) = "" Then Exit Sub

70    With frmRichText.rtb
80        .SelText = vbCrLf: CrCnt = CrCnt + 1
90        fSize = .SelFontSize
100       .SelBold = False
110       .SelFontSize = 8

120       If Trim$(tb!R) <> "" Then
130           .SelBold = False
140           .SelText = Space(10) & "Rota Virus : "
150           .SelBold = True
160           .SelText = tb!R
170           .SelText = vbCrLf: CrCnt = CrCnt + 1
180       End If

190       If Trim$(tb!a) <> "" Then
200           .SelBold = False
210           .SelText = Space(10) & "Adeno Virus : "
220           .SelBold = True
230           .SelText = tb!a
240           .SelText = vbCrLf: CrCnt = CrCnt + 1
250       End If

260       .SelBold = False
270       .SelFontSize = fSize

280   End With

290   UpdatePrintValidLog Val(SampleID) + SysOptMicroOffset(0), "ROTAADENO"

300   Exit Sub

PrintMicroRotaAdeno_Error:

      Dim strES As String
      Dim intEL As Integer

310   intEL = Erl
320   strES = Err.Description
330   LogError "modPrintMicro", "PrintMicroRotaAdeno", intEL, strES, sql

End Sub


Private Sub PrintMicroFluids(ByVal SID As String)

      Dim sql As String
      Dim tb As Recordset
      Dim SampleID As Double
      Dim Site As String
      Dim OB As Observation
      Dim OBS As New Observations

10    On Error GoTo PrintMicroFluids_Error

20    SampleID = SID + SysOptMicroOffset(0)

30    sql = "SELECT Site FROM MicroSiteDetails WHERE " & _
            "SampleID = '" & SampleID & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If Not tb.EOF Then
70        Site = Trim(tb!Site & "")
80    Else
90        Site = "Fluid"
100   End If

110   sql = "SELECT * FROM GenericResults WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "AND (TestName LIKE  'CSF%' " & _
            "     OR TestName LIKE  'Fluid%' )"
120   Set tb = New Recordset
130   RecOpenServer 0, tb, sql
140   If Not tb.EOF Then

150       With frmRichText.rtb
160           .SelColor = vbBlack
170           .SelBold = False

180           .SelText = Site & " Report: "
190           .SelText = vbCrLf: CrCnt = CrCnt + 1

200           ShowReportFor "FluidAppearance0", "Appearance", SampleID, "", 15
210           ShowReportFor "FluidAppearance1", "", SampleID, "", 15
220           ShowReportFor "FluidGram", "Gram", SampleID, "", 15
230           ShowReportFor "FluidGram(2)", "", SampleID, "", 15
240           ShowReportFor "FluidZN", "ZN Stain", SampleID, "", 15

250           ShowReportFor "FluidLeishmans", "Leishmans", SampleID, "", 15
260           ShowReportFor "FluidWetPrep", "Wet Prep", SampleID, "", 15
270           ShowReportFor "FluidCrystals", "Crystals", SampleID, "", 15

280           .SelText = vbCrLf: CrCnt = CrCnt + 1

290           ShowReportFor "FluidGlucose", "Glucose", SampleID, "mmol/L", 15
300           ShowReportFor "FluidProtein", "Protein", SampleID, "g/L", 15
310           ShowReportFor "FluidAlbumin", "Albumin", SampleID, "g/L", 15
320           ShowReportFor "FluidGlobulin", "Globulin", SampleID, "g/L", 15
330           ShowReportFor "FluidLDH", "LDH", SampleID, "IU/L", 15
340           ShowReportFor "FluidAmylase", "Amylase", SampleID, "IU/L", 15

350           ShowReportFor "CSFGlucose", "Glucose", SampleID, "mmol/L", 15
360           ShowReportFor "CSFProtein", "Protein", SampleID, "g/L", 15

370           .SelText = vbCrLf: CrCnt = CrCnt + 1

380           sql = "SELECT * FROM GenericResults WHERE " & _
                    "SampleID = '" & SampleID & "' " & _
                    "AND TestName LIKE  'CSFH%'"
390           Set tb = New Recordset
400           RecOpenServer 0, tb, sql
410           If Not tb.EOF Then

420               .SelBold = False
430               .SelText = "            Specimen         1         2         3" & vbCrLf & vbCrLf: CrCnt = CrCnt + 2
440               .SelText = "                 RCC   "
450               ShowReportForHaem 0, SampleID
460               .SelText = "/cmm" & vbCrLf: CrCnt = CrCnt + 1

470               .SelText = "                 WCC   "
480               ShowReportForHaem 3, SampleID
490               .SelText = "/cmm" & vbCrLf: CrCnt = CrCnt + 1

500               .SelText = "         Polymorphic   "
510               ShowReportForHaem 6, SampleID
520               .SelText = "%" & vbCrLf: CrCnt = CrCnt + 1

530               .SelText = "       Mononucleated   "
540               ShowReportForHaem 9, SampleID
550               .SelText = "%" & vbCrLf: CrCnt = CrCnt + 1
560           End If
570       End With
580   End If

590   ShowReportFor "PneumococcalAT", "Pneumococcal Antigen", SampleID, "", 25
600   ShowReportFor "LegionellaAT", "Legionella Antigen", SampleID, "", 25
610   ShowReportFor "FungalElements", "Fungal Elements", SampleID, "", 25

      'sql = "SELECT CSFFluid FROM Comments WHERE " & _
      '      "SampleID = '" & SampleID & "' " & _
      '      "AND CSFFluid IS NOT NULL AND RTRIM(LTRIM(CAST(CSFFluid AS nvarchar(4000)))) <> ''"
      'Set tb = New Recordset
      'RecOpenServer 0, tb, sql

620   Set OBS = OBS.Load(SampleID, "CSFFluid")
630   If Not OBS Is Nothing Then
640       With frmRichText.rtb
650           .SelText = vbCrLf
660           .SelColor = vbBlack
670           .SelBold = False
680           .SelText = "Comment:-"
690           .SelBold = True
700           .SelText = OB.Comment
710           .SelText = vbCrLf
720           .SelColor = vbBlack
730           .SelBold = False
740       End With
750   End If


760   UpdatePrintValidLog SampleID, "FLUIDS"

770   Exit Sub

PrintMicroFluids_Error:

      Dim strES As String
      Dim intEL As Integer

780   intEL = Erl
790   strES = Err.Description
800   LogError "modPrintMicroRTF", "PrintMicroFluids", intEL, strES, sql

End Sub
Private Sub ShowReportFor(ByVal Parameter As String, _
                          ByVal DisplayName As String, _
                          ByVal SampleIDWithOffset As Double, _
                          ByVal Units As String, _
                          ByVal Spacing As Integer)

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo ShowReportFor_Error

20    sql = "SELECT * FROM GenericResults WHERE " & _
            "SampleID = '" & SampleIDWithOffset & "' " & _
            "AND TestName = '" & Parameter & "'"

30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then Exit Sub

60    With frmRichText.rtb
70        .SelFontName = "Courier New"
80        .SelBold = False
90        .SelFontSize = 10
100       .SelText = "              "
110       .SelBold = False
120       .SelText = Left$(DisplayName & Space$(Spacing), Spacing)
130       .SelBold = True
140       .SelText = tb!Result & " " & Units & vbCrLf: CrCnt = CrCnt + 1
150       .SelBold = False
160   End With

170   Exit Sub

ShowReportFor_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "modPrintMicroRTF", "ShowReportFor", intEL, strES, sql

End Sub



Public Sub UpdatePrintValidLog(ByVal SampleID As Double, _
                               ByVal Dept As String)

      Dim tb As Recordset
      Dim sql As String
      Dim LogDept As String

      'A Rota/Adeno
      'B Biochemistry
      'C Fluids
      'D C and S
      'E Endocrinology
      'F FOB
      'G C.diff
      'H Haematology
      'I Immunology
      'M Micro
      'O Ova/Parasites
      'R Red Sub
      'S ESR
      'U Urine
      'V RSV
      'X External
      'Y H.Pylori

10    On Error GoTo UpdatePrintValidLog_Error

20    Select Case UCase$(Dept)
          Case "MICRO": LogDept = "M"
30        Case "RSV": LogDept = "V"
40        Case "OP": LogDept = "O"
50        Case "CDIFF": LogDept = "G"
60        Case "ROTAADENO": LogDept = "A"
70        Case "FOB": LogDept = "F"
80        Case "URINE": LogDept = "U"
90        Case "CANDS": LogDept = "D"
100       Case "REDSUB": LogDept = "R"
110       Case "HPYLORI": LogDept = "Y"
120       Case "FLUIDS": LogDept = "C"
130   End Select

140   sql = "SELECT * FROM PrintValidLog WHERE " & _
            "SampleID = '" & SampleID & "' " & _
            "AND Department = '" & LogDept & "'"
150   Set tb = New Recordset
160   RecOpenClient 0, tb, sql
170   If tb.EOF Then
180       tb.AddNew
190   Else
200       ValidatedBy = tb!ValidatedBy & ""
210       sql = "INSERT INTO PrintValidLogArc " & _
                "  SELECT PrintValidLog.*, " & _
                "  'PrintHandler', " & _
                "  '" & Format$(Now, "dd/MMM/yyyy hh:mm:ss") & "' " & _
                "  FROM PrintValidLog WHERE " & _
                "  SampleID = '" & SampleID & "' " & _
                "  AND Department = '" & LogDept & "' "
220       Cnxn(0).Execute sql
230   End If
240   tb!SampleID = SampleID
250   tb!Department = LogDept
260   tb!Printed = 1
270   tb!PrintedBy = RP.Initiator

280   If Not IsNull(tb!PrintedDateTime) Then
290       If Not IsDate(tb!PrintedDateTime) Then
300           tb!PrintedDateTime = Format$(Now, "dd/MMM/yyyy hh:mm:ss")
310       End If
320   Else
330       tb!PrintedDateTime = Format$(Now, "dd/MMM/yyyy hh:mm:ss")
340   End If

350   tb!ValidatedBy = ValidatedBy

360   If Not IsNull(tb!ValidatedDateTime) Then
370       If Not IsDate(tb!ValidatedDateTime) Then
380           tb!ValidatedDateTime = Format$(Now, "dd/MMM/yyyy hh:mm:ss")
390       End If
400   Else
410       tb!ValidatedDateTime = Format$(Now, "dd/MMM/yyyy hh:mm:ss")
420   End If
430   tb.Update

440   Exit Sub

UpdatePrintValidLog_Error:

      Dim strES As String
      Dim intEL As Integer

450   intEL = Erl
460   strES = Err.Description
470   LogError "modPrintMicroRTF", "UpdatePrintValidLog", intEL, strES, sql

End Sub


Public Function FillOrgGroups(ByRef strGroup() As OrgGroup, _
                              ByVal SampleIDWithOffset As Double) _
                              As Integer

      Dim tb As Recordset
      Dim tbO As Recordset
      Dim sql As String
      Dim n As Integer
      Dim IsoNum As Integer

10    On Error GoTo FillOrgGroups_Error

20    sql = "SELECT OrganismGroup, OrganismName, Qualifier, " & _
            "IsolateNumber, COALESCE(NonReportable, 0) NonReportable " & _
            "FROM Isolates WHERE " & _
            "SampleID = '" & SampleIDWithOffset & "'"
30    Set tb = New Recordset
40    RecOpenClient 0, tb, sql
50    n = 1
60    Do While Not tb.EOF
70        IsoNum = tb!IsolateNumber
80        With strGroup(IsoNum)
90            .OrgGroup = tb!OrganismGroup & ""
100           .OrgName = tb!OrganismName & ""
110           .Qualifier = tb!Qualifier & ""
120           .NonReportable = tb!NonReportable
130           sql = "Select ShortName, ReportName from Organisms where " & _
                    "Name = '" & tb!OrganismName & "'"
140           Set tbO = New Recordset
150           RecOpenClient 0, tbO, sql
160           If Not tbO.EOF Then
170               .ShortName = tbO!ShortName & ""
180               .ReportName = Trim$(tbO!ReportName & "")
190           Else
200               .ShortName = Trim$(tb!OrganismName & "")
210               .ReportName = Trim$(tb!OrganismName & "")
220           End If
230           If .ReportName = "" Then
240               .ReportName = .OrgName
250           End If
260       End With
270       n = n + 1
280       tb.MoveNext
290   Loop

300   FillOrgGroups = n - 1

310   Exit Function

FillOrgGroups_Error:

      Dim strES As String
      Dim intEL As Integer

320   intEL = Erl
330   strES = Err.Description
340   LogError "modPrintMicroRTF", "FillOrgGroups", intEL, strES, sql

End Function

Private Sub ShowReportForHaem(ByVal pNumber As Integer, ByVal SampleIDWithOffset As Double)

      Dim sql As String
      Dim tb As Recordset
      Dim s As String

10    On Error GoTo ShowReportForHaem_Error

20    s = "      "
30    sql = "SELECT Result FROM GenericResults WHERE " & _
            "SampleID = '" & SampleIDWithOffset & "' " & _
            "AND TestName = 'CSFHAEM" & Format$(pNumber) & "'"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    If tb.EOF Then
70        s = s & "          "
80    Else
90        s = s & Left$(tb!Result & Space$(10), 10)
100   End If

110   sql = "SELECT Result FROM GenericResults WHERE " & _
            "SampleID = '" & SampleIDWithOffset & "' " & _
            "AND TestName = 'CSFHAEM" & Format$(pNumber + 1) & "'"
120   Set tb = New Recordset
130   RecOpenServer 0, tb, sql
140   If tb.EOF Then
150       s = s & "          "
160   Else
170       s = s & Left$(tb!Result & Space$(10), 10)
180   End If

190   sql = "SELECT Result FROM GenericResults WHERE " & _
            "SampleID = '" & SampleIDWithOffset & "' " & _
            "AND TestName = 'CSFHAEM" & Format$(pNumber + 2) & "'"
200   Set tb = New Recordset
210   RecOpenServer 0, tb, sql
220   If tb.EOF Then
230       s = s & "          "
240   Else
250       s = s & Left$(tb!Result & Space$(10), 10)
260   End If

270   With frmRichText.rtb
280       .SelBold = True
290       .SelText = s
300       .SelBold = False
310   End With

320   Exit Sub

ShowReportForHaem_Error:

      Dim strES As String
      Dim intEL As Integer

330   intEL = Erl
340   strES = Err.Description
350   LogError "modPrintMicroRTF", "ShowReportForHaem", intEL, strES, sql

End Sub



Private Sub PrintMicroRSV(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo PrintMicroRSV_Error

20    sql = "SELECT " & _
            "COALESCE(LTRIM(RTRIM(LEFT(Result + '|', CHARINDEX('|', Result) - 1))), '') R " & _
            "FROM GenericResults WHERE " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "and TestName = 'RSV'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then Exit Sub

60    With frmRichText.rtb

70        .SelBold = False
80        .SelFontSize = 8

90        .SelText = Space(10) & "RSV : "
100       .SelBold = True
110       .SelText = tb!R & vbCrLf: CrCnt = CrCnt + 1
120       .SelBold = False
130       .SelFontSize = 8
140   End With
150   UpdatePrintValidLog Val(SampleID) + SysOptMicroOffset(0), "RSV"

160   Exit Sub

PrintMicroRSV_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "modPrintMicro", "PrintMicroRSV", intEL, strES, sql


End Sub

Private Sub PrintMicroRedSub(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo PrintMicroRedSub_Error

20    sql = "Select * from GenericResults where " & _
            "SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "and TestName = 'RedSub'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then Exit Sub
60    With frmRichText.rtb
70        .SelText = vbCrLf: CrCnt = CrCnt + 1
80        .SelBold = False
90        .SelFontSize = 8

100       .SelText = Space(5) & "Reducing Substances : "
110       .SelBold = True
120       .SelText = tb!Result & "" & vbCrLf: CrCnt = CrCnt + 1
130       .SelBold = False
140       .SelFontSize = 8
150   End With
160   UpdatePrintValidLog Val(SampleID) + SysOptMicroOffset(0), "REDSUB"

170   Exit Sub

PrintMicroRedSub_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "modPrintMicro", "PrintMicroRedSub", intEL, strES, sql

End Sub

Private Sub PrintMicroSensitivities4Wide(ByVal SampleID As String, _
                                         ByVal CommentsPresent As Boolean)

      Dim tb As Recordset
      Dim sql As String
      Dim strGroup(1 To 4) As OrgGroup
      Dim ABCount As Integer
      Dim n As Integer
      Dim X As Integer
      Dim Y As Integer
      Dim MaxIsolates As Integer
      Dim SampleIDWithOffset As Double

10    On Error GoTo PrintMicroSensitivities4Wide_Error

20    ReDim ResultArray(0 To 0) As AntibioticPrintLine

30    SampleIDWithOffset = Val(SampleID) + SysOptMicroOffset(0)

40    sql = "Select I.* from Isolates AS I, PrintValidLog AS P where " & _
            "I.SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "AND I.SampleID = P.SampleID " & _
            "AND P.Department = 'D' " & _
            "AND P.Valid = 1"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If tb.EOF Then
80        Exit Sub
90    End If

100   LoadResultArray SampleIDWithOffset, ResultArray()

110   ABCount = UBound(ResultArray())

120   MaxIsolates = FillOrgGroups(strGroup(), SampleIDWithOffset)

130   With frmRichText.rtb
140       If Not CommentsPresent Then
150           For n = 1 To 10 - ABCount
160               .SelText = vbCrLf: CrCnt = CrCnt + 1
170           Next
180       End If

190       .SelBold = True
200       .SelFontSize = 9
210       .SelText = vbCrLf: CrCnt = CrCnt + 1
220       .SelText = "Culture:"
230       For X = 1 To 4
240           If strGroup(X).OrgName <> "" Then
250               .SelText = strGroup(X).Qualifier & " " & strGroup(X).ReportName & vbCrLf & Space(8)
260           ElseIf strGroup(X).OrgGroup <> "" Then
270               .SelText = strGroup(X).OrgGroup & vbCrLf
280           End If
290       Next
300       .SelFontSize = 9
310       .SelBold = False
320       .SelText = vbCrLf: CrCnt = CrCnt + 1

330       If ABCount > 0 Then
340           .SelUnderline = True
350           .SelText = Left$("Sensitivities" & Space$(19), 20)

360           .SelText = Left$(IIf((strGroup(1).ShortName = ""), "     ", Left(strGroup(1).ShortName, 19)) & Space(20), 20)
370           .SelText = Left$(IIf(strGroup(2).ShortName = "", "     ", Left(strGroup(2).ShortName, 19)) & Space(20), 20)
380           .SelText = Left$(IIf(strGroup(3).ShortName = "", "     ", Left(strGroup(3).ShortName, 19)) & Space(20), 20)
390           .SelText = Left$(IIf(strGroup(4).ShortName = "", "     ", Left(strGroup(4).ShortName, 19)) & Space(20), 20) & vbCrLf: CrCnt = CrCnt + 1

400           .SelUnderline = False
410           .SelFontSize = 9
420           For Y = 1 To ABCount
430               .SelColor = vbBlack
440               .SelText = Left$(ResultArray(Y).AntibioticName & Space$(19), 19) & " "
450               For X = 1 To 4
460                   .SelText = Left$(ResultArray(Y).RSI(X) & Space$(20), 20)
470               Next
480               .SelText = vbCrLf: CrCnt = CrCnt + 1
490           Next
500           .SelText = vbCrLf: CrCnt = CrCnt + 1
510           .SelColor = vbBlack
520       End If
530   End With

540   UpdatePrintValidLog SampleIDWithOffset, "CANDS"

550   Exit Sub

PrintMicroSensitivities4Wide_Error:

      Dim strES As String
      Dim intEL As Integer

560   intEL = Erl
570   strES = Err.Description
580   LogError "modPrintMicroRTF", "PrintMicroSensitivities4Wide", intEL, strES, sql

End Sub

Private Sub PrintNegativeResults(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String
      Dim CultPrinted As Boolean

10    On Error GoTo PrintNegativeResults_Error

20    sql = "Select I.* from Isolates AS I, PrintValidLog AS P where " & _
            "I.SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "AND I.SampleID = P.SampleID " & _
            "AND P.Department = 'D' " & _
            "AND P.Valid = 1"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If Not tb.EOF Then
60        With frmRichText.rtb
70            .SelText = vbCrLf: CrCnt = CrCnt + 1
80            .SelText = vbCrLf: CrCnt = CrCnt + 1
90            .SelText = vbCrLf: CrCnt = CrCnt + 1
100           .SelText = vbCrLf: CrCnt = CrCnt + 1
110           .SelBold = True
120           .SelFontSize = 9

130           CultPrinted = False
140           Do While Not tb.EOF
150               If Not CultPrinted Then
160                   .SelText = Space(10) & "CULTURE: " & tb!OrganismName & " " & tb!Qualifier & vbCrLf: CrCnt = CrCnt + 1
170                   CultPrinted = True
180               Else
190                   .SelText = Space(10) & "         " & tb!OrganismName & " " & tb!Qualifier & vbCrLf: CrCnt = CrCnt + 1
200               End If
210               tb.MoveNext
220           Loop
230       End With
240   End If

250   UpdatePrintValidLog Val(SampleID) + SysOptMicroOffset(0), "CANDS"

260   Exit Sub

PrintNegativeResults_Error:

      Dim strES As String
      Dim intEL As Integer

270   intEL = Erl
280   strES = Err.Description
290   LogError "modPrintMicro", "PrintNegativeResults", intEL, strES, sql

End Sub

Public Sub PrintResultMicroRTF()

      Dim PageCount As Integer
      Dim PageNumber As Integer
      Dim pDefault As Integer
      Dim ABCount As Integer
      Dim CommentLineCount As Integer
      Dim MicroscopyLineCount As Integer
      Dim TotalLines As Integer
      Dim IsolateCount As Integer
      Dim NegativeResults As Boolean
      Dim CommentsPresent As Boolean
      'Dim TargetPrinter As String
      'Dim xFound As Boolean
      'Dim Px As Printer
      Dim MiscLineCount As Integer
      Dim FluidCount As Integer
      Dim Site As String
      Dim PrintTime As String

10    On Error GoTo PrintResultMicroRTF_Error

20    PrintTime = Format$(Now, "dd/MMM/yyyy HH:mm:ss")

30    ABCount = GetABCount(RP.SampleID)
40    FluidCount = GetFluidCount(RP.SampleID)
50    MiscLineCount = GetMiscLineCount(RP.SampleID)    'FOB+CDiff+Rota/Adeno+OP+RSV
60    CommentLineCount = GetCommentLineCount(RP.SampleID)
70    CommentsPresent = CommentLineCount > 0
80    MicroscopyLineCount = GetMicroscopyLineCount(RP.SampleID)
90    IsolateCount = GetIsolateCount(RP.SampleID)
100   NegativeResults = IsNegativeResults(RP.SampleID)
110   pDefault = GetPDefault(RP.SampleID)

120   Site = GetMicroSite(RP.SampleID)

      'gOriginalPrinter = Printer.DeviceName
      'If pForcePrintTo = "" Then
      '  xFound = False
      '  Select Case Site
      '    Case "URINE":
      '      TargetPrinter = PrinterName("CHURINE")
      '      If TargetPrinter = "" Then
      '        TargetPrinter = PrinterName("CHMICRO")
      '      End If
      '
      '    Case "FAECES":
      '      TargetPrinter = PrinterName("CHFAECES")
      '      If TargetPrinter = "" Then
      '        TargetPrinter = PrinterName("CHMICRO")
      '      End If
      '
      '    Case Else:  TargetPrinter = PrinterName("CHMICRO")
      '  End Select
      '  For Each Px In Printers
      '    If UCase(Px.DeviceName) = TargetPrinter Then
      '      Set Printer = Px
      '      xFound = True
      '      Exit For
      '    End If
      '  Next
      '  If Not xFound Then
      '    LogError "modPrintMicro", "PrintResultMicro", 350, "Can't find " & TargetPrinter
      '    Exit Sub
      '  End If
      'End If

130   If NegativeResults Then
140       PageNumber = 1
150       PrintMicroPage PrintTime, "HICADPLZGUMBRJTVEKNPYOQWXF", RP, PageNumber, PageCount, CommentsPresent
160       Printer.EndDoc
170   Else
180       TotalLines = FluidCount + CommentLineCount + MicroscopyLineCount + IsolateCount + ABCount + MiscLineCount
190       If TotalLines > 22 Then
200           PageCount = 2
210           PageNumber = 1
220           PrintMicroPage PrintTime, "HIDLZGJSBRTVEKPYOQWXF", RP, PageNumber, PageCount, CommentsPresent
230           Printer.EndDoc

240           PageNumber = 2
250           PrintMicroPage PrintTime, "HICADPLUMPYOQWXF", RP, PageNumber, PageCount, CommentsPresent
260           Printer.EndDoc
270       Else
280           PageCount = 1
290           PageNumber = 1
300           PrintMicroPage PrintTime, "HICADPLZGUMBTRVEKJSYOQWXF", RP, PageNumber, PageCount, CommentsPresent
310           Printer.EndDoc
320       End If

330   End If
      '
      'For Each Px In Printers
      '  If UCase$(Px.DeviceName) = UCase$(gOriginalPrinter) Then
      '    Set Printer = Px
      '    Exit For
      '  End If
      'Next

340   Exit Sub

PrintResultMicroRTF_Error:

      Dim strES As String
      Dim intEL As Integer

350   intEL = Erl
360   strES = Err.Description
370   LogError "modPrintMicroRTF", "PrintResultMicroRTF", intEL, strES

End Sub
Public Sub PrintSpecTypeRTF(ByRef RP As ReportToPrint, _
                            ByVal CurrentPage As Integer, _
                            ByVal TotalPages As Integer)

      Dim tbSite As Recordset
      Dim sql As String
      Dim SiteDetails As String
      Dim Site As String

10    On Error GoTo PrintSpecTypeRTF_Error

20    sql = "Select * from MicroSiteDetails where " & _
            "SampleID = '" & Val(RP.SampleID) + SysOptMicroOffset(0) & "'"
30    Set tbSite = New Recordset
40    RecOpenClient 0, tbSite, sql
50    If Not tbSite.EOF Then
60        Site = tbSite!Site & ""
70        SiteDetails = tbSite!SiteDetails & ""
80    End If

90    With frmRichText.rtb

          'Print Specimen type and its details
100       .SelColor = vbBlack
110       .SelFontSize = 8
120       .SelFontName = "Courier New"
130       .SelBold = False
140       .SelText = "Specimen Type:"
150       .SelBold = True
160       .SelText = Site & " " & SiteDetails & " "
          'Print page number if pages are more than 1
170       If TotalPages > 1 Then
180           .SelText = "Page " & CurrentPage & " of " & TotalPages & " " & vbCrLf: CrCnt = CrCnt + 1
190       End If
          'if print is a copy
200       If RP.ThisIsCopy Then
210           .SelBold = False
220           .SelText = "This is a COPY Report for Attention of "
230           .SelBold = True
240           .SelText = Trim$(RP.SendCopyTo)
250       End If
260       .SelText = vbCrLf: CrCnt = CrCnt + 1
270       .SelBold = False
280       .SelColor = vbBlack

290   End With

300   Exit Sub

PrintSpecTypeRTF_Error:

      Dim strES As String
      Dim intEL As Integer

310   intEL = Erl
320   strES = Err.Description
330   LogError "modPrintMicroRTF", "PrintSpecTypeRTF", intEL, strES, sql

End Sub




Public Sub PrintMicroscopyRTF(ByVal SampleID As String)

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo PrintMicroscopy_Error

20    sql = "SELECT COALESCE(LTRIM(RTRIM(LEFT(WCC + '|', CHARINDEX('|', WCC) - 1))), '') W, " & _
            "COALESCE(LTRIM(RTRIM(LEFT(RCC + '|', CHARINDEX('|', RCC) - 1))), '') R, " & _
            "Crystals, Casts, Misc0, Misc1, Misc2 " & _
            "FROM Urine U, PrintValidLog P WHERE " & _
            "U.SampleID = '" & Val(SampleID) + SysOptMicroOffset(0) & "' " & _
            "AND U.SampleID = P.SampleID " & _
            "AND P.Department = 'U' " & _
            "AND P.Valid = 1"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then Exit Sub
60    If Trim$(tb!W & tb!R & tb!Crystals & tb!Casts & _
               tb!Misc0 & tb!Misc1 & tb!Misc2 & "") = "" Then Exit Sub
70    If tb!Misc2 = tb!Misc0 Then tb!Misc2 = "": tb.Update
80    If tb!Misc1 = tb!Misc0 Then tb!Misc1 = "": tb.Update

90    With frmRichText.rtb
100       .SelFontSize = 8
110       .SelBold = False
120       .SelText = "    Microscopy:"
130       .SelText = Space(15) & "           Crystals:"
140       .SelBold = True
150       If Trim$(tb!Crystals & "") = "" Then
160           .SelText = "Nil" & vbCrLf: CrCnt = CrCnt + 1
170       Else
180           .SelText = Trim$(tb!Crystals & "") & vbCrLf: CrCnt = CrCnt + 1
190       End If
200       .SelFontSize = 8
210       .SelBold = False
220       .SelText = "           WCC:"
230       .SelBold = True
240       .SelText = Left(tb!W & " /cmm" & Space(15), 15)
250       .SelBold = False
260       .SelText = "              Casts:"
270       .SelBold = True
280       If Trim$(tb!Casts & "") = "" Then
290           .SelText = "Nil" & vbCrLf: CrCnt = CrCnt + 1
300       Else
310           .SelText = Trim$(tb!Casts & "") & vbCrLf: CrCnt = CrCnt + 1
320       End If
330       .SelFontSize = 8
340       .SelBold = False
350       .SelText = "           RCC:"
360       .SelBold = True
370       .SelText = Left(tb!R & " /cmm" & Space(15), 15)
380       .SelBold = False
390       .SelText = "               Misc:"
400       .SelBold = True
410       If Trim$(tb!Misc0 & tb!Misc1 & tb!Misc2 & "") = "" Then
420           .SelText = "Nil" & vbCrLf: CrCnt = CrCnt + 1
430       Else
440           .SelText = StrConv(Trim$(tb!Misc0 & " " & tb!Misc1 & " " & tb!Misc2 & ""), vbProperCase) & vbCrLf: CrCnt = CrCnt + 1
450       End If
460       .SelBold = False

470   End With

480   UpdatePrintValidLog Val(SampleID) + SysOptMicroOffset(0), "URINE"

490   Exit Sub

PrintMicroscopy_Error:

      Dim strES As String
      Dim intEL As Integer

500   intEL = Erl
510   strES = Err.Description
520   LogError "modPrintMicro", "PrintMicroscopy", intEL, strES, sql

End Sub


