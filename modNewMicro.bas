Attribute VB_Name = "modNewMicro"
Option Explicit

Public Type LineToPrint
    Title As String
    TestName As String
    Result As String
    Flag As String * 3
    Units As String * 17
    NormalRange As String * 13
    Fasting As String * 9
    ReasonAffected As String * 23
    Comment As String
    LongComment As String
    ValidatedBy As String
    LineToPrint As String
End Type
Public udtPL() As LineToPrint
Public udtPR() As LineToPrint

Private NumberOfTitles As Integer
Private NumberOfABToPrint As Integer

Private Const NORMALFONT As String = "^FontNameCourier New^^Bold-^^Underline-^^Italic-^^Colour0^^FontSize8^"
'Private Const NORMAL10 As String = "^FontNameCourier New^^Bold-^^Underline-^^Italic-^^Colour0^^FontSize10^"
Public NORMAL10 As String
Public BOLD10 As String
Public BOLD11 As String
Public TITLEFONT As String
'Private Const BOLD10 As String = "^FontNameCourier New^^Bold+^^Underline-^^Italic-^^Colour0^^FontSize10^"
Private Const BOLD12 As String = "^FontNameCourier New^^Bold+^^Underline-^^Italic-^^Colour0^^FontSize12^"
Private ABExists As Boolean
Private BCFanBottleInUse As Boolean
Private NotAccreditedText As String
Private MaxCommentLines As Integer


Public Sub GetPrintLineSemen()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim lpc As Integer
      Dim ColName As String
      Dim ShortName As String
      Dim Units As String
      Dim TestNameLength As Integer
      Dim ResultLength As Integer


10    On Error GoTo GetPrintLineSemen_Error

20    TestNameLength = 20
30    ResultLength = 22

40    sql = "SELECT * FROM SemenResults WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql

70    If Not tb.EOF Then
80        AddNotAccreditedTest "SA", True

90        If Trim$(tb!UserName & "") <> "" Then
100           RP.Initiator = tb!UserName
110       End If
120       For n = 1 To 8
130           ColName = Choose(n, "Volume", "Consistency", "SemenCount", _
                               "Motility", "MotilityPro", "MotilityNonPro", _
                               "MotilitySlow", "MotilityNonMotile")
140           ShortName = Choose(n, "Volume", "Consistency", "Spermatozoa Count", _
                                 "Motile", "Progressive", "Non Progressive", _
                                 "Slow Progressive", "Non Motile")
150           Units = Choose(n, "mL", "", "Million per mL", _
                             "%", "%", "%", "%", "%")
160           If Trim$(tb(ColName) & "") <> "" Then
170               lpc = UBound(udtPL)
180               lpc = lpc + 1
190               ReDim Preserve udtPL(0 To lpc)
200               udtPL(lpc).Title = "SEMEN ANALYSIS"
210               If n = 3 And InStr(UCase(tb!semenCount & ""), "SEEN") Then
220                   Units = ""
230               End If

240               udtPL(lpc).TestName = Left$(ShortName & Space$(TestNameLength), TestNameLength)
250               udtPL(lpc).Result = Left$(tb(ColName) & Units & Space$(ResultLength), ResultLength)
260               udtPL(lpc).LineToPrint = NORMAL10 & _
                                           Space$(4) & _
                                           FormatString(Left$(ShortName & Space$(TestNameLength), TestNameLength) & _
                                                        Trim$(tb(ColName)) & "  " & Units, 42)
270           End If
280       Next
290   End If

300   sql = "SELECT " & _
            "R = ( SELECT  Result FROM GenericResults WHERE " & _
            "      TestName = 'SemenMorphResult' " & _
            "      AND (SampleID = '" & RP.SampleID & "' )), " & _
            "D = ( SELECT  Result FROM GenericResults WHERE " & _
            "      TestName = 'SemenMorphDescription' " & _
            "      AND (SampleID = '" & RP.SampleID & "' )) "
310   Set tb = New Recordset
320   RecOpenServer 0, tb, sql
330   If Not tb.EOF Then
340       If Not IsNull(tb!R) And Not IsNull(tb!d) Then
350           lpc = UBound(udtPL)
360           lpc = lpc + 1
370           ReDim Preserve udtPL(0 To lpc)
380           udtPL(lpc).Title = "SEMEN ANALYSIS"
390           udtPL(lpc).TestName = Left$("Morphology" & Space$(TestNameLength), TestNameLength)
400           udtPL(lpc).Result = Left$(tb!R & " " & tb!d & Space$(ResultLength), ResultLength)
410           udtPL(lpc).LineToPrint = NORMAL10 & _
                                       Space$(4) & _
                                       FormatString(Left$("Morphology" & Space$(TestNameLength), TestNameLength) & _
                                                    Trim$(tb!R) & " " & tb!d, 42)
420       End If
430   End If

440   Exit Sub

GetPrintLineSemen_Error:

      Dim strES As String
      Dim intEL As Integer

450   intEL = Erl
460   strES = Err.Description
470   LogError "modNewMicro", "GetPrintLineSemen", intEL, strES, sql

End Sub



Private Sub GetPrintLineComments(ByVal CommentTitle As String, _
                                 ByVal FieldName As String, Optional Chrs As Integer = 46)

10    On Error GoTo GetPrintLineComments_Error

20    ReDim Comments(1 To MaxCommentLines) As String
      Dim sql As String
      Dim tb As Recordset
      Dim n As Integer
      Dim lpc As Integer
      Dim OB As Observation
      Dim OBS As New Observations


30    Set OBS = OBS.Load(RP.SampleID, FieldName)

40    If Not OBS Is Nothing Then
50        For Each OB In OBS

60            FillCommentLines OB.Comment, MaxCommentLines, Comments(), Chrs
70            For n = 1 To MaxCommentLines
80                If Trim(Comments(n) & "") <> "" Then
90                    lpc = UBound(udtPL)
100                   lpc = lpc + 1
110                   ReDim Preserve udtPL(0 To lpc)
120                   udtPL(lpc).Title = CommentTitle
130                   If FieldName = "MicroConsultant" Then
140                       udtPL(lpc).LineToPrint = BOLD11 & FormatString(Comments(n), Chrs, , AlignLeft)
150                   Else
160                       udtPL(lpc).LineToPrint = NORMAL10 & FormatString(Comments(n), Chrs, , AlignLeft)
170                   End If

180               End If
190           Next
200       Next
210   End If

220   Exit Sub

GetPrintLineComments_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "modNewMicro", "GetPrintLineComments", intEL, strES, sql

End Sub

Private Sub GetPrintLineNotAccreditedTests(ByVal Site As String)

      Dim lpc        As Integer
      Dim MicroSiteNote As String

10    On Error GoTo GetPrintLineNotAccreditedTests_Error

20    lpc = UBound(udtPL)
30    lpc = lpc + 1
40    ReDim Preserve udtPL(0 To lpc)
50    udtPL(lpc).Title = ""
60    udtPL(lpc).LineToPrint = NORMAL10 & FormatString(LTrim(RTrim(NotAccreditedText) & " test(s) not accredited"), 80, , AlignLeft)

70    MicroSiteNote = GetOptionSetting("MicroSiteNote" & Site, "")
80    If MicroSiteNote <> "" Then
90        lpc = UBound(udtPL)
100       lpc = lpc + 1
110       ReDim Preserve udtPL(0 To lpc)
120       udtPL(lpc).Title = ""
130       udtPL(lpc).LineToPrint = NORMAL10 & FormatString(LTrim(RTrim(MicroSiteNote)), 80, , AlignLeft)
140   End If

150   Exit Sub
GetPrintLineNotAccreditedTests_Error:

160   LogError "modNewMicro", "GetPrintLineNotAccreditedTests", Erl, Err.Description


End Sub

Private Function GetCommentLineCount(ByVal SampleID As String, Chrs As Integer) As Integer

      Dim sql As String
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
30    Set OBS = OBS.Load(Val(SampleID), "MicroCS", "Demographic", "MicroConsultant", "MicroGeneral", "CSFFluid", "Semen", "MicroCDiff")
40    If Not OBS Is Nothing Then
50        For Each OB In OBS
60            n = n + CountLines(OB.Comment, Chrs)
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
150   LogError "modPrintMicro", "GetCommentLineCount", intEL, strES, sql


End Function


Private Function CountLines(ByVal strIP As String, Chrs As Integer) As Integer

10    ReDim Comments(1 To MaxCommentLines) As String
      Dim n As Integer

20    FillCommentLines strIP, MaxCommentLines, Comments(), Chrs

30    For n = MaxCommentLines To 1 Step -1
40        If Trim$(Comments(n)) <> "" Then
50            CountLines = n
60            Exit For
70        End If
80    Next

End Function

Private Sub CommentsForOP()

      Dim lpc As Integer
      Dim ABEndLine As Integer

10    On Error GoTo CommentsForOP_Error

20    ABEndLine = (frmMain.g.Rows - 6) + 2

30    If UBound(udtPL) < ABEndLine Then
40        For lpc = UBound(udtPL) + 1 To ABEndLine
50            ReDim Preserve udtPL(0 To lpc)
60            udtPL(lpc).LineToPrint = FormatString("", 46)
70        Next lpc
80    End If


90    lpc = UBound(udtPL)
100   lpc = lpc + 1
110   ReDim Preserve udtPL(0 To lpc)
120   udtPL(lpc).Title = "OVA AND PARASITES"
130   udtPL(lpc).LineToPrint = NORMALFONT & _
                               "I wish to remind you that " & _
                               "^Italic+^Ova and Parasites^Italic-^" & _
                               " should be requested only when there is a"

140   lpc = lpc + 1
150   ReDim Preserve udtPL(0 To lpc)
160   udtPL(lpc).Title = "OVA AND PARASITES"
170   udtPL(lpc).LineToPrint = NORMALFONT & _
                               "high index of suspicion. The clinical details received " & _
                               "with this test request fail to meet the"

180   lpc = lpc + 1
190   ReDim Preserve udtPL(0 To lpc)
200   udtPL(lpc).Title = "OVA AND PARASITES"
210   udtPL(lpc).LineToPrint = NORMALFONT & _
                               "criteria for testing and as such has been deemed " & _
                               "unsuitable for analysis. Please refer to"

220   lpc = lpc + 1
230   ReDim Preserve udtPL(0 To lpc)
240   udtPL(lpc).Title = "OVA AND PARASITES"
250   udtPL(lpc).LineToPrint = NORMALFONT & _
                               "the following guidelines for requesting " & _
                               "^Italic+^Ova and Parasites.^Italic-^"

260   lpc = lpc + 1
270   ReDim Preserve udtPL(0 To lpc)
280   udtPL(lpc).Title = "OVA AND PARASITES"
290   udtPL(lpc).LineToPrint = NORMALFONT & _
                               "Submit one stool sample if: Persistent diarrhoea > 7 days"

300   lpc = lpc + 1
310   ReDim Preserve udtPL(0 To lpc)
320   udtPL(lpc).Title = "OVA AND PARASITES"
330   udtPL(lpc).LineToPrint = NORMALFONT & _
                               Space$(24) & "^Underline+^^Bold+^or^Bold-^^Underline-^" & _
                               "  Patient is immunocompromised"

340   lpc = lpc + 1
350   ReDim Preserve udtPL(0 To lpc)
360   udtPL(lpc).Title = "OVA AND PARASITES"
370   udtPL(lpc).LineToPrint = NORMALFONT & _
                               Space$(24) & "^Underline+^^Bold+^or^Bold-^^Underline-^" & _
                               "  Patient has visited a developing country"

380   Exit Sub

CommentsForOP_Error:

      Dim strES As String
      Dim intEL As Integer

390   intEL = Erl
400   strES = Err.Description
410   LogError "modNewMicro", "CommentsForOP", intEL, strES

End Sub

Private Sub GetPrintLineRSV()

      Dim tb As Recordset
      Dim sql As String
      Dim lpc As Integer
      Dim Title As String
      Dim ShortName As String
      Dim TestNameLength As Integer
      Dim ResultLength As Integer
      Dim MicroSite As String

10    On Error GoTo GetPrintLineRSV_Error

20    TestNameLength = 12
30    ResultLength = 34

40    sql = "Select COALESCE(LTRIM(RTRIM(LEFT(Result + '|', CHARINDEX('|', Result) - 1))), '') Result " & _
            "from GenericResults where " & _
            "SampleID = '" & Val(RP.SampleID) & "' " & _
            "and TestName = 'RSV'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80        MicroSite = GetMicroSite(RP.SampleID)
90        If Trim$(tb!Result & "") <> "" Then
100           ShortName = "RSV:"
110           AddResultToArray IIf(InStr(1, MicroSite, "FAECES") > 1, "FAECES", ""), tb!Result & "", udtPL, ResultLength, FormatString(ShortName, TestNameLength), ""
              '
              '        lpc = UBound(udtPL)
              '        lpc = lpc + 1
              '        ReDim Preserve udtPL(0 To lpc)
              '        udtPL(lpc).Title = ""
              '        udtPL(lpc).TestName = Left$(ShortName & Space$(TestNameLength), TestNameLength)
              '        udtPL(lpc).Result = Left$(tb!Result & Space$(ResultLength), ResultLength)
              '        udtPL(lpc).LineToPrint = NORMAL10 & _
                       '                                 Space$(4) & _
                       '                                 FormatString(Left$(ShortName & Space$(TestNameLength), TestNameLength) & _
                       '                                 tb!Result, 42)
120       End If
130       UpdatePrintValidLog Val(RP.SampleID), "RSV"
140   End If


150   Exit Sub

GetPrintLineRSV_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "modNewMicro", "GetPrintLineRSV", intEL, strES, sql

End Sub

Private Sub GetPrintLineRedSub()

      Dim tb As Recordset
      Dim sql As String
      Dim lpc As Integer
      Dim Title As String
      Dim ShortName As String
      Dim TestNameLength As Integer
      Dim ResultLength As Integer

10    On Error GoTo GetPrintLineRedSub_Error

20    TestNameLength = 22
30    ResultLength = 24

40    sql = "Select * from GenericResults where " & _
            "SampleID = '" & Val(RP.SampleID) & "' " & _
            "and TestName = 'RedSub'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80        If Trim$(tb!Result & "") <> "" Then
90            ShortName = "Reducing Substances:"
100           AddResultToArray "FAECES", tb!Result & "", udtPL, ResultLength, FormatString(ShortName, TestNameLength), ""
              '        lpc = UBound(udtPL)
              '        lpc = lpc + 1
              '        ReDim Preserve udtPL(0 To lpc)
              '        udtPL(lpc).Title = ""
              '        udtPL(lpc).TestName = Left$(ShortName & Space$(TestNameLength), TestNameLength)
              '        udtPL(lpc).Result = Left$(tb!Result & Space$(ResultLength), ResultLength)
              '        udtPL(lpc).LineToPrint = NORMAL10 & _
                       '                                 Space$(4) & _
                       '                                 FormatString(Left$(ShortName & Space$(TestNameLength), TestNameLength) & _
                       '                                 tb!Result, 42)
110       End If
120       UpdatePrintValidLog Val(RP.SampleID), "REDSUB"
130   End If


140   Exit Sub

GetPrintLineRedSub_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "modNewMicro", "GetPrintLineRedSub", intEL, strES, sql

End Sub

Private Sub GetPrintLineOrganismsFromG()

      Dim lpc As Integer
      Dim s As String
      Dim X As Integer
      Dim y As Integer
      Dim Found As Boolean
      Dim Org As String
      Dim OrgCount As Byte

10    On Error GoTo GetPrintLineOrganismsFromG_Error

20    Found = False

30    For X = 0 To frmMain.g.Cols - 1
40        For y = 0 To frmMain.g.Rows - 1
50            If frmMain.g.TextArray(y * frmMain.g.Cols + X) <> "" Then
60                Found = True
70                Exit For
80            End If
90        Next
100       If Found Then
110           Exit For
120       End If
130   Next
140   If Not Found Then Exit Sub

150   lpc = UBound(udtPL)
160   lpc = lpc + 1
170   ReDim Preserve udtPL(0 To lpc)
180   udtPL(lpc).Title = "CULTURE:"

190   s = ""
200   OrgCount = 1

210   With frmMain.g
220       For X = 1 To 6
230           s = ""
240           If .TextMatrix(1, X) <> "" Or .TextMatrix(4, X) <> "" Then
250               If .TextMatrix(4, X) <> "" Then
260                   Org = .TextMatrix(4, X)
270               Else
280                   Org = .TextMatrix(1, X)
290               End If
300               If .TextMatrix(2, X) <> "" Then
310                   Org = .TextMatrix(2, X) & " " & Org
320               End If
330               s = OrgCount & ". " & Org & "." & vbCrLf
340               udtPL(lpc).LineToPrint = s
350               lpc = lpc + 1
360               ReDim Preserve udtPL(0 To lpc)
370               udtPL(lpc).Title = "CULTURE:"
380               OrgCount = OrgCount + 1
                  '            If Len(s & Org) < 80 Then
                  '                s = s & Org & ". "
                  '            Else
                  '                s = s & vbCrLf
                  '                udtPL(lpc).LineToPrint = s
                  '                lpc = lpc + 1
                  '                ReDim Preserve udtPL(0 To lpc)
                  '                udtPL(lpc).Title = "CULTURE:"
                  '                s = Org & ". "
                  '            End If
390           ElseIf .TextMatrix(0, X) <> "" Then
400               s = OrgCount & ". " & .TextMatrix(0, X) & "." & vbCrLf
410               udtPL(lpc).LineToPrint = s
420               lpc = lpc + 1
430               ReDim Preserve udtPL(0 To lpc)
440               udtPL(lpc).Title = "CULTURE:"
450               OrgCount = OrgCount + 1
                  '            If Len(s & .TextMatrix(0, x)) < 80 Then
                  '                s = s & .TextMatrix(0, x) & ". "
                  '            Else
                  '                s = s & vbCrLf
                  '                udtPL(lpc).LineToPrint = s
                  '                lpc = lpc + 1
                  '                ReDim Preserve udtPL(0 To lpc)
                  '                udtPL(lpc).Title = "CULTURE:"
                  '                s = .TextMatrix(0, x) & ". "
                  '            End If
460           End If
470       Next
480   End With



490   s = s & vbCrLf

500   udtPL(lpc).LineToPrint = s

510   Exit Sub

GetPrintLineOrganismsFromG_Error:

      Dim strES As String
      Dim intEL As Integer

520   intEL = Erl
530   strES = Err.Description
540   LogError "modNewMicro", "GetPrintLineOrganismsFromG", intEL, strES

End Sub


Private Sub GetPrintLineSensitivityFromG(ByRef RowNumber As Integer, _
                                         ByVal StartCol As Integer, _
                                         ByVal EndCol As Integer)

      Dim s As String
      Dim X As Integer
      Dim lpc As Integer
      Dim Found As Boolean

10    On Error GoTo GetPrintLineSensitivityFromG_Error

20    lpc = UBound(udtPL)
30    lpc = lpc + 1
40    ReDim Preserve udtPL(0 To lpc)

50    udtPL(lpc).Title = "Sensitivities       "
60    For X = StartCol To EndCol
70        If frmMain.g.TextMatrix(3, X) <> "" Then
80            udtPL(lpc).Title = udtPL(lpc).Title & Left$(frmMain.g.TextMatrix(3, X) & Space$(20), 20) & " "
90        ElseIf frmMain.g.TextMatrix(4, X) <> "" Then
100           udtPL(lpc).Title = udtPL(lpc).Title & Left$(frmMain.g.TextMatrix(4, X) & Space$(20), 20) & " "
110       ElseIf frmMain.g.TextMatrix(1, X) <> "" Then
120           udtPL(lpc).Title = udtPL(lpc).Title & Left$(frmMain.g.TextMatrix(1, X) & Space$(20), 20) & " "
130       ElseIf frmMain.g.TextMatrix(0, X) <> "" Then
140           udtPL(lpc).Title = udtPL(lpc).Title & Left$(frmMain.g.TextMatrix(0, X) & Space$(20), 20) & " "
150       End If
160   Next
170   udtPL(lpc).Title = Trim$(udtPL(lpc).Title)

180   s = NORMAL10 & _
          Left$(frmMain.g.TextMatrix(RowNumber, 0) & Space(19), 19) & " "
190   Found = False

      '********************************************************************************************
      '***********No Color coding required in tullamore
      'if it is required on other site, code can be enabled with a condition

      'For x = StartCol To EndCol
      '    Select Case UCase$(Left$(frmMain.g.TextMatrix(RowNumber, x), 1))
      '    Case "R": s = s & "^Colour" & Format$(vbRed) & "^": Found = True
      '    Case "S": s = s & "^Colour" & Format$(vbGreen) & "^": Found = True
      '    Case "I": s = s & "^Colour" & Format$(vbBlue) & "^": Found = True
      '    End Select
      '    s = s & Left$(frmMain.g.TextMatrix(RowNumber, x) & Space$(20), 20)
      'Next
      '********************************************************************************************

200   For X = StartCol To EndCol
210       Select Case UCase$(Left$(frmMain.g.TextMatrix(RowNumber, X), 1))
          Case "R": Found = True
220       Case "S": Found = True
230       Case "I": Found = True
240       End Select
250       s = s & Left$(frmMain.g.TextMatrix(RowNumber, X) & Space$(20), 20) & " "
260   Next
270   s = s & vbCrLf
280   If Found Then
290       udtPL(lpc).LineToPrint = s
300   Else
310       lpc = lpc - 1
320       ReDim Preserve udtPL(0 To lpc)
330   End If

340   Exit Sub

GetPrintLineSensitivityFromG_Error:

      Dim strES As String
      Dim intEL As Integer

350   intEL = Erl
360   strES = Err.Description
370   LogError "modNewMicro", "GetPrintLineSensitivityFromG", intEL, strES

End Sub

Private Sub PrintSensitivities(ByVal SampleID As String)

      Dim ABCount As Integer
      Dim y As Integer
      Dim OldPrintLines As Integer

10    On Error GoTo PrintSensitivities_Error

20    FillG

30    OldPrintLines = UBound(udtPL)

40    ABCount = 0
50    If frmMain.g.Rows > 7 Then
60        ABCount = frmMain.g.Rows - 6
70    ElseIf frmMain.g.TextMatrix(6, 0) <> "" Then
80        ABCount = 1
90    End If
100   GetPrintLineOrganismsFromG
110   If ABCount > 0 Then

120       For y = 6 To frmMain.g.Rows - 1
130           GetPrintLineSensitivityFromG y, 1, 4
140       Next
150       For y = 6 To frmMain.g.Rows - 1
160           GetPrintLineSensitivityFromG y, 5, 6
170       Next



180   End If

190   If UBound(udtPL) > OldPrintLines Then
200       UpdatePrintValidLog RP.SampleID, "CANDS"
210   End If

220   Exit Sub

PrintSensitivities_Error:

      Dim strES As String
      Dim intEL As Integer

230   intEL = Erl
240   strES = Err.Description
250   LogError "modNewMicro", "PrintSensitivities", intEL, strES

End Sub

Public Function BloodCultureBottleExists(ByVal SampleIDWithOffset As Double, ByVal BottleType As String) As Boolean

      Dim tb As Recordset
      Dim sql As String
      Dim TypeOfTest As String

10    On Error GoTo BloodCultureBottleExists_Error

20    Select Case BottleType
          Case "Aerobic": TypeOfTest = GetOptionSetting("BcAerobicBottle", "BSA")
30        Case "Anaerobic": TypeOfTest = GetOptionSetting("BcAnarobicBottle", "BSN")
40        Case "Fan": TypeOfTest = GetOptionSetting("BcFanBottle", "BFA")
50        Case Else: TypeOfTest = ""
60    End Select
70    sql = "SELECT Count(*) AS Cnt FROM BloodCultureResults WHERE SampleID = " & SampleIDWithOffset & " AND TypeOfTest = '" & TypeOfTest & "'"
80    Set tb = New Recordset
90    RecOpenServer 0, tb, sql
100   BloodCultureBottleExists = (tb!Cnt > 0)

110   Exit Function

BloodCultureBottleExists_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "frmMicroReport", "BloodCultureBottleExists", intEL, strES, sql

End Function

Private Sub GetPrintLineBloodCultureBottle(BottleLine As Integer)

      Dim BottleName As String
      Dim GramStain As String
      Dim Interval As String
      Dim lpc As Integer

10    On Error GoTo GetPrintLineBloodCultureBottle_Error

20    Select Case BottleLine
      Case 1:
30    If BloodCultureBottleExists(RP.SampleID, "Aerobic") Then
'40        BottleName = FormatString("Bottle A", 9, , AlignLeft)
40        BottleName = FormatString("Aerobic Bottle", 9, , AlignLeft)
50        Interval = GetBloodCultureBottleInterval(RP.SampleID, "Aerobic")
60        If Interval <> "" Then
70            If Interval < 12 Then
80                Interval = Interval & " hr(s)"
90            ElseIf Interval >= 12 And Interval < 24 Then
100               Interval = "<1 day"
110           ElseIf Interval >= 24 Then
120               Interval = (Interval \ 24) & " days"
130           End If
140       End If
150       Interval = FormatString(Interval, 8, , AlignRight)
160   End If
170   Case 2:
180       BottleName = FormatString("", 9, , AlignCenter)
190       Interval = FormatString("", 8, , AlignCenter)
200   Case 3:
210    If BloodCultureBottleExists(RP.SampleID, "Anaerobic") Then
'220       BottleName = FormatString("Bottle B", 9, , AlignLeft)
220       BottleName = FormatString("Anaerobic Bottle", 9, , AlignLeft)
230       Interval = GetBloodCultureBottleInterval(RP.SampleID, "Anaerobic")
240       If Interval <> "" Then
250           If Interval < 12 Then
260               Interval = Interval & " hr(s)"
270           ElseIf Interval >= 12 And Interval < 24 Then
280               Interval = "<1 day"
290           ElseIf Interval >= 24 Then
300               Interval = (Interval \ 24) & " days"
310           End If
320       End If
330       Interval = FormatString(Interval, 8, , AlignRight)
340   End If
350   Case 4:
360       BottleName = FormatString("", 9, , AlignCenter)
370       Interval = FormatString("", 8, , AlignCenter)
380   Case 5:
390   If BloodCultureBottleExists(RP.SampleID, "Fan") Then
'400       BottleName = FormatString("Bottle C", 9, , AlignLeft)
400       BottleName = FormatString("Paediatric bottle", 9, , AlignLeft)
410       Interval = GetBloodCultureBottleInterval(RP.SampleID, "Fan")
420       If Interval <> "" Then
430           If Interval < 12 Then
440               Interval = Interval & " hr(s)"
450           ElseIf Interval >= 12 And Interval < 24 Then
460               Interval = "<1 day"
470           ElseIf Interval >= 24 Then
480               Interval = (Interval \ 24) & " days"
490           End If
500       End If
510       Interval = FormatString(Interval, 8, , AlignRight)
520   End If
530   Case 6:
540       BottleName = FormatString("", 9, , AlignCenter)
550       Interval = FormatString("", 8, , AlignCenter)
560   End Select

570   GramStain = FormatString(GetGramIdentification(RP.SampleID, BottleLine), 35, , AlignLeft)

580   If LTrim(RTrim((GramStain))) <> "" Then
590       lpc = UBound(udtPL) + 1
600       ReDim Preserve udtPL(0 To lpc)
610       udtPL(lpc).Title = "GRAM STAIN"
620       udtPL(lpc).LineToPrint = NORMAL10 & BottleName & BottleLine & "." & GramStain
630   End If


640   Exit Sub

GetPrintLineBloodCultureBottle_Error:

      Dim strES As String
      Dim intEL As Integer

650   intEL = Erl
660   strES = Err.Description
670   LogError "modNewMicro", "GetPrintLineBloodCultureBottle", intEL, strES

End Sub

Private Sub GetPrintLineBloodCultureOrganismResult()

      Dim lpc As Integer
      Dim i As Integer
      Dim s As String
      Dim Interval As String
      Dim BottleAPrinted As Boolean
      Dim BottleBPrinted As Boolean
      Dim BottleCPrinted As Boolean


10    On Error GoTo GetPrintLineBloodCultureOrganismResult_Error

20    BottleAPrinted = False
30    BottleBPrinted = False
40    BottleCPrinted = False



50    For i = 1 To 6
60        If (i = 1 Or i = 2) And Not BottleAPrinted And BloodCultureBottleExists(RP.SampleID, "Aerobic") Then
70            If BloodCultureBottleExists(RP.SampleID, "Aerobic") Then
80                If BloodCultureBottleIsPositive(RP.SampleID, "Aerobic") Then
'90                    s = "Bottle A Flagged Positive @ "
                       s = "Aerobic Bottle Flagged Positive @ "
100               Else
110                   s = "Bottle A Flagged Negative @ "
                      s = "Aerobic Bottle Flagged Negative @ "
120               End If
130               Interval = GetBloodCultureBottleInterval(RP.SampleID, "Aerobic")
140           End If
150           BottleAPrinted = True
160       ElseIf (i = 3 Or i = 4) And Not BottleBPrinted And BloodCultureBottleExists(RP.SampleID, "Anaerobic") Then
170           If BloodCultureBottleExists(RP.SampleID, "Anaerobic") Then
180               If BloodCultureBottleIsPositive(RP.SampleID, "Anaerobic") Then
'190                   s = "Bottle B Flagged Positive @ "
190                   s = " Anaerobic Bottle Flagged Positive @ "
200               Else
'210                   s = "Bottle B Flagged Negative @ "
210                   s = " Anaerobic Bottle Flagged Negative @ "
220               End If
230               Interval = GetBloodCultureBottleInterval(RP.SampleID, "Anaerobic")
240           End If
250           BottleBPrinted = True
260       ElseIf (i = 5 Or i = 6) And (Not BottleCPrinted) And BloodCultureBottleExists(RP.SampleID, "Fan") Then
270           If BloodCultureBottleExists(RP.SampleID, "Fan") Then
280               If BloodCultureBottleIsPositive(RP.SampleID, "Fan") Then
'290                   s = "Bottle C Flagged Positive @ "
290                   s = "Paediatric bottle Flagged Positive @ "
300               Else
'310                   s = "Bottle C Flagged Negative @ "
291                   s = "Paediatric bottle Flagged Positive @ "
320               End If
330               Interval = GetBloodCultureBottleInterval(RP.SampleID, "Fan")
340           End If
350           BottleCPrinted = True
360       End If

370       If Interval <> "" Then
380           If Interval < 12 Then
390               Interval = Interval & " hr(s)"
400           ElseIf Interval >= 12 And Interval < 24 Then
410               Interval = "<1 day"
420           ElseIf Interval >= 24 Then
430               Interval = (Interval \ 24) & " days"
440           End If
450       End If
460       If s <> "" Then
              'Zyam 1-2-23 added trim for both strings
              s = Trim(s)
              Interval = Trim(Interval)
              'MsgBox (Len(s))
              'Zyam 1-2-23
470           lpc = UBound(udtPL) + 1
480           ReDim Preserve udtPL(0 To lpc)
490           udtPL(lpc).Title = "RESULT"
500           udtPL(lpc).LineToPrint = FormatString(s & Interval, 46, , AlignLeft)
510       End If
520       Interval = ""
530       s = ""

540   Next i

550   Exit Sub

GetPrintLineBloodCultureOrganismResult_Error:

      Dim strES As String
      Dim intEL As Integer

560   intEL = Erl
570   strES = Err.Description
580   LogError "modNewMicro", "GetPrintLineBloodCultureOrganismResult", intEL, strES


End Sub

Private Sub GetPrintLineBloodCultureOrganisms()

      Dim lpc As Integer
      Dim i As Integer
      Dim s As String
      Dim Interval As String
      Dim OrgNo As String

10    On Error GoTo GetPrintLineBloodCultureOrganisms_Error

20    For i = 1 To 6
30        If ((i = 1 Or i = 2) And Not BloodCultureBottleIsPositive(RP.SampleID, "Aerobic")) _
             Or ((i = 3 Or i = 4) And Not BloodCultureBottleIsPositive(RP.SampleID, "Anaerobic")) _
             Or ((i = 5 Or i = 6) And Not BloodCultureBottleIsPositive(RP.SampleID, "Fan")) Then
40            OrgNo = i & "."
50            If i = 1 Or i = 2 Then
60                Interval = GetBloodCultureBottleInterval(RP.SampleID, "Aerobic")
70            ElseIf i = 3 Or i = 4 Then
80                Interval = GetBloodCultureBottleInterval(RP.SampleID, "Anaerobic")
90            ElseIf i = 5 Or i = 6 Then
100               Interval = GetBloodCultureBottleInterval(RP.SampleID, "Fan")
110           End If
120           If Interval <> "" Then
130               If Interval < 12 Then
140                   Interval = Interval & " hr(s)"
150               ElseIf Interval >= 12 And Interval < 24 Then
160                   Interval = "<1 day"
170               ElseIf Interval >= 24 Then
180                   Interval = (Interval \ 24) & " days"
190               End If
200           End If
210           s = FormatString("No growth at " & Interval, 35, , AlignLeft)
220           Select Case i
                  Case 1, 2:
230                   If BloodCultureBottleExists(RP.SampleID, "Aerobic") Then
'240                       s = FormatString("Bottle A", 9, , AlignLeft) & FormatString("", 2, , AlignLeft) & s
240                       s = FormatString(" Aerobic Bottle", 9, , AlignLeft) & FormatString("", 2, , AlignLeft) & s

250                   Else
260                       s = ""
270                   End If
280               Case 3, 4:
290                   If BloodCultureBottleExists(RP.SampleID, "Anaerobic") Then
'300                       s = FormatString("Bottle B", 9, , AlignLeft) & FormatString("", 2, , AlignLeft) & s
300                       s = FormatString("Anaerobic Bottle", 9, , AlignLeft) & FormatString("", 2, , AlignLeft) & s

310                   Else
320                       s = ""
330                   End If
340               Case 5, 6:
350                   If BloodCultureBottleExists(RP.SampleID, "Fan") Then
'360                       s = FormatString("Bottle C", 9, , AlignLeft) & FormatString("", 2, , AlignLeft) & s
360                       s = FormatString("Paediatric bottle", 9, , AlignLeft) & FormatString("", 2, , AlignLeft) & s
370                   Else
380                       s = ""
390                   End If

400           End Select
410           If Trim(s) <> "" Then
420               lpc = UBound(udtPL) + 1
430               ReDim Preserve udtPL(0 To lpc)
440               udtPL(lpc).Title = "CULTURE"
450               udtPL(lpc).LineToPrint = NORMAL10 & s
460           End If
470           OrgNo = ""
480           i = i + 1
490       Else
500           s = FormatString(frmMain.g.TextMatrix(1, i), 35, , AlignLeft)
510           OrgNo = i & "."

520           If Trim(s) <> "" Then
530               Select Case i
                      Case 1, 2:
540                       If BloodCultureBottleExists(RP.SampleID, "Aerobic") Then
'550                           s = FormatString("Bottle A", 9, , AlignLeft) & FormatString(Trim(OrgNo), 2, , AlignLeft) & s
550                           s = FormatString("Aerobic Bottle", 9, , AlignLeft) & FormatString(Trim(OrgNo), 2, , AlignLeft) & s
560                       Else
570                           s = ""
580                       End If
590                   Case 3, 4:
600                       If BloodCultureBottleExists(RP.SampleID, "Anaerobic") Then
'610                           s = FormatString("Bottle B", 9, , AlignLeft) & FormatString(Trim(OrgNo), 2, , AlignLeft) & s
610                           s = FormatString(" Anaerobic Bottle", 9, , AlignLeft) & FormatString(Trim(OrgNo), 2, , AlignLeft) & s
620                       Else
630                           s = ""
640                       End If
650                   Case 5, 6:
660                       If BloodCultureBottleExists(RP.SampleID, "Fan") Then
'670                           s = FormatString("Bottle C", 9, , AlignLeft) & FormatString(Trim(OrgNo), 2, , AlignLeft) & s
670                           s = FormatString("Paediatric bottle", 9, , AlignLeft) & FormatString(Trim(OrgNo), 2, , AlignLeft) & s
680                       Else
690                           s = ""
700                       End If


710               End Select
720           End If

730           If Trim(s) <> "" Then
740               lpc = UBound(udtPL) + 1
750               ReDim Preserve udtPL(0 To lpc)
760               udtPL(lpc).Title = "CULTURE"
770               udtPL(lpc).LineToPrint = NORMAL10 & s
780           End If
790           OrgNo = ""
800       End If
810   Next i

820   Exit Sub

GetPrintLineBloodCultureOrganisms_Error:

      Dim strES As String
      Dim intEL As Integer

830   intEL = Erl
840   strES = Err.Description
850   LogError "modNewMicro", "GetPrintLineBloodCultureOrganisms", intEL, strES


End Sub


Private Sub GetPrintLineBloodCultureSensitivities()

      Dim lpc As Integer
      Dim s As String
      Dim i As Integer
      Dim SIndex As Integer
      Dim SEndIndex As Integer

10    On Error GoTo GetPrintLineBloodCultureSensitivities_Error

20    SIndex = 6
30    SEndIndex = SIndex + NumberOfABToPrint - 1

40    lpc = UBound(udtPR) + 1
50    ReDim Preserve udtPR(0 To lpc)
60    udtPR(lpc).LineToPrint = TITLEFONT & "SUSCEPTIBILITIES" & BOLD10 & FormatString(" ", 4)

70    If IsolateHasAntibiotics(1, frmMain.g) Then
           'Zyam Changed the alignment of microsopy report 12-12-23
'          udtPR(lpc).LineToPrint = udtPR(lpc).LineToPrint & FormatString("1", 3, , AlignCenter)
80        udtPR(lpc).LineToPrint = udtPR(lpc).LineToPrint & FormatString("1", 3, , AlignRight)
           'Zyam
90    Else
          'Zyam Changed the alignment of microsopy report 12-12-23
          udtPR(lpc).LineToPrint = udtPR(lpc).LineToPrint & FormatString(" ", 3, , AlignRight)
'100       udtPR(lpc).LineToPrint = udtPR(lpc).LineToPrint & FormatString(" ", 3, , AlignCenter)
          'Zyam
110   End If
120   If IsolateHasAntibiotics(2, frmMain.g) Then
          udtPR(lpc).LineToPrint = udtPR(lpc).LineToPrint & FormatString("2", 3, , AlignRight)
140   Else
150       udtPR(lpc).LineToPrint = udtPR(lpc).LineToPrint & FormatString(" ", 3, , AlignRight)
160   End If
170   If IsolateHasAntibiotics(3, frmMain.g) Then
180       udtPR(lpc).LineToPrint = udtPR(lpc).LineToPrint & FormatString("3", 3, , AlignRight)
190   Else
200       udtPR(lpc).LineToPrint = udtPR(lpc).LineToPrint & FormatString(" ", 3, , AlignRight)
210   End If
220   If IsolateHasAntibiotics(4, frmMain.g) Then
230       udtPR(lpc).LineToPrint = udtPR(lpc).LineToPrint & FormatString("4", 3, , AlignRight)
240   Else
250       udtPR(lpc).LineToPrint = udtPR(lpc).LineToPrint & FormatString(" ", 3, , AlignRight)
260   End If
270   If IsolateHasAntibiotics(5, frmMain.g) Then
280       udtPR(lpc).LineToPrint = udtPR(lpc).LineToPrint & FormatString("5", 3, , AlignRight)
290   Else
300       udtPR(lpc).LineToPrint = udtPR(lpc).LineToPrint & FormatString(" ", 3, , AlignRight)
310   End If
320   If IsolateHasAntibiotics(6, frmMain.g) Then
330       udtPR(lpc).LineToPrint = udtPR(lpc).LineToPrint & FormatString("6", 3, , AlignRight)
340   Else
350       udtPR(lpc).LineToPrint = udtPR(lpc).LineToPrint & FormatString(" ", 3, , AlignRight)
360   End If


370   If frmMain.g.Rows - 1 < SEndIndex Then
380       SEndIndex = frmMain.g.Rows - 1
390   End If
400   For i = SIndex To SEndIndex
410       lpc = UBound(udtPR) + 1
420       ReDim Preserve udtPR(0 To lpc)
430       s = NORMAL10 & FormatString(frmMain.g.TextMatrix(i, 0), 21, , AlignLeft) & _
              FormatString(Left$(frmMain.g.TextMatrix(i, 1), 1), 3, , AlignCenter) & _
              FormatString(Left$(frmMain.g.TextMatrix(i, 2), 1), 2, , AlignRight) & _
              FormatString(Left$(frmMain.g.TextMatrix(i, 3), 1), 3, , AlignRight) & _
              FormatString(Left$(frmMain.g.TextMatrix(i, 4), 1), 3, , AlignRight) & _
              FormatString(Left$(frmMain.g.TextMatrix(i, 5), 1), 3, , AlignRight) & _
              FormatString(Left$(frmMain.g.TextMatrix(i, 6), 1), 3, , AlignRight)
              'Zyam changed the alignment of antibiotics 17-12-23
              'Zyam changed the format string from 2 to 3 of isolate 3 4 5 6 on 12-20-23
440       udtPR(lpc).LineToPrint = s
'          udtPR(lpc).LineToPrint = FormatString(s, 105, , AlignRight)
              'Zyam
450   Next i

460   Exit Sub

GetPrintLineBloodCultureSensitivities_Error:

      Dim strES As String
      Dim intEL As Integer

470   intEL = Erl
480   strES = Err.Description
490   LogError "modNewMicro", "GetPrintLineBloodCultureSensitivities", intEL, strES

End Sub




Private Sub PrintSensitivitiesBloodCulture(ByVal SampleID As String)

      Dim ABCount As Integer
      Dim lpc As Integer
      Dim ResultsPerPage As Integer

10    On Error GoTo PrintSensitivitiesBloodCulture_Error

20    BCFanBottleInUse = GetOptionSetting("BcFanBottleInUse", 0)
30    FillG
40    ResultsPerPage = Val(GetOptionSetting("ResultsPerPage", "25"))

50    ABCount = 0
60    If frmMain.g.Rows > 7 Then
70        ABCount = frmMain.g.Rows - 6
80    ElseIf frmMain.g.TextMatrix(6, 0) <> "" Then
90        ABCount = 1
100   End If

      'PRINT LINE FOR HEADING (PRINT ONLY WHEN ATLEAST ONE BOTTLE IS +VE
110   If (Not BloodCultureBottleIsPositive(RP.SampleID, "Aerobic")) And _
         (Not BloodCultureBottleIsPositive(RP.SampleID, "Anaerobic")) And _
         ((Not BloodCultureBottleIsPositive(RP.SampleID, "Fan")) And BCFanBottleInUse) Then
          'print negative culture
120       GetPrintLineBloodCultureOrganismResult
130       If ColHasValue(1, frmMain.g) Or ColHasValue(2, frmMain.g) Or ColHasValue(3, frmMain.g) Or ColHasValue(4, frmMain.g) Or _
             ColHasValue(5, frmMain.g) Or ColHasValue(6, frmMain.g) Then
140           GetPrintLineBloodCultureOrganisms
150       End If
160   Else
170       GetPrintLineBloodCultureOrganismResult
          'Get Print Bottle Line1
180       GetPrintLineBloodCultureBottle 1
190       GetPrintLineBloodCultureBottle 2
200       GetPrintLineBloodCultureBottle 3
210       GetPrintLineBloodCultureBottle 4
220       If BCFanBottleInUse Then
230           GetPrintLineBloodCultureBottle 5
240           GetPrintLineBloodCultureBottle 6
250       End If

260       If ColHasValue(1, frmMain.g) Or ColHasValue(2, frmMain.g) Or ColHasValue(3, frmMain.g) Or ColHasValue(4, frmMain.g) Or _
             ColHasValue(5, frmMain.g) Or ColHasValue(6, frmMain.g) Then
270           GetPrintLineBloodCultureOrganisms
280       End If


290       If ABCount > 0 Then
300           GetPrintLineBloodCultureSensitivities
310       End If

320       lpc = UBound(udtPL)

330   End If
      'Zyam 2-1-23 changed removed the two empty lines after the organism name
      'print 2 blank lines before comments:
      'Print blank line after heading
'340   lpc = UBound(udtPL) + 1
'350   ReDim Preserve udtPL(0 To lpc)
360   'udtPL(lpc).LineToPrint = FormatString("", 110)

      'Print blank line after heading
'370   lpc = UBound(udtPL) + 1
'380   ReDim Preserve udtPL(0 To lpc)
'390   udtPL(lpc).LineToPrint = FormatString("", 110)
      'Zyam 2-1-23


400   Exit Sub

PrintSensitivitiesBloodCulture_Error:

      Dim strES As String
      Dim intEL As Integer

410   intEL = Erl
420   strES = Err.Description
430   LogError "modNewMicro", "PrintSensitivitiesBloodCulture", intEL, strES

End Sub

Private Sub PrintSensitivitiesOther(ByVal SampleID As String)

      Dim ABCount As Integer
      Dim lpc As Integer
      Dim ResultsPerPage As Integer
      Dim StartIndex As Integer
      Dim EndIndex As Integer
      Dim OldPrintLines As Integer

10    On Error GoTo PrintSensitivitiesOther_Error

20    ResultsPerPage = Val(GetOptionSetting("ResultsPerPage", "25"))

30    ABCount = 0
40    If frmMain.g.Rows > 7 Then
50        ABCount = frmMain.g.Rows - 6
60    ElseIf frmMain.g.TextMatrix(6, 0) <> "" Then
70        ABCount = 1
80    End If

90    StartIndex = UBound(udtPL) + 1
100   OldPrintLines = StartIndex - 1

110   StartIndex = UBound(udtPL) + 1

120   If ColHasValue(1, frmMain.g) Or ColHasValue(2, frmMain.g) Or ColHasValue(3, frmMain.g) Or ColHasValue(4, frmMain.g) Then
130       GetPrintLineOrganismsOther
140   End If

150   EndIndex = UBound(udtPL)

160   If ABCount > 0 Then
170       GetPrintLineBloodCultureSensitivities
180   End If

190   lpc = UBound(udtPL)

      'If StartIndex < EndIndex Then
      '    For i = StartIndex To EndIndex
      '        udtPL(i).LineToPrint = udtPL(i).LineToPrint & vbCrLf
      '    Next i
      'End If

200   If UBound(udtPL) > OldPrintLines Then
210       UpdatePrintValidLog RP.SampleID, "CANDS"
220   End If


      ''Print blank line after headingcy
      'lpc = UBound(udtPL) + 1
      'ReDim Preserve udtPL(0 To lpc)
      'udtPL(lpc).LineToPrint = vbCrLf
      '
      'lpc = UBound(udtPL) + 1
      'ReDim Preserve udtPL(0 To lpc)
      'udtPL(lpc).LineToPrint = vbCrLf



230   Exit Sub

PrintSensitivitiesOther_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "modNewMicro", "PrintSensitivitiesOther", intEL, strES

End Sub

Private Sub GetPrintLineOrganismsOther()

      Dim lpc As Integer
      Dim i As Integer
      Dim s As String
      Dim Start As Integer
      Dim OrgNo As String
      Dim Multiline As String
      Dim LastWordIndex As Integer

10    On Error GoTo GetPrintLineOrganismsOther_Error

20    Start = 1
30    Multiline = ""
      

40    For i = 1 To 4
50        If IsItemInList(frmMain.g.TextMatrix(0, i), "MicroNotAccredited") Then
60            AddNotAccreditedTest frmMain.g.TextMatrix(0, i), False

70        End If

80        If frmMain.g.TextMatrix(2, i) <> "" Then
90            s = frmMain.g.TextMatrix(2, i) & " "
100       End If

110       If frmMain.g.TextMatrix(4, i) <> "" Then
120           s = i & ". " & s & frmMain.g.TextMatrix(4, i)
130       Else
140           If frmMain.g.TextMatrix(1, i) <> "" Then
150               If frmMain.g.TextMatrix(0, i) = "Microscopy Negative" Then
160                   s = s & frmMain.g.TextMatrix(1, i)
170               Else
180                   s = i & ". " & s & frmMain.g.TextMatrix(1, i)
190               End If
200           End If
210       End If
          
          'Zyam changed 100 to 46 and added a second multine for long organismNames 17-12-23
          s = Trim(s)
220       If Len(s) > 46 Then
              If Len(s) > 95 Then
                LastWordIndex = InStrRev(Left(s, 46), " ")
241             Multiline = FormatString(Mid(s, LastWordIndex + 1, 46), 46, , AlignLeft)
                Dim LastWordIndex2 As Integer
                LastWordIndex2 = LastWordIndex + 46
                Dim multiLine2  As String
                multiLine2 = FormatString(Mid(s, LastWordIndex2 - 1, 25), 46, , AlignLeft)
252             s = Left(s, LastWordIndex - 1)
              Else
                LastWordIndex = InStrRev(Left(s, 46), " ")
240             Multiline = FormatString(Mid(s, LastWordIndex + 1, Len(s)), 46, , AlignLeft)
250             s = Left(s, LastWordIndex - 1)
              End If
230
260       End If
270       If Trim(s) <> "" Then
280           lpc = UBound(udtPL) + 1
290           ReDim Preserve udtPL(0 To lpc)
300           udtPL(lpc).Title = "CULTURE"
310           udtPL(lpc).LineToPrint = NORMAL10 & FormatString(s, 46, , AlignLeft)
320           If Multiline <> "" Then
330               lpc = UBound(udtPL) + 1
340               ReDim Preserve udtPL(0 To lpc)
350               udtPL(lpc).Title = "CULTURE"
360               udtPL(lpc).LineToPrint = NORMAL10 & FormatString("  " & Multiline, 46, , AlignLeft)
                  
370           End If
              If multiLine2 <> "" Then
331               lpc = UBound(udtPL) + 1
342               ReDim Preserve udtPL(0 To lpc)
353               udtPL(lpc).Title = "CULTURE"

361               udtPL(lpc).LineToPrint = NORMAL10 & FormatString("  " & multiLine2, 46, , AlignLeft)
371           End If
            'Zyam 17-12-23
381       End If
          
          
390       OrgNo = ""
400       Multiline = ""
          multiLine2 = ""
410       s = ""
420   Next i

430   Exit Sub

GetPrintLineOrganismsOther_Error:

      Dim strES As String
      Dim intEL As Integer

440   intEL = Erl
450   strES = Err.Description
460   LogError "modNewMicro", "GetPrintLineOrganismsOther", intEL, strES


End Sub


Private Sub GetPrintLineBlank(udt() As LineToPrint)

      Dim lpc As Integer

10    On Error GoTo GetPrintLineBlank_Error

20    lpc = UBound(udtPL) + 1
30    ReDim Preserve udtPL(0 To lpc)
40    udtPL(lpc).Title = udtPL(lpc - 1).Title
50    udtPL(lpc).LineToPrint = FormatString("", 46)

60    Exit Sub

GetPrintLineBlank_Error:

      Dim strES As String
      Dim intEL As Integer

70    intEL = Erl
80    strES = Err.Description
90    LogError "modNewMicro", "GetPrintLineBlank", intEL, strES

End Sub





Public Sub FillG()

      Dim tb As Recordset
      Dim sql As String
      Dim R As Integer
      Dim y As Integer
      Dim X As Integer
      Dim Found As Boolean
      Dim IsolateCount As Integer
      Dim MicroSite As String
      Dim ABIndex As Integer
      Dim Organisms As String
      Dim i As Integer


10    On Error GoTo FillG_Error

20    MicroSite = GetMicroSite(RP.SampleID)

30    frmMain.Visible = True

40    With frmMain.g
50        .Clear
60        .TextMatrix(0, 0) = "Org Group"
70        .TextMatrix(1, 0) = "Org Name"
80        .TextMatrix(2, 0) = "Qualifier"
90        .TextMatrix(3, 0) = "Short Name"
100       .TextMatrix(4, 0) = "Report Name"
110       .TextMatrix(5, 0) = "Isolate #"

120       sql = "SET NOCOUNT ON " & _
                "DECLARE @Tab table " & _
                "( OrganismGroup nvarchar(100), OrganismName nvarchar(100), Qualifier nvarchar(50), " & _
                "  IsolateNumber nvarchar(50), ShortName nvarchar(50), ReportName nvarchar(100), RowIndex int identity) " & _
                "INSERT INTO @tab (OrganismGroup, OrganismName, Qualifier, IsolateNumber, ShortName, ReportName) " & _
                "SELECT DISTINCT I.OrganismGroup, I.OrganismName, I.Qualifier, I.IsolateNumber, O.ShortName, O.ReportName " & _
                "FROM Isolates I LEFT JOIN Organisms O ON O.Name = I.OrganismName " & _
                "WHERE I.SampleID = '" & RP.SampleID & "' " & _
                "ORDER BY IsolateNumber " & _
                "SELECT * FROM @Tab"
130       Set tb = New Recordset
140       RecOpenClient 0, tb, sql
150       If tb.EOF Then
160           .Clear
170       End If

180       IsolateCount = tb.RecordCount
190       Do While Not tb.EOF
200           R = tb!IsolateNumber

210           .TextMatrix(0, R) = tb!OrganismGroup & ""
220           .TextMatrix(1, R) = tb!OrganismName & ""
230           .TextMatrix(2, R) = tb!Qualifier & ""
240           .TextMatrix(3, R) = tb!ShortName & ""
250           .TextMatrix(4, R) = tb!ReportName & ""
260           .TextMatrix(5, R) = tb!IsolateNumber

270           Organisms = Organisms & "'" & tb!OrganismGroup & "',"
280           tb.MoveNext
290       Loop
300       If Len(Organisms) > 1 Then
310           Organisms = Left(Organisms, Len(Organisms) - 1)

320           .Rows = 7
330           .AddItem ""
340           .RemoveItem 6

350           sql = "SELECT DISTINCT S.Antibiotic, B.ListOrder, LTRIM(RTRIM(COALESCE(A.ReportName, ''))) ReportName, " & _
                    "CASE S.RSI WHEN 'R' THEN 'Resistant' " & _
                    "         WHEN 'S' THEN 'Sensitive' " & _
                    "         WHEN 'I' THEN 'Intermediate' " & _
                    "         ELSE '' END RSI, S.IsolateNumber " & _
                    "FROM Sensitivities S INNER JOIN Antibiotics A ON S.AntibioticCode = A.Code " & _
                    "Inner Join (Select AntibioticName, Listorder from ABDefinitions Where Site = '" & MicroSite & "' " & _
                    "And OrganismGroup = '" & .TextMatrix(0, 1) & "') B on S.Antibiotic = B.AntibioticName " & _
                    "Where S.SampleID = '" & RP.SampleID & "' AND S.Report = 1 AND COALESCE(S.Antibiotic,'') <> '' "
360           sql = sql & " UNION "
370           sql = sql & _
                    "SELECT DISTINCT S.Antibiotic, B.ListOrder, LTRIM(RTRIM(COALESCE(A.ReportName, ''))) ReportName, " & _
                    "CASE S.RSI WHEN 'R' THEN 'Resistant' " & _
                    "         WHEN 'S' THEN 'Sensitive' " & _
                    "         WHEN 'I' THEN 'Intermediate' " & _
                    "         ELSE '' END RSI, S.IsolateNumber " & _
                    "FROM Sensitivities S INNER JOIN Antibiotics A ON S.AntibioticCode = A.Code " & _
                    "Inner Join (Select AntibioticName, Listorder from ABDefinitions Where Site = '" & MicroSite & "' " & _
                    "And OrganismGroup = '" & .TextMatrix(0, 2) & "') B on S.Antibiotic = B.AntibioticName " & _
                    "Where S.SampleID = '" & RP.SampleID & "' AND S.Report = 1 AND COALESCE(S.Antibiotic,'') <> '' "
380           sql = sql & " UNION "
390           sql = sql & _
                    "SELECT DISTINCT S.Antibiotic, B.ListOrder, LTRIM(RTRIM(COALESCE(A.ReportName, ''))) ReportName, " & _
                    "CASE S.RSI WHEN 'R' THEN 'Resistant' " & _
                    "         WHEN 'S' THEN 'Sensitive' " & _
                    "         WHEN 'I' THEN 'Intermediate' " & _
                    "         ELSE '' END RSI, S.IsolateNumber " & _
                    "FROM Sensitivities S INNER JOIN Antibiotics A ON S.AntibioticCode = A.Code " & _
                    "Inner Join (Select AntibioticName, Listorder from ABDefinitions Where Site = '" & MicroSite & "' " & _
                    "And OrganismGroup = '" & .TextMatrix(0, 3) & "') B on S.Antibiotic = B.AntibioticName " & _
                    "Where S.SampleID = '" & RP.SampleID & "' AND S.Report = 1 AND COALESCE(S.Antibiotic,'') <> '' "
400           sql = sql & " UNION "
410           sql = sql & _
                    "SELECT DISTINCT S.Antibiotic, B.ListOrder, LTRIM(RTRIM(COALESCE(A.ReportName, ''))) ReportName, " & _
                    "CASE S.RSI WHEN 'R' THEN 'Resistant' " & _
                    "         WHEN 'S' THEN 'Sensitive' " & _
                    "         WHEN 'I' THEN 'Intermediate' " & _
                    "         ELSE '' END RSI, S.IsolateNumber " & _
                    "FROM Sensitivities S INNER JOIN Antibiotics A ON S.AntibioticCode = A.Code " & _
                    "Inner Join (Select AntibioticName, Listorder from ABDefinitions Where Site = '" & MicroSite & "' " & _
                    "And OrganismGroup = '" & .TextMatrix(0, 4) & "') B on S.Antibiotic = B.AntibioticName " & _
                    "Where S.SampleID = '" & RP.SampleID & "' AND S.Report = 1 AND COALESCE(S.Antibiotic,'') <> '' "
420           sql = sql & " UNION "
430           sql = sql & _
                    "SELECT DISTINCT S.Antibiotic, B.ListOrder, LTRIM(RTRIM(COALESCE(A.ReportName, ''))) ReportName, " & _
                    "CASE S.RSI WHEN 'R' THEN 'Resistant' " & _
                    "         WHEN 'S' THEN 'Sensitive' " & _
                    "         WHEN 'I' THEN 'Intermediate' " & _
                    "         ELSE '' END RSI, S.IsolateNumber " & _
                    "FROM Sensitivities S INNER JOIN Antibiotics A ON S.AntibioticCode = A.Code " & _
                    "Inner Join (Select AntibioticName, Listorder from ABDefinitions Where Site = '" & MicroSite & "' " & _
                    "And OrganismGroup = '" & .TextMatrix(0, 5) & "') B on S.Antibiotic = B.AntibioticName " & _
                    "Where S.SampleID = '" & RP.SampleID & "' AND S.Report = 1 AND COALESCE(S.Antibiotic,'') <> '' "
440           sql = sql & " UNION "
450           sql = sql & _
                    "SELECT DISTINCT S.Antibiotic, B.ListOrder, LTRIM(RTRIM(COALESCE(A.ReportName, ''))) ReportName, " & _
                    "CASE S.RSI WHEN 'R' THEN 'Resistant' " & _
                    "         WHEN 'S' THEN 'Sensitive' " & _
                    "         WHEN 'I' THEN 'Intermediate' " & _
                    "         ELSE '' END RSI, S.IsolateNumber " & _
                    "FROM Sensitivities S INNER JOIN Antibiotics A ON S.AntibioticCode = A.Code " & _
                    "Inner Join (Select AntibioticName, Listorder from ABDefinitions Where Site = '" & MicroSite & "' " & _
                    "And OrganismGroup = '" & .TextMatrix(0, 6) & "') B on S.Antibiotic = B.AntibioticName " & _
                    "Where S.SampleID = '" & RP.SampleID & "' AND S.Report = 1 AND COALESCE(S.Antibiotic,'') <> '' "
460           sql = sql & "Order By B.ListOrder"

470           Set tb = New Recordset
480           RecOpenServer 0, tb, sql

490           Do While Not tb.EOF
500               ABExists = True
510               Found = False
520               For X = 7 To .Rows - 1
530                   If .TextMatrix(X, 0) = tb!Antibiotic Or .TextMatrix(X, 0) = tb!ReportName Then
                          'antibiotic already added
540                       .Row = X
550                       For y = 1 To .Cols - 1
560                           If .TextMatrix(5, y) = tb!IsolateNumber Then
570                               .TextMatrix(.Row, tb!IsolateNumber) = tb!RSI
580                               Found = True
590                               Exit For
600                           End If
610                       Next
620                   End If
630               Next X
640               If Not Found Then
650                   .AddItem IIf(tb!ReportName <> "", tb!ReportName & "", tb!Antibiotic & "")
660                   .Row = frmMain.g.Rows - 1
670                   For y = 1 To .Cols - 1
680                       If .TextMatrix(5, y) = tb!IsolateNumber Then
690                           .TextMatrix(.Row, tb!IsolateNumber) = tb!RSI
700                           Exit For
710                       End If
720                   Next
730               End If

740               tb.MoveNext
750           Loop

              'FILL IN FORECED ONES ********************************

760           sql = "SELECT LTRIM(RTRIM(A.AntibioticName)) AS Antibiotic, LTRIM(RTRIM(COALESCE(A.ReportName,''))) ReportName, " & _
                    "CASE S.RSI WHEN 'R' THEN 'Resistant' " & _
                    "         WHEN 'S' THEN 'Sensitive' " & _
                    "         WHEN 'I' THEN 'Intermediate' " & _
                    "         ELSE '' END RSI, S.IsolateNumber, " & _
                    "S.Report, S.RSI, S.CPOFlag, S.Result, S.RunDateTime, S.UserName " & _
                    "FROM Sensitivities S, Antibiotics A " & _
                    "Where S.SampleID = '" & RP.SampleID & "' AND S.Report = 1 AND COALESCE(Antibiotic,'') <> '' " & _
                    "AND S.AntibioticCode = A.Code " & _
                    "AND S.Forced = 1"
770           Set tb = New Recordset
780           RecOpenServer 0, tb, sql
790           Do While Not tb.EOF
800               ABExists = True
810               Found = False
820               For X = 7 To .Rows - 1
830                   If .TextMatrix(X, 0) = tb!Antibiotic Or .TextMatrix(X, 0) = tb!ReportName Then
                          'antibiotic already added
840                       .Row = X
850                       For y = 1 To .Cols - 1
860                           If .TextMatrix(5, y) = tb!IsolateNumber Then
870                               .TextMatrix(.Row, tb!IsolateNumber) = tb!RSI
880                               Found = True
890                               Exit For
900                           End If
910                       Next
920                   End If
930               Next X
940               If Not Found Then
950                   .AddItem IIf(tb!ReportName <> "", tb!ReportName & "", tb!Antibiotic & "")
960                   .Row = frmMain.g.Rows - 1
970                   For y = 1 To .Cols - 1
980                       If .TextMatrix(5, y) = tb!IsolateNumber Then
990                           .TextMatrix(.Row, tb!IsolateNumber) = tb!RSI
1000                          Exit For
1010                      End If
1020                  Next
1030              End If

1040              tb.MoveNext
1050          Loop



1060          If frmMain.g.Rows > 7 Then
1070              frmMain.g.RemoveItem 6
1080          End If
1090      End If

1100  End With


1110  Exit Sub

FillG_Error:

      Dim strES As String
      Dim intEL As Integer

1120  intEL = Erl
1130  strES = Err.Description
1140  LogError "modNewMicro", "FillG", intEL, strES, sql

End Sub

Private Sub GetPrintLineFaeces()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim lpc As Integer
      Dim ColName As String
      Dim ShortName As String
      Dim Title As String
      Dim TestNameLength As Integer
      Dim ResultLength As Integer

10    On Error GoTo GetPrintLineFaeces_Error

20    TestNameLength = 22
30    ResultLength = 24

40    sql = "SELECT " & _
            "COALESCE(LTRIM(RTRIM(LEFT(F.OB0 + '|', CHARINDEX('|', F.OB0) - 1))), '') B0, " & _
            "COALESCE(LTRIM(RTRIM(LEFT(F.OB1 + '|', CHARINDEX('|', F.OB1) - 1))), '') B1, " & _
            "COALESCE(LTRIM(RTRIM(LEFT(F.OB2 + '|', CHARINDEX('|', F.OB2) - 1))), '') B2, " & _
            "COALESCE(LTRIM(RTRIM(LEFT(F.Rota + '|', CHARINDEX('|', F.Rota) - 1))), '') R, " & _
            "COALESCE(LTRIM(RTRIM(LEFT(F.Adeno + '|', CHARINDEX('|', F.Adeno) - 1))), '') A, " & _
            "COALESCE(LTRIM(RTRIM(LEFT(ToxinAB + '|', CHARINDEX('|', ToxinAB) - 1))), '') ToxAB, " & _
            "COALESCE(LTRIM(RTRIM(LEFT(CDiffCulture + '|', CHARINDEX('|', CDiffCulture) - 1))), '') ToxC, " & _
            "COALESCE(GDHDetail, '') GDHDetail, " & _
            "COALESCE(PCRDetail, '') PCRDetail, " & _
            "COALESCE(LTRIM(RTRIM(LEFT(F.Cryptosporidium + '|', CHARINDEX('|', F.Cryptosporidium) - 1))), '') Crypto, " & _
            "COALESCE(LTRIM(RTRIM(LEFT(F.GiardiaLambila + '|', CHARINDEX('|', F.GiardiaLambila) - 1))), '') GL, " & _
            "OP0, OP1, OP2, " & _
            "COALESCE(LTRIM(RTRIM(LEFT(F.HPylori + '|', CHARINDEX('|', F.HPylori) - 1))), '') H " & _
            "FROM Faeces F, PrintValidLog P WHERE " & _
            "F.SampleID = '" & RP.SampleID & "' " & _
            "AND F.SampleID = P.SampleID " & _
            "AND P.Valid = 1"
      '      "AND P.Department = 'F' "
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80        NumberOfTitles = NumberOfTitles + 1
90        For n = 1 To 15
100           ColName = Choose(n, "B0", "B1", "B2", "R", "A", "ToxAB", "ToxC", "GDHDetail", "PCRDetail", "Crypto", "GL", "OP0", "OP1", "OP2", "H")
110           ShortName = Choose(n, "Occult Blood (1):", "Occult Blood (2):", "Occult Blood (3):", _
                                 "Rota Virus:", "Adeno Virus:", _
                                 "C.difficile Toxin A/B:", "C.difficile Culture:", "GDH:", "PCR:", _
                                 "Cryptosporidium:", "Giardia Lambila", _
                                 "Ova and Parasites (1):", "                  (2):", "                  (3):", _
                                 "H.pylori Antigen Test:")
120           If Trim$(tb(ColName) & "") <> "" Then
130               If ColName = "H" Then
140                   AddNotAccreditedTest "HP", True
                     
150               End If
160               AddResultToArray "FAECES", tb(ColName), udtPL, ResultLength, FormatString(ShortName, TestNameLength), ""


170               If n = 9 Or n = 10 Or n = 11 Then
180                   If InStr(UCase$(tb("OP" & Format(n - 9)) & ""), "REJECTED") <> 0 Then
190                       CommentsForOP
200                   End If
210               End If

220           End If
230       Next
240       If tb!B0 <> "" Or tb!B1 <> "" Or tb!B2 <> "" Then
250           UpdatePrintValidLog RP.SampleID, "FOB"
260       End If
270       If tb!R <> "" Or tb!a <> "" Then
280           UpdatePrintValidLog RP.SampleID, "ROTAADENO"
290       End If
300       If tb!h <> "" Then
310           UpdatePrintValidLog RP.SampleID, "HPYLORI"
320       End If
330       If tb!ToxAB <> "" Or tb!ToxC <> "" Or tb!PCRDetail <> "" Or tb!GDHDetail <> "" Then
340           UpdatePrintValidLog RP.SampleID, "CDIFF"
350       End If
360       If tb!Crypto <> "" Or tb!OP0 <> "" Or tb!OP1 <> "" Or tb!OP2 <> "" Or tb!GL <> "" Then
370           UpdatePrintValidLog RP.SampleID, "OP"
380       End If
390   End If

400   Exit Sub

GetPrintLineFaeces_Error:

      Dim strES As String
      Dim intEL As Integer

410   intEL = Erl
420   strES = Err.Description
430   LogError "modNewMicro", "GetPrintLineFaeces", intEL, strES, sql

End Sub

Private Sub GetPrintLineMicroscopy()

      Dim tb As Recordset
      Dim sql As String
      Dim n As Integer
      Dim lpc As Integer
      Dim ColName As String
      Dim ShortName As String
      Dim Units As String
      Dim Title As String
      Dim TestNameLength As Integer
      Dim ResultLength As Integer

10    On Error GoTo GetPrintLineMicroscopy_Error

20    TestNameLength = 26
30    ResultLength = 27

40    sql = "SELECT COALESCE(LTRIM(RTRIM(LEFT(WCC + '|', CHARINDEX('|', WCC) - 1))), '') W, " & _
            "COALESCE(LTRIM(RTRIM(LEFT(RCC + '|', CHARINDEX('|', RCC) - 1))), '') R, " & _
            "Crystals, Casts, Misc0, Misc1, Misc2, " & _
            "Pregnancy = Case When Pregnancy = 'P' then 'Positive' When Pregnancy = 'N' Then 'Negative' Else COALESCE(Pregnancy, '') End, HCGLevel " & _
            "FROM Urine U, PrintValidLog P WHERE " & _
            "U.SampleID = '" & Val(RP.SampleID) & "' " & _
            "AND U.SampleID = P.SampleID " & _
            "AND P.Department = 'U' " & _
            "AND P.Valid = 1"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80        For n = 1 To 7
90            ColName = Choose(n, "W", "R", "Crystals", "Casts", "Misc0", "Misc1", "Misc2")
100           ShortName = Choose(n, "White Blood Cell", "Red Blood Cell", "Crystals", "Casts", "Miscellaneous", "", "")
110           Units = Choose(n, " /cmm", " /cmm", "", "", "", "", "")
120           If Trim$(tb(ColName) & "") <> "" Then
130               AddResultToArray "MICROSCOPY", tb(ColName) & "", udtPL, ResultLength, FormatString(ShortName, TestNameLength, , AlignLeft), Units
                  '            lpc = UBound(udtPL)
                  '            lpc = lpc + 1
                  '            ReDim Preserve udtPL(0 To lpc)
                  '            udtPL(lpc).Title = "MICROSCOPY"
                  '            udtPL(lpc).TestName = Left$(ShortName & Space$(TestNameLength), TestNameLength)
                  '            udtPL(lpc).Result = Left$(tb(ColName) & " /cmm" & Space$(ResultLength), ResultLength)
                  '            udtPL(lpc).LineToPrint = NORMAL10 & _
                               '                                    FormatString(" ", 4, , AlignLeft) & _
                               '                                    FormatString(ShortName & "", TestNameLength, , AlignLeft) & _
                               '                                    FormatString(tb(ColName) & " /cmm", ResultLength, , AlignLeft)

                  '                                    NORMAL10 & _
                                                       '                                     Space$(10) & _
                                                       '                                     Left$(ShortName & Space$(TestNameLength), TestNameLength) & _
                                                       '                                     tb(ColName) & " /cmm"
140           End If
150       Next

          '    For n = 1 To 5
          '        ColName = Choose(n, "Crystals", "Casts", "Misc0", "Misc1", "Misc2")
          '        ShortName = Choose(n, "Crystals", "Casts", "Miscellaneous", "", "")
          '        If Trim$(tb(ColName) & "") <> "" Then
          '            lpc = UBound(udtPL)
          '            lpc = lpc + 1
          '            ReDim Preserve udtPL(0 To lpc)
          '            udtPL(lpc).Title = "MICROSCOPY"
          '            udtPL(lpc).TestName = Left$(ShortName & Space$(TestNameLength), TestNameLength)
          '            udtPL(lpc).Result = Left$(tb(ColName) & Space$(ResultLength), ResultLength)
          '            udtPL(lpc).LineToPrint = NORMAL10 & _
                       '                                    FormatString(" ", 4, , AlignLeft) & _
                       '                                    FormatString(ShortName & "", TestNameLength, , AlignLeft) & _
                       '                                    FormatString(tb(ColName) & " /cmm", ResultLength, , AlignLeft)
          '        End If
          '    Next
160       UpdatePrintValidLog RP.SampleID, "URINE"

170   End If



180   Exit Sub

GetPrintLineMicroscopy_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "modNewMicro", "GetPrintLineMicroscopy", intEL, strES, sql

End Sub

Private Sub GetPrintLineIQ200()

      Dim tb As Recordset
      Dim sql As String
      Dim lpc As Integer
      Dim ShortName As String
      Dim Title As String
      Dim TestNameLength As Integer
      Dim ResultLength As Integer
      Dim UnitLength As Integer
      Dim WBCFound As Boolean
      Dim RBCFound As Boolean
      Dim CastFound As Boolean
      Dim CrystalFound As Boolean
      Dim EpithelialFound As Boolean
      Dim tempResult As String
      Dim X As Integer

10    On Error GoTo GetPrintLineIQ200_Error

20    TestNameLength = 26
30    ResultLength = 16
40    UnitLength = 5

50    sql = "SELECT ShortName,LongName,Result,Unit FROM IQ200 WHERE " & _
            "SampleID = '" & Val(RP.SampleID) & "' " & _
            "AND Result <> '[none]'"


60    Set tb = New Recordset
70    RecOpenServer 0, tb, sql
80    If Not tb.EOF Then
90        Do While Not tb.EOF


100           X = InStr(tb!Result & "", " ") - 1
110           If X > 0 Then
120               tempResult = Left(tb!Result & "", X)
130           Else
140               tempResult = tb!Result & ""
150           End If

160           If tb!ShortName & "" = "WBC" Then
170               WBCFound = True
180           ElseIf tb!ShortName & "" = "RBC" Then
190               RBCFound = True
200           ElseIf InStr(tb!LongName & "", "Cast") > 0 Then
210               CastFound = True
220           ElseIf InStr(tb!LongName & "", "Crystal") > 0 Then
230               CrystalFound = True
240           ElseIf InStr(tb!LongName & "", "Epithelial") > 0 Then
250               EpithelialFound = True
260           End If


270           If (InStr(tb!LongName & "", "Cast") > 0) Or _
                 (InStr(tb!LongName & "", "Crystal") > 0) Or _
                 (InStr(tb!LongName & "", "Epithelial") > 0) Then
280               If tb!Result > 0 Then
290                   lpc = UBound(udtPL)
300                   lpc = lpc + 1
310                   ReDim Preserve udtPL(0 To lpc)
320                   udtPL(lpc).Title = "MICROSCOPY"
330                   udtPL(lpc).TestName = tb!ShortName & ""
340                   udtPL(lpc).Result = tb!LongName & ""
350                   udtPL(lpc).LineToPrint = NORMAL10 & _
                                               FormatString(tb!LongName & "", TestNameLength, , AlignLeft) & _
                                               FormatString(tb!Result & "", ResultLength, , AlignLeft) & _
                                               FormatString(tb!Unit & "", UnitLength, , AlignLeft)
360               End If
370           Else
380               If tb!ShortName & "" <> "BACT" And tb!ShortName & "" <> "PC" Then
390                   lpc = UBound(udtPL)
400                   lpc = lpc + 1
410                   ReDim Preserve udtPL(0 To lpc)
420                   udtPL(lpc).Title = "MICROSCOPY"
430                   udtPL(lpc).TestName = tb!ShortName & ""
440                   udtPL(lpc).Result = tb!LongName & ""
450                   udtPL(lpc).LineToPrint = NORMAL10 & _
                                               FormatString(tb!LongName & "", TestNameLength, , AlignLeft) & _
                                               FormatString(tb!Result & "", ResultLength, , AlignLeft) & _
                                               FormatString(tb!Unit & "", UnitLength, , AlignLeft)
460               End If
470           End If
480           tb.MoveNext
490       Loop


      '    If Not WBCFound Then
      '        lpc = UBound(udtPL)
      '        lpc = lpc + 1
      '        ReDim Preserve udtPL(0 To lpc)
      '        udtPL(lpc).Title = "MICROSCOPY"
      '        udtPL(lpc).LineToPrint = NORMAL10 & _
      '                                 FormatString(" ", 4, , AlignLeft) & _
      '                                 FormatString("WBC", TestNameLength, , AlignLeft) & _
      '                                 FormatString("0 /uL", ResultLength, , AlignLeft)
      '
      '    End If
      '
      '    If Not RBCFound Then
      '        lpc = UBound(udtPL)
      '        lpc = lpc + 1
      '        ReDim Preserve udtPL(0 To lpc)
      '        udtPL(lpc).Title = "MICROSCOPY"
      '        udtPL(lpc).LineToPrint = NORMAL10 & _
      '                                 FormatString(" ", 4, , AlignLeft) & _
      '                                 FormatString("RBC", TestNameLength, , AlignLeft) & _
      '                                 FormatString("0 /uL", ResultLength, , AlignLeft)
      '    End If
      '
      '    If Not EpithelialFound Then
      '        lpc = UBound(udtPL)
      '        lpc = lpc + 1
      '        ReDim Preserve udtPL(0 To lpc)
      '        udtPL(lpc).Title = "MICROSCOPY"
      '        udtPL(lpc).LineToPrint = NORMAL10 & _
      '                                 FormatString(" ", 4, , AlignLeft) & _
      '                                 FormatString("Epithelial Cells", TestNameLength, , AlignLeft) & _
      '                                 FormatString("0 /uL", ResultLength, , AlignLeft)
      '    End If


          '    If Not CastFound Then
          '         udtPL(lpc).LineToPrint = NORMAL10 & _
                    '                                     FormatString(" ", 10, , AlignLeft) & _
                    '                                     FormatString("Casts", 20, , AlignLeft) & _
                    '                                     FormatString("None Seen", 15, , AlignLeft) & vbCrLf
          '    End If
          '
          '    If Not CrystalFound Then
          '         udtPL(lpc).LineToPrint = NORMAL10 & _
                    '                                     FormatString(" ", 10, , AlignLeft) & _
                    '                                     FormatString("Crystals", 20, , AlignLeft) & _
                    '                                     FormatString("None Seen", 15, , AlignLeft) & vbCrLf
          '    End If
          '




          '    sql = "SELECT Count(distinct d.sampleid) as DemoCheck " & _
               '      "FROM Demographics D " & _
               '      "INNER JOIN IQ200 I " & _
               '      "ON D.SampleId = I.SampleId " & _
               '      "WHERE D.Ward NOT IN ('OHIU', 'ROHDU', 'Oncology OPD', " & _
               '      "                   'Ante-natal Clinic',  'Haematology OPD' ) " & _
               '      "AND COALESCE(D.Pregnant, 0) <> 1 " & _
               '      "AND (DATEDIFF(dd, D.DOB, I.DateTimeOfRecord)  / 365) > 16 " & _
               '      "AND D.SampleId = '" & Val(RP.SampleID) & "' "
          '
          '    Set tb = New Recordset
          '    RecOpenServer 0, tb, sql
          '
          '    If Not tb.EOF Then
          '        'QMS Ref
          '        If PrintThis And _
                   '            tb!DemoCheck <> 0 And _
                   '            Val(GetOptionSetting("IQ200NegativeComment", "0")) = 1 Then
          '
          '            lpc = UBound(udtPL)
          '                        lpc = lpc + 1
          '                        ReDim Preserve udtPL(0 To lpc)
          '                        udtPL(lpc).Title = "MICROSCOPY"
          '                        udtPL(lpc).LineToPrint = "Urine Microscopy Negative - culture "
          '            lpc = UBound(udtPL)
          '                        lpc = lpc + 1
          '                        ReDim Preserve udtPL(0 To lpc)
          '                        udtPL(lpc).Title = "MICROSCOPY"
          '                        udtPL(lpc).LineToPrint = "not indicated."
          '
          '            MaxIsolate = 1
          '            If ColHasValue(1, frmMain.g) Then MaxIsolate = MaxIsolate + 1
          '            If ColHasValue(2, frmMain.g) Then MaxIsolate = MaxIsolate + 1
          '            If ColHasValue(3, frmMain.g) Then MaxIsolate = MaxIsolate + 1
          '            If ColHasValue(4, frmMain.g) Then MaxIsolate = MaxIsolate + 1
          '
          '            sql = "INSERT INTO Isolates (" & _
                       '                    "SampleID, IsolateNumber, OrganismName, Valid ) Values (" & _
                       '                    RP.SampleID & ", " & MaxIsolate & ", 'Urine Microscopy Negative - culture not indicated.', 1)"
          '
          '            Cnxn(0).Execute sql
          '        End If
          '    End If


          'UpdatePrintValidLog RP.SampleID, "URINE"

500   End If




510   Exit Sub

GetPrintLineIQ200_Error:

      Dim strES As String
      Dim intEL As Integer

520   intEL = Erl
530   strES = Err.Description
540   LogError "modNewMicro", "GetPrintLineIQ200", intEL, strES, sql


End Sub

Private Sub GetPrintLinePregnancy()

      Dim tb As Recordset
      Dim sql As String
      Dim lpc As Integer
      Dim ColName As String
      Dim ShortName As String
      Dim Title As String
      Dim TestNameLength As Integer
      Dim ResultLength As Integer

10    On Error GoTo GetPrintLinePregnancy_Error

20    TestNameLength = 16
30    ResultLength = 26

40    sql = "SELECT COALESCE(LTRIM(RTRIM(LEFT(WCC + '|', CHARINDEX('|', WCC) - 1))), '') W, " & _
            "COALESCE(LTRIM(RTRIM(LEFT(RCC + '|', CHARINDEX('|', RCC) - 1))), '') R, " & _
            "Crystals, Casts, Misc0, Misc1, Misc2, " & _
            "Pregnancy = Case When Pregnancy = 'P' then 'Positive' When Pregnancy = 'N' Then 'Negative' Else COALESCE(Pregnancy, '') End, HCGLevel " & _
            "FROM Urine U, PrintValidLog P WHERE " & _
            "U.SampleID = '" & Val(RP.SampleID) & "' " & _
            "AND U.SampleID = P.SampleID " & _
            "AND P.Department = 'U' " & _
            "AND P.Valid = 1"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80        ColName = "Pregnancy"
90        ShortName = "Pregnancy Test"
100       If Trim$(tb(ColName) & "") <> "" Then
110           lpc = UBound(udtPL)
120           lpc = lpc + 1
130           ReDim Preserve udtPL(0 To lpc)
140           udtPL(lpc).Title = "PREGNANCY"
150           udtPL(lpc).TestName = Left$(ShortName & Space$(TestNameLength), TestNameLength)
160           udtPL(lpc).Result = Left$(tb(ColName) & Space$(ResultLength), ResultLength)
170           udtPL(lpc).LineToPrint = NORMAL10 & _
                                       Space$(4) & _
                                       FormatString(Left$(ShortName & Space$(TestNameLength), TestNameLength) & _
                                                    tb(ColName), 42)
180       End If

190       ColName = "HCGLevel"
200       ShortName = "HCG Level"
210       If Trim$(tb(ColName) & "") <> "" Then
220           AddNotAccreditedTest "HCG", True
230           lpc = UBound(udtPL)
240           lpc = lpc + 1
250           ReDim Preserve udtPL(0 To lpc)
260           udtPL(lpc).Title = "PREGNANCY"
270           udtPL(lpc).TestName = Left$(ShortName & Space$(TestNameLength), TestNameLength)
280           udtPL(lpc).Result = Left$(tb(ColName) & " IU/L" & Space$(ResultLength), ResultLength)
290           udtPL(lpc).LineToPrint = NORMAL10 & _
                                       Space$(4) & _
                                       FormatString(Left$(ShortName & Space$(TestNameLength), TestNameLength) & _
                                                    tb(ColName) & " IU/L", 42)
300       End If
310       UpdatePrintValidLog RP.SampleID, "URINE"
320   End If

330   Exit Sub

GetPrintLinePregnancy_Error:

      Dim strES As String
      Dim intEL As Integer

340   intEL = Erl
350   strES = Err.Description
360   LogError "modNewMicro", "GetPrintLinePregnancy", intEL, strES, sql

End Sub


Private Sub GetPrintLineInterimHeading()

      Dim lpc As Integer
      'PRINT LINE FOR REPORT TYPE

10    On Error GoTo GetPrintLineInterimHeading_Error

20    lpc = UBound(udtPL) + 1
30    ReDim Preserve udtPL(0 To lpc)
40    If RP.FinalInterim = "F" Then
50        udtPL(lpc).LineToPrint = BOLD10 & Space(35) & _
                                   TITLEFONT & "^FontSize12^" & FormatString(UCase$("Final Report"), 12, , AlignCenter) & _
                                   BOLD10 & Space(35)
60    Else
70        udtPL(lpc).LineToPrint = BOLD10 & Space(34) & _
                                   TITLEFONT & "^FontSize12^" & FormatString(UCase$("Interim Report"), 14, , AlignCenter) & _
                                   BOLD10 & Space(34)
80    End If
90    lpc = UBound(udtPR) + 1
100   ReDim Preserve udtPR(0 To lpc)
110   udtPR(lpc).LineToPrint = ""

      'Add extra blank line after heading
120   lpc = UBound(udtPL) + 1
130   ReDim Preserve udtPL(0 To lpc)
140   udtPL(lpc).LineToPrint = ""
150   lpc = UBound(udtPR) + 1
160   ReDim Preserve udtPR(0 To lpc)
170   udtPR(lpc).LineToPrint = ""



180   Exit Sub

GetPrintLineInterimHeading_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "modNewMicro", "GetPrintLineInterimHeading", intEL, strES


End Sub


Public Function DoPrint(Optional PrintA4 As Boolean = False) As Boolean

      Dim TotalPages As Integer
      Dim PageNumber As Integer
      Dim RecordLine As Integer
      Dim StartLine  As Integer
      Dim StopLine   As Integer
      Dim rtb        As RichTextBox
      Dim PrintTitle As Boolean
      Dim ResultsPerPage As Integer
      Dim PrintTime  As String
      Dim f          As Integer
      Dim i          As Integer
      Dim TotalCommentLines
      Dim tb         As Recordset
      Dim sql        As String
      Dim MicroSite  As String



10    On Error GoTo DoPrint_Error

20    If PrintA4 Then
30        MaxCommentLines = 20
40    Else
50        MaxCommentLines = 8
60    End If
70    ClearUdtHeading
80    ABExists = False
90    NotAccreditedText = ""

100   NumberOfABToPrint = GetOptionSetting("NumberOfABToPrint", 10)

110   With udtHeading
120       .SampleID = RP.SampleID
130       .Dept = "Microbiology"
140       .Ward = RP.Ward & ""
150       .Clinician = RP.Clinician & ""
160       .GP = RP.GP & ""
170       .GpClin = RP.Clinician & ""
180       .DocumentNo = GetOptionSetting("MicroDocumentNo", "")

190   End With

      'Change font size for faxing
200   If RP.FaxNumber <> "" Then
210       NORMAL10 = "^FontNameCourier New^^Bold-^^Underline-^^Italic-^^Colour0^^FontSize9^"
220       BOLD10 = "^FontNameCourier New^^Bold+^^Underline-^^Italic-^^Colour0^^FontSize9^"
230       BOLD11 = "^FontNameCourier New^^Bold+^^Underline-^^Italic-^^Colour0^^FontSize9^"
240   Else
250       NORMAL10 = "^FontNameCourier New^^Bold-^^Underline-^^Italic-^^Colour0^^FontSize10^"
260       BOLD10 = "^FontNameCourier New^^Bold+^^Underline-^^Italic-^^Colour0^^FontSize10^"
270       BOLD11 = "^FontNameCourier New^^Bold+^^Underline-^^Italic-^^Colour0^^FontSize11^"
280       TITLEFONT = "^FontNameCourier New^^Bold+^^Italic-^^Colour0^^FontSize10^"
290   End If

300   PrintTime = Format$(Now, "dd/MMM/yyyy HH:nn:ss")

310   ReDim udtPL(0 To 0)
320   ReDim udtPR(0 To 0)

330   ResultsPerPage = Val(GetOptionSetting("ResultsPerPage", "24"))      'only max of 23 lines can be printed on a page.

340   GetPrintLineInterimHeading
350   MicroSite = GetMicroSite(RP.SampleID)

360   If UCase$(MicroSite) = UCase(ListText("MicroNotAccredited", "EF")) Then
370       AddNotAccreditedTest MicroSite, False
380   End If
390   If UCase$(MicroSite) = UCase(ListText("MicroNotAccredited", "BC")) Then
400       AddNotAccreditedTest MicroSite, False
410   End If
420   If UCase$(MicroSite) = UCase(ListText("MicroNotAccredited", "HVS")) Then
430       AddNotAccreditedTest "HVS hays grading", False
440   End If


450   If UCase$(MicroSite) = "BLOOD CULTURE" Then


460       If RP.WardPrint = False Or ValidStatus4MicroDept(RP.SampleID, "B") = True Then
470           PrintSensitivitiesBloodCulture RP.SampleID
480       End If
490       TotalCommentLines = GetCommentLineCount(RP.SampleID, 82)
500       NotAccreditedText = LTrim(RTrim(NotAccreditedText))
510       If NotAccreditedText <> "" Then TotalCommentLines = TotalCommentLines + 1
520       If GetOptionSetting("MicroSiteNote" & MicroSite, "") <> "" Then TotalCommentLines = TotalCommentLines + 1
530       PrintBlankLinesForComments TotalCommentLines
540       If TotalCommentLines > 0 Then
550           PrintBlankLinesForComments TotalCommentLines
560           GetPrintLineComments "COMMENTS", "Demographic", 82
570           GetPrintLineComments "COMMENTS", "MicroCS", 82
580           GetPrintLineComments "COMMENTS", "MicroConsultant", 82
590           If NotAccreditedText <> "" Then GetPrintLineNotAccreditedTests (MicroSite)
600       End If
610   Else

620       FillG
630       GetPrintLineRSV
640       GetPrintLineFaeces
650       GetPrintLineRedSub
660       PrintFluidAppearance
          '460       GetPrintLineBlank udtPL
670       PrintFluids
680       If RP.WardPrint = False Or ValidStatus4MicroDept(RP.SampleID, "U") = True Then
690           GetPrintLineMicroscopy

700           If SysOptShowIQ200(0) = True Then
710               GetPrintLineIQ200
720           End If
730       End If
740       If RP.WardPrint = False Or ValidStatus4MicroDept(RP.SampleID, "D") = True Then
750           PrintSensitivitiesOther RP.SampleID
760       End If
770       If RP.WardPrint = False Or ValidStatus4MicroDept(RP.SampleID, "U") = True Then
780           GetPrintLinePregnancy
790       End If
800       GetPrintLineSemen

810       TotalCommentLines = GetCommentLineCount(RP.SampleID, 82)
820       NotAccreditedText = LTrim(RTrim(NotAccreditedText))
830       If NotAccreditedText <> "" Then TotalCommentLines = TotalCommentLines + 1
840       If GetOptionSetting("MicroSiteNote" & MicroSite, "") <> "" Then TotalCommentLines = TotalCommentLines + 1
850       PrintBlankLinesForComments TotalCommentLines
860       If TotalCommentLines > 0 Then
              
870           GetPrintLineComments "COMMENTS", "Demographic", 82
880           GetPrintLineComments "COMMENTS", "CSFFluid", 82
890           GetPrintLineComments "COMMENTS", "MicroGeneral", 82
900           GetPrintLineComments "COMMENTS", "Semen", 82
910           GetPrintLineComments "COMMENTS", "MicroCS", 82           'Medical scientist comments
920           GetPrintLineComments "COMMENTS", "MicroCDiff", 82
930           GetPrintLineComments "COMMENTS", "MicroConsultant", 82
940           If NotAccreditedText <> "" Then GetPrintLineNotAccreditedTests (MicroSite)
950       End If

960   End If



      'Format Titles for printing
970   FormatTitles udtPL

      'Make both arrays same index.
980   If UBound(udtPL) > UBound(udtPR) Then
990       For i = UBound(udtPR) + 1 To UBound(udtPL)
1000          ReDim Preserve udtPR(0 To i)
1010          udtPR(i).LineToPrint = ""
1020      Next
1030  ElseIf UBound(udtPR) > UBound(udtPL) Then
1040      For i = UBound(udtPL) + 1 To UBound(udtPR)
1050          ReDim Preserve udtPL(0 To i)
1060          udtPL(i).LineToPrint = FormatString("", 46)
1070      Next
1080  End If

1090  Set rtb = frmRichText.rtb

1100  If UBound(udtPL) > 0 Then

1110      NumberOfTitles = CountNumberOfTitles()

1120      TotalPages = ((UBound(udtPL) - 1 + NumberOfTitles) \ ResultsPerPage) + 1

1130      For PageNumber = 1 To TotalPages
1140          If UCase$(GetMicroSite(RP.SampleID)) = "BLOOD CULTURE" Then
1150              PrintTitle = False
1160          Else
1170              PrintTitle = False
1180          End If

1190          PrintHeadingNew PageNumber, TotalPages
1200          StartLine = (PageNumber - 1) * ResultsPerPage + 1
1210          StopLine = StartLine + (ResultsPerPage - 1)

1220          For RecordLine = StartLine To StopLine
1230              If RecordLine <= UBound(udtPL) Then
1240                  If Not PrintTitle Then
1250                      If udtPL(RecordLine).Title <> udtPL(RecordLine - 1).Title Then
1260                          PrintTitle = True
1270                      End If
1280                  End If

1290                  PrintResultLine RecordLine, rtb, PrintTitle
1300                  PrintTitle = False
1310              End If
1320          Next
1330          PrintFooterMicroRTB RP.SampleDate, RP.Rundate
1340          rtb.SelStart = 0
1350          rtb.SelLength = 10000000#

1360          If RP.FaxNumber <> "" Then
1370              f = FreeFile
1380              Open SysOptFax(0) & RP.SampleID & "Micro.doc" For Output As f
1390              Print #f, rtb
1400              Close f
1410              SendFax RP.FaxNumber, RP.SampleID, SysOptFax(0) & RP.SampleID & "Micro.doc"
1420          Else
                  'Masood 19_Feb_2013
1430              If RP.PrintAction = "Print" Or RP.PrintAction = "" Or RP.PrintAction = "PrintSaveFinal" Or RP.PrintAction = "PrintSaveTemp" Then
      '                If IsIDE Then
      '                    Printer.PaperSize = 11
      '                    Printer.Orientation = 2    ' PORTRAIT = 1, Landscope=2
      '                End If
1440                  rtb.SelPrint Printer.hdc
1450              End If
1460          End If

              '        SaveRTF PageNumber, rtb, PrintTime 'Masood 19_Feb_2013 Commented
              'Masood 19_Feb_2013
              Dim status As String
1470          If RP.PrintAction = "" Then
1480              status = "PRINTED"
              
              
1490          ElseIf RP.PrintAction = "PrintSaveFinal" Or RP.PrintAction = "SaveFinal" Or RP.PrintAction = "SaveBlank" Then
1500              If RP.FinalInterim = "F" Then
1510                  status = "RELEASED"
1520              Else
1530                  status = "INTERIM"
1540              End If
1550          End If

1560          If RP.PrintAction = "" Or RP.PrintAction = "SaveFinal" Or RP.PrintAction = "PrintSaveFinal" Then

                  '            If RP.PrintAction = "" Then
                  '                'If report is printed from NetAcquire of Ward Enquiry do not overwrite existing report
                  '                Set tb = New Recordset
                  '                sql = "SELECT Count(*) AS Cnt FROM Reports WHERE SampleID = " & RP.SampleID
                  '                RecOpenServer 0, tb, sql
                  '
                  '                If tb!Cnt = 0 Then
                  '                    SaveRTF PageNumber, rtb, PrintTime, status
                  '                End If
                  '
                  '            Else
                  '                SaveRTF PageNumber, rtb, PrintTime, status
                  '            End If
1570              SaveRTF PageNumber, rtb, PrintTime, status
1580          ElseIf RP.PrintAction = "SaveBlank" Then
1590              rtb.SelText = ""
1600              rtb.Text = ""
1610              PrintTextRTB rtb, "                  " & vbCrLf
1620              PrintTextRTB rtb, "Report Amended" & vbCrLf, 18, True, , True, vbBlack
1630              PrintTextRTB rtb, "                  " & vbCrLf
1640              PrintTextRTB rtb, "Results Pending", 18, True, , True, vbBlack
1650              SaveRTF PageNumber, rtb, PrintTime, status
1660          ElseIf RP.PrintAction = "PrintSaveTemp" Or RP.PrintAction = "SaveTemp" Then
1670              SaveRTFTemp PageNumber, rtb, PrintTime
1680          End If


              'Masood 19_Feb_2013
1690      Next

1700      DoPrint = True

1710  End If

1720  Exit Function

DoPrint_Error:

      Dim strES      As String
      Dim intEL      As Integer

1730  intEL = Erl
1740  strES = Err.Description
1750  LogError "modNewMicro", "DoPrint", intEL, strES

End Function



Private Sub PrintFluidAppearance()

      Dim sql As String
      Dim tb As Recordset
      Dim Site As String
      Dim TestNameLength As Integer
      Dim ResultLength As Integer

10    On Error GoTo PrintFluidAppearance_Error

20    TestNameLength = 15
30    ResultLength = 30

40    sql = "SELECT Site FROM MicroSiteDetails WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80        Site = Trim(tb!Site & "")
90    Else
100       Site = "Fluid"
110   End If

120   sql = "SELECT * FROM GenericResults WHERE " & _
            "SampleID = '" & RP.SampleID & "' " & _
            "AND TestName =  'FluidAppearance1'"
130   Set tb = New Recordset
140   RecOpenServer 0, tb, sql
150   If Not tb.EOF Then
160       GetPrintLineFluid "FluidAppearance1", "Appearance:", "", "Appearance"
170   End If

180   Exit Sub

PrintFluidAppearance_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "modNewMicro", "PrintFluidAppearance", intEL, strES, sql

End Sub

Private Sub PrintFluids()

      Dim sql As String
      Dim tb As Recordset
      Dim Site As String
      Dim TestName As String
      Dim ShortName As String
      Dim Title As String
      Dim TestNameLength As Integer
      Dim ResultLength As Integer
      Dim n As Integer
      Dim Units As String
      Dim lpc As Integer
      Dim ABEndLine As Integer

10    On Error GoTo PrintFluids_Error

20    TestNameLength = 15
30    ResultLength = 30

40    sql = "SELECT Site FROM MicroSiteDetails WHERE " & _
            "SampleID = '" & RP.SampleID & "'"
50    Set tb = New Recordset
60    RecOpenServer 0, tb, sql
70    If Not tb.EOF Then
80        Site = Trim(tb!Site & "")
90    Else
100       Site = "Fluid"
110   End If

120   sql = "SELECT * FROM GenericResults WHERE " & _
            "SampleID = '" & RP.SampleID & "' " & _
            "AND (TestName LIKE  'CSF%' " & _
            "     OR TestName LIKE  'Fluid%' )"
130   Set tb = New Recordset
140   RecOpenServer 0, tb, sql
150   If Not tb.EOF Then
160       For n = 1 To 13
170           TestName = Choose(n, "FluidAppearance0", "FluidGram", "FluidGram(2)", _
                                "FluidZN", "FluidLeishmans", "FluidWetPrep", "FluidCrystals", _
                                "FluidGlucose", "FluidProtein", "FluidAlbumin", "FluidGlobulin", _
                                "FluidLDH", "FluidAmylase")
180           ShortName = Choose(n, "Cell Count:", "Gram:", "", _
                                 "ZN Stain:", "Leishmans:", "Wet Prep:", "Crystals:", _
                                 "Glucose:", "Protein:", "Albumin:", "Globulin:", _
                                 "LDH:", "Amylase:")
190           Units = Choose(n, "", "", "", "", "", "", "", _
                             "mmol/L", "g/L", "g/L", "g/L", _
                             "IU/L", "IU/L")
200           GetPrintLineFluid TestName, ShortName, Units, "MICROSCOPY"

210       Next
220       GetPrintLineFluid "CSFGlucose", "Glucose:", "mmol/L", "CSF"
230       GetPrintLineFluid "CSFProtein", "Protein:", "g/L", "CSF"


240       sql = "SELECT * FROM GenericResults WHERE " & _
                "SampleID = '" & RP.SampleID & "' " & _
                "AND TestName LIKE  'CSFH%'"
250       Set tb = New Recordset
260       RecOpenServer 0, tb, sql
270       If Not tb.EOF Then
280           ABEndLine = (frmMain.g.Rows - 6) + 2

290           If UBound(udtPL) < ABEndLine Then
300               For lpc = UBound(udtPL) + 1 To ABEndLine
310                   ReDim Preserve udtPL(0 To lpc)
320                   udtPL(lpc).LineToPrint = FormatString("", 46)
330               Next lpc
340           End If

350           lpc = UBound(udtPL)
360           lpc = lpc + 1
370           ReDim Preserve udtPL(0 To lpc)
380           udtPL(lpc).Title = "Specimen                  " & FormatString("1", 12, , AlignCenter) & _
                                 FormatString("2", 12, , AlignCenter) & _
                                 FormatString("3", 12, , AlignCenter)
390           udtPL(lpc).LineToPrint = NORMAL10 & Space$(4) & "RCC                   " & GetPrintLineHaemCSF(0) & "/cmm"

400           lpc = lpc + 1
410           ReDim Preserve udtPL(0 To lpc)
420           udtPL(lpc).Title = "Specimen                  " & FormatString("1", 12, , AlignCenter) & _
                                 FormatString("2", 12, , AlignCenter) & _
                                 FormatString("3", 12, , AlignCenter)
430           udtPL(lpc).LineToPrint = NORMAL10 & Space$(4) & "WCC                   " & GetPrintLineHaemCSF(3) & "/cmm"

440           lpc = lpc + 1
450           ReDim Preserve udtPL(0 To lpc)
460           udtPL(lpc).Title = "Specimen                  " & FormatString("1", 12, , AlignCenter) & _
                                 FormatString("2", 12, , AlignCenter) & _
                                 FormatString("3", 12, , AlignCenter)
470           udtPL(lpc).LineToPrint = NORMAL10 & Space$(4) & "Polymorphic           " & GetPrintLineHaemCSF(6) & "%"

480           lpc = lpc + 1
490           ReDim Preserve udtPL(0 To lpc)
500           udtPL(lpc).Title = "Specimen                  " & FormatString("1", 12, , AlignCenter) & _
                                 FormatString("2", 12, , AlignCenter) & _
                                 FormatString("3", 12, , AlignCenter)
510           udtPL(lpc).LineToPrint = NORMAL10 & Space$(4) & "Mononucleated         " & GetPrintLineHaemCSF(9) & "%"

520       End If
530   End If

540   GetPrintLineFluid "PneumococcalAT", "Pneumococcal Antigen:", "", "ANTIGEN TESTS"
550   GetPrintLineFluid "LegionellaAT", "Legionella Antigen:", "", "ANTIGEN TESTS"
560   GetPrintLineFluid "FungalElements", "Fungal Elements:", "", "KOH PREPARATION"

570   UpdatePrintValidLog RP.SampleID, "FLUIDS"

580   Exit Sub

PrintFluids_Error:

      Dim strES As String
      Dim intEL As Integer

590   intEL = Erl
600   strES = Err.Description
610   LogError "modNewMicro", "PrintFluids", intEL, strES, sql

End Sub

Private Function GetPrintLineHaemCSF(ByVal pNumber As Integer) _
        As String

      Dim sql As String
      Dim tb As Recordset
      Dim s As String

10    On Error GoTo GetPrintLineHaemCSF_Error

20    sql = "SELECT Result FROM GenericResults WHERE " & _
            "SampleID = '" & RP.SampleID & "' " & _
            "AND TestName = 'CSFHAEM" & Format$(pNumber) & "'"
30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60        s = s & FormatString("", 12, , AlignCenter)
70    Else
80        s = s & FormatString(tb!Result & "", 12, , AlignCenter)
90    End If

100   sql = "SELECT Result FROM GenericResults WHERE " & _
            "SampleID = '" & RP.SampleID & "' " & _
            "AND TestName = 'CSFHAEM" & Format$(pNumber + 1) & "'"
110   Set tb = New Recordset
120   RecOpenServer 0, tb, sql
130   If tb.EOF Then
140       s = s & FormatString("", 12, , AlignCenter)
150   Else
160       s = s & FormatString(tb!Result & "", 12, , AlignCenter)
170   End If

180   sql = "SELECT Result FROM GenericResults WHERE " & _
            "SampleID = '" & RP.SampleID & "' " & _
            "AND TestName = 'CSFHAEM" & Format$(pNumber + 2) & "'"
190   Set tb = New Recordset
200   RecOpenServer 0, tb, sql
210   If tb.EOF Then
220       s = s & FormatString("", 12, , AlignCenter)
230   Else
240       s = s & FormatString(tb!Result & "", 12, , AlignCenter)
250   End If

260   GetPrintLineHaemCSF = s
270   Exit Function

GetPrintLineHaemCSF_Error:

      Dim strES As String
      Dim intEL As Integer

280   intEL = Erl
290   strES = Err.Description
300   LogError "modNewMicro", "GetPrintLineHaemCSF", intEL, strES, sql

End Function

Private Sub GetPrintLineFluid(ByVal Parameter As String, _
                              ByVal DisplayName As String, _
                              ByVal Units As String, _
                              ByVal Title As String)

      Dim sql As String
      Dim tb As Recordset
      Dim TestNameLength As Integer
      Dim ResultLength As Integer
      Dim UnitLength As Integer
      Dim lpc As Integer

10    On Error GoTo GetPrintLineFluid_Error


20    Select Case UCase(Title)
      Case "APPEARANCE":
30        TestNameLength = 12
40        ResultLength = 29
50        UnitLength = 5
60    Case "CSF":
70        TestNameLength = 12
80        ResultLength = 28
90        UnitLength = 6
100   Case "ANTIGEN TESTS":
110       TestNameLength = 22
120       ResultLength = 19
130       UnitLength = 5
140   Case "KOH PREPARATION":
150       TestNameLength = 22
160       ResultLength = 19
170       UnitLength = 5
180   Case "MICROSCOPY"
190       TestNameLength = 12
200       ResultLength = 28
210       UnitLength = 6
220   Case Default
230       TestNameLength = 22
240       ResultLength = 19
250       UnitLength = 5

260   End Select

      'If Len(DisplayName) > 16 Then
      '    TestNameLength = 26
      'End If

270   sql = "SELECT * FROM GenericResults WHERE " & _
            "SampleID = '" & RP.SampleID & "' " & _
            "AND TestName = '" & Parameter & "'"

280   Set tb = New Recordset
290   RecOpenServer 0, tb, sql
300   If Not tb.EOF Then
' --------------farhan----------------
310       If UCase(Parameter) = UCase("CSFGlucose") Then
320           AddResultToArray Title, tb!Result & "", udtPL, ResultLength, FormatString(DisplayName, TestNameLength), "  " & FormatString("(2.22 - 3.89 )", 15) & "  " & FormatString(Units, UnitLength)
              'elseif
330       ElseIf UCase(Parameter) = UCase("CSFProtein") Then
340           AddResultToArray Title, tb!Result & "", udtPL, ResultLength, FormatString(DisplayName, TestNameLength), "  " & FormatString("(0.10 - 0.45 )", 15) & "  " & FormatString(Units, UnitLength)
'==============farhan==================
350       Else
360           AddResultToArray Title, tb!Result & "", udtPL, ResultLength, FormatString(DisplayName, TestNameLength), FormatString(Units, UnitLength)
370       End If
          '    lpc = UBound(udtPL)
          '    lpc = lpc + 1
          '    ReDim Preserve udtPL(0 To lpc)
          '    udtPL(lpc).Title = Title
          '    udtPL(lpc).LineToPrint = NORMAL10 & _
               '                             Space$(4) & _
               '                             FormatString(Left$(DisplayName & Space$(TestNameLength), TestNameLength) & tb!Result & " " & Units, 42)
380   End If

390   Exit Sub

GetPrintLineFluid_Error:

      Dim strES As String
      Dim intEL As Integer

400   intEL = Erl
410   strES = Err.Description
420   LogError "modNewMicro", "GetPrintLineFluid", intEL, strES, sql

End Sub

Private Function CountNumberOfTitles() As Integer

      Dim Counter As Integer
      Dim t As String
      Dim y As Integer

10    On Error GoTo CountNumberOfTitles_Error

20    t = ""
30    Counter = 0
40    For y = 1 To UBound(udtPL())
50        If udtPL(y).Title <> t Then
60            t = udtPL(y).Title
70            Counter = Counter + 1
80        End If
90    Next

100   CountNumberOfTitles = Counter

110   Exit Function

CountNumberOfTitles_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "modNewMicro", "CountNumberOfTitles", intEL, strES

End Function

Private Sub PrintControlCode(ByRef Start As Integer, ByVal s As String)

      Dim n As Integer

10    On Error GoTo PrintControlCode_Error
20    n = InStr(Start + 1, s, "^")
30    If n = 0 Then Exit Sub

40    s = Mid$(s, Start + 1, n - Start - 1)

50    Start = n

60    If UCase$(Left$(s, 8)) = "FONTNAME" Then
70        frmRichText.rtb.SelFontName = Mid$(s, 9)
80    ElseIf UCase$(Left$(s, 8)) = "FONTSIZE" Then
90        frmRichText.rtb.SelFontSize = Mid$(s, 9)
100   ElseIf UCase$(Left$(s, 6)) = "COLOUR" Then
110       frmRichText.rtb.SelColor = Val(Mid$(s, 7))
120   ElseIf UCase$(Left$(s, 4)) = "BOLD" Then
130       frmRichText.rtb.SelBold = Mid$(s, 5) = "+"
140   ElseIf UCase$(Left$(s, 9)) = "UNDERLINE" Then
150       frmRichText.rtb.SelUnderline = Mid$(s, 10) = "+"
160   ElseIf UCase$(Left$(s, 6)) = "ITALIC" Then
170       frmRichText.rtb.SelItalic = Mid$(s, 7) = "+"
180   End If


190   Exit Sub

PrintControlCode_Error:

      Dim strES As String
      Dim intEL As Integer

200   intEL = Erl
210   strES = Err.Description
220   LogError "modNewMicro", "PrintControlCode", intEL, strES

End Sub

Public Sub PrintResultLine(ByVal Index As Integer, _
                           ByRef rtb As RichTextBox, _
                           ByVal PrintTitle As Boolean)
      'Control Codes
      '  ^BOLD+^
      '  ^BOLD-^
      '  ^COLOURnn..^
      Dim n As Integer
      Dim s As String
      Dim TitleToPrint As String
      Dim LineToPrint As String

10    On Error GoTo PrintResultLine_Error

20    With rtb
30        If PrintTitle Then
40            .SelFontName = "Courier New"
50            .SelFontSize = 10
60            .SelBold = True
70            .SelColor = vbBlack
80            .SelUnderline = False
90            TitleToPrint = LTrim(RTrim(udtPL(Index).Title & "    " & udtPR(Index).Title))
100           For n = 1 To Len(TitleToPrint)
110               s = Mid$(TitleToPrint, n, 1)
120               If s = "^" Then
130                   PrintControlCode n, TitleToPrint
140               Else
150                   .SelText = s
160               End If
170           Next
180           If LTrim(RTrim(udtPR(Index).LineToPrint)) = "" Then
190               .SelText = vbCrLf
200           End If
210           .SelBold = False
220           .SelUnderline = False
230       End If
240       Debug.Print udtPL(Index).LineToPrint
250       udtPL(Index).LineToPrint = ApplyPrintRule(udtPL(Index).LineToPrint)
260       udtPR(Index).LineToPrint = ApplyPrintRule(udtPR(Index).LineToPrint)
270       LineToPrint = udtPL(Index).LineToPrint & "    " & udtPR(Index).LineToPrint

280       For n = 1 To Len(LineToPrint)

290           s = Mid$(LineToPrint, n, 1)
300           If s = "^" Then
310               PrintControlCode n, LineToPrint
320           Else
330               .SelText = s
340           End If
350       Next
360       .SelText = vbCrLf
370   End With

380   Exit Sub

PrintResultLine_Error:

      Dim strES As String
      Dim intEL As Integer

390   intEL = Erl
400   strES = Err.Description
410   LogError "modNewMicro", "PrintResultLine", intEL, strES

End Sub
Private Sub SaveRTF(ByVal PageNumber As Integer, _
                    ByVal rtb As RichTextBox, _
                    ByVal PrintTime As String, Optional status As String)

      Dim sql As String
      Dim tb As Recordset
      Dim Dept As String

10    On Error GoTo SaveRTF_Error

20    Dept = UCase$(Left$(RP.Department, 1))

30    sql = "SELECT * FROM Reports WHERE 0 = 1"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    tb.AddNew
70    tb!SampleID = RP.SampleID
80    tb!Name = udtHeading.Name
90    tb!Dept = Dept
100   tb!Initiator = RP.Initiator
110   tb!PrintTime = PrintTime
120   tb!RepNo = (PageNumber - 1) & Dept & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
130   tb!PageNumber = PageNumber - 1
140   tb!Report = rtb.TextRTF
150   tb!Printer = Printer.DeviceName
160   tb!status = status
170   tb.Update

180   Exit Sub

SaveRTF_Error:

      Dim strES As String
      Dim intEL As Integer

190   intEL = Erl
200   strES = Err.Description
210   LogError "modNewMicro", "SaveRTF", intEL, strES, sql

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SaveRTFTemp
' Author    : XPMUser
' Date      : 2/20/2014
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub SaveRTFTemp(ByVal PageNumber As Integer, _
                        ByVal rtb As RichTextBox, _
                        ByVal PrintTime As String)

      Dim sql As String
      Dim tb As Recordset
      Dim Dept As String



10    On Error GoTo SaveRTFTemp_Error


20    Dept = UCase$(Left$(RP.Department, 1))

30    sql = "SELECT * FROM UnauthorisedReports WHERE 0 = 1"
40    Set tb = New Recordset
50    RecOpenServer 0, tb, sql
60    tb.AddNew
70    tb!SampleID = RP.SampleID
80    tb!Name = udtHeading.Name
90    tb!Dept = Dept
100   tb!Initiator = RP.Initiator
110   tb!PrintTime = PrintTime
120   tb!RepNo = (PageNumber - 1) & Dept & RP.SampleID & Format(PrintTime, "ddMMyyyyhhmmss")
130   tb!PageNumber = PageNumber - 1
140   tb!Report = rtb.TextRTF
150   tb!Printer = Printer.DeviceName
160   tb.Update



170   Exit Sub


SaveRTFTemp_Error:

      Dim strES As String
      Dim intEL As Integer

180   intEL = Erl
190   strES = Err.Description
200   LogError "modNewMicro", "SaveRTFTemp", intEL, strES, sql

End Sub


Public Function MicroValidatedBy(SampleID As String) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo MicroValidatedBy_Error

20    If RP.Department = "Z" Then
30        sql = "Select Username As AuthorisedBy From SemenResults Where SampleID = '%sampleid' And Valid = 1 "
40    Else
50        sql = "Select Top 1 ValidatedBy As AuthorisedBy From PrintValidLog Where SampleID = '%sampleid' And Valid = 1 " & _
                "Order By ValidatedDateTime Desc"
60    End If
70    sql = Replace(sql, "%sampleid", SampleID)
80    Set tb = New Recordset
90    RecOpenClient 0, tb, sql
100   If tb.EOF Then
110       MicroValidatedBy = ""
120   Else
130       MicroValidatedBy = tb!AuthorisedBy & ""
140   End If

150   Exit Function

MicroValidatedBy_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "modNewMicro", "MicroValidatedBy", intEL, strES, sql

End Function


Private Function ColHasValue(Col As Integer, g As MSFlexGrid) As Boolean

      Dim i As Integer

10    On Error GoTo ColHasValue_Error

20    ColHasValue = False
30    With g

40        For i = 0 To g.Rows - 1
50            If .TextMatrix(i, Col) <> "" Then
60                ColHasValue = True
70                Exit For
80            End If
90        Next i

100   End With

110   Exit Function

ColHasValue_Error:

      Dim strES As String
      Dim intEL As Integer

120   intEL = Erl
130   strES = Err.Description
140   LogError "modNewMicro", "ColHasValue", intEL, strES

End Function

Private Function IsolateHasAntibiotics(Col As Integer, g As MSFlexGrid) As Boolean

      Dim i As Integer

10    On Error GoTo IsolateHasAntibiotics_Error

20    IsolateHasAntibiotics = False

30    If g.Rows < 6 Then Exit Function

40    With g

50        For i = 6 To g.Rows - 1
60            If .TextMatrix(i, Col) <> "" Then
70                IsolateHasAntibiotics = True
80                Exit For
90            End If
100       Next i

110   End With


120   Exit Function

IsolateHasAntibiotics_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "modNewMicro", "IsolateHasAntibiotics", intEL, strES

End Function

Private Function BottleLinesToPrint(BottleType As Integer) As Integer

      Dim i As Integer
      Dim ColIndex As Integer
      Dim LinesCountCol1 As Integer
      Dim LinesCountCol2 As Integer

10    On Error GoTo BottleLinesToPrint_Error

20    LinesCountCol1 = 0
30    LinesCountCol2 = 0

40    Select Case BottleType
      Case 1: ColIndex = 1
50    Case 2: ColIndex = 3
60    Case 3: ColIndex = 5
70    End Select
80    With frmMain.g
90        For i = 6 To .Rows - 1
100           If .TextMatrix(i, ColIndex) <> "" Then
110               LinesCountCol1 = LinesCountCol1 + 1
120           End If
130           If .TextMatrix(i, ColIndex + 1) <> "" Then
140               LinesCountCol2 = LinesCountCol2 + 1
150           End If
160       Next i
170   End With

180   If LinesCountCol1 > LinesCountCol2 Then
190       BottleLinesToPrint = LinesCountCol1 + 7
200   Else
210       BottleLinesToPrint = LinesCountCol2 + 7
220   End If

230   Exit Function

BottleLinesToPrint_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "modNewMicro", "BottleLinesToPrint", intEL, strES

End Function

Public Function GetGramIdentification(ByVal SampleIDWithOffset As Double, ByVal Isolate As Byte) As String

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GetGramIdentification_Error

20    sql = "Select Gram From UrineIdent Where SampleID = '%sampleid' And Isolate = %isolate"
30    sql = Replace(sql, "%sampleid", SampleIDWithOffset)
40    sql = Replace(sql, "%isolate", Isolate)

50    Set tb = New Recordset
60    RecOpenClient 0, tb, sql

70    If Not tb.EOF Then
80        GetGramIdentification = tb!Gram & ""
90    Else
100       GetGramIdentification = ""
110   End If

120   Exit Function

GetGramIdentification_Error:

      Dim strES As String
      Dim intEL As Integer

130   intEL = Erl
140   strES = Err.Description
150   LogError "modNewMicro", "GetGramIdentification", intEL, strES, sql

End Function

Public Function GetBloodCultureBottleInterval(ByVal SampleIDWithOffset As String, ByVal BottleType As String) As String

      Dim tb As Recordset
      Dim sql As String
      Dim TypeOfTest As String

10    On Error GoTo GetBloodCultureBottleInterval_Error


20    Select Case BottleType
      Case "Aerobic": TypeOfTest = GetOptionSetting("BcAerobicBottle", "BSA")
30    Case "Anaerobic": TypeOfTest = GetOptionSetting("BcAnarobicBottle", "BSN")
40    Case "Fan": TypeOfTest = GetOptionSetting("BcFanBottle", "BFA")
50    Case Else: TypeOfTest = ""
60    End Select

70    If TypeOfTest = "" Then
80        GetBloodCultureBottleInterval = ""
90        Exit Function
100   End If


110   sql = "Select TTD From BloodCultureResults Where SampleID = '%sampleid' And TypeOfTest = '%typeoftest'"
120   sql = Replace(sql, "%sampleid", SampleIDWithOffset)
130   sql = Replace(sql, "%typeoftest", TypeOfTest)

140   Set tb = New Recordset
150   RecOpenClient 0, tb, sql

160   If tb.EOF Then
170       GetBloodCultureBottleInterval = ""
180   Else
190       GetBloodCultureBottleInterval = tb!TTD & ""
200   End If

210   Exit Function

GetBloodCultureBottleInterval_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "modNewMicro", "GetBloodCultureBottleInterval", intEL, strES, sql

End Function

Public Function BloodCultureBottleIsPositive(ByVal SampleIDWithOffset As Double, ByVal BottleType As String) As Boolean

      Dim tb As Recordset
      Dim sql As String
      Dim TypeOfTest As String

10    On Error GoTo BloodCultureBottleIsPositive_Error

20    Select Case BottleType
      Case "Aerobic": TypeOfTest = GetOptionSetting("BcAerobicBottle", "BSA")
30    Case "Anaerobic": TypeOfTest = GetOptionSetting("BcAnarobicBottle", "BSN")
40    Case "Fan": TypeOfTest = GetOptionSetting("BcFanBottle", "BFA")
50    Case Else: TypeOfTest = ""
60    End Select

70    If TypeOfTest = "" Then
80        BloodCultureBottleIsPositive = False
90        Exit Function
100   End If

110   sql = "Select Result From BloodCultureResults Where SampleID = '%sampleid' And TypeOfTest = '%typeoftest'"
120   sql = Replace(sql, "%sampleid", SampleIDWithOffset)
130   sql = Replace(sql, "%typeoftest", TypeOfTest)

140   Set tb = New Recordset
150   RecOpenClient 0, tb, sql

160   If tb.EOF Then
170       BloodCultureBottleIsPositive = False
180   Else
190       BloodCultureBottleIsPositive = (tb!Result & "" = "+")
200   End If


210   Exit Function

BloodCultureBottleIsPositive_Error:

      Dim strES As String
      Dim intEL As Integer

220   intEL = Erl
230   strES = Err.Description
240   LogError "modNewMicro", "BloodCultureBottleIsPositive", intEL, strES, sql

End Function


Private Function GramIdentificationExists(ByVal SampleIDWithOffset As Double) As Boolean

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo GramIdentificationExists_Error

20    sql = "Select Count(*) as Cnt From UrineIdent Where SampleID = '%sampleid'"
30    sql = Replace(sql, "%sampleid", SampleIDWithOffset)

40    Set tb = New Recordset
50    RecOpenClient 0, tb, sql

60    GramIdentificationExists = (tb!Cnt > 0)
70    Exit Function

GramIdentificationExists_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "modNewMicro", "GramIdentificationExists", intEL, strES, sql

End Function

Private Sub FormatTitles(ByRef udt() As LineToPrint)

      Dim i As Integer
      Dim udtTemp() As LineToPrint
      Dim PrevTitle As String
      Dim lpc As Integer
      Dim MaxLength As Integer

10    On Error GoTo FormatTitles_Error

20    MaxLength = 46
30    If Not ABExists Then
40        MaxLength = MaxLength + 38
50    End If

60    ReDim udtTemp(0 To 0)
70    PrevTitle = ""
80    lpc = 0

90    For i = 1 To UBound(udt)

100       If PrevTitle <> udt(i).Title And udt(i).Title <> "" Then
110           PrevTitle = udt(i).Title
120           lpc = UBound(udtTemp) + 1
130           ReDim Preserve udtTemp(0 To lpc)
140           If Len(udt(i).Title) > MaxLength Then
150               udtTemp(lpc).LineToPrint = TITLEFONT & FormatString(udt(i).Title, Len(Trim(udt(i).Title)))
160           Else
170               udtTemp(lpc).LineToPrint = TITLEFONT & FormatString(udt(i).Title, Len(Trim(udt(i).Title))) & _
                                             BOLD10 & Space(MaxLength - Len(Trim(udt(i).Title)))
180           End If
190       End If
200       lpc = UBound(udtTemp) + 1
210       ReDim Preserve udtTemp(0 To lpc)
220       udtTemp(lpc).LineToPrint = udt(i).LineToPrint

230   Next i

240   udt = udtTemp


250   Exit Sub

FormatTitles_Error:

      Dim strES As String
      Dim intEL As Integer

260   intEL = Erl
270   strES = Err.Description
280   LogError "modNewMicro", "FormatTitles", intEL, strES

End Sub

Private Sub PrintBlankLinesForComments(TotalCommentLines)
      Dim PrintedLines As Integer
      Dim lpc As Integer
      Dim BlankLines As Integer
      Dim NumberOfTitles As Integer
      Dim StartLine As Integer
      Dim StopLine As Integer
      Dim LastLine As Integer

10    On Error GoTo PrintBlankLinesForComments_Error

      'TotalCommentLines = TotalCommentLines + 1       'add one line for comment title
20    NumberOfTitles = CountNumberOfTitles() + 1      'add one line for comment title
30    PrintedLines = UBound(udtPL)
40    LastLine = GetOptionSetting("ResultsPerPage", 24) - NumberOfTitles
50    BlankLines = LastLine - TotalCommentLines - 1


60    If PrintedLines < BlankLines Then
70        StartLine = UBound(udtPL) + 1
80        StopLine = BlankLines
90        If StopLine <= NumberOfABToPrint Then StopLine = NumberOfABToPrint + 2
100       For lpc = StartLine To StopLine
110           ReDim Preserve udtPL(0 To lpc)
120           udtPL(lpc).LineToPrint = NORMAL10 & FormatString("", 46)

130       Next lpc
140   End If

150   Exit Sub

PrintBlankLinesForComments_Error:

      Dim strES As String
      Dim intEL As Integer

160   intEL = Erl
170   strES = Err.Description
180   LogError "modNewMicro", "PrintBlankLinesForComments", intEL, strES

End Sub


'---------------------------------------------------------------------------------------
' Procedure : AddResultToArray
' DateTime  : 11/02/2011 14:52
' Author    : Babar Shahzad
' Purpose   : Providing line length, it will automatically spread results to multiline if necessory
'           Returns no of lines used.
'---------------------------------------------------------------------------------------
'
Private Function AddResultToArray(Title As String, Result As String, ByRef udt() As LineToPrint, _
                                  LineLength As Integer, Optional TestName As String, Optional Units As String) As Integer


      Dim LastWordIndex As Integer
      Dim lpc As Integer
      Dim s As String
      Dim TotalLines As Integer
      Dim TestNamePrinted As Boolean
      Dim UnitsPrinted As Boolean
      Dim MaxLength As Integer


10    On Error GoTo AddResultToArray_Error

20    MaxLength = 46
30    If Not ABExists Then
40        LineLength = LineLength + 38
50        MaxLength = MaxLength + 38
60    End If

70    TestNamePrinted = (TestName = "")
80    UnitsPrinted = (Units = "")


90    If Len(Result) > LineLength Then
100       While Not Len(Result) <= LineLength
110           LastWordIndex = InStrRev(Left(Result, LineLength), " ")
120           s = FormatString(Left(Result, LastWordIndex - 1), LineLength)
130           lpc = UBound(udt) + 1
140           ReDim Preserve udt(0 To lpc)
150           udt(lpc).Title = Title
160           If TestNamePrinted = False Then
170               s = TestName & s
180               TestNamePrinted = True
190           Else
200               s = Space(Len(TestName)) & s
210           End If
220           If UnitsPrinted = False Then
230               If Len(Trim(s)) < MaxLength + Len(Units) Then
240                   s = FormatString(Trim(s) & Units, MaxLength + Len(Units))
250               Else
260                   s = s & Units
270               End If

280               UnitsPrinted = True
290           Else
300               s = s & Space(Len(Units))
310           End If
320           udt(lpc).LineToPrint = NORMAL10 & FormatString(s, MaxLength)
330           TotalLines = TotalLines + 1
340           Result = Mid(Result, LastWordIndex + 1, Len(Result))

350       Wend
360   End If

370   s = FormatString(Result, LineLength)

380   lpc = UBound(udt) + 1
390   ReDim Preserve udt(0 To lpc)
400   udt(lpc).Title = Title
410   If TestNamePrinted = False Then
420       s = TestName & s
430       TestNamePrinted = True
440   Else
450       s = Space(Len(TestName)) & s
460   End If
470   If UnitsPrinted = False Then
480       If Len(Trim(s)) < MaxLength + Len(Units) Then
490           s = FormatString(Trim(s) & Units, MaxLength + Len(Units))
500       Else
510           s = s & Units
520       End If

530       UnitsPrinted = True
540   Else
550       s = s & Space(Len(Units))
560   End If
570   udt(lpc).LineToPrint = NORMAL10 & FormatString(s, MaxLength)
580   TotalLines = TotalLines + 1

590   AddResultToArray = TotalLines

600   Exit Function

AddResultToArray_Error:

      Dim strES As String
      Dim intEL As Integer

610   intEL = Erl
620   strES = Err.Description
630   LogError "modNewMicro", "AddResultToArray", intEL, strES

End Function

Public Function ValidStatus4MicroDept(ByVal SampleIDWithOffset As Double, ByVal strDept As String) As Boolean

      Dim tb As Recordset
      Dim sql As String

10    On Error GoTo ValidStatus4MicroDept_Error

20    sql = "SELECT COALESCE(Valid, 0) AS Valid FROM PrintValidLog WHERE " & _
            "SampleID = '" & SampleIDWithOffset & "' " & _
            "AND Department = '" & strDept & "'"

30    Set tb = New Recordset
40    RecOpenServer 0, tb, sql
50    If tb.EOF Then
60        ValidStatus4MicroDept = False
70    Else
80        ValidStatus4MicroDept = tb!Valid
90    End If

100   Exit Function

ValidStatus4MicroDept_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "basMicro", "ValidStatus4MicroDept", intEL, strES, sql

End Function

Private Sub GetPrintLineBlack(udt() As LineToPrint)

      Dim lpc As Integer

10    On Error GoTo GetPrintLineBlack_Error

20    lpc = UBound(udtPL) + 1
30    ReDim Preserve udtPL(0 To lpc)
40    udtPL(lpc).LineToPrint = NORMAL10 & FormatString("", 46)

50    Exit Sub

GetPrintLineBlack_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "modNewMicro", "GetPrintLineBlack", intEL, strES

End Sub

Private Sub AddNotAccreditedTest(ByVal TestToAdd As String, Optional ByVal GetListText As Boolean = False)

      Dim LT         As String
10    On Error GoTo AddNotAccreditedTest_Error

20    If GetListText Then
30        LT = ListText("MicroNotAccredited", TestToAdd)
40    Else
50        LT = TestToAdd
60    End If
70    If InStr(1, UCase(NotAccreditedText), UCase(LT)) = 0 Then
80        If LTrim(RTrim(NotAccreditedText)) = "" Then
90            NotAccreditedText = LT
100       Else
110           NotAccreditedText = NotAccreditedText & ", " & LT
120       End If
130   End If

140   Exit Sub
AddNotAccreditedTest_Error:

150   LogError "modNewMicro", "AddNotAccreditedTest", Erl, Err.Description


End Sub

