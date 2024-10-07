VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BIEResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private mcolBIEResults As New Collection
Private mBIEResults As New BIEResults

Public Function Add(br As BIEResult) As BIEResult

10    On Error GoTo Add_Error

20    mcolBIEResults.Add br

30    Set Add = br

40    Exit Function

Add_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "BIEResults", "Add", intEL, strES


End Function

Public Property Get BIEResults() As BIEResults
  
10    On Error GoTo BIEResults_Error

20    Set BIEResults = mBIEResults

30    Exit Property

BIEResults_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "BIEResults", "BIEResults", intEL, strES


End Property

Public Sub Clear()

      Dim n As Integer

10    On Error GoTo Clear_Error

20    For n = 1 To mcolBIEResults.Count
30      mcolBIEResults.Remove 1
40    Next

50    Exit Sub

Clear_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "BIEResults", "Clear", intEL, strES


End Sub

Public Function Count() As Long
  
10    On Error GoTo Count_Error

20    Count = mcolBIEResults.Count

30    Exit Function

Count_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "BIEResults", "Count", intEL, strES


End Function

Public Function Item(ByVal X As Long) _
                     As BIEResult
Attribute Item.VB_UserMemId = 0
  
10    On Error GoTo Item_Error

20    On Error Resume Next

30    Set Item = mcolBIEResults(X)

40    Exit Function

Item_Error:

      Dim strES As String
      Dim intEL As Integer

50    intEL = Erl
60    strES = Err.Description
70    LogError "BIEResults", "Item", intEL, strES


End Function

Public Function Load(ByVal Discipline As String, _
                     ByVal SampleID As String, _
                     ByVal ResultOrRepeat As String, _
                     ByVal Cn As Long, _
                     ByVal Cat As String, _
                     ByVal Rundate As String) _
                     As BIEResults
      'Discipline is either "Bio", "Imm" or "End"

      Dim BRs As New BIEResults
      Dim br As BIEResult
      Dim tb As New Recordset
      Dim sql As String
      Dim DoB As String
      Dim DaysOld As Long
      Dim SELECTNormalRange As String
      Dim SELECTFlagRange As String
      Dim TableName As String
      Dim Sex As String
      Dim tbRange As Recordset
      Dim tbNewIDX As Recordset

10    On Error GoTo Load_Error

20    DaysOld = 9125

30    If SampleID = "" Then Exit Function

40    If UCase(Discipline) = "BIO" Or UCase(Discipline) = "BGA" Then Cat = ""

50    TableName = Discipline & ResultOrRepeat

60    sql = "SELECT DoB, Sex,rundate from Demographics WHERE " & _
            "SampleID = '" & Val(SampleID) & "'"
70    Set tb = Cnxn(Cn).Execute(sql)
80    If Not tb.EOF Then
90        If IsDate(tb!DoB) Then
100           DoB = Format$(tb!DoB, "dd/mmm/yyyy")
110           DaysOld = DateDiff("d", DoB, tb!Rundate)
120           If DaysOld = 0 Then DaysOld = 1
130       End If
140       Sex = Left$(UCase$(Trim$(tb!Sex & "")), 1)
150       Select Case Left$(UCase$(Trim$(tb!Sex & "")), 1)
              Case "M": SELECTNormalRange = " MaleLow as Low, MaleHigh as High, "
160               SELECTFlagRange = " FlagMaleLow as FlagLow, FlagMaleHigh as FlagHigh, "
170           Case "F": SELECTNormalRange = " FemaleLow as Low, FemaleHigh as High, "
180               SELECTFlagRange = " FlagFemaleLow as FlagLow, FlagFemaleHigh as FlagHigh, "
190           Case Else: SELECTNormalRange = " FemaleLow as Low, MaleHigh as High, "
200               SELECTFlagRange = " FlagFemaleLow as FlagLow, FlagMaleHigh as FlagHigh, "
210       End Select
220   Else
230       SELECTNormalRange = " FemaleLow as Low, MaleHigh as High, "
240       SELECTFlagRange = " FlagFemaleLow as FlagLow, FlagMaleHigh as FlagHigh, "
250   End If


260   sql = "SELECT COALESCE(R.DefIndex,0) DefIndex, COALESCE(X.NormalLow, 0) Low, " & _
            "COALESCE(X.NormalHigh, 9999) High, COALESCE(X.FlagLow, 0) FlagLow, COALESCE(X.FlagHigh, 9999) FlagHigh, " & _
            "COALESCE(X.PlausibleLow, 0) PlausibleLow, COALESCE(X.PlausibleHigh, 9999) PlausibleHigh, " & _
            "LongName, ShortName, DoDelta, DeltaLimit, CheckTime, Printable, " & _
            "COALESCE(ShowLessThan, 0) ShowLessThan, COALESCE(ShowMoreThan, 0) ShowMoreThan, " & _
            "DP, PrintPriority, " & _
            "R.SampleID, R.Code, R.Result, " & _
            "COALESCE(R.Valid, 0) AS Valid, COALESCE(R.Printed, 0) Printed, " & _
            "R.RunTime, R.RunDate, R.Operator, R.Flags, R.Units, " & _
            "R.SampleType, R.Analyser, R.Faxed, R.Authorised, " & _
            "R.Comment AS Comment, R.PC,D.Hospital "
270   If UCase$(Discipline) = "IMM" Then
280       sql = sql & ", prnrr "
290   End If
300   sql = sql & "FROM " & TableName & " R JOIN " & Discipline & "TestDefinitions D ON R.Code = D.Code " & _
            "LEFT JOIN " & Discipline & "DefIndex X ON R.DefIndex = X.DefIndex " & _
            "WHERE " & _
            "SampleID = '" & SampleID & "' AND D.category = '" & Cat & "' " & _
            "AND R.Code = D.Code AND D.Inuse = 1 " & _
            "AND Valid =  1 " & _
            "AND AgeFromDays <= " & DaysOld & " " & _
            "AND AgeToDays >= " & DaysOld & " " & _
            "AND ShortName NOT IN " & _
            "  ( SELECT Parameter FROM PrintInhibit WHERE " & _
            "    SampleID = '" & SampleID & "' AND Discipline = '" & Discipline & "') "




310   sql = sql & "and R.SampleType = D.SampleType "
      'If P = gNOTPRINTED And v = gNOTVALID Then
      '    sql = sql & "and Printed = 0 and Valid = 0 "
      'ElseIf P = gNOTPRINTED And v = gVALID Then
      '    sql = sql & "and Printed = 0 and Valid = 1 "
      'ElseIf P = gNOTPRINTED And v = gDONTCARE Then
      '    sql = sql & "and Printed = 0 "
      'ElseIf P = gPRINTED And v = gNOTVALID Then
      '    sql = sql & "and Printed = 1 and Valid = 0 "
      'ElseIf P = gPRINTED And v = gVALID Then
      '    sql = sql & "and Printed = 1 and Valid = 1 "
      'ElseIf P = gPRINTED And v = gDONTCARE Then
      '    sql = sql & "and Printed = 1 "
      'ElseIf P = gDONTCARE And v = gNOTVALID Then
      '    sql = sql & "and Valid = 0 "
      'ElseIf P = gDONTCARE And v = gVALID Then
      '    sql = sql & "and Valid = 1 "
      'End If

320   sql = sql & "Order by PrintPriority asc"
330   Set tb = New Recordset
340   RecOpenServer Cn, tb, sql    '  RecOpenClient 0,tb, Sql
350   Do While Not tb.EOF
360       Set br = New BIEResult
370       With br
380           .SampleID = Trim(tb!SampleID & "")
390           .DefIndex = tb!DefIndex
400           .Code = Trim(tb!Code & "")
410           .Result = Trim(tb!Result & "")
420           .Operator = Trim(tb!Operator & "")
430           .Rundate = Format$(tb!Rundate, "dd/mm/yyyy")
440           .RunTime = Format$(tb!RunTime, "dd/mm/yyyy hh:mm")
450           .Units = Trim(tb!Units & "")
460           If Trim(tb!Printed & "") <> "" Then .Printed = IIf(tb!Printed, True, False)
470           If Trim(tb!Valid & "") <> "" Then .Valid = IIf(tb!Valid, True, False)
480           .Flags = Trim(tb!Flags & "")
490           .SampleType = Trim(tb!SampleType & "")
500           .Low = IIf(IsNull(tb!Low), 0, tb!Low)
510           .FlagLow = IIf(IsNull(tb!FlagLow), 0, tb!FlagLow)
520           .PlausibleLow = IIf(IsNull(tb!PlausibleLow), 0, tb!PlausibleLow)
530           .High = IIf(IsNull(tb!High), 9999, tb!High)
540           .FlagHigh = IIf(IsNull(tb!FlagHigh), 9999, tb!FlagHigh)
550           .PlausibleHigh = IIf(IsNull(tb!PlausibleHigh), 99999, tb!PlausibleHigh)
560           .Printformat = IIf(IsNull(tb!DP), 0, tb!DP)
570           .ShortName = Trim(tb!ShortName & "")
580           .LongName = Trim(tb!LongName & "")
590           If tb!DoDelta & "" <> "" Then .DoDelta = tb!DoDelta Else .DoDelta = False
600           .DeltaLimit = IIf(IsNull(tb!DeltaLimit), 9999, tb!DeltaLimit)
610           .Analyser = Trim(tb!Analyser & "")
620           If Discipline = "Imm" Then
630               .PrnRR = IIf(IsNull(tb!PrnRR), True, tb!PrnRR)
640           End If
650           .Comment = Trim(tb!Comment & "")
660           .Hospital = Trim(tb!Hospital & "")
670           .Pc = Trim(tb!Pc & "")
680           If IsNull(tb!CheckTime) Then
690               .CheckTime = 1
700           Else
710               .CheckTime = tb!CheckTime
720           End If
730           .ShowLessThan = tb!ShowLessThan
740           .ShowMoreThan = tb!ShowMoreThan
750           .Printable = tb!Printable
              '   .NormalLow = tb!NormalLow
              '   .NormalHigh = tb!NormalHigh
              '   .NormalUsed = tb!NormalUsed

760           If .DefIndex = 0 Then
770               If DoB <> "" And Sex <> "" Then
780                   sql = "SELECT " & _
                            SELECTNormalRange & SELECTFlagRange & _
                            "COALESCE(PlausibleLow, 0) PlausibleLow, " & _
                            "COALESCE(PlausibleHigh, 9999) PlausibleHigh " & _
                            "FROM " & Discipline & "TestDefinitions  " & _
                            "WHERE category = '" & Cat & "' " & _
                            "AND Code = '" & .Code & "' " & _
                            "AND AgeFromDays <= " & DaysOld & " " & _
                            "AND AgeToDays >= " & DaysOld & " "
790                   Set tbRange = New Recordset
800                   RecOpenServer 0, tbRange, sql
810                   If Not tbRange.EOF Then
820                       .Low = tbRange!Low
830                       .High = tbRange!High
840                       .FlagLow = tbRange!FlagLow
850                       .FlagHigh = tbRange!FlagHigh
860                       .PlausibleLow = tbRange!PlausibleLow
870                       .PlausibleHigh = tbRange!PlausibleHigh

880                       sql = "SELECT * FROM " & Discipline & "DefIndex " & _
                                "WHERE NormalLow = '" & .Low & "' " & _
                                "AND NormalHigh = '" & .High & "' " & _
                                "AND FlagLow = '" & .FlagLow & "' " & _
                                "AND FlagHigh = '" & .FlagHigh & "' " & _
                                "AND PlausibleLow = '" & .PlausibleLow & "' " & _
                                "AND PlausibleHigh = '" & .PlausibleHigh & "' "

890                       Set tbNewIDX = New Recordset
900                       RecOpenClient 0, tbNewIDX, sql
910                       If Not tbNewIDX.EOF Then
920                           .DefIndex = tbNewIDX!DefIndex
930                       Else

940                           sql = "INSERT INTO " & Discipline & "DefIndex " & _
                                    "( NormalLow, NormalHigh, FlagLow, FlagHigh, " & _
                                    "  PlausibleLow, PlausibleHigh, AutoValLow, AutoValHigh) " & _
                                    "VALUES ( " & _
                                    .Low & ", " & .High & ", " & .FlagLow & ", " & .FlagHigh & ", " & _
                                    .PlausibleLow & ", " & .PlausibleHigh & ", 0,9999) "
950                           Cnxn(0).Execute sql

960                           sql = "SELECT MAX(DefIndex) NewIndex FROM " & Discipline & "DefIndex"
970                           Set tbNewIDX = New Recordset
980                           RecOpenClient 0, tbNewIDX, sql
990                           .DefIndex = tbNewIDX!NewIndex

1000                      End If

1010                      sql = "UPDATE " & TableName & " " & _
                                "SET DefIndex = '" & .DefIndex & "' " & _
                                "WHERE SampleID = '" & .SampleID & "' " & _
                                "AND Code = '" & .Code & "'"
1020                      Cnxn(0).Execute sql

1030                  End If
1040              End If
1050          End If
1060          BRs.Add br
1070      End With
1080      tb.MoveNext
1090  Loop

1100  If BRs.Count <> 0 Then
1110      Set Load = BRs
1120  Else
1130      Set Load = Nothing
1140  End If
1150  Set br = Nothing
1160  Set BRs = Nothing


1170  Exit Function



1180  Exit Function

Load_Error:

      Dim strES As String
      Dim intEL As Integer


1190  intEL = Erl
1200  strES = Err.Description
1210  LogError "BIEResults", "Load", intEL, strES, sql


End Function

'Public Function Load(ByVal Discipline As String, _
 '                      ByVal SampleID As String, _
 '                      ByVal ResultOrRepeat As String, _
 '                      ByVal Cn As Integer, _
 '                      ByVal Cat As String, _
 '                      ByVal Rundate As String) _
 '                      As BIEResults
'      'Discipline is either "Bio", "Imm" or "End"
'
'      Dim BRs As New BIEResults
'      Dim br As BIEResult
'      Dim tb As New Recordset
'      Dim sql As String
'      Dim Dob As String
'      Dim DaysOld As Long
'      Dim SELECTNormalRange As String
'      Dim SELECTFlagRange As String
'      Dim TableName As String
'
'10    On Error GoTo Load_Error
'
'20    DaysOld = 9125
'
'30    If UCase(Discipline) = "BIO" Or UCase(Discipline) = "BGA" Then Cat = ""
'
'40    TableName = Discipline & ResultOrRepeat
'
'50    sql = "SELECT DoB, Sex, rundate, SampleDate FROM Demographics WHERE " & _
 '             "SampleID = '" & SampleID & "'"
'60    Set tb = Cnxn(Cn).Execute(sql)
'70    If Not tb.EOF Then
'80      If IsDate(tb!Dob) Then
'90        Dob = Format$(tb!Dob, "dd/mmm/yyyy")
'100       DaysOld = DateDiff("d", Dob, Format(tb!SampleDate, "dd/mmm/yyyy"))
'110     End If
'120     Select Case Left$(UCase$(Trim$(tb!sex & "")), 1)
'          Case "M":  SELECTNormalRange = " COALESCE(MaleLow, 0) AS Low, COALESCE(MaleHigh, 9999) AS High, "
'130                  SELECTFlagRange = " COALESCE(FlagMaleLow, 0) AS FlagLow, COALESCE(FlagMaleHigh, 9999) AS FlagHigh, "
'140       Case "F":  SELECTNormalRange = " COALESCE(FemaleLow, 0) AS Low, COALESCE(FemaleHigh, 9999) AS High, "
'150                  SELECTFlagRange = " COALESCE(FlagFemaleLow, 0) AS FlagLow, COALESCE(FlagFemaleHigh, 9999) AS FlagHigh, "
'160       Case Else: SELECTNormalRange = " COALESCE(FemaleLow, 0) AS Low, COALESCE(MaleHigh, 9999) AS High, "
'170                  SELECTFlagRange = " COALESCE(FlagFemaleLow, 0) AS FlagLow, COALESCE(FlagMaleHigh, 9999) AS FlagHigh, "
'180     End Select
'190   Else
'200     SELECTNormalRange = " COALESCE(FemaleLow, 0) AS Low, COALESCE(MaleHigh, 9999) AS High, "
'210     SELECTFlagRange = " COALESCE(FlagFemaleLow, 0) AS FlagLow, COALESCE(FlagMaleHigh, 9999) AS FlagHigh, "
'220   End If
'
'230   sql = "SELECT LongName, Printable, ShortName, DoDelta, COALESCE(DeltaLimit, 9999) AS DeltaLimit, " & _
 '               "COALESCE(ShowLessThan, 0) ShowLessThan, COALESCE(ShowMoreThan, 0) ShowMoreThan, " & _
 '             SELECTNormalRange & SELECTFlagRange & _
 '             "COALESCE(PlausibleLow, 0) AS PlausibleLow, COALESCE(PlausibleHigh, 9999) AS PlausibleHigh, " & _
 '             "COALESCE(DP, 0) AS DP, R.*, PrintPriority "
'240         If Discipline = "Imm" Then
'250           sql = sql & ", prnrr "
'260         End If
'270   sql = sql & "FROM " & TableName & " R, " & Discipline & "TestDefinitions D WHERE " & _
 '             "SampleID = '" & SampleID & "' AND D.category = '" & Cat & "' " & _
 '             "AND R.Code = D.Code " & _
 '             "AND D.AgeFromDays <= " & DaysOld & " " & _
 '             "AND D.AgeToDays >= " & DaysOld & " AND Printable = 1 " & _
 '             "AND R.SampleType = D.SampleType " & _
 '             "AND ShortName NOT IN " & _
 '             "  ( SELECT Parameter FROM PrintInhibit WHERE " & _
 '             "    SampleID = '" & SampleID & "' AND Discipline = '" & Discipline & "')" & _
 '             "ORDER BY PrintPriority asc"
'280   Set tb = New Recordset
'290   RecOpenServer Cn, tb, sql '  RecOpenClient 0,tb, Sql
'300   Do While Not tb.EOF
'310     Set br = New BIEResult
'320     With br
'330       .SampleID = Trim(tb!SampleID & "")
'340       .Code = Trim(tb!Code & "")
'350       .Result = Trim(tb!Result & "")
'360       .Operator = Trim(tb!Operator & "")
'370       .Rundate = Format$(tb!Rundate, "dd/mm/yyyy")
'380       .RunTime = Format$(tb!RunTime, "dd/mm/yyyy hh:mm")
'390       .Units = Trim(tb!Units & "")
'400       .Printed = IIf(IsNull(tb!Printed), 0, tb!Printed)
'410       .Valid = IIf(tb!Valid, True, False)
'420       .Flags = Trim(tb!Flags & "")
'430       .SampleType = Trim(tb!SampleType & "")
'440       .Low = tb!Low
'450       .FlagLow = tb!FlagLow
'460       .PlausibleLow = tb!PlausibleLow
'470       .High = tb!High
'480       .FlagHigh = tb!FlagHigh
'490       .PlausibleHigh = tb!PlausibleHigh
'500       .Printable = IIf(tb!Printable, True, False)
'510       .Printformat = tb!DP
'520       .ShortName = Trim(tb!ShortName & "")
'530       .LongName = Trim(tb!LongName & "")
'540       If tb!DoDelta & "" <> "" Then .DoDelta = tb!DoDelta Else .DoDelta = False
'550       .DeltaLimit = tb!DeltaLimit
'560       .Analyser = Trim(tb!Analyser & "")
'570       .ShowLessThan = tb!ShowLessThan
'580       .ShowMoreThan = tb!ShowMoreThan
'590       If Discipline = "Imm" Then
'600           If IsNull(tb!PrnRR) Then
'610               .PrnRR = True
'620           Else
'630               .PrnRR = tb!PrnRR
'640           End If
'650       End If
'660       .Comment = Trim(tb!Comment & "")
'670       .Pc = Trim(tb!Pc & "")
'680       BRs.Add br
'690     End With
'700     tb.MoveNext
'710   Loop
'
'720   If BRs.Count <> 0 Then
'730     Set Load = BRs
'740   Else
'750     Set Load = Nothing
'760   End If
'770   Set br = Nothing
'780   Set BRs = Nothing
'
'790   Exit Function
'
'Load_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'800   intEL = Erl
'810   strES = Err.Description
'820   LogError "BIEResults", "Load", intEL, strES, sql
'
'End Function
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
  
10    On Error GoTo NewEnum_Error

20    Set NewEnum = mcolBIEResults.[_NewEnum]

30    Exit Function

NewEnum_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "BIEResults", "NewEnum", intEL, strES


End Function



Public Sub RemoveItem(ByVal X As Long)

10    On Error GoTo RemoveItem_Error

20    mcolBIEResults.Remove X

30    Exit Sub

RemoveItem_Error:

      Dim strES As String
      Dim intEL As Integer

40    intEL = Erl
50    strES = Err.Description
60    LogError "BIEResults", "RemoveItem", intEL, strES


End Sub


Public Sub Save(ByVal Discipline As String, _
                ByVal BRs As BIEResults)
      'Discipline is either "Bio", "Imm" or "End"

      Dim tb As New Recordset
      Dim sql As String
      Dim br As BIEResult
      Dim lngSampleID As Double





10    On Error GoTo Save_Error

20    For Each br In BRs
30      With br
40        lngSampleID = Format$(Val(.SampleID))
50        sql = "SELECT * FROM " & Discipline & "Results WHERE " & _
                "SampleID = '" & lngSampleID & "' " & _
                "and Code = '" & br.Code & "'"
60        Set tb = New Recordset
70        RecOpenClient 0, tb, sql
80        If Not tb.EOF Then
90          sql = "SELECT * FROM " & Discipline & "Repeats WHERE " & _
                  "SampleID = '" & lngSampleID & "'"
100         Set tb = New Recordset
110         RecOpenClient 0, tb, sql
120       End If
130       tb.AddNew
140       tb!SampleID = lngSampleID
150       tb!Rundate = Format$(.Rundate, "dd/mmm/yyyy")
160       If Trim$(.RunTime) <> "" Then
170         tb!RunTime = Format$(.RunTime, "dd/mmm/yyyy hh:mm")
180       End If
190       tb!Code = .Code
200       tb!Result = .Result
210       tb!Units = .Units
220       tb!Printed = .Printed
230       tb!Valid = .Valid
240       tb!Faxed = 0
250       tb!Analyser = .Analyser
260       tb!SampleType = .SampleType
270       tb.Update

280       sql = "Delete FROM " & Discipline & "Requests WHERE " & _
                "SampleID = '" & lngSampleID & "' " & _
                "and Code = '" & br.Code & "'"
290       Cnxn(0).Execute sql
    
300     End With
310   Next

320   Set br = Nothing


330   Exit Sub

Save_Error:

      Dim strES As String
      Dim intEL As Integer

340   intEL = Erl
350   strES = Err.Description
360   LogError "BIEResults", "Save", intEL, strES, sql


End Sub




