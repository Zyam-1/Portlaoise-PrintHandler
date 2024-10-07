Attribute VB_Name = "modDbDesign"
Option Explicit


'Public Sub CheckIQ200RepeatsInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckIQ200RepeatsInDb_Error
'
'20    If IsTableInDatabase("IQ200Repeats") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE IQ200Repeats " & _
'              "( SampleID  numeric(18, 0) NOT NULL, " & _
'              "  TestCode  nvarchar(50), " & _
'              "  ShortName nvarchar(50), " & _
'              "  LongName  nvarchar(50), " & _
'              "  Range nvarchar(50), " & _
'              "  Result nvarchar(50), " & _
'              "  WorklistPrinted bit, " & _
'              "  DateTimeOfRecord datetime NOT NULL DEFAULT getdate(), " & _
'              "  Validated bit, " & _
'              "  ValidatedBy nvarchar(50), " & _
'              "  Printed bit, " & _
'              "  PrintedBy nvarchar(50), " & _
'              "  Counter numeric(18, 0) IDENTITY(1,1) NOT NULL )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckIQ200RepeatsInDb_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckIQ200RepeatsInDb", intEL, strES, sql
'
'End Sub
'
'
'Public Sub CheckPhoresisRequestsInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckPhoresisRequestsInDb_Error
'
'20    If IsTableInDatabase("PhoresisRequests") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE [dbo].[PhoresisRequests] ( " & _
'              "[AnalysisProgramCode] [char](1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & _
'              "[PhoresisSampleNumber] [char](4) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
'              "[PatientID] [nvarchar](15) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & _
'              "[PatientName] [nvarchar](30) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & _
'              "[DoB] [smalldatetime] NULL, " & _
'              "[Sex] [char](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
'              "[Age] [char](3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
'              "[Department] [nvarchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL, " & _
'              "[SampleDate] [smalldatetime] NULL, " & _
'              "[Concentration] [float] NULL, " & _
'              "[DateTimeOfRecord] [datetime] NOT NULL CONSTRAINT [DF_PhoresisRequests_DateTimeOfRecord]  DEFAULT (getdate()), " & _
'              "[Counter] [bigint] IDENTITY(1,1) NOT NULL, " & _
'              "[UserName] [nvarchar](50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL, " & _
'              "[Programmed] [tinyint] NOT NULL )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckPhoresisRequestsInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckPhoresisRequestsInDb", intEL, strES, sql
'
'End Sub
'
'
'
'
'
'Public Sub CheckPrintValidLogInDb()
'
'      Dim sql As String
'10    On Error GoTo CheckPrintValidLogInDb_Error
'
'20    If IsTableInDatabase("PrintValidLog") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE PrintValidLog " & _
'              "( SampleID  numeric(9), " & _
'              "  Department nvarchar(1), " & _
'              "  Printed tinyint, " & _
'              "  Valid tinyint, " & _
'              "  PrintedBy nvarchar(50), " & _
'              "  PrintedDateTime datetime, " & _
'              "  ValidatedBy nvarchar(50), " & _
'              "  ValidatedDateTime datetime )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckPrintValidLogInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckPrintValidLogInDb", intEL, strES, sql
'
'
'End Sub
'
'Public Sub CheckMicroExternalsInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckMicroExternalsInDb_Error
'
'20    If IsTableInDatabase("MicroExternals") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE MicroExternals " & _
'              "( MicroSID  numeric(9) NOT NULL, " & _
'              "  InHouseSID  numeric(9) NOT NULL, " & _
'              "  OrderGlu bit NOT NULL, " & _
'              "  OrderTP bit NOT NULL, " & _
'              "  OrderAlb bit NOT NULL, " & _
'              "  OrderGlo bit NOT NULL, " & _
'              "  OrderLDH bit NOT NULL, " & _
'              "  OrderAmy bit NOT NULL, " & _
'              "  DateTimeOfRecord datetime NOT NULL DEFAULT getdate(), " & _
'              "  UserName nvarchar(50) NOT NULL, " & _
'              "  RowGUID uniqueidentifier NOT NULL DEFAULT newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckMicroExternalsInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckMicroExternalsInDb", intEL, strES, sql
'
'
'End Sub
'
'Public Sub CheckMicroExternalResultsInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckMicroExternalResultsInDb_Error
'
'20    If IsTableInDatabase("MicroExternalResults") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE MicroExternalResults " & _
'              "( SampleID  numeric(9) NOT NULL, " & _
'              "  TestName  nvarchar(50) NOT NULL, " & _
'              "  SentTo nvarchar(50), " & _
'              "  SentDate datetime, " & _
'              "  InterimReportDate datetime, " & _
'              "  InterimReportComment nvarchar(50), " & _
'              "  FinalReportDate datetime, " & _
'              "  FinalReportComment nvarchar(50), " & _
'              "  DateTimeOfRecord datetime NOT NULL DEFAULT getdate(), " & _
'              "  UserName nvarchar(50) NOT NULL, " & _
'              "  RowGUID uniqueidentifier NOT NULL DEFAULT newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckMicroExternalResultsInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckMicroExternalResultsInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckMicroExternalResultsArcInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckMicroExternalResultsArcInDb_Error
'
'20    If IsTableInDatabase("MicroExternalResultsArc") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE MicroExternalResultsArc " & _
'              "( SampleID  numeric(9) NOT NULL, " & _
'              "  TestName  nvarchar(50) NOT NULL, " & _
'              "  SentTo nvarchar(50), " & _
'              "  SentDate datetime, " & _
'              "  InterimReportDate datetime, " & _
'              "  InterimReportComment nvarchar(50), " & _
'              "  FinalReportDate datetime, " & _
'              "  FinalReportComment nvarchar(50), " & _
'              "  DateTimeOfRecord datetime, " & _
'              "  UserName nvarchar(50) NOT NULL, " & _
'              "  ArchiveDateTime datetime NOT NULL DEFAULT getdate(), " & _
'              "  ArchivedBy nvarchar(50) , " & _
'              "  RowGUID uniqueidentifier NOT NULL DEFAULT newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckMicroExternalResultsArcInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckMicroExternalResultsArcInDb", intEL, strES, sql
'
'End Sub
'Public Sub CheckUrineRequestsInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckUrineRequestsInDb_Error
'
'20    If IsTableInDatabase("UrineRequests") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE UrineRequests " & _
'              "( SampleID  numeric(9) NOT NULL, " & _
'              "  CS bit, " & _
'              "  Pregnancy bit, " & _
'              "  RedSub bit, " & _
'              "  DoNotDisplayInBatchEntry bit, " & _
'              "  DateTimeOfRecord datetime NOT NULL DEFAULT getdate() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckUrineRequestsInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckUrineRequestsInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckUrineRequestsArcInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckUrineRequestsArcInDb_Error
'
'20    If IsTableInDatabase("UrineRequestsArc") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE UrineRequestsArc " & _
'              "( SampleID  numeric(9) NOT NULL, " & _
'              "  CS bit, " & _
'              "  Pregnancy bit, " & _
'              "  RedSub bit, " & _
'              "  DoNotDisplayInBatchEntry bit, " & _
'              "  DateTimeOfRecord datetime NOT NULL DEFAULT getdate(), " & _
'              "  UserName nvarchar(50), " & _
'              "  ArchivedBy nvarchar(50) NOT NULL, " & _
'              "  ArchiveDateTime datetime NOT NULL DEFAULT getdate(), " & _
'              "  RowGUID uniqueidentifier NOT NULL DEFAULT newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckUrineRequestsArcInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckUrineRequestsArcInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckLoggedOnUsersInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckLoggedOnUsersInDb_Error
'
'20    If IsTableInDatabase("LoggedOnUsers") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE dbo.[LoggedOnUsers](" & _
'              "  [MachineName] [nvarchar](50) NULL, " & _
'              "  [UserName] [nvarchar](50) NULL, " & _
'              "  [AppName] [nvarchar](50) NULL)"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckLoggedOnUsersInDb_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckLoggedOnUsersInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckMicroExtLabNameInDb()
'
'      Dim sql As String
'10    On Error GoTo CheckMicroExtLabNameInDb_Error
'
'20    If IsTableInDatabase("MicroExtLabName") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE MicroExtLabName " & _
'              "( LabName nvarchar(50), " & _
'              "  Address0 nvarchar(50), " & _
'              "  Address1 nvarchar(50), " & _
'              "  Address2 nvarchar(50), " & _
'              "  DateTimeOfRecord datetime DEFAULT getdate(), " & _
'              "  RowGUID uniqueidentifier DEFAULT newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckMicroExtLabNameInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckMicroExtLabNameInDb", intEL, strES, sql
'
'
'End Sub
'
'
'Public Sub CheckIsolatesRepeatsInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckIsolatesRepeatsInDb_Error
'
'20    If IsTableInDatabase("IsolatesRepeats") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE IsolatesRepeats " & _
'              "( SampleID  numeric(9), " & _
'              "  IsolateNumber int, " & _
'              "  OrganismGroup nvarchar(50), " & _
'              "  OrganismName nvarchar(50), " & _
'              "  Qualifier nvarchar(50) )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckIsolatesRepeatsInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckIsolatesRepeatsInDb", intEL, strES, sql
'
'
'End Sub
'
'Public Sub CheckLockStatusInDb()
'
'      Dim sql As String
'10    On Error GoTo CheckLockStatusInDb_Error
'
'20    If IsTableInDatabase("LockStatus") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE LockStatus " & _
'              "( SampleID  numeric(9), " & _
'              "  Lock bit, " & _
'              "  DeptIndex int, " & _
'              "  RowGUID uniqueidentifier default newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckLockStatusInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckLockStatusInDb", intEL, strES, sql
'
'
'End Sub
'Public Sub CheckSensitivitiesRepeatsInDb()
'
'      Dim sql As String
'10    On Error GoTo CheckSensitivitiesRepeatsInDb_Error
'
'20    If IsTableInDatabase("SensitivitiesRepeats") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE SensitivitiesRepeats " & _
'              "( SampleID  numeric(9), " & _
'              "  IsolateNumber int, " & _
'              "  AntibioticCode nvarchar(50), " & _
'              "  Result nvarchar(50), " & _
'              "  Report bit, " & _
'              "  CPOFlag nvarchar(1), " & _
'              "  RunDate datetime, " & _
'              "  RunDateTime datetime, " & _
'              "  RSI char(1), " & _
'              "  UserName nvarchar(50), " & _
'              "  Forced bit, " & _
'              "  Secondary bit, " & _
'              "  Valid bit, " & _
'              "  AuthoriserCode nvarchar(50) )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckSensitivitiesRepeatsInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckSensitivitiesRepeatsInDb", intEL, strES, sql
'
'
'End Sub
'
'Public Sub CheckFaecesResultsInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckFaecesResultsInDb_Error
'
'20    If IsTableInDatabase("FaecesResults") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE FaecesResults " & _
'              "( SampleID  numeric NOT NULL, " & _
'              "  TestName nvarchar(50) NOT NULL, " & _
'              "  Result nvarchar(50) NOT NULL, " & _
'              "  UserName nvarchar(50) NOT NULL, " & _
'              "  Valid bit NOT NULL, " & _
'              "  HealthLink tinyint NOT NULL, " & _
'              "  DateTimeOfRecord datetime NOT NULL )"
'40      Cnxn(0).Execute sql
'
'50    End If
'
'60    Exit Sub
'
'CheckFaecesResultsInDb_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckFaecesResultsInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckGenericResultsInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckGenericResultsInDb_Error
'
'20    If IsTableInDatabase("GenericResults") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE GenericResults " & _
'              "( SampleID  numeric(9), " & _
'              "  TestName nvarchar(50), " & _
'              "  Result nvarchar(50) )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckGenericResultsInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckGenericResultsInDb", intEL, strES, sql
'
'End Sub
'
'
'Public Sub CheckFaecesWorksheetInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckFaecesWorksheetInDb_Error
'
'20    If IsTableInDatabase("FaecesWorksheet") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE FaecesWorksheet " & _
'              "( SampleID  numeric(9), " & _
'              "  Day111 nvarchar(50), Day112 nvarchar(50), Day113 nvarchar(50), " & _
'              "  Day121 nvarchar(50), Day122 nvarchar(50), Day123 nvarchar(50), " & _
'              "  Day131 nvarchar(50), Day132 nvarchar(50), Day133 nvarchar(50), " & _
'              "  Day211 nvarchar(50), Day212 nvarchar(50), Day213 nvarchar(50), " & _
'              "  Day221 nvarchar(50), Day222 nvarchar(50), Day223 nvarchar(50), " & _
'              "  Day231 nvarchar(50), Day232 nvarchar(50), Day233 nvarchar(50), " & _
'              "  Day31 nvarchar(50), Day32 nvarchar(50), Day33 nvarchar(50), " & _
'              "  TimeOfRecord datetime, " & _
'              "  Operator nvarchar(50) )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckFaecesWorksheetInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckFaecesWorksheetInDb", intEL, strES, sql
'
'
'End Sub
'
'Public Sub CheckGenericResultsArcInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckGenericResultsArcInDb_Error
'
'20    If IsTableInDatabase("GenericResultsArc") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE GenericResultsArc " & _
'              "( SampleID numeric NOT NULL, " & _
'              "  TestName nvarchar(50), " & _
'              "  Result nvarchar(50), " & _
'              "  UserName nvarchar(50), " & _
'              "  ArchivedBy nvarchar(50), " & _
'              "  ArchiveDateTime datetime default getdate(), " & _
'              "  RowGUID uniqueidentifier default newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckGenericResultsArcInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckGenericResultsArcInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckMicroSiteDetailsArcInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckMicroSiteDetailsArcInDb_Error
'
'20    If IsTableInDatabase("MicroSiteDetailsArc") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE MicroSiteDetailsArc " & _
'              "( SampleID numeric NOT NULL, " & _
'              "  Site nvarchar(50), " & _
'              "  SiteDetails nvarchar(50), " & _
'              "  PCA0 nvarchar(50), " & _
'              "  PCA1 nvarchar(50), " & _
'              "  PCA2 nvarchar(50), " & _
'              "  PCA3 nvarchar(50), " & _
'              "  ArchiveDateTime datetime default getdate(), " & _
'              "  ArchivedBy nvarchar(50), " & _
'              "  rowguid uniqueidentifier default newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckMicroSiteDetailsArcInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckMicroSiteDetailsArcInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckSemenResultsArcInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckSemenResultsArcInDb_Error
'
'20    If IsTableInDatabase("SemenResultsArc") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE SemenResultsArc " & _
'              "( SampleID numeric NOT NULL, " & _
'              "  Volume nvarchar(50), " & _
'              "  SemenCount nvarchar(50), " & _
'              "  MotilityPro nvarchar(50), " & _
'              "  MotilityNonPro nvarchar(50), " & _
'              "  MotilityNonMotile nvarchar(50), " & _
'              "  Consistency nvarchar(50), " & _
'              "  Valid int, " & _
'              "  Operator nvarchar(50), " & _
'              "  Printed int, " & _
'              "  Motility nvarchar(50), " & _
'              "  ArchiveDateTime datetime default getdate(), " & _
'              "  ArchivedBy nvarchar(50), " & _
'              "  rowguid uniqueidentifier default newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckSemenResultsArcInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckSemenResultsArcInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckUrineArcInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckUrineArcInDb_Error
'
'20    If IsTableInDatabase("UrineArc") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE UrineArc " & _
'              "( SampleID numeric NOT NULL, " & _
'              "  Pregnancy nvarchar(50), " & _
'              "  HCGLevel nvarchar(50), " & _
'              "  BenceJones nvarchar(50), " & _
'              "  SG nvarchar(50), " & _
'              "  FatGlobules nvarchar(50), " & _
'              "  pH nvarchar(50), " & _
'              "  Protein nvarchar(50), " & _
'              "  Glucose nvarchar(50), " & _
'              "  Ketones nvarchar(50), " & _
'              "  Urobilinogen nvarchar(50), " & _
'              "  Bilirubin nvarchar(50), " & _
'              "  BloodHb nvarchar(50), " & _
'              "  WCC nvarchar(50), " & _
'              "  RCC nvarchar(50), " & _
'              "  Crystals nvarchar(50), " & _
'              "  Casts nvarchar(50), " & _
'              "  Misc0 nvarchar(50), " & _
'              "  Misc1 nvarchar(50), " & _
'              "  Misc2 nvarchar(50), " & _
'              "  Valid bit, " & _
'              "  Bacteria nvarchar(50), " & _
'              "  [Count] nvarchar(50), " & _
'              "  HealthLink tinyint, "
'
'40      sql = sql & "Printed int, " & _
'              "  ArchiveDateTime datetime default getdate(), " & _
'              "  ArchivedBy nvarchar(50), " & _
'              "  rowguid uniqueidentifier default newid(), " & _
'              "  UserName nvarchar(50) )"
'50      Cnxn(0).Execute sql
'60    End If
'
'70    Exit Sub
'
'CheckUrineArcInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'80    intEL = Erl
'90    strES = Err.Description
'100   LogError "modDbDesign", "CheckUrineArcInDb", intEL, strES, sql
'
'End Sub
'
'Public Sub CheckUrineIdentArcInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckUrineIdentArcInDb_Error
'
'20    If IsTableInDatabase("UrineIdentArc") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE UrineIdentArc " & _
'              "( SampleID numeric NOT NULL, " & _
'              "  Gram nvarchar(50), " & _
'              "  WetPrep nvarchar(50), " & _
'              "  Coagulase nvarchar(50), " & _
'              "  Catalase nvarchar(50), " & _
'              "  Oxidase nvarchar(50), " & _
'              "  API0 nvarchar(50), " & _
'              "  API1 nvarchar(50), " & _
'              "  Ident0 nvarchar(50), " & _
'              "  Ident1 nvarchar(50), " & _
'              "  Rapidec nvarchar(50), " & _
'              "  Chromogenic nvarchar(50), " & _
'              "  Reincubation nvarchar(50), " & _
'              "  UrineSensitivity nvarchar(50), " & _
'              "  ExtraSensitivity nvarchar(50), " & _
'              "  Valid bit, " & _
'              "  Isolate int, " & _
'              "  Notes nvarchar(500), " & _
'              "  UserName nvarchar(50), " & _
'              "  ArchiveDateTime datetime default getdate(), " & _
'              "  ArchivedBy nvarchar(50), " & _
'              "  RowGUID uniqueidentifier default newid() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckUrineIdentArcInDb_Error:
'
'      Dim strES As String
'      Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckUrineIdentArcInDb", intEL, strES, sql
'
'End Sub
Public Function IsTableInDatabase(ByVal TableName As String) As Boolean

      Dim tbExists As Recordset
      Dim sql As String
      Dim RetVal As Boolean

      'How to find if a table exists in a database
      'open a recordset with the following sql statement:
      'Code:SELECT name FROM sysobjects WHERE xtype = 'U' AND name = 'MyTable'
      'If the recordset it at eof then the table doesn't exist
      'if it has a record then the table does exist.

10    On Error GoTo IsTableInDatabase_Error

20    sql = "SELECT OBJECT_ID('dbo." & TableName & "', 'U') E"
'20    sql = "SELECT name FROM sysobjects WHERE " & _
'            "xtype = 'U' " & _
'            "AND name = 'dbo." & TableName & "'"
30    Set tbExists = New Recordset
40    Set tbExists = Cnxn(0).Execute(sql)

50    RetVal = True

60    If tbExists.EOF Then 'There is no table <TableName> in database
70      RetVal = False
80    Else
90      If IsNull(tbExists!e) Then
100       RetVal = False
110     Else
120       RetVal = True
130     End If
140   End If
150   IsTableInDatabase = RetVal

160   Exit Function

IsTableInDatabase_Error:

      Dim strES As String
      Dim intEL As Integer

170   intEL = Erl
180   strES = Err.Description
190   LogError "modDbDesign", "IsTableInDatabase", intEL, strES, sql
  
End Function

Public Function EnsureColumnExists(ByVal TableName As String, _
                                   ByVal ColumnName As String, _
                                   ByVal Definition As String) _
                                   As Boolean

      'Return 1 if column created
      '       0 if column already exists

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo EnsureColumnExists_Error

20    sql = "IF NOT EXISTS " & _
            "    (SELECT * FROM syscolumns WHERE " & _
            "    id = object_id('" & TableName & "') " & _
            "    AND name = '" & ColumnName & "') " & _
            "  BEGIN " & _
            "    ALTER TABLE " & TableName & " " & _
            "    ADD " & ColumnName & " " & Definition & " " & _
            "    SELECT 1 AS RetVal " & _
            "  END " & _
            "ELSE " & _
            "  SELECT 0 AS RetVal"

30    Set tb = Cnxn(0).Execute(sql)

40    EnsureColumnExists = tb!RetVal

50    Exit Function

EnsureColumnExists_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "modDbDesign", "EnsureColumnExists", intEL, strES, sql

End Function

Public Sub CheckPrintHandlerLogInDb()

      Dim sql As String

On Error GoTo CheckPrintHandlerLogInDb_Error

20    If IsTableInDatabase("PrintHandlerLog") = False Then 'There is no table  in database
30      sql = "CREATE TABLE PrintHandlerLog " & _
              " ( DateTimeOfRecord datetime NOT NULL DEFAULT getdate(), " & _
              "   Description nvarchar(500) NOT NULL, " & _
              "   OptionalParameter nvarchar(500)) "
        Cnxn(0).Execute sql
50    End If

Exit Sub

CheckPrintHandlerLogInDb_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "modDbDesign", "CheckPrintHandlerLogInDb", intEL, strES, sql

End Sub

'Public Sub CheckAutoCommentsInDb()
'
'      Dim sql As String
'
'10    On Error GoTo CheckAutoCommentsInDb_Error
'
'20    If IsTableInDatabase("AutoComments") = False Then 'There is no table  in database
'30      sql = "CREATE TABLE AutoComments " & _
'              "( Discipline nvarchar(50) NOT NULL, " & _
'              "  Parameter nvarchar(50) NOT NULL, " & _
'              "  Criteria nvarchar(50) NOT NULL, " & _
'              "  Value0 nvarchar(50), " & _
'              "  Value1 nvarchar(50), " & _
'              "  Comment nvarchar(80), " & _
'              "  DateStart smalldatetime, " & _
'              "  DateEnd smalldatetime, " & _
'              "  ListOrder tinyint, " & _
'              "  DateTimeOfRecord datetime NOT NULL DEFAULT getdate() )"
'40      Cnxn(0).Execute sql
'50    End If
'
'60    Exit Sub
'
'CheckAutoCommentsInDb_Error:
'
'Dim strES As String
'Dim intEL As Integer
'
'70    intEL = Erl
'80    strES = Err.Description
'90    LogError "modDbDesign", "CheckAutoCommentsInDb", intEL, strES, sql
'
'End Sub

Public Function EnsureIndexExists(ByVal TableName As String, _
                                  ByVal ColumnName As String, _
                                  ByVal IndexName As String) _
                                  As Boolean

      Dim sql As String
      Dim tb As Recordset

10    On Error GoTo EnsureIndexExists_Error

20    sql = "IF NOT EXISTS " & _
            "    (SELECT name FROM sysindexes WHERE " & _
            "    name = '" & IndexName & "')" & _
            "  BEGIN " & _
            "    CREATE index [" & IndexName & "] " & _
            "    ON [" & TableName & "] ([" & ColumnName & "]) " & _
            "    SELECT 1 AS RetVal " & _
            "  END " & _
            "ELSE " & _
            "  SELECT 0 AS RetVal"

30    Set tb = Cnxn(0).Execute(sql)

40    EnsureIndexExists = tb!RetVal

50    Exit Function

EnsureIndexExists_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "modDbDesign", "EnsureIndexExists", intEL, strES, sql

End Function
