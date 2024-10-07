Attribute VB_Name = "modPrinter"
Option Explicit

    
Public Declare Function FindFirstPrinterChangeNotificationLong Lib "winspool.drv" Alias "FindFirstPrinterChangeNotification" _
  (ByVal hPrinter As Long, ByVal fdwFlags As Long, ByVal fdwOptions As Long, ByVal lpPrinterNotifyOptions As Long) As Long
  
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Declare Function FindNextPrinterChangeNotificationByLong Lib "winspool.drv" Alias "FindNextPrinterChangeNotification" _
    (ByVal hChange As Long, pdwChange As Long, pPrinterOptions As PRINTER_NOTIFY_OPTIONS, ppPrinterNotifyInfo As Long) As Long


Public Enum PrinterChangeNotifications
    PRINTER_CHANGE_ADD_PRINTER = &H1
    PRINTER_CHANGE_SET_PRINTER = &H2
    PRINTER_CHANGE_DELETE_PRINTER = &H4
    PRINTER_CHANGE_FAILED_CONNECTION_PRINTER = &H8
    PRINTER_CHANGE_PRINTER = &HFF
    PRINTER_CHANGE_ADD_JOB = &H100
    PRINTER_CHANGE_SET_JOB = &H200
    PRINTER_CHANGE_DELETE_JOB = &H400
    PRINTER_CHANGE_WRITE_JOB = &H800
    PRINTER_CHANGE_JOB = &HFF00
    PRINTER_CHANGE_ADD_FORM = &H10000
    PRINTER_CHANGE_SET_FORM = &H20000
    PRINTER_CHANGE_DELETE_FORM = &H40000
    PRINTER_CHANGE_FORM = &H70000
    PRINTER_CHANGE_ADD_PORT = &H100000
    PRINTER_CHANGE_CONFIGURE_PORT = &H200000
    PRINTER_CHANGE_DELETE_PORT = &H400000
    PRINTER_CHANGE_PORT = &H700000
    PRINTER_CHANGE_ADD_PRINT_PROCESSOR = &H1000000
    PRINTER_CHANGE_DELETE_PRINT_PROCESSOR = &H4000000
    PRINTER_CHANGE_PRINT_PROCESSOR = &H7000000
    PRINTER_CHANGE_ADD_PRINTER_DRIVER = &H10000000
    PRINTER_CHANGE_SET_PRINTER_DRIVER = &H20000000
    PRINTER_CHANGE_DELETE_PRINTER_DRIVER = &H40000000
    PRINTER_CHANGE_PRINTER_DRIVER = &H70000000
    PRINTER_CHANGE_TIMEOUT = &H80000000
End Enum

Public Enum JobChangeNotificationFields
    JOB_NOTIFY_FIELD_PRINTER_NAME = &H0
    JOB_NOTIFY_FIELD_MACHINE_NAME = &H1
    JOB_NOTIFY_FIELD_PORT_NAME = &H2
    JOB_NOTIFY_FIELD_USER_NAME = &H3
    JOB_NOTIFY_FIELD_NOTIFY_NAME = &H4
    JOB_NOTIFY_FIELD_DATATYPE = &H5
    JOB_NOTIFY_FIELD_PRINT_PROCESSOR = &H6
    JOB_NOTIFY_FIELD_PARAMETERS = &H7
    JOB_NOTIFY_FIELD_DRIVER_NAME = &H8
    JOB_NOTIFY_FIELD_DEVMODE = &H9
    JOB_NOTIFY_FIELD_STATUS = &HA
    JOB_NOTIFY_FIELD_STATUS_STRING = &HB
    JOB_NOTIFY_FIELD_SECURITY_DESCRIPTOR = &HC
    JOB_NOTIFY_FIELD_DOCUMENT = &HD
    JOB_NOTIFY_FIELD_PRIORITY = &HE
    JOB_NOTIFY_FIELD_POSITION = &HF
    JOB_NOTIFY_FIELD_SUBMITTED = &H10
    JOB_NOTIFY_FIELD_START_TIME = &H11
    JOB_NOTIFY_FIELD_UNTIL_TIME = &H12
    JOB_NOTIFY_FIELD_TIME = &H13
    JOB_NOTIFY_FIELD_TOTAL_PAGES = &H14
    JOB_NOTIFY_FIELD_PAGES_PRINTED = &H15
    JOB_NOTIFY_FIELD_TOTAL_BYTES = &H16
    JOB_NOTIFY_FIELD_BYTES_PRINTED = &H17
End Enum


'\\ Declarations
Public Type PRINTER_NOTIFY_OPTIONS
    Version As Long '\\should be set to 2
    Flags As Long
    Count As Long
    lpPrintNotifyOptions As Long
End Type

Public Type PRINTER_NOTIFY_OPTIONS_TYPE
    Type As Integer
    Reserved_0 As Integer
    Reserved_1 As Long
    Reserved_2 As Long
    Count As Long
    pFields As Long
End Type

Private PrintOptions As PRINTER_NOTIFY_OPTIONS
Private PrinterNotifyOptions(0 To 1) As PRINTER_NOTIFY_OPTIONS_TYPE

'\\ Initialising the PrintOptions
Private Sub InitialiseNotifyOptions()

With PrintOptions
  .Version = 2 '\\ This must be set to 2
  .Count = 2 '\\ There is job notification and printer notification
  '\\ The type of printer events we are interested in...
  With PrinterNotifyOptions(0)
    .Type = PRINTER_NOTIFY_TYPE
    ReDim pFieldsPrinter(0 To 19) As Integer
    '\\ Add the list of printer events you are interested in being notified about
    '\\ to this list. Note that the fewer notifications you ask for the less of a
    '\\ burden your app place upon the system.
    pFieldsPrinter(0) = PRINTER_NOTIFY_FIELD_PRINTER_NAME
    pFieldsPrinter(1) = PRINTER_NOTIFY_FIELD_SHARE_NAME
    pFieldsPrinter(2) = PRINTER_NOTIFY_FIELD_STATUS
    .Count = (UBound(pFieldsPrinter) - LBound(pFieldsPrinter)) + 1 '\\ Add one as the array is zero based
    .pFields = VarPtr(pFieldsPrinter(0))
  End With
  '\\ The type of print job events we are interested in...
  With PrinterNotifyOptions(1)
    .Type = JOB_NOTIFY_TYPE
    '\\ Add the list of print job events you are interested in being notified about
    '\\ to this list. Note that the fewer notifications you ask for the less of a
    '\\ burden your app place upon the system.
    ReDim pFieldsJob(0 To 22) As Integer
    pFieldsJob(0) = JOB_NOTIFY_FIELD_PRINTER_NAME
    pFieldsJob(1) = JOB_NOTIFY_FIELD_MACHINE_NAME
    pFieldsJob(2) = JOB_NOTIFY_FIELD_PORT_NAME
    pFieldsJob(3) = JOB_NOTIFY_FIELD_USER_NAME
    pFieldsJob(4) = JOB_NOTIFY_FIELD_NOTIFY_NAME
    pFieldsJob(5) = JOB_NOTIFY_FIELD_DATATYPE
    pFieldsJob(6) = JOB_NOTIFY_FIELD_PRINT_PROCESSOR
    pFieldsJob(7) = JOB_NOTIFY_FIELD_PARAMETERS
    pFieldsJob(8) = JOB_NOTIFY_FIELD_DRIVER_NAME
    pFieldsJob(9) = JOB_NOTIFY_FIELD_DEVMODE
    pFieldsJob(10) = JOB_NOTIFY_FIELD_STATUS
    pFieldsJob(11) = JOB_NOTIFY_FIELD_STATUS_STRING
    pFieldsJob(12) = JOB_NOTIFY_FIELD_DOCUMENT
    pFieldsJob(13) = JOB_NOTIFY_FIELD_PRIORITY
    pFieldsJob(14) = JOB_NOTIFY_FIELD_POSITION
    pFieldsJob(15) = JOB_NOTIFY_FIELD_SUBMITTED
    pFieldsJob(16) = JOB_NOTIFY_FIELD_START_TIME
    pFieldsJob(17) = JOB_NOTIFY_FIELD_UNTIL_TIME
    pFieldsJob(18) = JOB_NOTIFY_FIELD_TIME
    pFieldsJob(19) = JOB_NOTIFY_FIELD_TOTAL_PAGES
    pFieldsJob(20) = JOB_NOTIFY_FIELD_PAGES_PRINTED
    pFieldsJob(21) = JOB_NOTIFY_FIELD_TOTAL_BYTES
    .Count = (UBound(pFieldsJob) - LBound(pFieldsJob)) + 1 '\\ Add one as the array is zero based
    .pFields = VarPtr(pFieldsJob(0))
  End With
  .lpPrintNotifyOptions = VarPtr(PrinterNotifyOptions(0))
End With

End Sub
