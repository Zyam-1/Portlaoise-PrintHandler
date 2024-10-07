Attribute VB_Name = "modPrinterControl"
Option Explicit

Public gOriginalPrinter As String

Private Declare Function ClosePrinter _
                     Lib "winspool.drv" _
                  (ByVal lngHandle As Long) As Long
                  
Private Declare Function OpenPrinter _
                     Lib "winspool.drv" _
                   Alias "OpenPrinterA" _
                  (ByVal strPrinter_Name As String, _
                   ByRef lngHandle As Long, _
                   ByRef udtPRINTER_DEFAULTS As Any) As Long

Private Function IsPrinterAvailable(ByVal strPrinterName As String) As Boolean

      Dim RetVal As Boolean
      Dim lngHandle As Long

10    On Error GoTo IsPrinterAvailable_Error

20    RetVal = False
  
30    If (OpenPrinter(strPrinterName, lngHandle, ByVal 0&)) Then
40      If lngHandle <> 0& Then
50        RetVal = ClosePrinter(lngHandle)
60      End If
70    End If
  
80    IsPrinterAvailable = RetVal

90    Exit Function

IsPrinterAvailable_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modPrinterControl", "IsPrinterAvailable", intEL, strES, , strPrinterName

End Function

Public Function SetCurrentPrinter(ByVal ForceTo As String) As Boolean

    Dim RetVal As Boolean
    Dim TargetPrinter As String
    Dim Site As String

10    On Error GoTo SetCurrentPrinter_Error

20    RetVal = False
30    TargetPrinter = ""

40    If Trim$(ForceTo) <> "" Then
50      TargetPrinter = ForceTo
60    Else
70      Select Case RP.Department
        Case "B", "G", "Q", "R", "T", "S": TargetPrinter = colPRNs("CHBIO").PrinterName
80      Case "C", "D": TargetPrinter = colPRNs("CHCOAG").PrinterName
90      Case "E": TargetPrinter = colPRNs("CHEND").PrinterName
100     Case "F": TargetPrinter = colPRNs("CHFEA").PrinterName
110     Case "H", "K": TargetPrinter = colPRNs("CHHAEM").PrinterName
120     Case "I", "J": TargetPrinter = colPRNs("CHIMM").PrinterName
130     Case "M": TargetPrinter = colPRNs("CHCOAG").PrinterName
140     Case "W": TargetPrinter = colPRNs("CHALLERGY").PrinterName

150     Case "N", "U":
160         Site = GetMicroSite(RP.SampleID)
170         Select Case Site

            Case "URINE":
180             TargetPrinter = PrinterName("CHURINE")
190             If TargetPrinter = "" Then
200                 TargetPrinter = PrinterName("CHMICRO")
210             End If

220         Case "SWAB":
230             TargetPrinter = PrinterName("CHSWAB")
240             If TargetPrinter = "" Then
250                 TargetPrinter = PrinterName("CHMICRO")
260             End If

270         Case "FAECES":
280             TargetPrinter = PrinterName("CHFAECES")
290             If TargetPrinter = "" Then
300                 TargetPrinter = PrinterName("CHMICRO")
310             End If

320         Case Else
330             If InStr(Site, "SWAB") Then
340                 TargetPrinter = PrinterName("CHSWAB")
350                 If TargetPrinter = "" Then
360                     TargetPrinter = PrinterName("CHMICRO")
370                 End If
380             Else
390                 TargetPrinter = PrinterName("CHMICRO")
400             End If
410         End Select

420     Case "P", "Y": TargetPrinter = colPRNs("CHHIST").PrinterName
430     Case "V", "X": TargetPrinter = colPRNs("CHEXT").PrinterName
440     Case "Z": TargetPrinter = colPRNs("CHSEMEN").PrinterName
450     End Select
460   End If

470   If TargetPrinter <> "" Then
480     If SetPrinter(TargetPrinter) Then
490         RetVal = True
500     End If
510   End If

520   SetCurrentPrinter = RetVal

530   Exit Function

SetCurrentPrinter_Error:

    Dim strES As String
    Dim intEL As Integer

540   intEL = Erl
550   strES = Err.Description
560   LogError "modPrinterControl", "SetCurrentPrinter", intEL, strES

End Function


Public Sub ResetPrinter()

      Dim Px As Printer

10    On Error GoTo ResetPrinter_Error

20    If gOriginalPrinter <> "" Then
30      For Each Px In Printers
40        If Px.DeviceName = gOriginalPrinter Then
50          Set Printer = Px
60          Exit For
70        End If
80      Next
90    End If

100   Exit Sub

ResetPrinter_Error:

      Dim strES As String
      Dim intEL As Integer

110   intEL = Erl
120   strES = Err.Description
130   LogError "modPrinterControl", "ResetPrinter", intEL, strES

End Sub


Public Function SetPrinter(ByVal DeviceName As String) As Boolean

      Dim Px As Printer
      Dim RetVal As Boolean

10    On Error GoTo SetPrinter_Error

20    RetVal = False

30    DeviceName = UCase$(DeviceName)

40    gOriginalPrinter = Printer.DeviceName
50    For Each Px In Printers
60      If UCase(Px.DeviceName) = DeviceName Then
' Px.ScaleLeft = -300
70        Set Printer = Px
80        RetVal = IsPrinterAvailable(DeviceName)
90        Exit For
100     End If
110   Next
  
120   SetPrinter = RetVal

130   Exit Function

SetPrinter_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "modPrinterControl", "SetPrinter", intEL, strES, , DeviceName

End Function


