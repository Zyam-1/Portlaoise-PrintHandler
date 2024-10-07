Attribute VB_Name = "PrinterCollection"


Option Explicit

      Const PRINTER_ENUM_CONNECTIONS = &H4
      Const PRINTER_ENUM_LOCAL = &H2

      Type PRINTER_INFO_1
         Flags As Long
         pDescription As String
         PName As String
         PComment As String
      End Type

      Type PRINTER_INFO_4
         pPrinterName As String
         pServerName As String
         Attributes As Long
      End Type

      Declare Function EnumPrinters Lib "winspool.drv" Alias _
         "EnumPrintersA" (ByVal Flags As Long, ByVal Name As String, _
         ByVal Level As Long, pPrinterEnum As Long, ByVal cdBuf As Long, _
         pcbNeeded As Long, pcReturned As Long) As Long
      Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyA" _
         (ByVal RetVal As String, ByVal Ptr As Long) As Long
      Declare Function StrLen Lib "kernel32" Alias "lstrlenA" _
         (ByVal Ptr As Long) As Long


Public Type InstalledPrinters
    Flags As String
    Description As String
    Name As String
    Comment As String
End Type

Public LocalPrinters() As InstalledPrinters
Public NetworkPrinters() As InstalledPrinters
Public AllPrinters() As InstalledPrinters


Public Function LoadPrinters() As Boolean
          Dim Success As Boolean, cbRequired As Long, cbBuffer As Long
          Dim Buffer() As Long, nEntries1 As Long, nEntries2 As Long
          Dim i As Long, PFlags As Long, PDesc As String, PName As String
          Dim PComment As String, Temp As Long
          Dim TotalPrinters As Integer
10    On Error GoTo LoadPrinters_Error

20        cbBuffer = 3072
30        ReDim Buffer((cbBuffer \ 4) - 1) As Long
    
          'Get local Printers
    
40        Success = EnumPrinters(PRINTER_ENUM_LOCAL, _
                                vbNullString, _
                                1, _
                                Buffer(0), _
                                cbBuffer, _
                                cbRequired, _
                                nEntries1)
50        If Success And nEntries1 > 0 Then
60           If cbRequired > cbBuffer Then
70                cbBuffer = cbRequired
80                Debug.Print "Buffer too small.  Trying again with " & _
                           cbBuffer & " bytes."
  
90                ReDim Buffer(cbBuffer \ 4) As Long
100           End If
110           ReDim LocalPrinters(nEntries1 - 1) As InstalledPrinters
120           Debug.Print "There are " & nEntries1 & _
                          " local and connected printers."
130           For i = 0 To nEntries1 - 1
140               PFlags = Buffer(4 * i)
150               PDesc = Space$(StrLen(Buffer(i * 4 + 1)))
160               Temp = PtrToStr(PDesc, Buffer(i * 4 + 1))
170               PName = Space$(StrLen(Buffer(i * 4 + 2)))
180               Temp = PtrToStr(PName, Buffer(i * 4 + 2))
190               PComment = Space$(StrLen(Buffer(i * 4 + 2)))
200               Temp = PtrToStr(PComment, Buffer(i * 4 + 2))
210               Debug.Print PFlags, PDesc, PName, PComment
220               LocalPrinters(i).Flags = PFlags
230               LocalPrinters(i).Description = PDesc
240               LocalPrinters(i).Name = PName
250               LocalPrinters(i).Comment = PComment

260          Next i
270       Else
280          Debug.Print "Error enumerating local printers."
290       End If
    
          'Get network printers
300       ReDim Buffer((cbBuffer \ 4) - 1) As Long
310       Success = EnumPrinters(PRINTER_ENUM_CONNECTIONS, _
                                vbNullString, _
                                1, _
                                Buffer(0), _
                                cbBuffer, _
                                cbRequired, _
                                nEntries2)
320       If Success And nEntries2 > 0 Then
 
330          If cbRequired > cbBuffer Then
340             cbBuffer = cbRequired
350             Debug.Print "Buffer too small.  Trying again with " & _
                         cbBuffer & " bytes."
360             ReDim Buffer(cbBuffer \ 4) As Long
370           End If
380           ReDim NetworkPrinters(nEntries2 - 1) As InstalledPrinters
390           Debug.Print "There are " & nEntries2 & _
                          " local and connected printers."
400           For i = 0 To nEntries2 - 1
410               PFlags = Buffer(4 * i)
420               PDesc = Space$(StrLen(Buffer(i * 4 + 1)))
430               Temp = PtrToStr(PDesc, Buffer(i * 4 + 1))
440               PName = Space$(StrLen(Buffer(i * 4 + 2)))
450               Temp = PtrToStr(PName, Buffer(i * 4 + 2))
460               PComment = Space$(StrLen(Buffer(i * 4 + 2)))
470               Temp = PtrToStr(PComment, Buffer(i * 4 + 2))
480               Debug.Print PFlags, PDesc, PName, PComment
490               NetworkPrinters(i).Flags = PFlags
500               NetworkPrinters(i).Description = PDesc
510               NetworkPrinters(i).Name = PName
520               NetworkPrinters(i).Comment = PComment
530          Next i
540       Else
550          Debug.Print "Error enumerating network printers."
560       End If
570       TotalPrinters = nEntries1 + nEntries2
580       If TotalPrinters > 0 Then
590           ReDim AllPrinters(TotalPrinters - 1) As InstalledPrinters
              Dim Ind As Integer
600           Ind = 0
610           If nEntries1 > 0 Then
620               For i = 0 To UBound(LocalPrinters)
630                   AllPrinters(Ind).Flags = LocalPrinters(i).Flags
640                   AllPrinters(Ind).Description = LocalPrinters(i).Description
650                   AllPrinters(Ind).Name = LocalPrinters(i).Name
660                   AllPrinters(Ind).Comment = LocalPrinters(i).Comment
670                   Ind = Ind + 1
680               Next
690           End If
700           If nEntries2 > 0 Then
710               For i = 0 To UBound(NetworkPrinters)
720                   AllPrinters(Ind).Flags = NetworkPrinters(i).Flags
730                   AllPrinters(Ind).Description = NetworkPrinters(i).Description
740                   AllPrinters(Ind).Name = NetworkPrinters(i).Name
750                   AllPrinters(Ind).Comment = NetworkPrinters(i).Comment
760                   Ind = Ind + 1
770               Next
780           End If
790           LoadPrinters = (TotalPrinters > 0)
800       Else
810           LoadPrinters = False
820       End If

830   Exit Function

LoadPrinters_Error:

      Dim strES As String
      Dim intEL As Integer

840   intEL = Erl
850   strES = Err.Description
860   LogError "PrinterCollection", "LoadPrinters", intEL, strES

End Function

Public Sub FillLocalPrinters(List As ListBox)
      Dim i As Integer
10    If LoadPrinters Then
20        With List
30            .Clear
40            For i = 0 To UBound(LocalPrinters)
50                List.AddItem LocalPrinters(i).Name
60            Next i
  
70        End With
80    End If
End Sub

Public Sub FillNetworkPrinters(List As ListBox)
      Dim i As Integer
10    If LoadPrinters Then
20        With List
30            .Clear
40            For i = 0 To UBound(NetworkPrinters)
50                List.AddItem NetworkPrinters(i).Name
60            Next i
  
70        End With
80    End If
End Sub
Public Sub FillAllPrinters(List As ListBox)
      Dim i As Integer
10    If LoadPrinters Then
20        With List
30            .Clear
40            For i = 0 To UBound(AllPrinters)
50                List.AddItem AllPrinters(i).Name
60            Next i
  
70        End With
80    End If
End Sub

      Sub EnumeratePrinters4()
            Dim Success As Boolean, cbRequired As Long, cbBuffer As Long
            Dim Buffer() As Long, nEntries As Long
            Dim i As Long, PName As String, SName As String
            Dim Attrib As Long, Temp As Long
10             cbBuffer = 3072
20             ReDim Buffer((cbBuffer \ 4) - 1) As Long
30             Success = EnumPrinters(PRINTER_ENUM_CONNECTIONS Or _
                                     PRINTER_ENUM_LOCAL, _
                                     vbNullString, _
                                     4, _
                                     Buffer(0), _
                                     cbBuffer, _
                                     cbRequired, _
                                     nEntries)
40             If Success Then
50                If cbRequired > cbBuffer Then
60                   cbBuffer = cbRequired
70                   Debug.Print "Buffer too small.  Trying again with " & _
                              cbBuffer & " bytes."
80                   ReDim Buffer(cbBuffer \ 4) As Long
90                   Success = EnumPrinters(PRINTER_ENUM_CONNECTIONS Or _
                                         PRINTER_ENUM_LOCAL, _
                                         vbNullString, _
                                         4, _
                                         Buffer(0), _
                                         cbBuffer, _
                                         cbRequired, _
                                         nEntries)
100                  If Not Success Then
110                     Debug.Print "Error enumerating printers."
120                     Exit Sub
130                  End If
140               End If
150               Debug.Print "There are " & nEntries & _
                            " local and connected printers."
160               For i = 0 To nEntries - 1
170               PName = Space$(StrLen(Buffer(i * 3)))
180               Temp = PtrToStr(PName, Buffer(i * 3))
190               SName = Space$(StrLen(Buffer(i * 3 + 1)))
200               Temp = PtrToStr(SName, Buffer(i * 3 + 1))
210               Attrib = Buffer(i * 3 + 2)
220               Debug.Print "Printer: " & PName, "Server: " & SName, _
                              "Attributes: " & Hex$(Attrib)
230               Next i
240            Else
250               Debug.Print "Error enumerating printers."
260            End If
      End Sub




