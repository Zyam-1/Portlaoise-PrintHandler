Attribute VB_Name = "PrintRTFwithMargin"
Option Explicit

Private Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type

Private Type FORMATRANGE
    hdc As Long
    hdcTarget As Long
    rc As Rect
    rcPage As Rect
    chrg As CHARRANGE
End Type

Private Const WM_USER As Long = &H400
Private Const EM_FORMATRANGE As Long = WM_USER + 57
Private Const EM_SETTARGETDEVICE As Long = WM_USER + 72
Private Const PHYSICALOFFSETX As Long = 112
Private Const PHYSICALOFFSETY As Long = 113

Private Declare Function GetDeviceCaps Lib "gdi32" _
                                       (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Declare Function SendMessage Lib "USER32" _
                                     Alias "SendMessageA" (ByVal hWnd As Long, ByVal msg As Long, _
                                                           ByVal wp As Long, lp As Any) As Long


Public Function PrintRTFwithMargins(RTFControl As Object, _
                                    ByVal LeftMargin As Single, ByVal TopMargin As Single, _
                                    ByVal RightMargin As Single, ByVal BottomMargin As Single) _
                                    As Boolean

      '********************************************************8
      'PURPOSE: Prints Contents of RTF Control with Margins

      'PARAMETERS:
      'RTFControl: RichTextBox Control For Printing
      'LeftMargin: Left Margin in Inches
      'TopMargin: TopMargin in Inches
      'RightMargin: RightMargin in Inches
      'BottomMargin: BottomMargin in Inches

      '***************************************************************

10    On Error GoTo ErrorHandler


      '*************************************************************
      'I DO THIS BECAUSE IT IS MY UNDERSTANDING THAT
      'WHEN CALLING A SERVER DLL, YOU CAN RUN INTO
      'PROBLEMS WHEN USING EARLY BINDING WHEN A PARAMETER
      'IS A CONTROL OR A CUSTOM OBJECT.  IF YOU JUST PLUG THIS INTO
      'A FORM, YOU CAN DECLARE RTFCONTROL AS RICHTEXTBOX
      'AND COMMENT OUT THE FOLLOWING LINE

20    If Not TypeOf RTFControl Is RichTextBox Then Exit Function
      '**************************************************************

      Dim lngLeftOffset As Long
      Dim lngTopOffSet As Long
      Dim lngLeftMargin As Long
      Dim lngTopMargin As Long
      Dim lngRightMargin As Long
      Dim lngBottomMargin As Long

      Dim typFr As FORMATRANGE
      Dim rectPrintTarget As Rect
      Dim rectPage As Rect
      Dim lngTxtLen As Long
      Dim lngPos As Long
      Dim lngRet As Long
      Dim iTempScaleMode As Integer

30    iTempScaleMode = Printer.ScaleMode

      ' needed to get a Printer.hDC
40    Printer.Print ""
50    Printer.ScaleMode = vbTwips

      ' Get the offsets to printable area in twips
60    lngLeftOffset = Printer.ScaleX(GetDeviceCaps(Printer.hdc, _
                                                   PHYSICALOFFSETX), vbPixels, vbTwips)
70    lngTopOffSet = Printer.ScaleY(GetDeviceCaps(Printer.hdc, _
                                                  PHYSICALOFFSETY), vbPixels, vbTwips)

      ' Get Margins in Twips
80    lngLeftMargin = InchesToTwips(LeftMargin) - lngLeftOffset
90    lngTopMargin = InchesToTwips(TopMargin) - lngTopOffSet
100   lngRightMargin = (Printer.Width - _
                        InchesToTwips(RightMargin)) - lngLeftOffset

110   lngBottomMargin = (Printer.Height - _
                         InchesToTwips(BottomMargin)) - lngTopOffSet

      ' Set printable area rect
120   rectPage.Left = 0
130   rectPage.Top = 0
140   rectPage.Right = Printer.ScaleWidth
150   rectPage.Bottom = Printer.ScaleHeight

      ' Set rect in which to print, based on margins passed in
160   rectPrintTarget.Left = lngLeftMargin
170   rectPrintTarget.Top = lngTopMargin
180   rectPrintTarget.Right = lngRightMargin
190   rectPrintTarget.Bottom = lngBottomMargin

      ' Set up the printer for this print job
200   typFr.hdc = Printer.hdc    'for rendering
210   typFr.hdcTarget = Printer.hdc    'for formatting
220   typFr.rc = rectPrintTarget
230   typFr.rcPage = rectPage
240   typFr.chrg.cpMin = 0
250   typFr.chrg.cpMax = -1

      ' Get length of text in the RichTextBox Control
260   lngTxtLen = Len(RTFControl.Text)

      ' print page by page
270   Do
          ' Print the page by sending EM_FORMATRANGE message
          'Allows you to range of text within a specific device
          'here, the device is the printer, which must be specified
          'as hdc and hdcTarget of the FORMATRANGE structure

280       lngPos = SendMessage(RTFControl.hWnd, EM_FORMATRANGE, _
                               True, typFr)

290       If lngPos >= lngTxtLen Then Exit Do  'Done
300       typFr.chrg.cpMin = lngPos    ' Starting position next page
310       Printer.NewPage             ' go to next page
320       Printer.Print ""   'to get hDC again
330       typFr.hdc = Printer.hdc
340       typFr.hdcTarget = Printer.hdc
350   Loop

      ' Done
360   Printer.EndDoc

      ' This frees memory
370   lngRet = SendMessage(RTFControl.hWnd, EM_FORMATRANGE, _
                           False, Null)
380   Printer.ScaleMode = iTempScaleMode
390   PrintRTFwithMargins = True
400   Exit Function

ErrorHandler:
410   Err.Raise Err.Number, , Err.Description

End Function

Public Function InchesToTwips(ByVal Inches As Single) As Single
10    InchesToTwips = 1440 * Inches
End Function
Public Function MillimetersToInches(ByVal Mm As Single) As Single
10    MillimetersToInches = 0.0393701 * Mm
End Function

