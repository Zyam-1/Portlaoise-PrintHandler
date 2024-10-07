Attribute VB_Name = "modNoConstant"
Option Explicit
  
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
    "GetPrivateProfileStringA" _
    (ByVal lpSectionName As String, ByVal lpKeyName As String, _
     ByVal lpDefault As String, ByVal lpbuffurnedString As String, _
     ByVal nBuffSize As Long, ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileSectionNames Lib "Kernel32.dll" Alias _
    "GetPrivateProfileSectionNamesA" _
    (ByVal lpszReturnBuffer As String, _
     ByVal nSize As Long, _
     ByVal lpFileName As String) As Long

Public Sub ConnectToDatabase()
      'MsgBox "Hello"
      Dim Con As String
      Dim ConBB As String

10    On Error GoTo ConnectToDatabase_Error

20    HospName(0) = GetcurrentConnectInfo(Con, ConBB)

30    Set Cnxn(0) = New Connection
40    Cnxn(0).Open Con

50    Exit Sub

ConnectToDatabase_Error:

      Dim strES As String
      Dim intEL As Integer

60    intEL = Erl
70    strES = Err.Description
80    LogError "modNoConstant", "ConnectToDatabase", intEL, strES

End Sub
Public Function GetConnectInfo(ByVal ConnectTo As String, _
                               ByRef ReturnConnectionString As String, _
                               Optional ByRef HospName As Variant) As Boolean

      'ConnectTo = "Active"
      '            "BB"
      '            "Active" & n - HospitalGroup
      '            "BB" & n - HospitalGroup

10    On Error GoTo GetConnectInfo_Error

20    GetConnectInfo = False

30    If Not IsMissing(HospName) Then
40      HospName = GetSetting("NetAcquire", "HospName", ConnectTo, "")
50      If Left$(UCase$(HospName), 5) = "LOCAL" Then
60        HospName = Mid$(HospName, 6)
70      End If
80    End If

90    ReturnConnectionString = GetSetting("NetAcquire", "Cnxn", ConnectTo, "")

100   If Trim$(ReturnConnectionString) <> "" Then
  
110     ReturnConnectionString = Obfuscate(ReturnConnectionString)
  
120     GetConnectInfo = True
  
130   End If

140   Exit Function

GetConnectInfo_Error:

      Dim strES As String
      Dim intEL As Integer

150   intEL = Erl
160   strES = Err.Description
170   LogError "modNoConstant", "GetConnectInfo", intEL, strES


End Function


Public Function GetcurrentConnectInfo(ByRef Con As String, ByRef ConBB As String) As String

      'Returns Hospital Name

      Dim HospitalNames() As String
      Dim n As Long
      Dim HospitalName As String
      Dim retHospitalName As String
      Dim ServerName As String
      Dim NetAcquireDB As String
      Dim TransfusionDB As String
      Dim UID As String
      Dim PWD As String
      Dim CurrentPath As String

10    On Error GoTo GetcurrentConnectInfo_Error
''Comment
'20    If IsIDE Then
'30      CurrentPath = "C:\ClientCode\NetAcquire.INI"
'40    Else
'50      CurrentPath = App.Path & "\NetAcquire.INI"
'60    End If

70    HospitalNames = GetINISectionNames(CurrentPath, n)
80    HospitalName = HospitalNames(0)
90    If Left$(UCase$(HospitalName), 5) = "LOCAL" Then
100     retHospitalName = Mid$(HospitalName, 6)
110   Else
120     retHospitalName = HospitalName
130   End If

140   ServerName = ProfileGetItem(HospitalName, "N", "", CurrentPath)
150   NetAcquireDB = ProfileGetItem(HospitalName, "D", "", CurrentPath)
160   TransfusionDB = ProfileGetItem(HospitalName, "T", "", CurrentPath)
170   PWD = GetPass(UID)
'PWD = "DfySiywtgtw$1>)="
'180   Con = "DRIVER={SQL Server};" & _
'            "Server=" & Obfuscate(ServerName) & ";" & _
'            "Database=" & Obfuscate(NetAcquireDB) & ";" & _
'            "uid=" & UID & ";" & _
'            "pwd=" & PWD & ";"
190                   Con = "Provider=SQLOLEDB;" & _
              "Data Source=" & "DESKTOP-3OMS1N5\SQLEXPRESS" & ";" & _
              "Initial Catalog=" & "PortLive" & ";" & _
              "Integrated Security=SSPI;"
'
'181   Con = "DRIVER={SQL Server};" & _
'            "Server=" & "192.168.20.83" & ";" & _
'            "Database=" & "PortLive" & ";" & _
'            "uid=" & "zyam" & ";" & _
'            "pwd=" & "zyam123" & ";"

191   If TransfusionDB <> "" Then
'200     ConBB = "DRIVER={SQL Server};" & _
'                "Server=" & Obfuscate(ServerName) & ";" & _
'                "Database=" & Obfuscate(TransfusionDB) & ";" & _
'                "uid=" & UID & ";" & _
'                "pwd=" & PWD & ";"

210   End If

220   GetcurrentConnectInfo = retHospitalName

230   Exit Function

GetcurrentConnectInfo_Error:

      Dim strES As String
      Dim intEL As Integer

240   intEL = Erl
250   strES = Err.Description
260   LogError "modNoConstant", "GetcurrentConnectInfo", intEL, strES

End Function
Private Function GetPass(ByRef UID As String) As String

      Dim p As String
      Dim a As String
      Dim n As Integer

10    a = ""
20    For n = 97 To 122
30      a = a & Chr$(n)
40    Next
50    For n = 65 To 90
60      a = a & Chr$(n)
70    Next

80    a = a & "!£$%^&*()<>-_+={}[]:@~||;'#,./?"
90    For n = 48 To 57
100     a = a & Chr$(n)
110   Next

      '    abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!£$%^&*()<>-_+={}[]:@~||;'#,./?0123456789"
      '    123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123
      '             1         2         3         4         5         6         7         8         9

      'p = ""
      'UID = "sa"

      'LabUser
120   UID = Mid$(a, 38, 1) & Mid$(a, 1, 1) & Mid$(a, 2, 1) & Mid$(a, 47, 1) & _
            Mid$(a, 19, 1) & Mid$(a, 5, 1) & Mid$(a, 18, 1)
    
      'DfySiywtgtw$1>)=
130   p = Mid$(a, 30, 1) & Mid$(a, 6, 1) & Mid$(a, 25, 1) & Mid$(a, 45, 1) & _
          Mid$(a, 9, 1) & Mid$(a, 25, 1) & Mid$(a, 23, 1) & Mid$(a, 20, 1) & _
          Mid$(a, 7, 1) & Mid$(a, 20, 1) & Mid$(a, 23, 1) & Mid$(a, 55, 1) & _
          Mid$(a, 85, 1) & Mid$(a, 63, 1) & Mid$(a, 61, 1) & Mid$(a, 67, 1)

140   GetPass = p

End Function


Private Function ProfileGetItem(ByRef sSection As String, _
                                ByRef sKeyName As String, _
                                ByRef sDefValue As String, _
                                ByRef sIniFile As String) As String

          'retrieves a value FROM an ini file
          'corresponding to the section and
          'key name passed.

      Dim dwSize As Integer
      Dim nBuffSize As Integer
      Dim buff As String
      Dim RetVal As String

      'Call the API with the parameters passed.
      'nBuffSize is the length of the string
      'in buff, including the terminating null.
      'If a default value was passed, and the
      'section or key name are not in the file,
      'that value is returned. If no default
      'value was passed (""), then dwSize
      'will = 0 if not found.
      '
      'pad a string large enough to hold the data
10    On Error GoTo ProfileGetItem_Error

20    buff = Space(2048)
30    nBuffSize = Len(buff)
40    dwSize = GetPrivateProfileString(sSection, sKeyName, sDefValue, buff, nBuffSize, sIniFile)

50    If dwSize > 0 Then
60      RetVal = Left$(buff, dwSize)
70    End If

80    ProfileGetItem = RetVal

90    Exit Function

ProfileGetItem_Error:

      Dim strES As String
      Dim intEL As Integer

100   intEL = Erl
110   strES = Err.Description
120   LogError "modNoConstant", "ProfileGetItem", intEL, strES


End Function

Private Function GetINISectionNames(ByRef inFile As String, ByRef outCount As Long) As String()

      Dim StrBuf As String
      Dim BufLen As Long
      Dim RetVal() As String
      Dim Count As Long

10    On Error GoTo GetINISectionNames_Error

20    BufLen = 16

30    Do
40      BufLen = BufLen * 2
50      StrBuf = Space$(BufLen)
60      Count = GetPrivateProfileSectionNames(StrBuf, BufLen, inFile)
70    Loop While Count = BufLen - 2

80    If (Count) Then
90      RetVal = Split(Left$(StrBuf, Count - 1), vbNullChar)
100     outCount = UBound(RetVal) + 1
110   End If

120   GetINISectionNames = RetVal

130   Exit Function

GetINISectionNames_Error:

      Dim strES As String
      Dim intEL As Integer

140   intEL = Erl
150   strES = Err.Description
160   LogError "modNoConstant", "GetINISectionNames", intEL, strES


End Function
  

Public Function Obfuscate(ByVal strData As String) As String

      Dim lngI As Long
      Dim lngJ As Long
   
10    On Error GoTo Obfuscate_Error

20    For lngI = 0 To Len(strData) \ 4
30      For lngJ = 1 To 4
40         Obfuscate = Obfuscate & Mid$(strData, (4 * lngI) + 5 - lngJ, 1)
50      Next
60    Next

70    Exit Function

Obfuscate_Error:

      Dim strES As String
      Dim intEL As Integer

80    intEL = Erl
90    strES = Err.Description
100   LogError "modNoConstant", "Obfuscate", intEL, strES


End Function

