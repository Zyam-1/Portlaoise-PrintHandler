Attribute VB_Name = "basOptions"
Option Explicit

Public SysOptFaxServer() As String
Public SysOptFax() As String 'Fax Path
Public SysOptPrintMiddle() As Boolean

Public SysOptBioMaskSym() As String

Public SysOptBioCodeForRF() As String

'departments
Public SysOptDeptHaem() As Boolean 'Haematology In Use
Public SysOptDeptBio() As Boolean  'Biochemistry in Use
Public SysOptDeptCoag() As Boolean 'Coagulation In Use
Public SysOptDeptMicro() As Boolean 'Microbiology in use
Public SysOptDeptImm() As Boolean  'Immunology in Use
Public SysOptDeptEnd() As Boolean  'Endocriology in use
Public SysOptDeptBga() As Boolean  'Blood Gas in Use
Public SysOptDeptExt() As Boolean  'Externals in Use
Public SysOptDeptSemen() As Boolean 'Semen anaysis Un Use
Public SysOptDeptCyto() As Boolean  'Cytology In Use
Public SysOptDeptHisto() As Boolean 'Histology in Use

'added by myles
Public SysOptExp() As Boolean

'Bio Codes
Public SysOptBioCodeForCreat() As String 'creatinine
Public SysOptBioCodeForUCreat() As String 'urinay creat
Public SysOptBioCodeForUProt() As String ' urinary prot
Public SysOptBioCodeForAlb() As String 'albumin
Public SysOptBioCodeForGlob() As String 'glob
Public SysOptBioCodeForTProt() As String 'total protein
Public SysOptBioCodeFor24UProt() As String  '24 urinary protein
Public SysOptBioCodeFor24Vol() As String '24 Volumne
Public SysOptBioCodeForChol() As String  'Cholestrol Code
Public SysOptBioCodeForHDL() As String   'HDL Code
Public SysOptBioCodeForTrig() As String  'Triglyceride code
Public SysOptBioCodeForCholHDLRatio() As String  'CholHdl Ratio
Public SysOptBioCodeForHbA1c() As String  'HBA1c Code
Public SysOptBioCodeBNP() As String
Public SysOptBioCodeForEGFR() As String 'eGfr
Public SysOptBioCodeForGlucose() As String  'BioGlucose Code
Public SysOptBioCodeForGlucose1() As String  'BioGlucose Code
Public SysOptBioCodeForGlucose2() As String  'BioGlucose Code
Public SysOptBioCodeForGlucose3() As String  'BioGlucose Code
Public SysOptBioCodeForFastGlucose() As String  'Bio fast Glucose Code
Public SysOptBioCodeForGlucoseP() As String  'BioGlucose Code Plasma
Public SysOptBioCodeForGlucose1P() As String  'BioGlucose 1Hr Code Plasma
Public SysOptBioCodeForGlucose2P() As String  'BioGlucose 2 Hr Code Plasma
Public SysOptBioCodeForGlucose3P() As String  'BioGlucose 3 Hr Code Plasma
Public SysOptBioCodeForFastGlucoseP() As String  'Bio fast Glucose Code Plasma
Public SysOptBioCodeForCholP() As String  'Cholestrol Code Plasma
Public SysOptBioCodeForTrigP() As String  'Triglyceride code Plasma

'End Codes
Public SysOptEndCodeHBA1C() As String
Public SysOptEndCodeBNP() As String
Public SysOptEndCodeB12() As String
Public SysOptEndCodeCalcA1C() As String
Public SysOptEndCodeB12New() As String
Public SysOptEndCodeCo() As String
Public SysOptEndCodeFSH() As String
Public SysOptEndCodeLH() As String
Public SysOptEndCodePRO() As String
Public SysOptEndCodeOES() As String
Public SysOptEndCodePRL() As String
Public SysOptEndCodeTHC() As String
Public SysOptEndCodeTRO() As String

Public SysOptEndCodeVITD() As String


'End Comments
Public SysOptEndHbA1cComment() As String
Public SysOptEndCalcA1cComment() As String

'Urines
Public SysOptBioCodeForUNa() As String 'Sodium
Public SysOptBioCodeForUUrea() As String  'Urea
Public SysOptBioCodeForUK() As String 'Pota
Public SysOptBioCodeForUChol() As String 'Chol
Public SysOptBioCodeForUCA() As String 'Calcium
Public SysOptBioCodeForUPhos() As String 'Phos
Public SysOptBioCodeForUMag() As String 'Mag
Public SysOptShowIQ200() As Boolean

'phone Numbers
Public SysOptHaemPhone() As String  'haematology
Public SysOptBioPhone() As String   'Biochemistry
Public SysOptCoagPhone() As String  'Coagulation
Public SysOptBloodPhone() As String 'blood Trans
Public SysOptImmPhone() As String   'Immunology
Public SysOptEndPhone() As String   'Endocrinology
Public SysOptMicroPhone() As String   'Microbiology
Public SysOptCytoPhone() As String   'Cytology
Public SysOptHistoPhone() As String   'Histology
Public SysOptExtPhone() As String   'Histology


Public SysOptHaemAddress() As String  'haematology
Public SysOptBioAddress() As String   'Biochemistry
Public SysOptCoagAddress() As String  'Coagulation
Public SysOptBloodAddress() As String 'blood Trans
Public SysOptImmAddress() As String   'Immunology
Public SysOptEndAddress() As String   'Endocrinology
Public SysOptMicroAddress() As String   'Microbiology
Public SysOptCytoAddress() As String   'Cytology
Public SysOptHistoAddress() As String   'Histology
Public SysOptExtAddress() As String   'Histology


'offsets
Public SysOptSemenOffset() As Double '10,000,000
Public SysOptMicroOffset() As Double '20,000,000
Public SysOptHistoOffset() As Double '30,000,000
Public SysOptCytoOffset() As Double '40,000,000
Public Sub LoadOptions()

Dim tb As New Recordset
Dim sql As String
Dim n As Integer

On Error GoTo LoadOptions_Error

ReDimOptions

For n = 0 To intOtherHospitalsInGroup
  SysOptBioMaskSym(n) = " XXXXXX"
  sql = "SELECT * FROM Options " & _
  "order by ListOrder"
  
  Set tb = New Recordset
  RecOpenClient n, tb, sql
  Do While Not tb.EOF
    Select Case UCase$(Trim$(tb!Description & ""))
        Case "BIOMASKSYM": SysOptBioMaskSym(n) = " " & Trim(tb!Contents & "")
        Case "PRINTMIDDLE": SysOptPrintMiddle(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
        Case "ENDB12": SysOptEndCodeB12(n) = Trim(tb!Contents & "")
        Case "ENDB12NEW": SysOptEndCodeB12New(n) = Trim(tb!Contents & "")
        Case "ENDCO": SysOptEndCodeCo(n) = Trim(tb!Contents & "")
        Case "ENDFSH": SysOptEndCodeFSH(n) = Trim(tb!Contents & "")
        Case "ENDLH": SysOptEndCodeLH(n) = Trim(tb!Contents & "")
        Case "ENDPRO": SysOptEndCodePRO(n) = Trim(tb!Contents & "")
        Case "ENDOES": SysOptEndCodeOES(n) = Trim(tb!Contents & "")
        Case "ENDPRL": SysOptEndCodePRL(n) = Trim(tb!Contents & "")
        Case "ENDTHC": SysOptEndCodeTHC(n) = Trim(tb!Contents & "")
        Case "ENDTRO": SysOptEndCodeTRO(n) = Trim(tb!Contents & "")
        Case "ENDHBA1C": SysOptEndCodeHBA1C(n) = Trim(tb!Contents & "")
        Case "ENDCALCA1C": SysOptEndCodeCalcA1C(n) = Trim(tb!Contents & "")
        Case "ENDHBA1CCOMMENT": SysOptEndHbA1cComment(n) = Trim(tb!Contents & "")
        Case "ENDCALCA1CCOMMENT": SysOptEndCalcA1cComment(n) = Trim(tb!Contents & "")
        Case "ENDBNP": SysOptEndCodeBNP(n) = Trim(tb!Contents & "")
        Case "BIOBNP": SysOptBioCodeBNP(n) = Trim(tb!Contents & "")

        Case "ENDVITD": SysOptEndCodeVITD(n) = Trim(tb!Contents & "")

        Case "FAX": SysOptFax(n) = Trim(tb!Contents & "")
        Case "FAXSERVER": SysOptFaxServer(n) = Trim(tb!Contents & "")
        Case "BIOADD": SysOptBioAddress(n) = Trim(tb!Contents & "")
        Case "COAGADD": SysOptCoagAddress(n) = Trim(tb!Contents & "")
        Case "HAEMADD": SysOptHaemAddress(n) = Trim(tb!Contents & "")
        Case "IMMADD": SysOptImmAddress(n) = Trim(tb!Contents & "")
        Case "MICROADD": SysOptMicroAddress(n) = Trim(tb!Contents & "")
        Case "HISTOADD": SysOptHistoAddress(n) = Trim(tb!Contents & "")
        Case "CYTOADD": SysOptCytoAddress(n) = Trim(tb!Contents & "")
        Case "BIOCODEFORGLUCOSE": SysOptBioCodeForGlucose(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORFASTGLUCOSE": SysOptBioCodeForFastGlucose(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORGLUCOSE1": SysOptBioCodeForGlucose1(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORGLUCOSE2": SysOptBioCodeForGlucose2(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORGLUCOSE3": SysOptBioCodeForGlucose3(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORGLUCOSEP": SysOptBioCodeForGlucoseP(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORGLUCOSE1P": SysOptBioCodeForGlucose1P(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORGLUCOSE2P": SysOptBioCodeForGlucose2P(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORGLUCOSE3P": SysOptBioCodeForGlucose3P(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORFASTGLUCOSEP": SysOptBioCodeForFastGlucoseP(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORCHOL": SysOptBioCodeForChol(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORCHOLP": SysOptBioCodeForCholP(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORHDL": SysOptBioCodeForHDL(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORCHOLHDLRATIO": SysOptBioCodeForCholHDLRatio(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORTRIG": SysOptBioCodeForTrig(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORTRIGP": SysOptBioCodeForTrigP(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORHBA1C": SysOptBioCodeForHbA1c(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORCREAT": SysOptBioCodeForCreat(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORUCREAT": SysOptBioCodeForUCreat(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORUPROT": SysOptBioCodeForUProt(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORALB": SysOptBioCodeForAlb(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORGLOB": SysOptBioCodeForGlob(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORTPROT": SysOptBioCodeForTProt(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFOR24VOL": SysOptBioCodeFor24Vol(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFOR24UPROT": SysOptBioCodeFor24UProt(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFOREGFR": SysOptBioCodeForEGFR(n) = Trim$(tb!Contents & "")
  
        Case "BIOCODEFORRF": SysOptBioCodeForRF(n) = Trim$(tb!Contents & "")
  
        Case "BIOCODEFORUNA": SysOptBioCodeForUNa(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORUK": SysOptBioCodeForUK(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORUMAG": SysOptBioCodeForUMag(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORUPHOS": SysOptBioCodeForUPhos(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORUCA": SysOptBioCodeForUCA(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORUCHOL": SysOptBioCodeForUChol(n) = Trim$(tb!Contents & "")
        Case "BIOCODEFORUUREA": SysOptBioCodeForUUrea(n) = Trim$(tb!Contents & "")
        Case "BIOPHONE": SysOptBioPhone(n) = Trim$(tb!Contents & "")
        Case "COAGPHONE": SysOptCoagPhone(n) = Trim$(tb!Contents & "")
        Case "CYTOOFFSET": SysOptCytoOffset(n) = Val(Trim$(tb!Contents & ""))
        Case "CYTOPHONE": SysOptCytoPhone(n) = Trim$(tb!Contents & "")
        Case "DEPTBGA": SysOptDeptBga(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
        Case "DEPTBIO": SysOptDeptBio(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
        Case "DEPTCOAG": SysOptDeptCoag(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
        Case "DEPTCYTO": SysOptDeptCyto(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
        Case "DEPTEND": SysOptDeptEnd(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
        Case "DEPTEXT": SysOptDeptExt(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
        Case "DEPTHAEM": SysOptDeptHaem(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
        Case "DEPTHISTO": SysOptDeptHisto(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
        Case "DEPTIMM": SysOptDeptImm(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
        Case "DEPTMICRO": SysOptDeptMicro(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
        Case "DEPTSEMEN": SysOptDeptSemen(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)
        Case "ENDPHONE": SysOptEndPhone(n) = Trim$(tb!Contents & "")
        Case "ENDADD": SysOptEndAddress(n) = Trim(tb!Contents & "")
        Case "EXP": SysOptExp(n) = Trim$(tb!Contents & "")
        Case "EXTADD": SysOptExtAddress(n) = Trim(tb!Contents & "")
        Case "EXTPHONE": SysOptExtPhone(n) = Trim$(tb!Contents & "")
        Case "HAEMPHONE": SysOptHaemPhone(n) = Trim$(tb!Contents & "")
        Case "HISTOOFFSET": SysOptHistoOffset(n) = Val(Trim$(tb!Contents & ""))
        Case "HISTOPHONE": SysOptHistoPhone(n) = Trim$(tb!Contents & "")
        Case "IMMPHONE": SysOptImmPhone(n) = Trim$(tb!Contents & "")
        Case "MICROOFFSET": SysOptMicroOffset(n) = Val(Trim$(tb!Contents & ""))
        Case "MICROPHONE": SysOptMicroPhone(n) = Trim$(tb!Contents & "")
        Case "SEMENOFFSET": SysOptSemenOffset(n) = Val(Trim$(tb!Contents & ""))
        Case "SHOWIQ200":  SysOptShowIQ200(n) = IIf(Trim$(tb!Contents & "") = "1", True, False)

    End Select
    tb.MoveNext
  Loop
Next

Exit Sub

LoadOptions_Error:

Dim strES As String
Dim intEL As Integer

intEL = Erl
strES = Err.Description
LogError "basOptions", "LoadOptions", intEL, strES, sql

End Sub


Private Sub ReDimOptions()

ReDim SysOptBioCodeForEGFR(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioMaskSym(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeBNP(0 To intOtherHospitalsInGroup) As String
ReDim SysOptExtPhone(0 To intOtherHospitalsInGroup) As String   'External Phone
ReDim SysOptExtAddress(0 To intOtherHospitalsInGroup) As String   'External address
ReDim SysOptPrintMiddle(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysOptEndCodeB12(0 To intOtherHospitalsInGroup) As String
ReDim SysOptEndCodeHBA1C(0 To intOtherHospitalsInGroup) As String
ReDim SysOptEndCodeCalcA1C(0 To intOtherHospitalsInGroup) As String
ReDim SysOptEndHbA1cComment(0 To intOtherHospitalsInGroup) As String
ReDim SysOptEndCalcA1cComment(0 To intOtherHospitalsInGroup) As String
ReDim SysOptEndCodeBNP(0 To intOtherHospitalsInGroup) As String
ReDim SysOptEndCodeB12New(0 To intOtherHospitalsInGroup) As String
ReDim SysOptEndCodeCo(0 To intOtherHospitalsInGroup) As String
ReDim SysOptEndCodeFSH(0 To intOtherHospitalsInGroup) As String
ReDim SysOptEndCodeLH(0 To intOtherHospitalsInGroup) As String
ReDim SysOptEndCodePRO(0 To intOtherHospitalsInGroup) As String
ReDim SysOptEndCodeOES(0 To intOtherHospitalsInGroup) As String
ReDim SysOptEndCodePRL(0 To intOtherHospitalsInGroup) As String
ReDim SysOptEndCodeTHC(0 To intOtherHospitalsInGroup) As String
ReDim SysOptEndCodeTRO(0 To intOtherHospitalsInGroup) As String

ReDim SysOptEndCodeVITD(0 To intOtherHospitalsInGroup) As String


ReDim SysOptFax(0 To intOtherHospitalsInGroup) As String  'Fax
ReDim SysOptFaxServer(0 To intOtherHospitalsInGroup) As String  'Faxserver
ReDim SysOptHaemAddress(0 To intOtherHospitalsInGroup) As String  'haematology
ReDim SysOptBioAddress(0 To intOtherHospitalsInGroup) As String   'Biochemistry
ReDim SysOptCoagAddress(0 To intOtherHospitalsInGroup) As String  'Coagulation
ReDim SysOptBloodAddress(0 To intOtherHospitalsInGroup) As String 'blood Trans
ReDim SysOptImmAddress(0 To intOtherHospitalsInGroup) As String   'Immunology
ReDim SysOptEndAddress(0 To intOtherHospitalsInGroup) As String   'Endocrinology
ReDim SysOptMicroAddress(0 To intOtherHospitalsInGroup) As String   'Microbiology
ReDim SysOptCytoAddress(0 To intOtherHospitalsInGroup) As String   'Cytology
ReDim SysOptCytoPhone(0 To intOtherHospitalsInGroup) As String   'Cytology
ReDim SysOptHistoAddress(0 To intOtherHospitalsInGroup) As String   'Histology
ReDim SysOptNoCumShow(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysOptBlankSid(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysOptDemoVal(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysOptCommVal(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysSetFoc(0 To intOtherHospitalsInGroup) As String
ReDim SysOptDemo(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysOptDeptHaem(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysOptDeptBio(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysOptDeptCoag(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysOptDeptMicro(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysOptDeptImm(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysOptDeptEnd(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysOptDeptBga(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysOptDeptExt(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysOptDeptSemen(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysOptDeptCyto(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysOptDeptHisto(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysOptBioCodeForGlucoseP(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForCholP(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForTrigP(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForGlucose1P(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForGlucose2P(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForGlucose3P(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForFastGlucoseP(0 To intOtherHospitalsInGroup) As String

ReDim SysOptExp(0 To intOtherHospitalsInGroup) As Boolean
ReDim SysOptHaemN1(0 To intOtherHospitalsInGroup) As String
ReDim SysOptHaemN2(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioN1(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioN2(0 To intOtherHospitalsInGroup) As String
ReDim SysOptChange(0 To intOtherHospitalsInGroup) As Boolean

ReDim SysOptBioCodeForCreat(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForUCreat(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForUProt(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForAlb(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForGlob(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForTProt(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeFor24UProt(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeFor24Vol(0 To intOtherHospitalsInGroup) As String

ReDim SysOptBioCodeForRF(0 To intOtherHospitalsInGroup) As String

ReDim SysOptBioCodeForUCA(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForUChol(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForUMag(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForUK(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForUNa(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForUPhos(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForUUrea(0 To intOtherHospitalsInGroup) As String

ReDim SysOptHaemPhone(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioPhone(0 To intOtherHospitalsInGroup) As String
ReDim SysOptCoagPhone(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBloodPhone(0 To intOtherHospitalsInGroup) As String
ReDim SysOptImmPhone(0 To intOtherHospitalsInGroup) As String
ReDim SysOptMicroPhone(0 To intOtherHospitalsInGroup) As String
ReDim SysOptCytoPhone(0 To intOtherHospitalsInGroup) As String
ReDim SysOptHistoPhone(0 To intOtherHospitalsInGroup) As String
ReDim SysOptEndPhone(0 To intOtherHospitalsInGroup) As String

ReDim SysOptSemenOffset(0 To intOtherHospitalsInGroup) As Double '10,000,000
ReDim SysOptMicroOffset(0 To intOtherHospitalsInGroup) As Double '20,000,000
ReDim SysOptHistoOffset(0 To intOtherHospitalsInGroup) As Double '30,000,000
ReDim SysOptCytoOffset(0 To intOtherHospitalsInGroup) As Double '30,000,000

ReDim SysOptBioCodeForGlucose(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForFastGlucose(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForGlucose1(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForGlucose2(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForGlucose3(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForChol(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForHDL(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForTrig(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForCholHDLRatio(0 To intOtherHospitalsInGroup) As String
ReDim SysOptBioCodeForHbA1c(0 To intOtherHospitalsInGroup) As String
ReDim SysOptShowIQ200(0 To intOtherHospitalsInGroup) As Boolean

End Sub




