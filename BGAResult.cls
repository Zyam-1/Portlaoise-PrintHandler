VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BGAResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Rundate As String
Public RunDateTime As String
Public pH As String
Public PCO2 As String
Public PO2 As String
Public HCO3 As String
Public BE As String
Public O2SAT As String
Public TotCO2 As String
Public Valid As Boolean
Public Printed As Boolean
Public NOPAS As String
Public SampleID As String

