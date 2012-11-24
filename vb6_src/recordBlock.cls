VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "recordBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public blockContents As Object
Public name As String
Public CSVHeader As Boolean ' if a csv file do we use the fieldnames to build a header?



Private Sub Class_Initialize()
    On Error Resume Next
    name = "record"
    CSVHeader = False
End Sub
