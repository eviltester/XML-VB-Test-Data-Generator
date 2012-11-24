VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "repeatBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public blockContents As Object
Public rMin As Variant
Public rMax As Variant
Public rRnd As Boolean ' either random repeat within limits, or mandatory use these limits

Private Sub Class_Initialize()

    On Error Resume Next
    rMin = -1
    rMax = -1
    rRnd = True
    
End Sub

