VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rangeBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public rFrom As Variant
Public rTo As Variant
Public rWidth As Integer
Public rPadded As Boolean
Public rPadChar As String
Public rPaddedDir As String
Public rType As String 'int, char
Public rFormat As String

Private Sub Class_Initialize()

    On Error Resume Next
    
    rFrom = -1
    rTo = -1
    rWidth = -1
    rPadded = False
    rPadChar = "0"
    rType = "int"
    rFormat = ""
    
End Sub

