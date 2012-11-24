VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "setOperationBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'public enum setOpTypes
Const UNION_OP As Integer = 1
Const INTERSECTION_OP As Integer = 2
Const DIFFERENCE_OP As Integer = 3

Public setsInvolved As Collection
Public operationType As Integer



Private Sub Class_Initialize()
    On Error Resume Next
    
    Set setsInvolved = New Collection
    operationType = -1
    
End Sub

Public Sub setType(aName As String)
    On Error Resume Next
    Select Case UCase$(aName)
    Case "UNION"
        operationType = UNION_OP
    Case "INTERSECTION"
        operationType = INTERSECTION_OP
    Case "DIFFERENCE"
        operationType = DIFFERENCE_OP
    End Select
    
End Sub

Public Sub addSetName(aName As String)
    On Error Resume Next
    setsInvolved.Add aName, aName
End Sub

