VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRuleDefn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarName As String
Private mvarValues As Collection


Private Sub Class_Initialize()

    On Error Resume Next
    mvarName = "<DefaultRuleName>"
    Set mvarValues = New Collection
    
End Sub

Public Property Get contents() As Collection
    On Error Resume Next
    Set contents = mvarValues
End Property

Public Property Let name(aNameString As String)
    On Error Resume Next
    mvarName = aNameString
    
End Property

Public Property Get name() As String
    On Error Resume Next
    name = mvarName
End Property
