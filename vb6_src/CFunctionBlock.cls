VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFunctionBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public name As String
Private mvarParams As Collection
Private mvarParamNames As Collection
Public blockContents As Object


Private Sub Class_Initialize()
    On Error Resume Next
    name = "<unnamedFunction>"
    Set mvarParams = New Collection
    Set mvarParamNames = New Collection
End Sub

Public Sub addParam(aparamname As String, aValue As Variant)
    On Error Resume Next
    mvarParams.Remove aparamname
    mvarParams.Add aValue, aparamname
    mvarParamNames.Add aparamname, aparamname
End Sub

Public Function getValue(aparamname As String) As Variant
    On Error Resume Next
    getValue = mvarParams.Item(aparamname)
End Function

Public Function paramExists(aparamname As String) As Boolean
    On Error Resume Next
    Dim aSt As Variant
    aSt = ""
    paramExists = False
    aSt = mvarParamNames.Item(aparamname)
    If aSt <> "" Then
        paramExists = True
    End If
End Function




