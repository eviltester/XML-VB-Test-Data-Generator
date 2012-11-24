VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGUIDGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CoCreateGuid_Alt Lib "OLE32.DLL" Alias "CoCreateGuid" (pGuid _
    As Any) As Long
Private Declare Function StringFromGUID2_Alt Lib "OLE32.DLL" Alias "StringFromGUID2" _
    (pGuid As Any, ByVal address As Long, ByVal Max As Long) As Long

Public bracesOn As Boolean

Public Function CreateGUID() As String
    Dim res As String, resLen As Long, guid(15) As Byte
    res = Space$(128)
    CoCreateGuid_Alt guid(0)
    resLen = StringFromGUID2_Alt(guid(0), ByVal StrPtr(res), 128)
    CreateGUID = Left$(res, resLen - 1)
    If bracesOn = False Then
        'remove braces
        If Left$(CreateGUID, 1) = "{" Then
            CreateGUID = Mid$(CreateGUID, 2)
        End If
        If Right$(CreateGUID, 1) = "}" Then
            CreateGUID = Left$(CreateGUID, Len(CreateGUID) - 1)
        End If
    End If
    
End Function

Private Sub Class_Initialize()
    On Error Resume Next
    bracesOn = False
End Sub
