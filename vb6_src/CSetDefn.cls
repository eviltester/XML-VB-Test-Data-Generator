VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSetDefn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarName As String
Private mvarFileName As String

' <SET name="aRowFile" filename="c:\datafiles\rows.txt" filetype="ROWS" />
' file path is assumed to be in the same place as the xml file when no path is given
' <SET name="aRowFile" filename="rows.txt" filetype="ROWS" />

' a set made from column 3 in the csv file datavals.csv
' <SET name="aRowFile" filename="datavals.csv" filetype="CSVCOLS" colID="3"/>

' a set made from column 3 in the csv file datavals.csv with a header row - headerrow defaults to false
' <SET name="aRowFile" filename="datavals.csv" filetype="CSVCOLS" colID="3" headerRow="TRUE"/>


Private mvarFileType As String ' ROWS   'NOT IMPLEMENTED YET, CSVCOLS  ... future DATABASE? XML

' for CSVCOLS - choose either the colname
'Private mvarHeaderRow As Boolean ' does the file have a header row?
'Private mvarColName As String
Private mvarColID As Integer

Private mvarValues As Collection


Private Sub Class_Initialize()

    On Error Resume Next
    mvarName = "<DefaultSetName>"
    Set mvarValues = New Collection
    
End Sub

Public Property Get values() As Collection
    On Error Resume Next
    
    Dim aCSVParser As CSVParse
    
    If mvarFileName <> "" Then
        'if we are a filetype set then we need to create a set and return all the values in the file
        If mvarValues.Count = 0 Then
            'load the file in to our collection
            Select Case mvarFileType
            Case "ROWS"
            
                Set aCSVParser = New CSVParse
                aCSVParser.FileName = mvarFileName
                aCSVParser.FieldSeperator = ","
                Set mvarValues = aCSVParser.getAllLinesAsCollection()
                
            Case "CSVCOLS"

                Set aCSVParser = New CSVParse
                aCSVParser.FileName = mvarFileName
                aCSVParser.FieldSeperator = ","
                Set mvarValues = aCSVParser.getAllFieldsAsCollection(mvarColID)
            
            End Select
        End If
    End If
    
    Set values = mvarValues
    
End Property

Public Property Let name(aNameString As String)
    On Error Resume Next
    mvarName = aNameString
    
End Property

Public Property Get name() As String
    On Error Resume Next
    name = mvarName
End Property

Public Property Let FileName(aNameString As String)
    On Error Resume Next
    mvarFileName = aNameString
    
End Property

Public Property Get FileName() As String
    On Error Resume Next
    FileName = mvarFileName
End Property

Public Property Let FileType(aNameString As String)
    On Error Resume Next
    mvarFileType = aNameString
    
End Property

Public Property Get FileType() As String
    On Error Resume Next
    FileType = mvarFileType
End Property

Public Property Let ColID(aCol As Integer)
    On Error Resume Next
    mvarColID = aCol
    
End Property

Public Property Get ColID() As Integer
    On Error Resume Next
    ColID = mvarColID
End Property
