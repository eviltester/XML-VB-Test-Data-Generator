VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSVParse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim fileID As Integer                           ' A FreeFile fileID
Private mvarLineCount       As Integer          ' initial count of the number of lines
Private mvarFileName        As String           ' Path to the target file
Private mvarFieldCount      As Integer          ' Count of fields in this record
Private mvarStatus          As Boolean          ' Are we able to provide data?
Private strRecord           As String           ' Current record
Private Fields()            As String           ' Current record field array
Private strErrMsg           As String           ' Last error message
Private mvarFieldSeperator  As String           ' User defined field seperator

' this code does not handle embedded , in quoted strings
' and will return quoted strings

'------------------------------------------------------------------------------
'   CSV Parser
'   This class handles retrieving elements from a CSV (C_omma S_eperated V_alues)
'   file. In the CSV file each line is a record and each field in the record is
'   seperated from its neighbor by a delimiter character. The character is usually
'   a comma (,) but can be any character.
'
'   This class requires a reference to the MS Scripting Runtime.
'
'   Create an instance of the class (Dim CSVP as New CSVParse)
'   Set the FieldSeperator property if it is not comma.
'   Set the FileName property using the full path to the target file.
'      a. Read the Status property. If it is false, the file was not
'         accessed so call the GetErrorMessage function to retrieve the
'         descripition of the problem
'   Process the file as follows:
'
'       While CSVP.LoadNextLine = True
'           MyString = CSVP.GetField(n) <- for each field you want to read
'                                          where n is the field number where
'           .                              1 is the first field.
'           .
'           .
'       Wend
'----------------------------------------------------------------------------
'   Return the message string
Public Function GetErrorMessage() As String
    GetErrorMessage = strErrMsg
End Function

'   Set the field delimiter character. Default is the comma.
Public Property Let FieldSeperator(ByVal vData As String)
    mvarFieldSeperator = Trim(vData)
End Property

Public Property Get FieldSeperator() As String
    FieldSeperator = mvarFieldSeperator
End Property

'   Internal status set
Private Property Let Status(ByVal vData As Boolean)
    mvarStatus = vData
End Property

'   Tell caller the status
Public Property Get Status() As Boolean
    Status = mvarStatus
End Property

'   Give out the number of fields in this record
Public Property Get FieldCount() As Integer
    FieldCount = mvarFieldCount
End Property

'   Set the target file name
Public Property Let FileName(ByVal vData As String)
    mvarFileName = vData                            'Set the file path
    mvarStatus = AccessTargetFile()                 'Open it
End Property

' how many lines the file has
Public Property Get lineCount() As Integer
    On Error Resume Next
    lineCount = mvarLineCount
End Property

'   Reads the next line of text and parses it into fields array
Public Function LoadNextLine() As Boolean
'    On Error GoTo LoadNextLine_Err
'    If TS.AtEndOfStream Then
'        LoadNextLine = False
'        Exit Function
'    End If
'    strRecord = TS.ReadLine
'    ReDim Fields(0)
'    Fields = Split(strRecord, FieldSeperator)       'Break out the string of fields
'    mvarFieldCount = UBound(Fields) + 1             '# of fields available
'    LoadNextLine = True
'    Exit Function
'LoadNextLine_Err:
'    LoadNextLine = False                        'Should happen at EOF...
End Function

Public Function getFieldValueAtLine(fieldID As Integer, lineID As Integer) As String
        
    On Error Resume Next
    Dim aFieldValue As String
    Dim aLineValue As String
    
    aLineValue = getLineID(lineID)
    
    ReDim Fields(0)
    Fields = Split(aLineValue, FieldSeperator)       'Break out the string of fields
    mvarFieldCount = UBound(Fields) + 1             '# of fields available
    getFieldValueAtLine = GetField(fieldID)

End Function

Public Function getRandomLine() As String

    On Error Resume Next
    AccessTargetFile
    getRandomLine = getLineID(randomNumberCode.randomInt(1, mvarLineCount))
    
End Function

Public Function getFieldIDFromRandomLine(aFieldID As Integer) As String

    On Error Resume Next
    AccessTargetFile
    getFieldIDFromRandomLine = getFieldValueAtLine(aFieldID, randomNumberCode.randomInt(1, mvarFieldCount))
    
End Function

'return the line given a line id
Public Function getLineID(anID As Integer) As String

    On Error GoTo getLineID_Err
    
    fileID = FreeFile
    
    Dim NextLine As String
    Dim aLineCount As Integer
        
    getLineID = ""
    aLineCount = 0
    Open mvarFileName For Input Shared As #fileID
    Do Until EOF(fileID)
        Line Input #fileID, NextLine
        aLineCount = aLineCount + 1
        If aLineCount = anID Then
            Exit Do
        End If
    Loop
    getLineID = NextLine
    Close #fileID
    
    fileID = -1
    
getLineID_Err:
    If fileID <> -1 Then
        Close #fileID
    End If
    fileID = -1
    

End Function

'   Pass back the specified field
Public Function GetField(FieldNum As Integer) As String
    If FieldNum < 1 Or FieldNum > FieldCount Then
        GetField = ""
    Else
        GetField = Trim(Fields(FieldNum - 1))
    End If
End Function

'   Open the target file...
Private Function AccessTargetFile() As Boolean
    On Error Resume Next
    
    'close if open
    If fileID <> -1 Then
        Close #fileID
    End If
    
    On Error GoTo AccessTargetFile_Err
    Status = True
    strErrMsg = ""
    fileID = FreeFile
    
    Dim NextLine As String

    ' count the number of lines
    mvarLineCount = 0
    Open mvarFileName For Input Shared As #fileID
    Do Until EOF(fileID)
        Line Input #fileID, NextLine
        mvarLineCount = mvarLineCount + 1
    Loop
    Close #fileID
    
    fileID = -1
    AccessTargetFile = True
    Status = True
    Exit Function
AccessTargetFile_Err:
    strErrMsg = CStr(Err.Number) & " " & Err.Description & " in AccessTargetFile."
    AccessTargetFile = False
    fileID = -1
End Function

Public Function getAllLinesAsCollection() As Collection
    On Error Resume Next
    
    'close if open
    If fileID <> -1 Then
        Close #fileID
    End If
    
    On Error GoTo getAllLinesAsCollection_Err
    
    Set getAllLinesAsCollection = New Collection
        
    Status = True
    strErrMsg = ""
    fileID = FreeFile
    
    Dim NextLine As String

    ' count the number of lines
    mvarLineCount = 0
    Open mvarFileName For Input Shared As #fileID
    Do Until EOF(fileID)
        Line Input #fileID, NextLine
        mvarLineCount = mvarLineCount + 1
        getAllLinesAsCollection.Add NextLine, NextLine
    Loop
    Close #fileID
    
    fileID = -1

    Status = True
    Exit Function
getAllLinesAsCollection_Err:
    strErrMsg = CStr(Err.Number) & " " & Err.Description & " in getAllLinesAsCollection."
    Status = False
    fileID = -1
End Function

Public Function getAllFieldsAsCollection(aColID As Integer) As Collection
    On Error Resume Next
    
    'close if open
    If fileID <> -1 Then
        Close #fileID
    End If
    
    On Error GoTo getAllLinesAsCollection_Err
    
    Set getAllFieldsAsCollection = New Collection
        
    Status = True
    strErrMsg = ""
    fileID = FreeFile
    
    Dim NextLine As String

    ' count the number of lines
    mvarLineCount = 0
    Open mvarFileName For Input Shared As #fileID
    Do Until EOF(fileID)
        Line Input #fileID, NextLine
        mvarLineCount = mvarLineCount + 1
        
        ReDim Fields(0)
        Fields = Split(NextLine, FieldSeperator)       'Break out the string of fields
        mvarFieldCount = UBound(Fields) + 1             '# of fields available
        NextLine = GetField(aColID)
        getAllFieldsAsCollection.Add NextLine, NextLine
        
    Loop
    Close #fileID
    
    fileID = -1

    Status = True
    Exit Function
getAllLinesAsCollection_Err:
    strErrMsg = CStr(Err.Number) & " " & Err.Description & " in getAllFieldsAsCollection."
    Status = False
    fileID = -1
End Function
'   Normal VB class initialization
Private Sub Class_Initialize()
    Status = False                          'Not open for business yet
    FieldSeperator = ","                    'Default to comma
    mvarFileName = ""                       'No file yet
    fileID = -1
End Sub

'   Normal VB class termination
Private Sub Class_Terminate()
    If fileID <> -1 Then
        Close #fileID
    End If
    fileID = -1
End Sub

