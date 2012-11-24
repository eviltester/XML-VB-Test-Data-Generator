Attribute VB_Name = "ModCommandLine"
Option Explicit

Public silentMode As Boolean

Sub Main()
   Dim a_strArgs() As String
   Dim blnDebug As Boolean
   Dim strFilename As String
   
   Dim i As Integer
   
   silentMode = False
   
   'command lines are
   ' -s silent mode
   ' -i input.xml file
   ' -o name of output rule
   ' -of name of output file
   ' these can be repeated as often as necessary
   
   Dim processingWhat As Integer    ' 0 - nothing
                                    ' 1 - inputxml file name
                                    ' 2 - output rule name
                                    ' 3 - output file name
                                    
    processingWhat = 0
    
    Dim inputXMLFile As String
    Dim outputRuleName As String
    Dim outputFileName As String
    
   If Command$ <> "" Then
        a_strArgs = Split(Command$, " ")
        For i = LBound(a_strArgs) To UBound(a_strArgs)
           Select Case LCase(a_strArgs(i))
           Case "-s", "/s"
                silentMode = True
           Case "-i", "/i"
                ifAllThenDoIt inputXMLFile, outputRuleName, outputFileName
                inputXMLFile = ""
                outputRuleName = ""
                outputFileName = ""
                processingWhat = 1
            Case "-o", "/o"
                'if we already have an -o and a -of then
                ' do it and reset
                ifAllThenDoIt inputXMLFile, outputRuleName, outputFileName
                outputRuleName = ""
                outputFileName = ""
                
                processingWhat = 2
           Case "-of", "/of"
                'if we already have an -o and a -of then
                ' do it and reset
                ifAllThenDoIt inputXMLFile, outputRuleName, outputFileName
                outputFileName = ""
                processingWhat = 3
                
           Case Else
                'reconstruct arguments with spaces
                Select Case processingWhat
                Case 1
                    If inputXMLFile = "" Then
                        inputXMLFile = a_strArgs(i)
                    Else
                        inputXMLFile = inputXMLFile & " " & a_strArgs(i)
                    End If
                Case 2
                    If outputRuleName = "" Then
                        outputRuleName = a_strArgs(i)
                    Else
                        outputRuleName = outputRuleName & " " & a_strArgs(i)
                    End If
                Case 3
                    If outputFileName = "" Then
                        outputFileName = a_strArgs(i)
                    Else
                        outputFileName = outputFileName & " " & a_strArgs(i)
                    End If
                Case Else
                    MsgBox "Invalid argument: " & a_strArgs(i)
                End Select
           End Select
           
        Next i
        
        ifAllThenDoIt inputXMLFile, outputRuleName, outputFileName
    
    Else
        Form1.Show
    End If
    
End Sub

Public Sub ifAllThenDoIt(xmlPath As String, outputRule As String, outputFilePath As String)

    If xmlPath = "" Then Exit Sub
    If outputRule = "" Then Exit Sub
    If outputFilePath = "" Then Exit Sub
    
    doOutputOnFile xmlPath, outputRule, outputFilePath
    
End Sub
Public Sub doOutputOnFile(xmlPath As String, outputRule As String, outputFilePath As String)

    On Error Resume Next
    Dim aC As xmlDataParser
    Dim aDM As dataModel
    
    Dim msgB As String
    
    msgB = ""
    msgB = msgB & "running in batch mode with: " & vbCrLf
    msgB = msgB & "         xml: " & xmlPath & vbCrLf
    msgB = msgB & "  outputRule: " & outputRule & vbCrLf
    msgB = msgB & "  outputFile: " & outputFilePath & vbCrLf
    
    If Not silentMode Then
        MsgBox msgB, vbOKOnly
    End If
    
    Set aDM = New dataModel
    Set aC = New xmlDataParser
    
    aC.Init xmlPath, aDM
    aDM.outputRule outputRule, outputFilePath, vbCrLf
    
End Sub
