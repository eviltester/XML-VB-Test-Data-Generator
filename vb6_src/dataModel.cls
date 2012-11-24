VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "dataModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public theSets As Collection
'Public theSetNames As Collection
Public theRules As Collection
'Public theRuleNames As Collection
Public theOutputRules As Collection
'Public theOutputRuleNames As Collection
Public theVariables As Collection   ' constructed with the <As varname="bob">
Public theVariableNames As Collection

Private outputFileID As Integer
Private theRecEnd As String

Public MAXREPEAT As Integer

Private outputAs As String

' private for CSV header output
Private mvarLastOutputRecordName As String
Private mvarCSVHeaderString As String
Private mvarCSVHeaders As Collection  ' key will be outputRecordName|recordName
Private mvarCSVState As Integer     ' 0 - not processing
                                    ' 1 - building header
                                    ' 2 - header built and output
Private mvarGlobalCSVHeader As Boolean

Private Sub Class_Initialize()

    On Error Resume Next
    
        Set theSets = New Collection
        'Set theSetNames = New Collection
        
        Set theRules = New Collection
        'Set theRuleNames = New Collection
        
        Set theOutputRules = New Collection
        'Set theOutputRuleNames = New Collection
        
        Set theVariables = New Collection
        Set theVariableNames = New Collection
        
        MAXREPEAT = 100
        outputFileID = -1
        outputAs = "TEXT"
        mvarLastOutputRecordName = ""
        mvarCSVHeaderString = ""
        Set mvarCSVHeaders = New Collection
        mvarCSVState = 0
        mvarGlobalCSVHeader = False
        
End Sub

Public Sub outputRule(aName As String, afilePath As String, aRecordTerminator As String)

    On Error GoTo outputRuleFailed
    
    outputFileID = FreeFile
    Open afilePath For Output Access Write Lock Read Write As #outputFileID
    
    Dim aRule As Collection
    Dim aBlock As Object
    
    Dim anOutputRule As COutputRule
    Set anOutputRule = theOutputRules.Item(aName)
    If anOutputRule Is Nothing Then GoTo outputRuleFailed
    
    outputAs = anOutputRule.theType
    
    Set aRule = anOutputRule.contents
    
    If aRule Is Nothing Then GoTo outputRuleFailed
    
    mvarLastOutputRecordName = aName
    theRecEnd = aRecordTerminator
    For Each aBlock In aRule
        performBlock aBlock
    Next
    
    ' remove any csv headers created in my name
    ' at this point they would all be created in my name so delete them all
    Set mvarCSVHeaders = New Collection
    
outputRuleFailed:

    If outputFileID <> -1 Then
        Close #outputFileID
    End If
    outputFileID = -1
    
End Sub
Public Function exampleRules() As String

    On Error Resume Next
    Dim aSetName As Variant
    Dim aSetValues As Collection
    Dim aSetValue As Variant
    Dim aCount As Integer
    
    Dim aSet As CSetDefn
    Dim aRule As CRuleDefn
    Dim anOutputRule As COutputRule
    
    Dim aVarName As Variant
    Dim aVariable As String
    
    exampleRules = ""
    
    'print the sets
    If Not theSetNames Is Nothing Then
        If theSetNames.Count > 0 Then
            exampleRules = exampleRules & "Sets:" & vbCrLf
            'For Each aSetName In theSetNames
            For Each aSet In theSets
                'Set aSet = theSets.Item(aSetName & "")
                aSetName = aSet.name
                
                exampleRules = exampleRules & aSetName & " := {"
                aCount = 0
                
                Set aSetValues = aSet.values
                For Each aSetValue In aSetValues
                    exampleRules = exampleRules & """" & aSetValue & """"
                    aCount = aCount + 1
                    If aCount < aSetValues.Count Then
                        exampleRules = exampleRules & ","
                    End If
                Next
                exampleRules = exampleRules & "};" & vbCrLf
            Next
            
            exampleRules = exampleRules & vbCrLf & "rules:" & vbclrf
            'For Each aSetName In theRuleNames
            For Each aRule In theRules
                aSetName = aRule.name
                exampleRules = exampleRules & aSetName & " := ["
                exampleRules = exampleRules & performRule(aSetName & "")
                exampleRules = exampleRules & "]" & vbCrLf
            Next
            
            If theVariables.Count > 0 Then
                exampleRules = exampleRules & vbCrLf & "Variables:" & vbclrf
                For Each aVarName In theVariableNames
                    exampleRules = exampleRules & aVarName & " := "
                    exampleRules = exampleRules & theVariables.Item(aVarName & "") & vbCrLf
                Next
            End If
        End If
    End If
    


End Function
Public Sub debugPrint()

    On Error Resume Next
    Debug.Print exampleRules
        
End Sub

Public Function randomInt(lowerbound As Integer, upperbound As Integer) As Integer
    On Error Resume Next
    randomInt = randomNumberCode.randomInt(lowerbound, upperbound)
    
End Function

Public Function randomLong(lowerbound As Long, upperbound As Long) As Long
    On Error Resume Next
    
    If upperbound < lowerbound Then
        randomLong = Int((lowerbound - upperbound + 1) * Rnd + upperbound)
    Else
        randomLong = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
    End If

End Function

Public Function randomDouble(lowerbound As Double, upperbound As Double) As Double
    On Error Resume Next
    
    If upperbound < lowerbound Then
        randomDouble = (lowerbound - upperbound) * Rnd + upperbound
    Else
        randomDouble = (upperbound - lowerbound) * Rnd + lowerbound
    End If

End Function

Public Function performRule(aName As String) As String

    On Error Resume Next
    
    Dim aRuleContents As Collection
    Dim aRule As CRuleDefn
    Dim aBlock As Object
    
    Set aRule = theRules.Item(aName)
    If aRule Is Nothing Then Exit Function
    
    Set aRuleContents = aRule.contents
    If aRuleContents Is Nothing Then Exit Function
    
    For Each aBlock In aRuleContents
        performRule = performRule & performBlock(aBlock)
    Next
    
End Function

Public Function performSet(aName As String) As String

    On Error Resume Next
    
    Dim aSetValues As Collection
    Dim aSet As CSetDefn
    
    Set aSet = theSets.Item(aName)
    If aSet Is Nothing Then Exit Function
    
    Dim aFileSet As Boolean
    aFileSet = False
    Set aSetValues = aSet.values
    ' following code not required 22/03/05 as the values are
    ' pulled in to memory when required (only done once)
'    If aSetValues Is Nothing Then
'        ' is it a fileset?
'        aFileSet = True
'    End If
'    If aSetValues.Count = 0 Then
'        aFileSet = True
'    End If
'    If aFileSet Then
'        If aSet.FileName <> "" Then
'            'this will be horribly slow when big files are used on big output rules
'            ' either keep it open with the len in place or load it into memory
'            Select Case aSet.FileType
'            Case "ROWS"
'                Dim aCSVParser As CSVParse
'                Set aCSVParser = New CSVParse
'                aCSVParser.FileName = aSet.FileName
'                aCSVParser.FieldSeperator = ","
'                performSet = aCSVParser.getRandomLine()
'            End Select
'        End If
'        Exit Function
'    End If
       
    performSet = performSetCollection(aSetValues)
    
End Function

Public Function performSetCollection(aSet As Collection) As String

    On Error Resume Next

    Dim lB As Long
    Dim uB As Long

    If aSet Is Nothing Then Exit Function

    lB = 1
    uB = aSet.Count
    
    performSetCollection = aSet.Item(randomLong(lB, uB))


End Function
Public Function performSetOperation(aBlock As setOperationBlock) As String

    On Error Resume Next
    'build a collection
    Dim aSet As Collection
    Dim anOpSet As Collection
    Dim aSetName As Variant
    Dim aSetElement As Variant
    Dim aCheckSetElement As Variant
    
    Dim setCount As Integer
    Dim itemCount As Integer
    
    Set aSet = New Collection
    
    setCount = 1
    
    For Each aSetName In aBlock.setsInvolved
        Select Case aBlock.operationType
        Case -1
        Case 1 ' union
        
            On Error Resume Next
            Set anOpSet = Nothing
            Set anOpSet = theSets.Item(aSetName).values
            If Not anOpSet Is Nothing Then
                For Each aSetElement In anOpSet
                    aSet.Add aSetElement, aSetElement
                Next
            End If
            
        Case 2 ' intersection
        
            Err.Clear
            On Error Resume Next
            
            If setCount = 1 Then
                'first time around just add everything
                Set anOpSet = Nothing
                Set anOpSet = theSets.Item(aSetName).values
                If Not anOpSet Is Nothing Then
                    For Each aSetElement In anOpSet
                        aSet.Add aSetElement, aSetElement
                    Next
                End If
            Else
                Set anOpSet = Nothing
                Set anOpSet = theSets.Item(aSetName).values
                If Not anOpSet Is Nothing Then
                    itemCount = aSet.Count
                    Do While itemCount > 1
                        aSetElement = aSet.Item(itemCount)
                        Err.Clear
                        aCheckSetElement = anOpSet.Item(aSetElement)
                        If Err.Number > 0 Then
                            'it does not exist in anopset
                            'so remove it from aset
                            aSet.Remove itemCount
                        End If
                        itemCount = itemCount - 1
                    Loop
                End If
            End If
        Case 3 ' difference
            'x E A, x !E B, x !E C, ...
            On Error Resume Next
            
            If setCount = 1 Then
                'first time around just add everything
                Set anOpSet = Nothing
                Set anOpSet = theSets.Item(aSetName).values
                If Not anOpSet Is Nothing Then
                    For Each aSetElement In anOpSet
                        aSet.Add aSetElement, aSetElement
                    Next
                End If
            Else
                Set anOpSet = Nothing
                Set anOpSet = theSets.Item(aSetName).values
                If Not anOpSet Is Nothing Then
                    itemCount = aSet.Count
                    Do While itemCount > 1
                        aSetElement = aSet.Item(itemCount)
                        Err.Clear
                        aCheckSetElement = anOpSet.Item(aSetElement)
                        If Err.Number = 0 Then
                            'it exists in anopset
                            'so remove it from aset
                            aSet.Remove itemCount
                        End If
                        itemCount = itemCount - 1
                    Loop
                End If
            End If
        End Select
    Next
    
    performSetOperation = performSetCollection(aSet)
    
End Function
Public Function performBlock(aBlock As Object) As String

    'given a block, evaluate it and return the result
    Dim aBlockColEntry As Object
    
    
    On Error Resume Next
    
    Dim repcount As Long
    Dim lCount As Long
    Dim mCount As Long
    Dim retStr As String
    Dim aCol As Collection
    
    Dim aRule As CRuleDefn
    Dim aSet As CSetDefn
    Dim aVar As String
    
    Dim tempString As String
        
    Select Case True
    Case TypeOf aBlock Is Collection
    
        For Each aBlockColEntry In aBlock
            performBlock = performBlock & performBlock(aBlockColEntry)
        Next
        
    Case TypeOf aBlock Is optionalBlock
        
        Dim retDbl As Double
        retDbl = randomDouble(0.00001, 1#)
        If (retDbl) <= aBlock.probability Then
            ' use the probability and if less than or equal then do
            performBlock = performBlock & performBlock(aBlock.blockContents)
        Else
            'Debug.Print "Not probable"
        End If
        
    Case TypeOf aBlock Is setOperationBlock
    
        performBlock = performBlock & performSetOperation(aBlock)
    
    Case TypeOf aBlock Is recordBlock
        ' by only blanking out the output when writing to a file, it should allow us to test the rule
        If outputFileID <> -1 Then
            performBlock = ""
        End If
        
        Select Case outputAs
        Case "TEXT"
            'output as is
        Case "XML"
            'output the record name as the header
            performBlock = performBlock & "<" & ModStringFuncs.stringAsXML(aBlock.name) & ">" & vbCrLf
        Case "CSV"
            ' if it is CSV and we have been asked to output a CSVHEADER then we should build the header
            ' as we process the fields then write the header, then write the contents of the record
            ' but only do this once for each record name - in each outputRecord
            
            'if header already exists then we do not do anything
            tempString = mvarCSVHeaders(mvarLastOutputRecordName & "|" & aBlock.name)
            If Err.Number = 0 Then
                'header exists
                mvarCSVState = 0
            Else
                'header does not exist
                'if we are supposed to build a header
                If aBlock.CSVHeader Or mvarGlobalCSVHeader Then
                    mvarCSVState = 1
                    mvarCSVHeaderString = ""
                End If
            End If
        End Select
        
        tempString = ""
        
        For Each aBlockColEntry In aBlock
            tempString = tempString & performBlock(aBlock.blockContents)
        Next
        
        If mvarCSVState = 1 Then
            mvarCSVHeaderString = removeTrailingComma(mvarCSVHeaderString)
            performBlock = performBlock & mvarCSVHeaderString & theRecEnd & tempString
            mvarCSVState = 2
            mvarCSVHeaders.Add mvarCSVHeaderString, mvarLastOutputRecordName & "|" & aBlock.name
        Else
            performBlock = performBlock & tempString
        End If
        
        Select Case outputAs
        Case "TEXT"
            'output as is
        Case "CSV"
            'remove trailing ,
            performBlock = removeTrailingComma(performBlock)
        Case "XML"
            'output the record name as the header
            performBlock = performBlock & "</" & ModStringFuncs.stringAsXML(aBlock.name) & ">"
        End Select

        If outputFileID <> -1 Then
            Print #outputFileID, performBlock & theRecEnd;
        Else
            ' to allow us to test the rule
            performBlock = performBlock & theRecEnd
        End If
        
    Case TypeOf aBlock Is CFieldBlock
        
        Select Case outputAs
        Case "TEXT"
        Case "CSV"
            If mvarCSVState = 1 Then
                mvarCSVHeaderString = mvarCSVHeaderString & ModStringFuncs.stringAsCSVField(aBlock.name) & ","
            End If
            'output as is
        Case "XML"
            performBlock = performBlock & "<" & ModStringFuncs.stringAsXML(aBlock.name) & ">"
        End Select
        

        tempString = ""
        
        For Each aBlockColEntry In aBlock
            tempString = performBlock(aBlock.blockContents)
        Next
        
        Select Case outputAs
        Case "TEXT"
            'output as is
             performBlock = performBlock & tempString
        Case "XML"
            performBlock = performBlock & ModStringFuncs.stringAsXML(tempString)
        Case "CSV"
           ' output as is here
           performBlock = performBlock & ModStringFuncs.stringAsCSVField(tempString) & ","
        End Select
        
        
        Select Case outputAs
        Case "TEXT", "CSV"
            'output as is
        Case "XML"
            'output the record name as the header
            performBlock = performBlock & "</" & ModStringFuncs.stringAsXML(aBlock.name) & ">" & vbCrLf
        End Select
                
    Case TypeOf aBlock Is rangeBlock
    
        ' generate a value in the range
        If StrComp(UCase$(aBlock.rType), "INT") = 0 Then
            lCount = aBlock.rFrom
            mCount = aBlock.rTo
            repcount = randomLong(lCount, mCount)
            
            'now format it
            If aBlock.rWidth <> -1 Then
                retStr = String$(aBlock.rWidth, aBlock.rPadChar)
                If Left$(UCase$(aBlock.rPaddedDir), 1) = "L" Then
                    retStr = retStr & repcount
                    performBlock = Right$(retStr, aBlock.rWidth)
                Else
                    retStr = repcount & retStr
                    performBlock = Left$(retStr, aBlock.rWidth)
                End If
            End If
        ElseIf StrComp(UCase$(aBlock.rType), "CHAR") = 0 Then
            lCount = Asc(aBlock.rFrom & "")
            mCount = Asc(aBlock.rTo & "")
            
            
            repcount = randomLong(lCount, mCount)
            
            'now format it
            If aBlock.rWidth <> -1 Then
                retStr = String$(aBlock.rWidth, aBlock.rPadChar)
                If Left$(UCase$(aBlock.rPaddedDir), 1) = "L" Then
                    retStr = retStr & Chr$(repcount)
                    performBlock = Right$(retStr, aBlock.rWidth)
                Else
                    retStr = Chr$(repcount) & retStr
                    performBlock = Left$(retStr, aBlock.rWidth)
                End If
            Else
                performBlock = Chr$(repcount)
            End If
        ElseIf StrComp(UCase$(aBlock.rType), "DATE") = 0 Then
            performBlock = rangeDate(aBlock.rFrom, aBlock.rTo, aBlock.rFormat)
        End If
        
    Case TypeOf aBlock Is repeatBlock
        
        'decide how often to loop then loop around it
        If aBlock.rMin = -1 Then
            lCount = 1
        Else
            lCount = aBlock.rMin
        End If
        
        If aBlock.rMax = -1 Then
            'randomly decide on a max
            Dim tLong As Long
            tLong = MAXREPEAT
            mCount = randomLong(lCount, tLong)
        Else
            If aBlock.rRnd Then
                mCount = randomLong(lCount, aBlock.rMax)
            Else
                mCount = aBlock.rMax
            End If
        End If
        
        For repcount = 1 To mCount
            performBlock = performBlock & performBlock(aBlock.blockContents)
        Next
        
    Case TypeOf aBlock Is termBlock
    
        If aBlock.isLiteral Then
            performBlock = performBlock & aBlock.blockContents
        Else
            ' is it a rule or a set?
            Set aRule = Nothing
            Set aRule = theRules.Item(aBlock.blockContents)
            If aRule Is Nothing Then
                'Set aCol = theSets.Item(aBlock.blockContents & "").values
                Set aSet = Nothing
                Set aSet = theSets.Item(aBlock.blockContents)
                If aSet Is Nothing Then
                    'check for var
                    aVar = ""
                    aVar = theVariables.Item(aBlock.blockContents)
                    performBlock = aVar
                Else
                    performBlock = performBlock & performSet(aBlock.blockContents)
                End If
            Else
                performBlock = performBlock & performRule(aBlock.blockContents)
            End If
            
        End If
        
    Case TypeOf aBlock Is choiceBlock
    
        'work out which of the options to take
        Dim choiceID As Integer
        Dim anyBlock As Object
        Dim totalID As Integer
        
        choiceID = randomInt(1, aBlock.totalWeighting)
        
        totalID = 0
        For Each anyBlock In aBlock.blockContents
            totalID = totalID + anyBlock.aWeighting
            If choiceID <= totalID Then
                'pick this one
                performBlock = performBlock & performBlock(anyBlock)
                Exit For
            End If
        Next
        
    Case TypeOf aBlock Is optionBlock
        performBlock = performBlock & performBlock(aBlock.blockContents)
        
    Case TypeOf aBlock Is CFunctionBlock
        performBlock = performBlock & performFunctionOn(aBlock, aBlock.blockContents)
    End Select

End Function

Private Function removeTrailingComma(aSTring As String) As String
    On Error Resume Next
    'remove trailing ,
    If Right$(aSTring, 1) = "," Then
        removeTrailingComma = Left$(aSTring, Len(aSTring) - 1)
    Else
        removeTrailingComma = aSTring
    End If
            
End Function
Public Function performFunctionOn(aBlock As CFunctionBlock, aCol As Collection) As String
    On Error Resume Next
    
    Dim retStr As String
    Dim aBorder As String
    Dim aLen As Integer     ' length of string from right, or left or substr
    Dim aStart As Integer   ' substr start char
    Dim maxlenV As Variant  ' variant to convert nums
    Dim aName As String
    
    
    Select Case UCase$(aBlock.name)
    Case "CSVHEADER"
        ' turn on the global CSVheader if appropriate
        aName = "<unnamed>"
        If aBlock.paramExists("show") Then
            aName = aBlock.getValue("show")
            If StrComp(UCase$(aName), "TRUE") = 0 Then
                ' turn csv on
                mvarGlobalCSVHeader = True
            Else
                mvarGlobalCSVHeader = False
            End If
        End If
     End Select
    
    ' perform the block
    retStr = performBlock(aCol)
    
    ' now perform the function on that string
    Select Case UCase$(aBlock.name)
    Case "GUID"
        'guid children are ignored
        Dim aGUID As CGUIDGenerator
        Dim bracesOn As Boolean
        Set aGUID = New CGUIDGenerator
        
        
        If aBlock.paramExists("braces") Then
            If StrComp(UCase$(aBlock.getValue("braces")), "YES") = 0 Then
                aGUID.bracesOn = True
            End If
        End If
        
        performFunctionOn = aGUID.CreateGUID()
        
    Case "TRIM"
        'Trim children are processed first
        Dim doL As Boolean
        Dim doR As Boolean
        
        doL = True
        doR = True
        If aBlock.paramExists("border") Then
            aBorder = UCase$(aBlock.getValue("border"))
            'default is full trim, but we are specifying l or right
            doL = False: doR = False
            If InStr(1, aBorder, "L") > 0 Then
                doL = True
            End If
            If InStr(1, aBorder, "R") > 0 Then
                doR = True
            End If
        End If
        If doL And doR Then
            performFunctionOn = Trim(retStr)
        ElseIf doL Then
            performFunctionOn = LTrim(retStr)
        ElseIf doR Then
            performFunctionOn = RTrim(retStr)
        Else
            'what else?!?!??!
            performFunctionOn = retStr
        End If
        
    Case "PADSTRING"
        ' vars maxlen=int border=LorR with=" "
        Dim withStr As String
        withStr = " "
        maxlen = -1

        aBorder = "L"
        
        If aBlock.paramExists("border") Then
            aBorder = UCase$(aBlock.getValue("border"))
            'default is full trim, but we are specifying l or right
            aBorder = Left$(aBorder, 1)
            Select Case UCase$(aBorder)
            Case "L"
                aBorder = "L"
            Case "R"
                aBorder = "R"
            Case Else
                aBorder = "L"
            End Select
        End If
        
        If aBlock.paramExists("maxlen") Then
            maxlenV = aBlock.getValue("maxlen")
            maxlen = maxlenV
        End If
        
        If aBlock.paramExists("with") Then
            withStr = aBlock.getValue("with")
            withStr = Left$(withStr, 1)
            If withStr = "" Then
                withStr = " "
            End If
        End If
        
        If maxlen > 0 Then
            Dim paddy As String
            paddy = String(maxlen, withStr)
            Select Case aBorder
            Case "L"
                retStr = paddy & retStr
                retStr = Right$(retStr, maxlen)
            Case "R"
                retStr = retStr & paddy
                retStr = Left$(retStr, maxlen)
            Case Else
                retStr = paddy & retStr
                retStr = Right$(retStr, maxlen)
            End Select
            performFunctionOn = retStr
        Else
            performFunctionOn = retStr
        End If
        
    Case "RIGHT", "LEFT"
        'get the rightmost or leftmost len characters
        
        aLen = Len(retStr)
        If aBlock.paramExists("len") Then
            maxlenV = aBlock.getValue("len")
            aLen = maxlenV
        End If
        
        If Len(retStr) <= aLen Then
            performFunctionOn = retStr
        Else
            Select Case UCase$(aBlock.name)
            Case "RIGHT"
                performFunctionOn = Right$(retStr, aLen)
            Case "LEFT"
                performFunctionOn = Left$(retStr, aLen)
            End Select
        End If
        
    Case "SUBSTR"
        'get the rightmost or leftmost len characters
        
        aLen = Len(retStr)
        If aBlock.paramExists("len") Then
            maxlenV = aBlock.getValue("len")
            aLen = maxlenV
        End If
        
        aStart = 1
        If aBlock.paramExists("start") Then
            maxlenV = aBlock.getValue("start")
            aStart = maxlenV
        End If
        
        performFunctionOn = Mid$(retStr, aStart, aLen)
    
    Case "AS"
        ' create a variable and add the enclosed rule text as the variable, and pass it back
        aName = "<unnamed>"
        If aBlock.paramExists("name") Then
            aName = aBlock.getValue("name")
            If aName = "" Then
                aName = "<unnamed>"
            End If
        End If
    
        performFunctionOn = retStr
        theVariables.Remove (aName)
        theVariables.Add retStr, aName
        theVariableNames.Add aName, aName
        
    Case "CSVHEADER"
        'turnoff the header as the function is finished now
        mvarGlobalCSVHeader = False
        performFunctionOn = retStr
    End Select
    
    
End Function

Public Function rangeDate(fromDate As Variant, toDate As Variant, Optional aFormat As String = "dddddd ttttt") As String

    On Error Resume Next
    Dim fDate As Date
    Dim tDate As Date
    
    fDate = fromDate
    tDate = toDate
    
    Dim lCount As Double
    Dim mCount As Double
    
    lCount = fDate
    mCount = tDate
    
    rangeDate = Format$(randomDouble(lCount, mCount), aFormat) '"Long Date")

End Function

