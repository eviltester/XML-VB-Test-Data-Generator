VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xmlDataParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public aDataModel As dataModel
Private theFileName As String


Public Sub Init(aFileName As String, ByRef theDataModel As dataModel)
    'open and parse the file and setup the class
    

        Dim oDoc As MSXML2.DOMDocument
        Dim fSuccess As Boolean
        Dim oRoot As MSXML2.IXMLDOMNode
        Dim oChild As MSXML2.IXMLDOMNode

        
        On Error GoTo HandleErr
        
        Set oDoc = New MSXML2.DOMDocument
        
        theFileName = aFileName     ' make the filename available to other functions in the parser

        Set aDataModel = theDataModel
        
        ' Load the  XML from disk, without validating it. Wait
        ' for the load to finish before proceeding.
        oDoc.async = False
        oDoc.validateOnParse = False
        fSuccess = oDoc.Load(aFileName)
         ' If anything went wrong, quit now.
        If Not fSuccess Then
            MsgBox "Error during document load of " & aFileName
          GoTo ExitHere
        End If
    
    
 
        ' Get the root of the XML tree.
        Set oRoot = oDoc.documentElement
         ' Go through all children of the root.
        For Each oChild In oRoot.childNodes
            'what is it?
            Select Case oChild.baseName
            Case "Sets"
                inputGenericItems oChild, "SET"
                'inputSets oChild
            Case "Rules"
                inputGenericItems oChild, "RULE"
                'inputRules oChild
            Case "OutputRules"
                inputGenericItems oChild, "OUTPUTRULE"
                'inputOutputRules oChild
            End Select
        Next
             
    
    

ExitHere:
        Exit Sub
HandleErr:
        MsgBox "Error " & Err.Number & ": " & Err.Description
        Resume ExitHere
        Resume
    
End Sub

' theType is SET or RULE or OUTPUTRULE
Public Sub inputGenericItems(theRoot As MSXML2.IXMLDOMNode, theType As String)

    On Error GoTo HandleErr
    
    Dim oThing As MSXML2.IXMLDOMNode
    Dim oAttributes As MSXML2.IXMLDOMNamedNodeMap
    Dim oAttribute As MSXML2.IXMLDOMNode
    Dim oElement As MSXML2.IXMLDOMNode
    
    Dim aGlobalCol As Collection
    Dim aCol As Collection
    Dim aName As String
    Dim aValue As String
    
    Dim aThing As Object
    Dim aVar As Variant
    
    For Each oThing In theRoot.childNodes
        'read in the definition
        'get the name
        aName = ""
        Set oAttribute = Nothing
        
        Set oAttributes = oThing.Attributes
          ' Extract the name
          
        If TypeOf oThing Is IXMLDOMComment Then
            GoTo skipComment
        End If
        
          Set oAttribute = oAttributes.getNamedItem("name")
          aName = oAttribute.Text
        
        Select Case theType
        Case "SET"
            Set aThing = New CSetDefn
            Set aGlobalCol = aDataModel.theSets
            Set aCol = aThing.values
            
            'set can have:
            '   filename
            '   filetype
            ' Extract the name
            Dim aFileName As String
            Dim aFileType As String
            Set oAttribute = Nothing
            Set oAttribute = oAttributes.getNamedItem("fileName")
            If oAttribute Is Nothing Then
                ' it is not a file type set
                
            Else
                'get all the file details
                aFileName = oAttribute.Text
                
                Set oAttribute = Nothing
                Set oAttribute = oAttributes.getNamedItem("fileType")
                If oAttribute Is Nothing Then
                    ' it is not a file type set
                    aFileName = ""
                    aFileType = ""
                Else
                    'get the file type details
                    aFileType = oAttribute.Text
                End If
                
                
                Select Case aFileType
                Case "ROWS"
                Case "CSVCOLS"
                    'get the colID
                    
                    Set oAttribute = Nothing
                    Set oAttribute = oAttributes.getNamedItem("colID")
                    If oAttribute Is Nothing Then
                        aThing.ColID = 1    'default
                    Else
                        'convert it to an int from text
                        aVar = oAttribute.Text
                        aThing.ColID = aVar
                    End If
                    
                End Select
                
                If aFileName <> "" Then
                    Dim fCheck As String
                    fCheck = Dir(aFileName)
                    If fCheck = "" Then
                        Dim aPath As String
                        aPath = Left$(theFileName, _
                                        Len(theFileName) - _
                                        InStr(StrReverse(theFileName), "\") + 1)
                    End If
                    fCheck = Dir(aPath & aFileName)
                    If fCheck <> "" Then
                        aFileName = aPath & aFileName
                    End If
                    'if the file could not be found then carry on regardless
                    aThing.FileName = aFileName
                    aThing.FileType = aFileType
                End If
                
            End If
            
        Case "RULE"
            Set aThing = New CRuleDefn
            Set aGlobalCol = aDataModel.theRules
            Set aCol = aThing.contents
        Case "OUTPUTRULE"
            Set aThing = New COutputRule
            Set aGlobalCol = aDataModel.theOutputRules
            Set aCol = aThing.contents
            
                'get optional type for output rule
                Set oAttribute = Nothing
                Set oAttribute = oAttributes.getNamedItem("type")
                If oAttribute Is Nothing Then
                    ' it is not a file type set
                    aFileName = ""
                    aFileType = ""
                Else
                    'get the file type details
                    aThing.theType = UCase$(oAttribute.Text)
                End If
                
            
            
        End Select
        
        aThing.name = aName
        aGlobalCol.Add aThing, aName
        
        'now read in the thing details
        
        ' we can have set operations union, intersection etc.
        ' also <Defn>
        ' but at the moment we have only coded for <Element>
        For Each oElement In oThing.childNodes
            Select Case theType
                Case "SET"
                    If StrComp("Element", oElement.nodeName) = 0 Then
                        aValue = ""
                        aValue = oElement.nodeTypedValue
                        On Error Resume Next
                        ' it is a set so we do not allow duplicate values
                        aCol.Add aValue, aValue
                        
                        If Err.Number <> 0 Then
                            Err.Clear
                        End If
                        
                        On Error GoTo HandleErr
                    End If
                
                Case "RULE", "OUTPUTRULE"
                    readRuleDetails aCol, oElement
            End Select
        Next
        
skipComment:
    Next
    


            
ExitHere:
        Exit Sub
HandleErr:
        MsgBox "Error in input" & theType & " " & Err.Number & ": " & Err.Description
        Resume ExitHere
        Resume

End Sub


Public Sub readRuleDetails(ByRef theRule As Collection, theRuleElement As MSXML2.IXMLDOMNode)

    On Error GoTo HandleErr
    
    Dim anOptionalBlock As optionalBlock
    Dim aRecordBlock As recordBlock
    Dim aRepeatBlock As repeatBlock
    Dim aTermBlock As termBlock
    Dim aRangeBlock As rangeBlock
    Dim aChoiceBlock As choiceBlock
    Dim anOptionBlock As optionBlock
    Dim aSetOperation As setOperationBlock
    Dim aFunctionBlock As CFunctionBlock
    Dim aFieldBlock As CFieldBlock
    Dim anyBlock As Object
    
    Dim aBlockChild  As MSXML2.IXMLDOMNode
    
    Dim oAttributes As MSXML2.IXMLDOMNamedNodeMap
    Dim oAttribute As MSXML2.IXMLDOMNode

    Dim aVar As Variant
    
    'skip xml comments
    If TypeOf theRuleElement Is IXMLDOMComment Then Exit Sub
    
    Select Case theRuleElement.nodeName
    Case "Optional"
        'can contain stuff
        Set anOptionalBlock = Nothing
        Set anOptionalBlock = New optionalBlock
        Set anOptionalBlock.blockContents = New Collection
        
         ' Extract the probability
        Set oAttribute = Nothing
        Set oAttributes = theRuleElement.Attributes
        Set oAttribute = oAttributes.getNamedItem("probability")
        If oAttribute Is Nothing Then
            'already defaulted to 0.5
        Else
            aVar = oAttribute.Text
            anOptionalBlock.probability = oAttribute.Text
        End If
        
        theRule.Add anOptionalBlock
        For Each aBlockChild In theRuleElement.childNodes
            readRuleDetails anOptionalBlock.blockContents, aBlockChild
        Next
        
    Case "Record"
        'can contain stuff, when output acts as a record
        Set aRecordBlock = Nothing
        Set aRecordBlock = New recordBlock
        Set aRecordBlock.blockContents = New Collection
        
         ' Extract the optional name
        Set oAttribute = Nothing
        Set oAttributes = theRuleElement.Attributes
        Set oAttribute = oAttributes.getNamedItem("name")
        If oAttribute Is Nothing Then
            'already defaulted to 0.5
        Else
            aVar = oAttribute.Text
            aRecordBlock.name = oAttribute.Text
        End If
        
        Set oAttribute = oAttributes.getNamedItem("csvheader")
        If oAttribute Is Nothing Then
            'already defaulted to false
        Else
            If StrComp(UCase$(oAttribute.Text), "TRUE") = 0 Then
                aRecordBlock.CSVHeader = True
            End If
        End If
        
        theRule.Add aRecordBlock
        For Each aBlockChild In theRuleElement.childNodes
            readRuleDetails aRecordBlock.blockContents, aBlockChild
        Next
        
        
    Case "Field"
        'can contain stuff in a record
        Set aFieldBlock = Nothing
        Set aFieldBlock = New CFieldBlock
        Set aFieldBlock.blockContents = New Collection
        
        
         ' Extract the optional name
        Set oAttribute = Nothing
        Set oAttributes = theRuleElement.Attributes
        Set oAttribute = oAttributes.getNamedItem("name")
        If oAttribute Is Nothing Then
            'already defaulted to 0.5
        Else
            aVar = oAttribute.Text
            aFieldBlock.name = oAttribute.Text
        End If
        
        'get the optional type
        Set oAttribute = oAttributes.getNamedItem("type")
        If oAttribute Is Nothing Then
            'already defaulted to 0.5
        Else
            aVar = oAttribute.Text
            aFieldBlock.fieldType = oAttribute.Text
        End If
        
        theRule.Add aFieldBlock
        For Each aBlockChild In theRuleElement.childNodes
            readRuleDetails aFieldBlock.blockContents, aBlockChild
        Next
        
        
    Case "Choice"
        'can contain stuff
        Set aChoiceBlock = Nothing
        Set aChoiceBlock = New choiceBlock
        Set aChoiceBlock.blockContents = New Collection
        
        theRule.Add aChoiceBlock
        For Each aBlockChild In theRuleElement.childNodes
            readRuleDetails aChoiceBlock.blockContents, aBlockChild
        Next
        
        'calculate the global weighting
        For Each anyBlock In aChoiceBlock.blockContents
            If TypeOf anyBlock Is optionBlock Then
                aChoiceBlock.totalWeighting = aChoiceBlock.totalWeighting + anyBlock.aWeighting
            End If
        Next
        
    Case "Option"
        'can contain stuff
        Set anOptionBlock = Nothing
        Set anOptionBlock = New optionBlock
        Set anOptionBlock.blockContents = New Collection
        
        'get the repeat details
        Set oAttributes = Nothing
        Set oAttributes = theRuleElement.Attributes
        
        ' Extract the name
        Set oAttribute = Nothing
        Set oAttribute = oAttributes.getNamedItem("weighting")
        If oAttribute Is Nothing Then
            anOptionBlock.aWeighting = 1
        Else
            anOptionBlock.aWeighting = oAttribute.Text
        End If
        
        theRule.Add anOptionBlock
        
        For Each aBlockChild In theRuleElement.childNodes
            readRuleDetails anOptionBlock.blockContents, aBlockChild
        Next
        
    Case "SetOperation"
    
        'can contain stuff
        Set aSetOperation = Nothing
        Set aSetOperation = New setOperationBlock
        
        'get the type details
        Set oAttributes = Nothing
        Set oAttributes = theRuleElement.Attributes
        
        ' Extract the name
        Set oAttribute = Nothing
        Set oAttribute = oAttributes.getNamedItem("type")
        If Not oAttribute Is Nothing Then
            aSetOperation.setType (oAttribute.Text)
        End If
        
        theRule.Add aSetOperation
            
        For Each aBlockChild In theRuleElement.childNodes
            If StrComp("OperatesOn", aBlockChild.nodeName) = 0 Then
                aSetOperation.addSetName aBlockChild.Text
            End If
        Next
    
    Case "Repeat"
        'can contain stuff
        Set aRepeatBlock = Nothing
        Set aRepeatBlock = New repeatBlock
        Set aRepeatBlock.blockContents = New Collection
        
        theRule.Add aRepeatBlock
        
        'get the repeat details
        Set oAttributes = Nothing
        Set oAttributes = theRuleElement.Attributes
        

        
        ' Extract the name
        Set oAttribute = Nothing
        Set oAttribute = oAttributes.getNamedItem("from")
        If oAttribute Is Nothing Then
            aRepeatBlock.rMin = 1
        Else
            aRepeatBlock.rMin = oAttribute.Text
        End If
        
        
        Set oAttribute = Nothing
        Set oAttribute = oAttributes.getNamedItem("to")
        If oAttribute Is Nothing Then
            aRepeatBlock.rMax = aDataModel.MAXREPEAT
        Else
            ' it is not a random as we have set a max
            aRepeatBlock.rMax = oAttribute.Text
        End If
          
        Set oAttribute = Nothing
        Set oAttribute = oAttributes.getNamedItem("fixed")
        If oAttribute Is Nothing Then
            aRepeatBlock.rRnd = False
        Else
            Select Case UCase$(oAttribute.Text)
            Case "TRUE"
                aRepeatBlock.rRnd = True
            Case "FALSE"
                aRepeatBlock.rRnd = False
            End Select
        End If

        
        'repeat what?
        For Each aBlockChild In theRuleElement.childNodes
            readRuleDetails aRepeatBlock.blockContents, aBlockChild
        Next
        
    Case "Term"
    
        
        Set aTermBlock = Nothing
        Set aTermBlock = New termBlock
    
        theRule.Add aTermBlock
        
        Set oAttributes = Nothing
        Set oAttributes = theRuleElement.Attributes
        
        ' Extract the name or the value
        Set oAttribute = Nothing
        Set oAttribute = oAttributes.getNamedItem("name")
        If oAttribute Is Nothing Then
            'no name, it is a value term
            aTermBlock.blockContents = theRuleElement.nodeTypedValue
            aTermBlock.isLiteral = True
        Else
            'it is a reference to something else (either a var, or a rule or a set)
            aTermBlock.blockContents = oAttribute.Text
            aTermBlock.isLiteral = False
        End If
        
    Case "Range"
    
        Set aRangeBlock = Nothing
        Set aRangeBlock = New rangeBlock
    
        theRule.Add aRangeBlock
        
        Set oAttributes = Nothing
        Set oAttributes = theRuleElement.Attributes
        
        'width="5" padded="left" paddedWith="0" from="0" to="99999"
        '<Range from="a" to="Z" />
        
        ' Extract the type
        Set oAttribute = Nothing
        Set oAttribute = oAttributes.getNamedItem("type")
        If Not oAttribute Is Nothing Then
            aRangeBlock.rType = oAttribute.nodeTypedValue
        End If
        
        Set oAttribute = Nothing
        Set oAttribute = oAttributes.getNamedItem("format")
        If Not oAttribute Is Nothing Then
            aRangeBlock.rFormat = oAttribute.nodeTypedValue
        End If
        
        ' Extract the width
        Set oAttribute = Nothing
        Set oAttribute = oAttributes.getNamedItem("width")
        If Not oAttribute Is Nothing Then
            aRangeBlock.rWidth = oAttribute.nodeTypedValue
        End If
    
        ' Extract the padded
        Set oAttribute = Nothing
        Set oAttribute = oAttributes.getNamedItem("padded")
        If Not oAttribute Is Nothing Then
            aRangeBlock.rPadded = True
            aRangeBlock.rPaddedDir = oAttribute.Text
        End If
    
        ' Extract the paddedWith char
        Set oAttribute = Nothing
        Set oAttribute = oAttributes.getNamedItem("paddedWith")
        If Not oAttribute Is Nothing Then
            aRangeBlock.rPadChar = Left$(oAttribute.Text, 1)
        End If
    
        ' Extract the from value
        Set oAttribute = Nothing
        Set oAttribute = oAttributes.getNamedItem("from")
        If Not oAttribute Is Nothing Then
            aRangeBlock.rFrom = oAttribute.Text
        End If
    
        ' Extract the to value
        Set oAttribute = Nothing
        Set oAttribute = oAttributes.getNamedItem("to")
        If Not oAttribute Is Nothing Then
            aRangeBlock.rTo = oAttribute.Text
        End If
        
    Case Else
        ' we will add it as a function block
        Set oAttributes = Nothing
        Set oAttributes = theRuleElement.Attributes
        
        Set aFunctionBlock = New CFunctionBlock
        
        theRule.Add aFunctionBlock
        
        aFunctionBlock.name = theRuleElement.nodeName
        Set aFunctionBlock.blockContents = New Collection
        
        For Each oAttribute In oAttributes
            ' add the attribute as a parameter
            aFunctionBlock.addParam oAttribute.nodeName, oAttribute.Text
        Next

        'apply function to what?
        For Each aBlockChild In theRuleElement.childNodes
            readRuleDetails aFunctionBlock.blockContents, aBlockChild
        Next
    End Select
    
    

ExitHere:
        Exit Sub
HandleErr:
        MsgBox "Error in readRuleDetails " & Err.Number & ": " & Err.Description
        Resume ExitHere
        Resume
        
End Sub




