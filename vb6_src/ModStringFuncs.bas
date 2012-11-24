Attribute VB_Name = "ModStringFuncs"
Option Explicit

Public Function stringAsCSVField(aText As String, Optional alwaysDoubleQuote As Boolean = False)

    Dim quoteIt As Boolean
    
    ' based on info from http://www.edoceo.com/utilis/csv-file-format.php
'# Each record is one line - Line separator may be LF (0x0A) or CRLF (0x0D0A), a line seperator may also be embedded in the data (making a record more than one line but still acceptable).
'# Fields are separated with commas. - Duh.
'# Leading and trailing whitespace is ignored - Unless the field is delimited with double-quotes in that case the whitespace is preserved.
'# Embedded commas - Field must be delimited with double-quotes.
'# Embedded double-quotes - Embedded double-quote characters must be doubled, and the field must be delimited with double-quotes.
'# Embedded line-breaks - Fields must be surounded by double-quotes.
'# Always Delimiting - Fields may always be delimited with double quotes, the delimiters will be parsed and discarded by the reading applications
    
    stringAsCSVField = aText
    quoteIt = alwaysDoubleQuote
    
    If Not quoteIt Then
        'check to see if need to quote it
        If Len(Trim(stringAsCSVField)) <> Len(stringAsCSVField) Then
            quoteIt = True
        End If
        ' does it have a double quote?
        If InStr(1, stringAsCSVField, """") > 0 Then
            'turn all " in to ""
            stringAsCSVField = Replace(stringAsCSVField, """", """" & """")
            quoteIt = True
        End If
        If InStr(1, stringAsCSVField, ",") > 0 Then
            quoteIt = True
        End If
        If InStr(1, stringAsCSVField, vbCrLf) > 0 Then
            quoteIt = True
        End If
    End If
    
    If quoteIt Then
        stringAsCSVField = """" & stringAsCSVField & """"
    End If

End Function

Public Function stringAsXML(aText As String) As String

    On Error Resume Next
    
 '&lt; = <
'&gt; = >
'&apos; = '
'&quot; = "
'&amp; = &

    stringAsXML = aText
    stringAsXML = Replace(stringAsXML, "&", "&amp;")
    stringAsXML = Replace(stringAsXML, "<", "&lt;")
    stringAsXML = Replace(stringAsXML, ">", "&gt;")
    stringAsXML = Replace(stringAsXML, "'", "&apos;")
    stringAsXML = Replace(stringAsXML, """", "&quot;")

End Function
