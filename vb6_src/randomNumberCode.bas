Attribute VB_Name = "randomNumberCode"
Option Explicit

Public Function randomInt(lowerbound As Integer, upperbound As Integer) As Integer

    On Error Resume Next
    
    If upperbound < lowerbound Then
        randomInt = Int((lowerbound - upperbound + 1) * Rnd + upperbound)
    Else
        randomInt = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
    End If
    
End Function
