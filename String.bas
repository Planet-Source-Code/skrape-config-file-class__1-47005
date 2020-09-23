Attribute VB_Name = "modString"
Option Explicit

Public Function ReturnBeforeBackSlash(strInput As String) As String

    Dim intCounter As Integer
    
    For intCounter = 1 To Len(strInput)
        If Mid(strInput, intCounter, 1) = "\" Then
            ReturnBeforeBackSlash = Left(strInput, intCounter - 1)
            Exit For
        End If
    Next
            
End Function

Public Function ReturnToLastBackSlash(strInput As String) As String

    Dim intCounter As Integer
    
    For intCounter = Len(strInput) To 1 Step -1
        If Mid(strInput, intCounter, 1) = "\" Then
            ReturnToLastBackSlash = Left(strInput, intCounter - 1)
            Exit For
        End If
    Next
            
End Function

Public Function ReturnAfterBackSlash(strInput As String) As String

    Dim intCounter As Integer
    
    For intCounter = 1 To Len(strInput)
        If Mid(strInput, intCounter, 1) = "\" Then
            ReturnAfterBackSlash = Right(strInput, Len(strInput) - intCounter)
            Exit For
        End If
    Next
            
End Function

Public Function ReturnAfterLastBackslash(strInput As String) As String

    Dim intCounter As Integer
    
    For intCounter = Len(strInput) To 1 Step -1
        If Mid(strInput, intCounter, 1) = "\" Then
            ReturnAfterLastBackslash = Right(strInput, Len(strInput) - intCounter)
            Exit For
        End If
    Next
    
End Function

Public Function StripNulls(strInput As String) As String

    Dim intCounter As Integer
    Dim strTempString As String
    
    strTempString = ""
    
    For intCounter = 1 To Len(strInput)
        If Mid(strInput, intCounter, 1) <> Chr(0) Then
            strTempString = strTempString + Mid(strInput, intCounter, 1)
        End If
    Next
    
    StripNulls = strTempString

End Function

Public Function StripNullsAndSpaces(strInput As String) As String

    Dim intCounter As Integer
    Dim strTempString As String
    
    strTempString = ""
    
    For intCounter = 1 To Len(strInput)
        If Mid(strInput, intCounter, 1) <> Chr(0) And Mid(strInput, intCounter, 1) <> Chr(32) Then
            strTempString = strTempString + Mid(strInput, intCounter, 1)
        End If
    Next
    
    StripNullsAndSpaces = strTempString

End Function


' NOTE TO SELF - DOES NOT WORK, DO NOT USE
Public Function StripCRAndLF(strInput As String) As String

    Dim intCounter As Integer
    Dim strTempString As String
    
    strTempString = ""
    
    For intCounter = 1 To Len(strInput)
        If Mid(strInput, intCounter, 1) <> Chr(0) And Mid(strInput, intCounter, 1) <> Chr(32) Then
            strTempString = strTempString + Mid(strInput, intCounter, 1)
        End If
    Next
    
    'StripNullsAndSpaces = strTempString

End Function

Public Function CountCommas(strInput As String) As Integer

    Dim intCounter As Integer
    
    For intCounter = 1 To Len(strInput)
        If Mid(strInput, intCounter, 1) = "," Then
            CountCommas = CountCommas + 1
        End If
    Next
    
End Function

Public Function ReturnToComma(strInput As String) As String

    Dim intCounter As Integer
    
    For intCounter = 1 To Len(strInput)
        If Mid(strInput, intCounter, 1) = "," Then
            ReturnToComma = Left(strInput, intCounter - 1)
            strInput = Right(strInput, Len(strInput) - intCounter)
            Exit For
        End If
    Next
            
End Function

Public Function ReturnToString(strInput As String, strTargetString) As String

    Dim intCounter As Integer
    
    For intCounter = 1 To Len(strInput)
        If Mid(strInput, intCounter, Len(strTargetString)) = strTargetString Then
            ReturnToString = Left(strInput, intCounter - 1)
            strInput = Right(strInput, Len(strInput) - intCounter - Len(strTargetString) + 1)
            Exit Function
        End If
    Next
            
    ReturnToString = ""
    
End Function

Public Function ReturnToChar(strInput As String, strTarget As String) As String

    Dim intCounter As Integer
    
    For intCounter = 1 To Len(strInput)
        If Mid(strInput, intCounter, 1) = strTarget Then
            ReturnToChar = Left(strInput, intCounter - 1)
            strInput = Right(strInput, Len(strInput) - intCounter)
            Exit Function
        End If
    Next
            
    ReturnToChar = strInput
    strInput = ""

End Function

Public Function CountOccurances(strInput As String, strTarget As String) As Integer

    Dim intCounter As Integer
    Dim intLength As Integer
    
    
    intLength = Len(strTarget)
    
    For intCounter = 1 To Len(strInput)
        If Mid(strInput, intCounter, intLength) = strTarget Then
            CountOccurances = CountOccurances + 1
        End If
    Next
    
End Function

