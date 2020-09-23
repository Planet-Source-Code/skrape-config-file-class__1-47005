Attribute VB_Name = "modFile"
Option Explicit

Public Function Locate(strSearchString As String, intFileHandle As Integer) As Boolean
    
    Dim strInput As String
    
    While Not EOF(intFileHandle)
        Line Input #intFileHandle, strInput
        If strInput = strSearchString Then
            Locate = True
            Exit Function
        End If
    Wend

    Locate = False
    Exit Function
    
End Function

Public Function BinaryReadToComma(intHandle As Integer, lngpointer As Long) As String

    Dim strTemp As String
    Dim bytInput As Byte
    
    Do While (bytInput <> 44)
        Get intHandle, lngpointer, bytInput
        strTemp = strTemp + Chr(bytInput)
        lngpointer = lngpointer + 1
    Loop
    
    BinaryReadToComma = strTemp
    
End Function

Public Function BinaryRead(intHandle As Integer, lngpointer As Long, lngLength As Long) As String

    Dim strTemp As String
    Dim bytInput As Byte
    Dim intCounter As Integer
    
    For intCounter = 1 To lngLength
        Get intHandle, lngpointer, bytInput
        strTemp = strTemp + Chr(bytInput)
        lngpointer = lngpointer + 1
    Next
        
    BinaryRead = strTemp
    
End Function


Public Sub BinaryWriteString(intHandle As Integer, lngpointer As Long, strInput As String)
    
    Dim intCounter As Integer
    Dim bytOutput As Byte
        
        
    For intCounter = 1 To Len(strInput)
        bytOutput = Asc(Mid(strInput, intCounter, 1))
        Put #intHandle, lngpointer, bytOutput
        lngpointer = lngpointer + 1
    Next
        
End Sub


' DOES NOT WORK DO NOT USE
Public Sub BinaryWriteStringInsert(intHandle As Integer, lngpointer As Long, strInput As String, lngEndOfFile As Long)
    
    Dim intCounter As Integer
    Dim bytOutput As Byte
    
    For intCounter = 1 To Len(strInput)
        bytOutput = Asc(Mid(strInput, intCounter, 1))
        Put intHandle, lngpointer, bytOutput
        lngpointer = lngpointer + 1
    Next
        
End Sub

' returns true is the file exsists, otherwise it returns false
Public Function FileExsists(strFilename As String) As Boolean

    Dim intHandle As Integer ' var for temporary file handle
    
    On Error GoTo ErrorHandler ' set up error catching
    
    intHandle = FreeFile ' get next available file number

    ' try to open the file, if there is an error the error handler will return false
    Open strFilename For Input As intHandle
    Close intHandle ' close the file, we just wanted to validate it, not use it
    
    FileExsists = True ' yes the file is really there
    
    Exit Function
    
    
ErrorHandler:
    'Call modErrorHandler.ErrorHandler
    'Exit Function
    
    Select Case Err.Number   ' Evaluate error number.
        Case 51 ' file does not exsist
            Close intHandle
            'MsgBox "File not found", , "Error"
            FileExsists = False
            Exit Function
        Case Else ' handle other situations here...
            'MsgBox "Error " + Str(Err.Number) + " - " + Err.Description
            FileExsists = False
            Exit Function
   End Select
   
End Function

Public Function BinaryReadToSmile(intHandle As Integer, lngpointer As Long) As String

    Dim strTemp As String
    Dim bytInput As Byte
       
    Do While (bytInput <> 1)
        Get intHandle, lngpointer, bytInput
        If bytInput = 0 Then Exit Do
        strTemp = strTemp + Chr(bytInput)
        lngpointer = lngpointer + 1
    Loop
    
    If Len(strTemp) > 1 Then BinaryReadToSmile = Left(strTemp, Len(strTemp) - 1)
    
End Function

Public Function BinaryReadToChar(intHandle As Integer, lngpointer As Long, intChar As Integer) As String

    Dim strTemp As String
    Dim bytInput As Byte
       
    Do While (bytInput <> intChar)
        Get intHandle, lngpointer, bytInput
        If bytInput = 0 Then Exit Do
        strTemp = strTemp + Chr(bytInput)
        lngpointer = lngpointer + 1
    Loop
    
    If Len(strTemp) > 1 Then BinaryReadToChar = Left(strTemp, Len(strTemp) - 1)
    
End Function


