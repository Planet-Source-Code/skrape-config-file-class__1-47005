VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim intHandle As Integer
Dim lngPointer As Long
Private strFilename As String


Public Sub Load(Filename As String)

    lngPointer = 1
    intHandle = FreeFile
    Open strFilename For Binary Access Read Write As intHandle
    strFilename = Filename
    
End Sub

Public Sub Create(Filename As String)

    On Error Resume Next
    Kill Filename
    
    lngPointer = 1
    intHandle = FreeFile
    Open Filename For Binary Access Read Write As intHandle
    strFilename = Filename
    
End Sub

Public Sub Add(Description As String, Parameter As Variant)

    Dim lngTempPointer As Long
    Dim strTemp As String
    Dim strError As String
    
    
    If InStr(Description, "[") Then
        MsgBox "Description contains '['.", vbOKOnly + vbCritical, "Invalid Description"
        Exit Sub
    ElseIf InStr(Description, "]") Then
        MsgBox "Description contains ']'.", vbOKOnly + vbCritical, "Invalid Description"
        Exit Sub
    ElseIf InStr(Parameter, "]") Then
        MsgBox "Parameter contains ']'.", vbOKOnly + vbCritical, "Invalid Parameter"
        Exit Sub
    ElseIf InStr(Parameter, "[") Then
        MsgBox "Parameter contains '['.", vbOKOnly + vbCritical, "Invalid Parameter"
        Exit Sub
    End If
    
    lngTempPointer = 11
    Do While lngTempPointer < LOF(intHandle)
        strTemp = BinaryReadToChar(intHandle, lngTempPointer, Asc("["))
        strTemp = BinaryReadToChar(intHandle, lngTempPointer, Asc("]"))
        If strTemp = Description Then
            strError = "'" + Description + "'" + " is already declared in the config file.  This entry will not be added."
            MsgBox strError, vbOKOnly + vbCritical, "Duplicate Entry"
            Exit Sub
        End If
    Loop
    
    
    lngPointer = LOF(intHandle) + 1
    Call BinaryWriteString(intHandle, lngPointer, "[" + Description + "]")
    Call BinaryWriteString(intHandle, lngPointer, Parameter + vbCr + vbLf)
    
    
End Sub


Public Function Find(Description As String) As String

    Dim lngTempPointer As Long
    Dim strTemp As String
    Dim strError As String
    
    
    lngTempPointer = 1
    Do While lngTempPointer < LOF(intHandle)
        strTemp = BinaryReadToChar(intHandle, lngTempPointer, Asc("["))
        strTemp = BinaryReadToChar(intHandle, lngTempPointer, Asc("]"))
        If strTemp = Description Then
            Find = BinaryReadToChar(intHandle, lngTempPointer, Asc(vbCr + vbLf))
            Exit Function
        End If
    Loop
    
    strError = "'" + Description + "' was not found in the config file."
    MsgBox strError, vbCritical + vbOKOnly, "Parameter Not Found"
    
End Function

Private Sub Class_Terminate()

    Close intHandle

End Sub

Public Sub Update(Description As String, Parameter As String)
    
    Delete Description
    Add Description, Parameter
    
End Sub

Public Sub Delete(Description As String)

    Dim lngTempFilePointer As Long
    Dim strTemp As String
    Dim strError As String
    Dim lngPlaceHolder As Long
    Dim intCounter As Integer
    Dim intTempHandle As Integer
    
    
    
    intTempHandle = FreeFile
    Open App.Path + "\temp.cfg" For Binary Access Read Write As intTempHandle
    
    
    lngPointer = 1
    lngTempFilePointer = 1
    intCounter = 1
    
    Do While intCounter < LOF(intHandle)
        strTemp = BinaryReadToChar(intHandle, lngPointer, Asc(vbLf))
        If Len(strTemp) > 0 And InStr(strTemp, Description) = 0 Then
            Call BinaryWriteString(intTempHandle, lngTempFilePointer, strTemp + vbLf)
        End If
        intCounter = intCounter + 1
    Loop
    
    Close intHandle
    Close intTempHandle
    
    FileCopy App.Path + "\temp.cfg", strFilename
    Kill App.Path + "\temp.cfg"
    intHandle = FreeFile
    Open strFilename For Binary Access Read Write As intHandle
    
End Sub
