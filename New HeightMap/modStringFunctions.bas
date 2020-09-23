Attribute VB_Name = "modStringFunctions"
Private strGetExt As String
Private strGetDir As String
Private strGetFileName As String
Private strExtractName As String

Option Explicit

Public Function GetExt(str As String) As String
'Retrieves the exstention of a given file

    Dim strLen As String
    Dim i As Integer
    
    strLen = Len(str)
    Do Until Mid(str, strLen - i, 1) = "."
        i = i + 1
    Loop
    strGetExt = Right(str, i)
    GetExt = Right(str, i)
    
End Function

Public Function GetDir(str As String) As String
'Retrieves the directory path of a given file path

    Dim strLen As String
    Dim i As Integer
    strLen = Len(str)
    Do Until (Mid(str, (strLen - i), 1) = "\")
        i = i + 1
    Loop
    strGetDir = Left(str, (strLen - i) - 1)
    GetDir = Left(str, (strLen - i) - 1)
End Function

Public Function ReturnColorString(str As String, delim As String) As String
'Retrieves the exstention of a given file
    Dim strLen As String
    Dim i As Integer
    strLen = Len(str)
    Do Until (Mid(str, (strLen - i), 1) = delim)
        i = i + 1
    Loop
    ReturnColorString = Left(str, (strLen - i))
End Function

Public Function ReturnColor(str As String, delim As String) As String
'Retrieves the exstention of a given file

    Dim strLen As String
    Dim i As Integer
    i = 1
    strLen = Len(str)
    Do Until Mid(str, i, 1) = delim
        i = i + 1
    Loop
    ReturnColor = Left(str, i - 1)
End Function

Public Function GetFileName(str As String) As String
'Retrieves the file name without extstention from a given file

    Dim strLen As String
    Dim i As Integer
    
    strLen = Len(str)
    Do Until Mid(str, strLen - i, 1) = "."
        i = i + 1
    Loop
    strGetFileName = Left(str, strLen - (i + 1))
    GetFileName = Left(str, strLen - (i + 1))
End Function

Public Function ExtractName(str As String) As String
'Retrieves the file name of a given path
    Dim strTempPath As String
    Dim strLen As String
    Dim i As Integer
    strTempPath = str
    strLen = Len(strTempPath)
    If strTempPath <> "" And strTempPath <> "<None>" Then
    
        Do Until (Mid(strTempPath, (strLen - i), 1) = "\")
            i = i + 1
        Loop
        strExtractName = Right(strTempPath, (i))
        ExtractName = Right(strTempPath, (i))
    End If
    If strLen = "" Then
        strExtractName = "<NoName>"
    End If
    
End Function

Public Function SetExt() As String
    SetExt = strGetExt
End Function
Public Function SetDir() As String
    SetDir = strGetDir
End Function
Public Function SetFileName() As String
    SetFileName = strGetFileName
End Function
Public Function setExtractName() As String
    setExtractName = strExtractName
End Function




