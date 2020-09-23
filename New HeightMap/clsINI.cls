VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Note: I modified this a bit. I added the subs ReadINILine, WriteINILine to
'test if the lines exsist in the inifile.

'**************************************
' Name: clsINI.cls
' Description:Lately I've seen some post
'     s for editing ini files that involve ope
'     ning the ini file directly as a text fil
'     e, looping line by line thru the file un
'     til locating the line desired and then a
'     ltering that line. There is a much easie
'     r and more reliable way using the Window
'     s API. This class module makes that easy
'     . It also shows the proper way to handle
'     errors that happen in a class module by
'     raising custom error codes to be handled
'     by the application that using using the
'     class.
' By: Bryan Johns
'
'
' Inputs:See the comments in the code.
'
' Returns:See the comments in the code.
'
'Assumes:It assumes a basic familiarity


'     with how to use class modules.
'
'Side Effects:None
'This code is copyrighted and has limite
'     d warranties.
'Please see http://www.Planet-Source-Cod
'     e.com/xq/ASP/txtCodeId.33363/lngWId.1/qx
'     /vb/scripts/ShowCode.htm
'for details.
'**************************************

Option Explicit

' API Declarations
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

' Local variables to hold property value
Private mstrINIPath As String
Private mstrFileName As String
Private mstrWindowsPath As String
Private mboolVal As Boolean

Public Property Get WindowsPath() As String
    WindowsPath = mstrWindowsPath
End Property


Private Property Let WindowsPath(ByVal strWindowsPath As String)
    mstrWindowsPath = strWindowsPath
End Property
'***************************
'* Procedure: WindowsPathGet
'* Copyright: (C) 2002, Bryan Johns
'* Purpose : Uses an API call to set the
'     read only WindowsPath property.
'****************************

Private Sub WindowsPathGet()
    Dim Y As String
    On Error GoTo error
    mstrWindowsPath = Space(255)
    Y = GetWindowsDirectory(mstrWindowsPath, 255)
    mstrWindowsPath = Left$(mstrWindowsPath, Y)
    Exit Sub
error:
    Err.Raise 10001, "clsINI.cls", "Unable to read the windows path."
End Sub

Public Property Get FileName() As String
    FileName = mstrFileName
End Property

Public Property Let FileName(ByVal strFileName As String)
    mstrFileName = strFileName
End Property

'***************************
'* Procedure: WriteINI
'* Copyright: (C) 2002, Bryan Johns
'* Purpose : Exposes the private WriteTo sub.
'****************************

Public Sub WriteINI(Section As String, Field As String, Value As String)
    WriteTo Section, Field, Value
End Sub

'***************************
'* Function : ReadINI
'* Copyright: (C) 2002, Bryan Johns
'* Purpose : Exposes the Private ReadFro
'     m function.
'***************************

Public Function ReadINI(Section As String, Field As String) As String
    ReadINI = ReadFrom(Section, Field)
End Function

Public Property Get INIPath() As String
    INIPath = mstrINIPath
End Property

Public Property Let INIPath(ByVal strINIPath As String)
    mstrINIPath = strINIPath
End Property

'***************************
'* Function : ReadFrom
'* Copyright: (C) 2002, Bryan Johns
'* Purpose : Returns values read from th
'     e INI file.
'***************************

Private Function ReadFrom(lstrSection As String, lstrField As String) As String
    Dim varReturnedString As Integer
    Dim lstrResults As String
    'IniReadErr = False
    lstrResults = Space(255)
    varReturnedString = GetPrivateProfileString&(lstrSection, lstrField, "", lstrResults, 255, mstrINIPath & "\" & mstrFileName)
    lstrResults = Left$(lstrResults, varReturnedString)

    If Len(lstrResults) < 1 Then
        mboolVal = False
        'Err.Raise 10000, "ReadFrom()", "Unable to read ini file entry."
        Err.Clear
        Exit Function
    Else
        mboolVal = True
    End If
    
    ReadFrom = lstrResults
End Function
'***************************
'* Procedure: WriteTo
'* Copyright: (C) 2002, Bryan Johns
'* Purpose : Writes values to the INI fi
'     le.
'****************************

Private Sub WriteTo(lstrSection As String, lstrField As String, lstrDefaultValue As String)
    Dim X As Boolean
    X = WritePrivateProfileString&(lstrSection, lstrField, lstrDefaultValue, mstrINIPath & "\" & mstrFileName)
    

    If X = False Then
        Err.Raise 10002, "WriteTo()", "There was a critical error writing to the" & mstrFileName & " file."
        Err.Clear
        Exit Sub
    Else
    
    End If
End Sub

Private Sub Class_Initialize()
    ' get the windows path and assign it to
    '     the INIPath property so that if the user
    '     of this
    ' class module doesn't supply a path it'
    '     s defaulted to the windows path.
  
    mboolVal = False
    WindowsPathGet
    mstrINIPath = mstrWindowsPath
End Sub

Public Property Get ReadINILine() As Boolean
    ReadINILine = mboolVal
End Property

Private Property Let ReadINILine(ByVal val As Boolean)
    mboolVal = val
End Property







