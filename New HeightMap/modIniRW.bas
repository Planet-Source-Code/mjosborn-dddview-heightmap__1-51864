Attribute VB_Name = "modIniRW"
Public INI As New clsINI 'Ini read/write class
Dim answer As Boolean 'Answer, does directory exist
Dim HomeDir As String
Dim INIDir As String
Dim ININame As String
Dim ScreenShots As String

Option Explicit

Public Function IsDirExist(path As String) As Boolean
'Check if directory with given path exists
    IsDirExist = Dir$(path, vbDirectory) <> ""
End Function

Public Sub CheckDirectories()
    
    ININame = "Paths.ini"
    answer = False 'Initialize the answer/nothing has been checked yet
    HomeDir = App.path  'Set home directory
    INIDir = HomeDir & "\" & ININame 'Set ini directory
    answer = IsDirExist(INIDir) 'Determine if the INI directory exists
    
    If answer = False Then
        INI.INIPath = HomeDir
        INI.FileName = ININame
        
        INI.WriteINI "Notes", "Important", "Please Do Not Modify This INI File"
        INI.WriteINI "Created", "Date", FormatDateTime(Now, vbLongDate)
        INI.WriteINI "Paths", "CDPath", HomeDir
    End If
    
    If answer = True Then
        INI.INIPath = HomeDir
        INI.FileName = ININame
    End If
    
End Sub

Private Sub MakeDir(path As String)
'Make the given directory
    MkDir (path)
End Sub

Public Function INIExist() As Boolean
'Determine if ini file exist. If does assign its path and name
'If it does not, the create the file

    Dim answer As Boolean 'Answer, does directory exist
    answer = False 'initialize answer to false
    answer = IsDirExist(INIDir) 'Determine if the INI directory exists
    
    If answer = False Then 'Ini does not exist
        CheckDirectories 'call checkdirectories to create the ini file
    End If
    
    If answer = True Then 'Ini file does exist
        INIExist = answer 'set INIExist to true
    End If

End Function

Public Function ReadTheINI(strSection As String, strField As String, strValue As String) As String
'Read information that has been written to the ini file
    Dim INITest As Boolean 'Test: Does ini file exist
    Dim TestLine As Boolean 'Test: Does line to read exist
    
    INITest = False 'Initalize the test to false
    
    INITest = INIExist 'Test if ini exsists before we read it
    
    If INITest = True Then
        INI.INIPath = HomeDir
        INI.FileName = ININame
    End If
    
    ReadTheINI = INI.ReadINI(strSection, strField) 'Line to read, returns the value
    INITest = INI.ReadINILine
    'Test: Does line to read from exist
    
    If INITest = False Then 'Line Does not exsist, create it
        INI.WriteINI strSection, strField, strValue 'Create the line
        ReadTheINI = INI.ReadINI(strSection, strField) 'Return the line value
    End If
    
    
End Function
Public Sub WriteToINI(strSection As String, strField As String, strValue As String)
'Write information to the ini file
    Dim INITest As Boolean 'Does ini file exist
    
    INITest = IsDirExist(INIDir) 'Test if ini exsists before we read it
    If INITest = True Then
        INI.INIPath = HomeDir
        INI.FileName = ININame
        INI.WriteINI strSection, strField, strValue
    End If
    
    If INITest = False Then 'Line Does not exsist, create it
        Call CheckDirectories 'create the ini file
        'Call frmMainViewer.ProgramStart 'Restore ini to program start
        INI.WriteINI strSection, strField, strValue 'Create the line
    End If
    
    
End Sub


