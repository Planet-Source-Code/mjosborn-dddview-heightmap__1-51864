Attribute VB_Name = "modComDialog"
Private CD As CommonDialog
Private strImagePath As String

Option Explicit

Public Function GetImage(strTheFilter As String, strTheInitDir As String, ComD As CommonDialog) As String
    GetImage = Initalize(strTheFilter, strTheInitDir, ComD)
    Call SetPath(GetImage)
End Function

Private Function Initalize(strfilter As String, strInitDir As String, ComD As CommonDialog) As String
On Error GoTo HandleError
    Set CD = ComD
    With CD
        .CancelError = True
        .Filter = strfilter
        .InitDir = strInitDir
        .ShowOpen
        Initalize = .FileName
    End With
HandleError:
    If Err.Number = 32755 Then 'cd cancel was selected
        Err.Clear 'clear cancel error
    End If
End Function

Public Sub SavePic(picSave As PictureBox, ComD As CommonDialog)
On Error GoTo HandleError
    Set CD = ComD
    With CD
        .Flags = cdlOFNOverwritePrompt 'File alread exsists error
        .CancelError = True
        .Filter = "Bitmap (bmp)|*.bmp" 'Set the save filter type of the image
        .ShowSave
        SavePicture picSave.Image, CD.FileName
    End With
HandleError:
    If Err.Number = 32755 Then 'cd cancel was selected
        Err.Clear 'clear cancel error
    End If
End Sub

Public Sub SetPath(strPath As String)
    strImagePath = strPath
End Sub

Public Function GetPath() As String
    GetPath = strImagePath
End Function

