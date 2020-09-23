Attribute VB_Name = "modOpenEditor"

Const SW_RESTORE        As Long = &H9&
Option Explicit

Public Sub OpenEditor(ByVal frm As frmMain, strPath As String)
    Dim lRet    As Long
    lRet = ShellExecute(frm.hWnd, "Open", strPath, &H0&, &H0&, SW_RESTORE)
End Sub

