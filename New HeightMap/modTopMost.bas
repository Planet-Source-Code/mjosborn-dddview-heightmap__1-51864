Attribute VB_Name = "modTopMost"
' ------------------------------------------------------------------------
'
'                       BulletSoft Solutions
'
'  You have a royalty-free right to use, modify, reproduce and distribute
'  this file (and/or any modified version) in any way you find useful,
'  provided that you agree that BulletSoft Solutions has no
'  warranty, obligation or liability for its contents.
'  Refer to the http://www.bulletsoftsolutions.com for more project
'  examples like this one.
'
' ------------------------------------------------------------------------
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_FLAGS = SWP_NOMOVE + SWP_NOSIZE

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Public Sub SetTopMost(frmTopMost As Form, iIndex As Integer)

    Dim l As Long
    Dim rc As Long
    
    'Option iIndex 0 = TopMost translates to HWND_TOPMOST
    'Option iIndex 1 = TopMost translates to HWND_NOTOPMOST
    
    l = (iIndex + 1) * -1
    
    rc = SetWindowPos(frmTopMost.hWnd, l, 0&, 0&, 0&, 0&, SWP_FLAGS)
    
    
End Sub


