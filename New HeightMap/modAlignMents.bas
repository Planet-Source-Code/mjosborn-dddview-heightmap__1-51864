Attribute VB_Name = "modAlignMents"
Option Explicit

Public Sub EnableScrolls(f As Form, VA As PictureBox, SI As PictureBox, m As Double)
    Dim vaw As Double
    Dim vah As Double
    Dim sh As Double
    Dim sw As Double
    Dim hsmax As Double
    Dim vsmax As Double
    
    sw = SI.ScaleWidth * m
    sh = SI.ScaleHeight * m
    vaw = VA.ScaleWidth
    vah = VA.ScaleHeight
    
    If sw <= vaw And sh <= vah Then
        f.hsScrScroll.Enabled = False
        f.vsSrcScroll.Enabled = False
        f.cmdCenter.Enabled = False 'enable cmdcenter
    End If
    
    If sw > vaw Then
        f.hsScrScroll.Max = sw - vaw
        f.hsScrScroll.LargeChange = sw - vaw
        f.hsScrScroll.Enabled = True
        f.cmdCenter.Enabled = True 'enable cmdcenter
    End If
    
    If sh > vah Then
        f.vsSrcScroll.Max = sh - vah
        f.vsSrcScroll.LargeChange = sh - vah
        f.vsSrcScroll.Enabled = True
        f.cmdCenter.Enabled = True 'enable cmdcenter
    End If
    
    If sw > vaw And sh < vah Then
        f.hsScrScroll.Max = sw - vaw
        f.hsScrScroll.LargeChange = sw - vaw
        f.hsScrScroll.Enabled = True
        f.vsSrcScroll.Enabled = False
        f.cmdCenter.Enabled = True 'enable cmdcenter
    End If
    
    If sw < vaw And sh > vah Then
        f.vsSrcScroll.Max = sh - vah
        f.vsSrcScroll.LargeChange = sh - vah
        f.vsSrcScroll.Enabled = True
        f.hsScrScroll.Enabled = False
        f.cmdCenter.Enabled = True 'enable cmdcenter
    End If
    
End Sub

Public Sub DisableScrolls(f As Form)
    If f.hsScrScroll.Enabled = True Then
        f.hsScrScroll.Enabled = False 'disable horizontal scroll
        f.vsSrcScroll.Enabled = False 'disable vertical scroll
        f.cmdCenter.Enabled = False 'Hide cmdcenter
    End If
End Sub


Public Sub MoveControls(frm As Form, pArea As PictureBox, pDest As PictureBox, pTop As PictureBox, pLeft As PictureBox, sbStat As StatusBar, hScr As HScrollBar, vScr As VScrollBar, cmdBut As CommandButton)
    If frm.WindowState <> 1 Then 'Window is not minimized then do
    
        pTop.Top = 0
        pTop.Left = pLeft.Width
        
        If frm.ScaleWidth - pLeft.Width > 0 Then
            pTop.Width = frm.ScaleWidth - pLeft.Width
        End If
        
        If frm.ScaleWidth - pLeft.Width - vScr.Width > 0 Then
            pArea.Width = frm.ScaleWidth - pLeft.Width - vScr.Width
        End If
        
        If frm.ScaleHeight - pTop.Height - hScr.Height - sbStat.Height > 0 Then
            pArea.Height = frm.ScaleHeight - pTop.Height - hScr.Height - sbStat.Height
        End If
        
        If frm.ScaleWidth - pLeft.Width - vScr.Width > 0 Then
            hScr.Width = frm.ScaleWidth - pLeft.Width - vScr.Width
        End If
        
        If frm.ScaleHeight - hScr.Height - pTop.Height - sbStat.Height > 0 Then
            vScr.Height = frm.ScaleHeight - hScr.Height - pTop.Height - sbStat.Height
        End If
        
        pArea.Top = pTop.Height
        pArea.Left = pLeft.Width
        
        hScr.Top = frm.ScaleHeight - hScr.Height - sbStat.Height
        hScr.Left = pLeft.Width
        
        vScr.Top = pTop.Height
        vScr.Left = frm.ScaleWidth - vScr.Width
        
        cmdBut.Left = vScr.Left
        cmdBut.Top = frm.ScaleHeight - cmdBut.Height - sbStat.Height
        
    End If
End Sub

Public Sub UpdateScrolls(pArea As PictureBox, pDest As PictureBox, hScr As HScrollBar, vScr As VScrollBar, cmdBut As CommandButton)
'Determine if scroll bars hScr/vScr will be enabled according to pDest w/h
'Determine if cmdBut button will be enabled according to pDest w/h
'Assign hScr/vScr max according to pDest w/h minus pArea w/h

    If pDest.Width <= pArea.Width Then 'Canvass width is <= View area width
        hScr.Enabled = False 'disable horizontal scroll bar
        cmdBut.Enabled = False 'disable cmdCenter button
    Else
        hScr.Enabled = True 'enable horizontal scroll bar
        cmdBut.Enabled = True 'enable cmdCenter button
    End If

    If pDest.Height <= pArea.Height Then 'Canvass height is <= View area height
        vScr.Enabled = False 'disable vertical scroll bar
        cmdBut.Enabled = False 'disable cmdCenter button
    Else
        vScr.Enabled = True 'enable vertical scroll bar
        cmdBut.Enabled = True 'enable cmdCenter button
    End If
    
    vScr.Max = pDest.Height - pArea.Height
    hScr.Max = pDest.Width - pArea.Width
End Sub
