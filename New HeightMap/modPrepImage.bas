Attribute VB_Name = "modPrepImage"
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Const HALFTONE = 4
Private Type COLORADJUSTMENT
        caSize As Integer
        caFlags As Integer
        caIlluminantIndex As Integer
        caRedGamma As Integer
        caGreenGamma As Integer
        caBlueGamma As Integer
        caReferenceBlack As Integer
        caReferenceWhite As Integer
        caContrast As Integer
        caBrightness As Integer
        caColorfulness As Integer
        caRedGreenTint As Integer
End Type
Public Const ScrCopy = &HCC0020

Private hsmax As Double 'Hor width differance
Private vsmax As Double 'Ver height differance
Private strImageWidth As String
Private strImageHeight As String
Private strImageScaleW As String
Private strImageScaleH As String
Private strSWP As String
Private strSHP As String

Option Explicit

Public Sub PrepImage(f As Form, m As Double, picSource As PictureBox, picViewAreaWindow As PictureBox, picDisplay As PictureBox)
    Dim dblSW As Double 'source scale width
    Dim dblSH As Double 'source scale height
    Dim dblFW As Double 'form scale width
    Dim dblFH As Double 'form scale height
    Dim dblVAW As Double 'view area scale width
    Dim dblVAH As Double 'view area scale height
    Call f.cmdCenter_Click
    'Set f = f
    dblSW = picSource.ScaleWidth * m 'Assign Source Width
    dblSH = picSource.ScaleHeight * m 'Assign Source height
    
    If f.optImageStyle(0).Value = True Then 'determine how to display the image
    'display the image scaled to view area
        'Call AlignVA(f, f.picViewAreaWindow, f.sb) 'Realign the view area to the form, hide scrolls
        Call DisableScrolls(f) 'Hide hor/vert scrolls
        dblFW = f.Width 'Assign new form width
        dblFH = f.Height 'Assign new form height
        dblVAW = picViewAreaWindow.ScaleWidth 'Assign new view area width
        dblVAH = picViewAreaWindow.ScaleHeight 'Assign new view area height
        Call BestFit(dblSW, dblSH, dblVAW, dblVAH, picSource, picDisplay)   'Preform calcs to best fit image
    End If
    
    If f.optImageStyle(1).Value = True Then 'determine how to display the image
    'display the image normaly
        Call EnableScrolls(f, picViewAreaWindow, picSource, m)
        dblFW = f.Width
        dblFH = f.Height
        dblVAW = picViewAreaWindow.ScaleWidth
        dblVAH = picViewAreaWindow.ScaleHeight
        Call Normal(dblSW, dblSH, dblVAW, dblVAH, picSource, picDisplay, f)
    End If
    
    If f.optImageStyle(2).Value = True Then
    'diplay the image scaled to width of view area
        'Call AlignVA(f, f.picViewAreaWindow, f.sb) 'Realign the view area to the form, hide scrolls
        Call DisableScrolls(f) 'Hide hor/vert scrolls
        dblFW = f.Width
        dblFH = f.Height
        dblVAW = picViewAreaWindow.ScaleWidth
        dblVAH = picViewAreaWindow.ScaleHeight
        Call FitToWidth(dblSW, dblSH, dblVAW, dblVAH, picSource, picDisplay)
    End If
    
End Sub
Private Sub BestFit(sw As Double, sh As Double, vaw As Double, vah As Double, picSource As PictureBox, picDisplay As PictureBox)
    Dim dblSR As Double 'Source ratio
    Dim dblVAR As Double 'View area ratio
    Dim draww As Double
    Dim drawh As Double
    
    dblSR = (sw / sh)
    dblVAR = (vaw / vah)
   
    
    If (dblVAR > dblSR) Then 'Test if source ratio is less than view area ratio
    'True: source is smaller than view area. Assign true to bool val
    'Calculate the new ratio width, by multiply the source width by
    'viewing area ratio.
        draww = (vah * dblSR)  'calulate new width, then assign
        drawh = vah
    Else
    'false: source is larger than view area. Assign false to bool val
    'Calculate the new ratio heigth, by divideing the source height by
    'viewing area ratio.
        drawh = (vaw / dblSR) 'calulate new height, then assign
        draww = vaw
    End If
    
     Call RepaintImage(draww, drawh, picSource, picDisplay)
End Sub

Private Sub FitToWidth(sw As Double, sh As Double, vaw As Double, vah As Double, picSource As PictureBox, picDisplay As PictureBox)
'FitImage to width of view area
    Dim dblSRW As Double
    Dim dblSRH As Double
    Dim dblVARW As Double
    Dim dblVARH As Double
    Dim draww As Double
    Dim drawh As Double
    
    dblSRW = (sw / vaw) 'Percent Source width is to view area width
    dblVARW = (vaw / sw) 'Percent view area width is to source width
    dblSRH = (sh / vah) 'percent source height is to view area height
    dblVARH = (vah / sh) 'percent view area height is to source height
    
    'If dblSRW > dblVARW Then 'Souce Width is > than view area width
    '    draww = sw * dblVARW
    'End If
    'If dblSRW <= dblVARW Then
        draww = sw * dblVARW
   ' End If
   ' If dblSRH > dblVARH Then
        drawh = sh * dblVARH
    'End If
    'If dblSRH <= dblVARH Then
    '    drawh = sh * dblVARH
    'End If
    
    Call RepaintImage(draww, drawh, picSource, picDisplay)
    
End Sub

Private Sub Normal(sw As Double, sh As Double, vaw As Double, vah As Double, picSource As PictureBox, picDisplay As PictureBox, f As Form)
    If vaw >= sw And vah >= sh Then
        Call DisableScrolls(f)
        Call RepaintImage(sw, sh, picSource, picDisplay)
        Exit Sub
    Else
        Call RepaintImage(sw, sh, picSource, picDisplay)
    End If
End Sub

Private Sub RepaintImage(w As Double, h As Double, picSource As PictureBox, picDisplay As PictureBox)
'Repaint the image to the picture box picViewedImage to be viewed
    picDisplay.Cls 'Clear old image
    picDisplay.Width = w 'Set Width of image picture box
    picDisplay.Height = h 'Set height of image picture box
    'Repaint the image
    SetStretchBltMode picDisplay.hdc, HALFTONE
    StretchBlt picDisplay.hdc, 0, 0, w, h, picSource.hdc, 0, 0, picSource.Width, picSource.Height, ScrCopy
    'picDisplay.PaintPicture picSource.Picture, 0, 0, w, h
    picDisplay.Visible = True
    strImageWidth = picSource.Width
    strImageHeight = picSource.Height
    strImageScaleW = w
    strImageScaleH = h
    strSWP = (w / strImageWidth) * 100
    strSHP = (h / strImageHeight) * 100
End Sub

Public Sub SampleImage(ByRef pSource As PictureBox, pDisplay As PictureBox)
    Dim origRatio As Double
    Dim statRatio As Double
    Dim sw As Double
    Dim sh As Double
    Dim w As Integer
    Dim h As Integer
    Dim SampleW As Double
    Dim SampleH As Double
    Dim SampleRatio As Double
    
    pDisplay.Cls
    
    'get original size
    sw = pSource.ScaleWidth 'Soucre width
    sh = pSource.ScaleHeight 'Source height

    
    SampleW = pDisplay.ScaleWidth
    SampleH = pDisplay.ScaleHeight
    
    origRatio = sw / sh
    'scale width & Height of destination pix
    
   
    SampleRatio = SampleW / SampleH

    
    If (SampleW / SampleH) > origRatio Then
        SampleW = origRatio * SampleH
    Else
        SampleH = SampleW / origRatio
    End If
    w = ((pDisplay.ScaleWidth - SampleW) / 2)
    h = ((pDisplay.ScaleHeight - SampleH) / 2)
    'pDisplay.PaintPicture pSource.Image, w, h, SampleW, SampleH
    SetStretchBltMode pDisplay.hdc, HALFTONE
    StretchBlt pDisplay.hdc, w, h, SampleW, SampleH, pSource.hdc, 0, 0, sw, sh, ScrCopy
End Sub
Public Function GetHSMax() As Double
    GetHSMax = hsmax
End Function

Public Function GetVSMax() As Double
    GetVSMax = vsmax
End Function

Public Function GetImageWidth() As String
    GetImageWidth = strImageWidth
End Function
Public Function GetImageHeight() As String
    GetImageHeight = strImageHeight
End Function
Public Function GetImageScaleW() As String
    GetImageScaleW = strImageScaleW
End Function
Public Function GetImageScaleH() As String
    GetImageScaleH = strImageScaleH
End Function
Public Function GetSWP() As String
    GetSWP = strSWP
End Function
Public Function GetSHP() As String
    GetSHP = strSHP
End Function



