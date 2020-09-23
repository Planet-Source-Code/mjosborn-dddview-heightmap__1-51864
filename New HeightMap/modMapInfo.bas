Attribute VB_Name = "modMapInfo"
'Structure Delcares
Private biRct As RECT 'rectangle structure


'Declares
Private lngBlue As Long 'extracted blue value
Private lngGreen As Long 'extracted green value
Private lngRed As Long 'extracted red value
Private intH As Integer 'desired height
Private intlx As Integer 'line start x point
Private intly As Integer 'line start y point
Private intShearX As Integer 'Shear X value
Private intShearY As Integer 'Shear y value
Private intShear As Integer 'Shearing incrementing value
Private lngHigh As Long 'Highest point of the picture
Private lngLow As Long 'Lowest point of the picture
Private lngTmpH As Long 'Temporary high point storage
Private lngTmpL As Long 'Temporary low point storage
Private lngPixel As Long 'Extracted lngPixel from source
Private lngHO As Long 'Calculated height offset to move y pixel
Private lngCArray(0 To 2) As Long 'Color array Note may not use
Public lngRA() As Long
Public lngGA() As Long
Public lngBA() As Long
Public lngColorA() As Long
Private dblCSX As Double 'center point x of source image
Private dblCSY As Double 'center point y of source image
Private dblCDX As Double 'center point x of destination image
Private dblCDY As Double 'center point x of destination image
Private dblSinA As Double 'x angle point
Private dblCosA As Double 'y angle point
Private dblGXI As Double 'Global x index for 2d array
Private dblGYI As Double 'Global y index for 2d array
Public strXYZ() As String
Dim blnContinue As Boolean 'Test: Do we continue to draw
Public intStep As Integer

'Object Declares
Private picDest As PictureBox 'Global destination picture box
Private picSrc As PictureBox 'Global source picture box

'Constants
Private Const Deg2Rad As Double = 0.017453292519943 'PI/180


Option Explicit

Public Sub InitialiseHM(ByRef pSource As PictureBox, pCanvass As PictureBox, ByRef angle As Double, ByRef intDM As Integer)

    Set picDest = pCanvass 'set the global destination picture box
    Set picSrc = pSource 'set the global source picture box
    intStep = 1
    intShear = -picDest.Width / 2 'Center picture if sheared
    intShearX = 0 'Reset shearing x point
    intShearY = 0 'Reset shearing y point
    lngHigh = 0 'Reset high point
    lngLow = 0 'Reset low point
    lngTmpH = 0 'Reset temp hight point
    lngTmpL = 0 'Reset temp low point
    dblGXI = 0 'Reset global index
    dblGYI = 0 'Reset global index
    
    dblCSX = CLng(picSrc.Width) * 0.5  'get center point x of source image
    dblCSY = CLng(picSrc.Height) * 0.5  'get center point y of source image
    dblCDX = CLng(picDest.Width) * 0.5 'get center point x of destination image
    dblCDY = CLng(picDest.Height) * 0.5 'get center point x of destination image
    dblCosA = Cos(angle * Deg2Rad * -1) 'Convert x angle
    dblSinA = Sin(angle * Deg2Rad * -1) 'convert y angle
    
    ReDim strXYZ(0 To (picSrc.Width), 0 To (picSrc.Height))
    ReDim lngRA(0 To (picSrc.Width), 0 To picSrc.Height)
    ReDim lngGA(0 To (picSrc.Width), 0 To picSrc.Height)
    ReDim lngBA(0 To (picSrc.Width), 0 To picSrc.Height)
    ReDim lngColorA(0 To (picSrc.Width), 0 To picSrc.Height)
    
    'Get bounds of source picture
    SetRect biRct, 0, 0, CLng(picSrc.Width), CLng(picSrc.Height)
    
    Call CyclePoints(intDM) 'cycle the points of the picture
End Sub

Private Sub CleanUp()
    Set picDest = Nothing 'clear picture box from memory
    Set picSrc = Nothing 'clear picture box from memory
End Sub

Public Sub CyclePoints(ByVal intDM As Integer)
    Dim intX As Integer 'X image cycle
    Dim intY As Integer 'Y image cycle
    Dim dblDestX As Double 'Starting point x
    Dim dblDestY As Double 'Starting point y
    Dim intRotX As Integer 'Rotated destination X point
    Dim intRotY As Integer 'Rotated destination y point
    
    For intY = 0 To CLng(picDest.Height) - 1 Step intStep  'Cycle y's
        DoEvents
        If GetDrawState = False Then Exit Sub
        dblDestY = (intY - dblCDY)  'calculate y destination
        For intX = 0 To CLng(picDest.Width) - 1 Step intStep 'cycle x's
            dblDestX = (intX - dblCDX)  'calculate x destination
            'Rotate destination X and Y according to angle and
            'round off rotated lngPixel at X, Y coordinates
            intRotX = Int((dblDestX) * dblCosA - (dblDestY) * dblSinA + dblCSX)
            intRotY = Int((dblDestX) * dblSinA + (dblDestY) * dblCosA + dblCSY)
            If (PtInRect(biRct, intRotX, intRotY)) Then   'if x,y is in destination area then draw
                
                Call PixelInfo(intRotX, intRotY) 'Get pixel information
                Call Shearing 'Determine camera view
                Call DrawHeightMap(intX - (-1 * intShearX), intY - intShearY, intDM) 'draw heightmap
            End If
        Next intX
        intShear = intShear + 1 'increment new camera view each cycle of y
        intlx = 0 'new line needs to start, x's have cycled
        intly = 0 'new line needs to start, x's have cycled
        

        Call UpdateDrawInfo(intX, intY) 'Display percent completed..etc
    Next intY

    Call CleanUp 'clean up memory
End Sub

Private Sub PixelInfo(ByRef xRot As Integer, ByRef yRot As Integer)
    lngPixel = GetPixel(picSrc.hdc, xRot, yRot)   'get lngPixel from source
    Call BreakUpRGB(lngPixel)  'break up it's red, green, blue values
    lngHO = GetHO 'get calculated height offset
    lngPixel = OutPutColors(lngPixel) 'get color of pixel to draw
    
    Call HighLowPoints 'Assign highest and lowest points
End Sub

Private Sub Shearing()
'Assing variables to angle the camera left,right, top, side
    Select Case frmMain.cboRotOptions.ListIndex
        Case 0 'Aerial
            intShearX = 0
            intShearY = intShearX
        Case 1 'Left
            intShearX = -intShear
            intShearY = 0
        Case 2 'Right
            intShearX = intShear
            intShearY = 0
        Case 3 'Profile
            intShearX = intShear
            intShearY = intShearX
    End Select
End Sub

Private Sub HighLowPoints()
    lngTmpH = lngHO 'temp point
    lngTmpL = lngHO 'temp low point
        
    If lngTmpH >= lngHigh Then 'if temp high is higher that high point
        lngHigh = lngTmpH 'assign highest point
    End If
        
    If lngTmpL <= lngLow Then 'if temp low is lower that low point
        lngLow = lngTmpL 'assign lowest point
    End If
End Sub

Private Sub DrawHeightMap(ByRef xPoint As Integer, yPoint As Integer, intDrawMode As Integer)
   Dim bl As Integer
   Dim br As Integer
   Dim lPoint As POINTAPI
   Dim poly(1 To 3) As COORD
   Dim NumCoords As Long ', hBrush As Long, hRgn As Long
   
    'SetPixelV picDest.hdc, xPoint, yPoint, vbBlack  'Set bottom color
        Select Case intDrawMode
            Case 0 'Dots
                SetPixelV picDest.hdc, xPoint, (yPoint - lngHO), lngPixel   'draw dot
            Case 1 'Draw Lines
                If intlx = 0 Then 'assign line start x,y a starting point
                    intlx = xPoint 'starting point is x
                    intly = yPoint - lngHO 'starting point is y minus height offset
                End If
                picDest.ForeColor = lngPixel
                lPoint.x = CLng(intlx): lPoint.y = CLng(intly)
                MoveToEx picDest.hdc, intlx, intly, lPoint
                LineTo picDest.hdc, xPoint, yPoint - lngHO
                intlx = xPoint 'next start point x was last starting point x, keep the line going
                intly = yPoint - lngHO 'next start point y was last starting point y, keep the line going
                'picDest.Line (intlx, intly )-(xPoint, yPoint - lngHO), lngPixel 'Draw Line
            Case 2 'Draw Bars
                picDest.ForeColor = lngPixel
                lPoint.x = xPoint: lPoint.y = yPoint
                MoveToEx picDest.hdc, xPoint, yPoint, lPoint
                LineTo picDest.hdc, xPoint, yPoint - lngHO
                'picDest.Line (xPoint, yPoint)-(xPoint, yPoint - lngHO), lngPixel 'draw Spike line
            Case 3 'Draw Circles
                picDest.ForeColor = lngPixel
                Ellipse picDest.hdc, xPoint - 1, (yPoint - lngHO) - 1, xPoint + 1, (yPoint - lngHO) + 1
                'picDest.Circle (xPoint, yPoint - lngHO), 1, lngPixel
            Case 4 'Triangle
                picDest.ForeColor = lngPixel
                NumCoords = 3
                poly(1).x = xPoint: poly(1).y = yPoint
                poly(2).x = xPoint + 1: poly(2).y = yPoint - lngHO
                poly(3).x = xPoint + 2: poly(3).y = yPoint
                Polygon picDest.hdc, poly(1), NumCoords
                'picDest.Line (xPoint, yPoint)-(xPoint, yPoint - lngHO), lngPixel
                'picDest.Line (xPoint, yPoint - lngHO)-(xPoint - 1, yPoint), lngPixel
                'picDest.Line (xPoint - 1, yPoint)-((xPoint - 1) + 2, yPoint), lngPixel
            Case 5 'cross
                picDest.Line (xPoint, yPoint)-(xPoint, yPoint - lngHO), lngPixel
                picDest.Line (xPoint, yPoint)-(xPoint, yPoint + 1), lngPixel
                picDest.Line (xPoint, yPoint)-(xPoint - 1, yPoint), lngPixel
                picDest.Line (xPoint, yPoint)-(xPoint + 1, yPoint), lngPixel
        End Select
        'Determine if color map will be used
        If frmMain.chkColorMap.Value = 1 Then Call GetColorMap(xPoint, yPoint, lngHO)
    
End Sub

Private Sub GetColorMap(x As Integer, y As Integer, z As Long)

    If dblGXI > (CInt(picSrc.Width) - 1) / intStep Then 'x index is greater than x max
        dblGXI = 0 'reset x to begining of next x row
        If (dblGYI + 1) < picSrc.Height / intStep Then
            dblGYI = dblGYI + 1 'increment y to next colum
        End If
    End If
    
    strXYZ(dblGXI, dblGYI) = x & "," & y & "," & z
    lngRA(dblGXI, dblGYI) = lngCArray(0)
    lngGA(dblGXI, dblGYI) = lngCArray(1)
    lngBA(dblGXI, dblGYI) = lngCArray(2)
    lngColorA(dblGXI, dblGYI) = lngPixel
    dblGXI = dblGXI + 1 'increment global x index value

End Sub

Private Sub BreakUpRGB(ByRef Color As Long)
'Get Color Values from lngPixel
'Formula of color: Color = Red + (Green * 256) + (Blue * 256 ^ 2)
    lngBlue = (Color And &HFF0000) / 65536 'Extract blue
    lngGreen = ((Color And &HFF00) / 256&) Mod 256& 'Extract green
    lngRed = Color Mod 256& 'Extract Red
    
    
    'If frmMain.cboColorStyle.ListIndex = 2 Then
    '        lngC = (lngRed + lngGreen + lngBlue) / 3
    '        lngCArray(0) = lngC 'Assign red color to array
    '        lngCArray(1) = lngC 'Assign green color to array
    '        lngCArray(2) = lngC 'Assign blue color to array
    'Else
            lngCArray(0) = lngRed 'Assign red color to array
            lngCArray(1) = lngGreen 'Assign green color to array
            lngCArray(2) = lngBlue 'Assign blue color to array
    'End If
End Sub

Private Function GetHO() As Long
    Dim lngHO As Long 'height offset
    Dim intI As Integer 'counting index

    For intI = 0 To frmMain.lstRGBH.ListCount - 2 'cycle r,g,b height
        If frmMain.lstRGBH.Selected(intI) = True Then 'if item is selected
            lngCArray(intI) = 0 'set it's height to zero
        End If
    Next intI
    
    'if you leave out any one color then you'll see a hue, of the other two
    lngHO = (lngCArray(0) + lngCArray(1) + lngCArray(2)) / 3 'Average of colors 'average of all 3 colors
    lngHO = lngHO / GetHeight 'determine height offset at desired height value "5" change 5 to change height offset
    GetHO = lngHO 'assign height offset
    
    If frmMain.chkInvert.Value = 1 Then GetHO = -lngHO 'invert the image
    

End Function

Public Sub SetHeight(ByVal h As Integer)
    intH = h 'set height offset
End Sub

Private Function GetHeight() As Integer
    GetHeight = 255 / intH 'assign height offset
End Function

Private Function FadeColor(ByRef ho As Long) As Long
    Dim lngColors(0 To 2) As Long
    Dim intI As Integer
    
    For intI = 0 To UBound(lngColors) 'cycle r,g,b, shaded
        If frmMain.chkRGB(intI).Value = 1 Then lngColors(intI) = ho * val(frmMain.txtHSensitivity.Text)
        If frmMain.chkRGB(intI).Value = 0 Then lngColors(intI) = 0 'lngCArray(intI)
    Next intI
    
    FadeColor = RGB(lngColors(0), lngColors(1), lngColors(2)) 'return color
End Function

Private Function OutPutColors(ByRef pix As Long) As Long
'Define the color to draw the image.
'optOpc(1)=custom, optOpc(0)=define, otherwise use original colors
    Dim lngC As Long
    'Draw image at custom color
    lngC = (lngRed + lngGreen + lngBlue) / 3 'Average of colors
    
    If frmMain.optOPC(0).Value = True Then OutPutColors = FadeColor(GetHO) 'Fadecolors
    If frmMain.optOPC(1).Value = True Then OutPutColors = frmMain.picDefineColor.BackColor 'User defined color
    If frmMain.optOPC(2).Value = True Then 'Default Colors

        Select Case frmMain.cboColorStyle.ListIndex
            Case 0 'Defualt colors
                OutPutColors = pix 'default color
            Case 1 'Grey Scale
                OutPutColors = RGB(lngC, lngC, lngC) 'Assign average color
            Case 2 To 3 'Color Height
                OutPutColors = ColorHeight(lngC)
        End Select
    End If
End Function

Private Function ColorHeight(ByVal lngColor As Long) As Long
'Divide 255/3 to get color range, ranges are 85,170,255
'Multiply by the number of times the sprecific color range goes into 255.
'Ex: 85 * 3 = 255, 170 * 1.5 = 255, 255 * 1 = 255, just to increase color intensity,
'according to value, so there is a smoother transition to color change.
    Dim lngRGB As Long
    Select Case frmMain.cboColorStyle.ListIndex
    Case 2
        If lngHO < 0 Then 'invert
            Select Case lngColor
                Case 0 To 85
                    lngRGB = RGB(0, 0, lngColor)
                Case 86 To 170
                    lngRGB = RGB(0, lngColor * 1.5, 0)
                Case 171 To 255
                    lngRGB = RGB(lngColor * 3, 0, 0)
            End Select
        Else 'not inverted
        Select Case lngColor
            Case 0 To 85
                lngRGB = RGB(lngColor * 3, 0, 0)
            Case 86 To 170
                lngRGB = RGB(0, lngColor * 1.5, 0)
            Case 171 To 255
                lngRGB = RGB(0, 0, lngColor)
        End Select
        End If
    Case 3
        If lngHO < 0 Then 'invert
            Select Case lngColor
            Case 0 To 42
                lngRGB = RGB(0, 0, lngColor)
            Case 43 To 85
                lngRGB = RGB(128, 128, lngColor * 1.2)
            Case 86 To 127
                lngRGB = RGB(0, lngColor, 0)
            Case 128 To 169
                lngRGB = RGB(128, lngColor * 2, 128)
            Case 170 To 211
                lngRGB = RGB(lngColor, 0, 0)
            Case 212 To 255
                lngRGB = RGB(lngColor * 6, lngColor * 6, lngColor * 6)
        End Select
        Else 'not inverted
            Select Case lngColor
                Case 0 To 42
                    lngRGB = RGB(lngColor * 6, lngColor * 6, lngColor * 6)
                Case 43 To 85
                    lngRGB = RGB(lngColor, 0, 0)
                Case 86 To 127
                    lngRGB = RGB(128, lngColor * 2, 128)
                Case 128 To 169
                    lngRGB = RGB(0, lngColor, 0)
                Case 170 To 211
                    lngRGB = RGB(128, 128, lngColor * 1.2)
                Case 212 To 255
                    lngRGB = RGB(0, 0, lngColor)
            End Select
        End If
    End Select
    
    'Call BreakUpRGB(lngRGB)
    ColorHeight = lngRGB 'return color calculated to height of point
End Function

Public Function SetDrawState(ByVal blnDrawState As Boolean)
    blnContinue = blnDrawState 'Test: Do we continue to draw
End Function

Private Function GetDrawState() As Boolean
    GetDrawState = blnContinue 'Assign draw state, stop or start
End Function

Private Sub UpdateDrawInfo(ByVal x As Integer, ByVal y As Integer)
'Display percent of image completed. As well as show the image being drawn
        frmMain.lblPixelCycles.Caption = "(" & x & "," & y + 1 & ")"
        frmMain.lblPixelCycles.Refresh
        frmMain.pg.Value = y + 1
        frmMain.lblPerc.Caption = CInt(((frmMain.pg.Value / frmMain.pg.Max) * 100)) & "%" 'percent completed
        frmMain.lblPerc.Refresh 'refresh label to view it
        If frmMain.chkShowDraw.Value = 1 Then frmMain.picHMDest.Refresh 'redraw
        frmMain.lblHP.Caption = lngHigh
        frmMain.lblLP.Caption = lngLow
End Sub

