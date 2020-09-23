Attribute VB_Name = "modRotatePIExample"

'Constants
Private Const Deg2Rad As Double = 0.017453292519943 'PI/180
Private biRct As RECT

Option Explicit

Public Sub RotatePIPicture(ByRef pSource As PictureBox, ByRef pCanvass As PictureBox, ByRef angle As Double)
   Dim intX As Integer 'X image cycle
    Dim intY As Integer 'Y image cycle
    Dim dblDestX As Double 'Starting point x
    Dim dblDestY As Double 'Starting point y
    Dim lngPixel As Long 'Extracted lngPixel from source
    Dim dblCSX As Double 'center point x of source image
    Dim dblCSY As Double 'center point y of source image
    Dim dblCDX As Double 'center point x of destination image
    Dim dblCDY As Double 'center point x of destination image
    Dim dblSinA As Double 'x angle point
    Dim dblCosA As Double 'y angle point
    Dim intRotX As Integer 'Rotated destination X point
    Dim intRotY As Integer 'Rotated destination y point

    pCanvass.Cls 'clear old image
    dblCSX = CLng(pSource.Width) * 0.5  'get center point x of source image
    dblCSY = CLng(pSource.Height) * 0.5  'get center point y of source image
    dblCDX = CLng(pCanvass.Width) * 0.5 'get center point x of destination image
    dblCDY = CLng(pCanvass.Height) * 0.5 'get center point x of destination image
    dblCosA = Cos(angle * Deg2Rad * -1) 'Convert x angle
    dblSinA = Sin(angle * Deg2Rad * -1) 'convert y angle
    
    'Get bounds of source picture
    SetRect biRct, 0, 0, CLng(pSource.Width - 1), CLng(pSource.Height - 1)
    For intY = 0 To CLng(pCanvass.Height) - 1 'Cycle y's
        dblDestY = intY - dblCDY 'calculate y destination
        For intX = 0 To CLng(pCanvass.Width) - 1 'cycle x's
            dblDestX = intX - dblCDX  'calculate x destination
            'Rotate destination X and Y according to angle and
            'round off rotated lngPixel at X, Y coordinates
            intRotX = ((dblDestX) * dblCosA - dblDestY * dblSinA + dblCSX)
            intRotY = ((dblDestX) * dblSinA + dblDestY * dblCosA + dblCSY)
            If (PtInRect(biRct, intRotX, intRotY)) Then 'if x,y is in destination area then draw
            lngPixel = GetPixel(pSource.hdc, intRotX, intRotY) 'get lngPixel from source
            SetPixelV pCanvass.hdc, intX, intY, lngPixel   'draw dot
            End If
        Next intX
    Next intY
    pCanvass.Refresh 'Make sure image is shown
End Sub
