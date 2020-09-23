Attribute VB_Name = "modDrawShape"
'Local Vars
Private sglStopX As Single 'Ending x point
Private sglStopY As Single 'Ending y point
Private sglStartX As Single 'x where mouse down is
Private sglStartY As Single 'y where mouse down is
Private TheShape As Shape
Private DrawArea As PictureBox
Option Explicit

Public Sub StartShape(ByRef Square As Shape, ByRef xStart As Single, ByRef yStart As Single, ByRef Area As PictureBox)
        Set TheShape = Square
        Set DrawArea = Area
        TheShape.Visible = True
        sglStartX = xStart 'Assign starting x point
        sglStartY = yStart 'Assign starting y point
        'Move shape to start x,y. Reset shape to no widht or height
        TheShape.Move sglStartX, sglStartY, 0, 0
        Call ShapeInfo(TheShape.Width, TheShape.Height) 'display height and width of selection
End Sub

Public Sub StopShape(ByRef CurX As Single, ByRef CurY As Single)

    If CurX - sglStartX > 0 Then 'x is increasing (mouse moving down)
        If CurX <= DrawArea.ScaleWidth Then 'x is not beyond area
            sglStopX = CurX - sglStartX 'calc stopping x point
            TheShape.Width = sglStopX 'Assign shape width
        End If
    End If
    
    If CurY - sglStartY > 0 Then 'y is increasing (mouse moving down)
        If CurY <= DrawArea.ScaleHeight Then 'y is not beyond area
            sglStopY = CurY - sglStartY 'calc stopping y point
            TheShape.Height = sglStopY 'Assign shape Height
        End If
    End If

    If CurX - sglStartX < 0 Then 'x is decreasing (mouse moving up)
        If CurX >= 0 Then 'x is greater than starting point
            TheShape.Left = CurX 'invert shape
            sglStopX = sglStartX - CurX 'calc stopping x point
            TheShape.Width = sglStopX 'Assign shape width
        End If
    End If
    If CurY - sglStartY < 0 Then 'y is decreasing (mouse moving up)
        If CurY >= 0 Then 'y is greater than starting point
            TheShape.Top = CurY 'invert shape
            sglStopY = sglStartY - CurY 'calc stopping y point
            TheShape.Height = sglStopY 'Assign shape Height
        End If
    End If
    Call ShapeInfo(TheShape.Width, TheShape.Height) 'display height and width of selection
End Sub

Public Sub DrawShapeArea(ByRef pDestination As PictureBox)
On Error GoTo ResolveError
    pDestination.Cls 'clear old picture
    'draw image to destination
    pDestination.Width = TheShape.Width
    pDestination.Height = TheShape.Height
    pDestination.PaintPicture DrawArea.Image, 0, 0, TheShape.Width, TheShape.Height, TheShape.Left, TheShape.Top, TheShape.Width, TheShape.Height
    TheShape.Visible = False
    Call ClearMemory 'clean up
ResolveError:
    If Err.Number = 91 Then ErrorNumber = Err.Number 'object not set error
    If Err.Number = 0 Then ErrorNumber = Err.Number 'no error
    If Err.Number = 91 Then Err.Clear 'clear error
End Sub

Public Sub ShapeInfo(ByRef w As Long, ByRef h As Long)
'display height and width of selection
    frmMain.lblSelection.Caption = w & "," & h
End Sub

Public Sub ClearMemory()
    Set TheShape = Nothing 'clear used memeory
    Set DrawArea = Nothing 'clear used memeory
End Sub

