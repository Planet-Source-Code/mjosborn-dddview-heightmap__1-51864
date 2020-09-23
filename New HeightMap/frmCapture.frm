VERSION 5.00
Begin VB.Form frmCapture 
   AutoRedraw      =   -1  'True
   Caption         =   "Selection"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   Icon            =   "frmCapture.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   430
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   4680
      TabIndex        =   9
      Top             =   1080
      Width           =   1695
      Begin VB.PictureBox picSample 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000B&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   79
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   95
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Selection"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   3960
      Width           =   1695
   End
   Begin VB.PictureBox picArea 
      Height          =   3690
      Left            =   120
      ScaleHeight     =   242
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   266
      TabIndex        =   4
      Top             =   240
      Width           =   4050
      Begin VB.PictureBox picSource 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.HScrollBar hs 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   100
      Left            =   120
      Max             =   0
      TabIndex        =   3
      Top             =   3960
      Width           =   4020
   End
   Begin VB.VScrollBar vs 
      Enabled         =   0   'False
      Height          =   3660
      LargeChange     =   100
      Left            =   4200
      Max             =   0
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton cmdCenter 
      Caption         =   "C"
      Enabled         =   0   'False
      Height          =   255
      Left            =   4200
      TabIndex        =   1
      Top             =   3960
      Width           =   255
   End
   Begin VB.Frame fraSelection 
      Caption         =   "Selection"
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CheckBox chkTop 
         Caption         =   "Stay on top"
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   0
         Value           =   1  'Checked
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CurX 'mouse x position on image
Dim CurY 'mouse y postion on image

Option Explicit

Private Sub chkTop_Click()
'Determine if we keep form on top
    'Write check value, to deterimne next startup type
    WriteToINI "Top", "Value", chkTop.Value
    
    Select Case chkTop.Value
        Case 1 'on top
            SetTopMost Me, 0
        Case 0 'not on top
            SetTopMost Me, 1
    End Select
End Sub

Private Sub cmdCenter_Click()
'Center source to (x,y) cordinates of the picture box area
    If hs.Enabled = True Then hs.Value = hs.Max / 2 'calc center x
    If vs.Enabled = True Then vs.Value = vs.Max / 2 'calc center y
End Sub

Private Sub cmdClose_Click()
'Close form
    Call ExitForm
End Sub

Private Sub cmdSave_Click()
    Call SavePic(picSource, frmMain.CD1)
End Sub

Private Sub Form_Activate()
    If Me.Visible = True Then
        'update scrolls max
        Call UpdateScrolls(picArea, picSource, hs, vs, cmdCenter)
        Call cmdCenter_Click 'center image
        Call DisplayInfo 'display info about selection
        picArea.SetFocus
        chkTop.Value = ReadTheINI("Top", "Value", 1) 'get check value
        Call chkTop_Click 'check if form stays on top
        modPrepImage.SampleImage picSource, picSample
    End If
End Sub

Private Sub Form_Load()
    EnableMaxButton Me.hWnd, False 'Disable maximize button
End Sub

Private Sub Form_Resize()
    If Me.Width > 6570 Then Me.Width = 6570 'maintain initial width of form
    'maintain initial height of form
    If Me.WindowState <> 1 Then If Me.Height > 4830 Then Me.Height = 4830
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ExitForm 'exit form
End Sub

Private Sub hs_Change()
'move image left/right
    picSource.Left = 0 - hs.Value
End Sub

Private Sub picSource_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    CurX = x 'X where mouse is down on source
    CurY = y 'Y where mouse is down on source
    If Button = vbLeftButton Then Screen.MousePointer = 5 'direction arrow
End Sub

Private Sub picSource_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If hs.Enabled = True And (hs.Value + CurX - x) >= 0 And (hs.Value + CurX - x) <= hs.Max Then hs.Value = hs.Value + (CurX - x)
        If vs.Enabled = True And (vs.Value + CurY - y) >= 0 And (vs.Value + CurY - y) <= vs.Max Then vs.Value = vs.Value + (CurY - y)
    End If
    
    'Theres no sence in constantly writing this while moveing the mouse
    If frmMain.SB1.Panels(2).Text <> "Tip: Hold left mouse button to move image!" Then
        frmMain.SB1.Panels(2).Text = "Tip: Hold left mouse button to move image!"
        frmMain.SB1.Panels(2).Visible = True
    End If
End Sub

Private Sub picSource_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Screen.MousePointer = 0 'Default arrow
End Sub

Private Sub vs_Change()
'move picture up/down
    picSource.Top = 0 - vs.Value
End Sub

Private Sub DisplayInfo()
'Display info about selection
    'image size, width and height
    fraSelection.Caption = "Selection: W(" & picSource.Width & ")" & " H(" & picSource.Height & ")"
End Sub

Public Sub ExitForm()
    picSource.Cls 'clear selected image
    picArea.Cls 'clear picarea
    Set picSource = Nothing 'clear memory of picsource
    Set picArea = Nothing 'clear memory of picarea
    modDrawShape.ShapeInfo 0, 0 'reset label caption
    Unload Me 'unload form from memory
End Sub
 
