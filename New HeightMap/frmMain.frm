VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000A&
   Caption         =   "DDDV"
   ClientHeight    =   9015
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11130
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   601
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   742
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrReset 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6000
      Top             =   6600
   End
   Begin MSComctlLib.StatusBar SB1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   40
      Top             =   8760
      Width           =   11130
      _ExtentX        =   19632
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.ToolTipText     =   "Canvass size can be changed via User Options!"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   0
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picTop 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   4560
      ScaleHeight     =   105
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   705
      TabIndex        =   18
      Top             =   0
      Width           =   10575
      Begin VB.Frame Frame9 
         Caption         =   "User Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   0
         TabIndex        =   88
         Top             =   870
         Width           =   8415
         Begin VB.ComboBox cboRotOptions 
            Height          =   315
            ItemData        =   "frmMain.frx":0CCA
            Left            =   4200
            List            =   "frmMain.frx":0CDA
            TabIndex        =   99
            Text            =   "cboRotOptions"
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox cboColorStyle 
            Height          =   315
            ItemData        =   "frmMain.frx":0D08
            Left            =   1920
            List            =   "frmMain.frx":0D18
            TabIndex        =   93
            Text            =   "Combo1"
            Top             =   240
            Width           =   2175
         End
         Begin VB.ComboBox cboOptions 
            Height          =   315
            ItemData        =   "frmMain.frx":0D63
            Left            =   120
            List            =   "frmMain.frx":0D76
            TabIndex        =   89
            Text            =   "Combo1"
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "H-Map View"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   4200
            TabIndex        =   130
            Top             =   0
            Width           =   1035
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "H-Map Filter"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   1920
            TabIndex        =   129
            Top             =   0
            Width           =   1050
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000000&
         Caption         =   "Height Map"
         ForeColor       =   &H00000000&
         Height          =   1545
         Left            =   9000
         TabIndex        =   66
         Top             =   0
         Width           =   1575
         Begin VB.PictureBox picSample 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000B&
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   120
            ScaleHeight     =   79
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   87
            TabIndex        =   67
            Top             =   240
            Width           =   1335
            Begin VB.Shape Shape4 
               BorderColor     =   &H000000FF&
               Height          =   135
               Left            =   0
               Top             =   0
               Visible         =   0   'False
               Width           =   135
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000004&
         Caption         =   "Draw Status"
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   3615
         Begin VB.CheckBox chkShowDraw 
            BackColor       =   &H80000000&
            Caption         =   "Show Drawing"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2160
            TabIndex        =   39
            Top             =   0
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin MSComctlLib.ProgressBar pg 
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   250
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label lblPerc 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   3000
            TabIndex        =   32
            Top             =   240
            Width           =   255
         End
         Begin VB.Label lblPixelCycles 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(0,0)"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1440
            TabIndex        =   31
            Top             =   600
            Width           =   315
         End
         Begin VB.Label lblPC 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(X,Y) Pixel-Cycles:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   600
            Width           =   1275
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000004&
         Caption         =   "Viewing / Cordinates"
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   3720
         TabIndex        =   19
         Top             =   0
         Width           =   4695
         Begin VB.Label lblSelection 
            AutoSize        =   -1  'True
            BackColor       =   &H80000000&
            Caption         =   "0,0"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   2650
            TabIndex        =   91
            Top             =   0
            Width           =   225
         End
         Begin VB.Label Label1 
            Caption         =   "Selection:"
            Height          =   195
            Left            =   1920
            TabIndex        =   90
            Top             =   0
            Width           =   730
         End
         Begin VB.Label lblTopLeft 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(X,Y)"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   910
            TabIndex        =   27
            Top             =   240
            Width           =   345
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Top, Left):"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   780
         End
         Begin VB.Label lblMCords 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(X,Y) Cordinates:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   480
            Width           =   1185
         End
         Begin VB.Label lblMouseCords 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0,0"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1330
            TabIndex        =   24
            Top             =   480
            Width           =   225
         End
         Begin VB.Label lblBR 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Bottom,Right):"
            Height          =   195
            Left            =   2640
            TabIndex        =   23
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label lblBottomRight 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(X,Y)"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   3720
            TabIndex        =   22
            Top             =   240
            Width           =   345
         End
         Begin VB.Label lblVAXY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(X,Y)"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   3720
            TabIndex        =   21
            Top             =   480
            Width           =   345
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(X,Y) View-Window:"
            Height          =   195
            Left            =   2280
            TabIndex        =   20
            Top             =   480
            Width           =   1410
         End
      End
   End
   Begin VB.PictureBox picHMViewArea 
      BackColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   4560
      ScaleHeight     =   317
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   293
      TabIndex        =   16
      Top             =   1560
      Width           =   4455
      Begin VB.PictureBox picHMDest 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   12000
         Left            =   0
         ScaleHeight     =   800
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   800
         TabIndex        =   17
         Top             =   0
         Width           =   12000
         Begin VB.Shape shpFollow 
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   90
            Left            =   960
            Shape           =   3  'Circle
            Top             =   360
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Line yL 
            BorderColor     =   &H00FFFFFF&
            Visible         =   0   'False
            X1              =   128
            X2              =   128
            Y1              =   0
            Y2              =   224
         End
         Begin VB.Line xL 
            BorderColor     =   &H00FFFFFF&
            Visible         =   0   'False
            X1              =   0
            X2              =   152
            Y1              =   176
            Y2              =   176
         End
         Begin VB.Shape shCapture 
            BorderColor     =   &H00FFFFFF&
            BorderStyle     =   3  'Dot
            Height          =   615
            Left            =   240
            Top             =   1320
            Visible         =   0   'False
            Width           =   1455
         End
      End
   End
   Begin VB.CommandButton cmdCanvassCenter 
      Caption         =   "C"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6960
      TabIndex        =   15
      Top             =   3360
      Width           =   255
   End
   Begin VB.VScrollBar vsCScroll 
      Enabled         =   0   'False
      Height          =   2175
      LargeChange     =   100
      Left            =   6960
      Max             =   1
      TabIndex        =   14
      Top             =   1200
      Width           =   255
   End
   Begin VB.HScrollBar hsCScroll 
      Enabled         =   0   'False
      Height          =   255
      LargeChange     =   100
      Left            =   5880
      Max             =   1
      TabIndex        =   13
      Top             =   3600
      Width           =   1095
   End
   Begin VB.PictureBox picLeft 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   8760
      Left            =   0
      ScaleHeight     =   584
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.Frame fraOptions 
         Caption         =   "Color Map "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4575
         Index           =   4
         Left            =   0
         TabIndex        =   92
         Top             =   5760
         Visible         =   0   'False
         Width           =   4455
         Begin VB.Frame Frame10 
            Caption         =   "Color Height Graph"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   3015
            Left            =   120
            TabIndex        =   112
            Top             =   1440
            Width           =   4280
            Begin VB.PictureBox picOriginalColor 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   215
               Left            =   1160
               ScaleHeight     =   12
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   23
               TabIndex        =   136
               Top             =   2760
               Width           =   375
            End
            Begin VB.ComboBox cboGraph 
               Height          =   315
               ItemData        =   "frmMain.frx":0DBA
               Left            =   120
               List            =   "frmMain.frx":0DC7
               TabIndex        =   132
               Text            =   "Combo1"
               Top             =   2160
               Width           =   1095
            End
            Begin VB.PictureBox pColorMap 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   215
               Left            =   2880
               ScaleHeight     =   12
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   87
               TabIndex        =   121
               Top             =   2760
               Width           =   1335
            End
            Begin VB.CheckBox chkShowPos 
               Caption         =   "Show Position"
               Height          =   255
               Left            =   2760
               TabIndex        =   120
               Top             =   0
               Width           =   1335
            End
            Begin VB.PictureBox pGraphFrame 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   1695
               Left            =   720
               ScaleHeight     =   111
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   231
               TabIndex        =   113
               Top             =   345
               Width           =   3495
               Begin VB.PictureBox pGraph 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  FillColor       =   &H0000FFFF&
                  ForeColor       =   &H80000008&
                  Height          =   1665
                  Left            =   0
                  ScaleHeight     =   111
                  ScaleMode       =   3  'Pixel
                  ScaleWidth      =   231
                  TabIndex        =   119
                  Top             =   0
                  Width           =   3465
               End
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Displayed Color"
               Height          =   195
               Left            =   2880
               TabIndex        =   135
               Top             =   2520
               Width           =   1095
            End
            Begin VB.Label lblCGX 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   195
               Left            =   2280
               TabIndex        =   128
               Top             =   2160
               Width           =   90
            End
            Begin VB.Line Line17 
               X1              =   1560
               X2              =   3120
               Y1              =   2280
               Y2              =   2280
            End
            Begin VB.Label lblGraphMax 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               Height          =   195
               Left            =   3120
               TabIndex        =   127
               Top             =   2160
               Width           =   90
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               Height          =   195
               Left            =   1440
               TabIndex        =   126
               Top             =   2160
               Width           =   90
            End
            Begin VB.Label lblColorAt 
               AutoSize        =   -1  'True
               Caption         =   "(0,0,0)"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   1560
               TabIndex        =   125
               Top             =   2760
               Width           =   450
            End
            Begin VB.Label lblCV 
               AutoSize        =   -1  'True
               Caption         =   "Map (X,Y,Height):"
               Height          =   195
               Left            =   120
               TabIndex        =   124
               Top             =   2520
               Width           =   1260
            End
            Begin VB.Label lblListValue 
               AutoSize        =   -1  'True
               Caption         =   "0"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   1440
               TabIndex        =   123
               Top             =   2520
               Width           =   90
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Original Color:"
               Height          =   195
               Left            =   120
               TabIndex        =   122
               Top             =   2760
               Width           =   975
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "255"
               Height          =   195
               Left            =   120
               TabIndex        =   118
               Top             =   240
               Width           =   270
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "127.5"
               Height          =   195
               Left            =   120
               TabIndex        =   117
               Top             =   1065
               Width           =   405
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "0"
               Height          =   195
               Left            =   120
               TabIndex        =   116
               Top             =   1905
               Width           =   90
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "85"
               Height          =   195
               Left            =   120
               TabIndex        =   115
               Top             =   1515
               Width           =   180
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "170"
               Height          =   195
               Left            =   120
               TabIndex        =   114
               Top             =   660
               Width           =   270
            End
            Begin VB.Line Line12 
               BorderColor     =   &H00000000&
               X1              =   725
               X2              =   405
               Y1              =   345
               Y2              =   345
            End
            Begin VB.Line Line13 
               BorderColor     =   &H00000000&
               X1              =   725
               X2              =   405
               Y1              =   795
               Y2              =   795
            End
            Begin VB.Line Line14 
               BorderColor     =   &H00000000&
               X1              =   715
               X2              =   555
               Y1              =   1185
               Y2              =   1185
            End
            Begin VB.Line Line15 
               BorderColor     =   &H00000000&
               X1              =   720
               X2              =   330
               Y1              =   1605
               Y2              =   1605
            End
            Begin VB.Line Line16 
               BorderColor     =   &H00000000&
               X1              =   720
               X2              =   240
               Y1              =   2025
               Y2              =   2025
            End
         End
         Begin VB.Frame fraColorValue 
            Caption         =   "Color Value At: (X,Y)"
            Enabled         =   0   'False
            Height          =   1095
            Left            =   120
            TabIndex        =   101
            Top             =   240
            Width           =   4215
            Begin VB.CheckBox chkFollowMode 
               Caption         =   "Enable Follow Mode"
               Height          =   255
               Left            =   2280
               TabIndex        =   131
               ToolTipText     =   "Place Currsor on Source Image!"
               Top             =   720
               Value           =   1  'Checked
               Width           =   1815
            End
            Begin VB.HScrollBar hsY 
               Height          =   255
               Left            =   1200
               TabIndex        =   111
               Top             =   720
               Width           =   975
            End
            Begin VB.HScrollBar hsX 
               Height          =   255
               Left            =   120
               TabIndex        =   110
               Top             =   720
               Width           =   975
            End
            Begin VB.CommandButton cmdGet 
               Caption         =   "Show Value"
               Height          =   255
               Left            =   2280
               TabIndex        =   107
               Top             =   360
               Width           =   1815
            End
            Begin VB.TextBox txtY 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               Left            =   1200
               TabIndex        =   103
               Text            =   "0"
               ToolTipText     =   "Y Location"
               Top             =   360
               Width           =   855
            End
            Begin VB.TextBox txtX 
               Appearance      =   0  'Flat
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               Left            =   240
               TabIndex        =   102
               Text            =   "0"
               ToolTipText     =   "X location"
               Top             =   360
               Width           =   855
            End
            Begin VB.Label lblMax 
               AutoSize        =   -1  'True
               Caption         =   "(0 - 0 , 0 - 0)"
               Height          =   195
               Left            =   2400
               TabIndex        =   109
               Top             =   0
               Width           =   855
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Max:"
               Height          =   195
               Left            =   2040
               TabIndex        =   108
               Top             =   0
               Width           =   345
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ")"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   2040
               TabIndex        =   106
               Top             =   240
               Width           =   120
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ","
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Left            =   1080
               TabIndex        =   105
               Top             =   120
               Width           =   135
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "("
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               TabIndex        =   104
               Top             =   240
               Width           =   120
            End
         End
         Begin VB.CheckBox chkColorMap 
            Caption         =   "Enable Color Map"
            Height          =   255
            Left            =   2760
            TabIndex        =   100
            ToolTipText     =   "Disable to increase draw time!"
            Top             =   0
            Width           =   1575
         End
      End
      Begin VB.Frame fraOptions 
         BackColor       =   &H80000000&
         Caption         =   "Canvass Size"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2055
         Index           =   3
         Left            =   0
         TabIndex        =   56
         Top             =   5760
         Width           =   4455
         Begin VB.CheckBox chkAutoSize 
            Caption         =   "Auto Size to height map"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   1200
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CommandButton cmdDefault2 
            Caption         =   "Default"
            Height          =   375
            Left            =   120
            TabIndex        =   60
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CommandButton cmdSetCHW 
            Caption         =   "Set"
            Height          =   375
            Left            =   2760
            TabIndex        =   59
            Top             =   1560
            Width           =   1575
         End
         Begin VB.TextBox txtCanSize 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   960
            TabIndex        =   58
            Text            =   "800"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox txtCanSize 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   57
            Text            =   "800"
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lblCanXY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Y 800"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   2400
            TabIndex        =   64
            Top             =   240
            Width           =   420
         End
         Begin VB.Label lblCanXY 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "X 800"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   3840
            TabIndex        =   63
            Top             =   960
            Width           =   420
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H00000000&
            BorderStyle     =   3  'Dot
            Height          =   495
            Left            =   2520
            Top             =   600
            Width           =   975
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00000000&
            Height          =   255
            Left            =   2520
            Top             =   840
            Width           =   375
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00000000&
            BorderStyle     =   3  'Dot
            Height          =   615
            Left            =   2520
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblCH 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(Y) Height:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   480
            Width           =   750
         End
         Begin VB.Label lblCW 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "(X) Width:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   840
            Width           =   705
         End
      End
      Begin VB.Frame fraOptions 
         BackColor       =   &H80000000&
         Caption         =   "Output Colors"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2535
         Index           =   2
         Left            =   0
         TabIndex        =   43
         Top             =   5760
         Width           =   4455
         Begin VB.OptionButton optOPC 
            BackColor       =   &H80000000&
            Caption         =   "RGB (shaded)"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   1335
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H80000000&
            ForeColor       =   &H00FFFFFF&
            Height          =   1455
            Left            =   2040
            TabIndex        =   49
            Top             =   120
            Width           =   2295
            Begin VB.TextBox txtHSensitivity 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1440
               TabIndex        =   53
               Text            =   "5"
               Top             =   720
               Width           =   615
            End
            Begin VB.CheckBox chkRGB 
               BackColor       =   &H80000000&
               Caption         =   "Blue"
               ForeColor       =   &H00FF8080&
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   52
               Top             =   1080
               Width           =   615
            End
            Begin VB.CheckBox chkRGB 
               BackColor       =   &H80000000&
               Caption         =   "Green"
               ForeColor       =   &H00008000&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   51
               Top             =   720
               Width           =   735
            End
            Begin VB.CheckBox chkRGB 
               BackColor       =   &H80000000&
               Caption         =   "Red"
               ForeColor       =   &H000000FF&
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   50
               Top             =   360
               Width           =   615
            End
            Begin VB.Line Line5 
               BorderColor     =   &H00000000&
               X1              =   1080
               X2              =   600
               Y1              =   1200
               Y2              =   1200
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00000000&
               X1              =   1080
               X2              =   720
               Y1              =   840
               Y2              =   840
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00000000&
               X1              =   1080
               X2              =   600
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Line Line2 
               BorderColor     =   &H00000000&
               X1              =   1080
               X2              =   1440
               Y1              =   840
               Y2              =   840
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Color Sensitivity"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   1065
               TabIndex        =   54
               Top             =   240
               Width           =   1110
            End
            Begin VB.Line Line1 
               BorderColor     =   &H00000000&
               X1              =   1080
               X2              =   1080
               Y1              =   480
               Y2              =   1200
            End
         End
         Begin VB.OptionButton optOPC 
            BackColor       =   &H80000000&
            Caption         =   "Define Colors"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   48
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton optOPC 
            BackColor       =   &H80000000&
            Caption         =   "Default Image Colors"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   47
            Top             =   600
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.PictureBox picDefineColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   15
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   111
            TabIndex        =   46
            ToolTipText     =   "Click to choose color!"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H80000000&
            Caption         =   "Canvass Back-Color / Click to Change"
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   120
            TabIndex        =   44
            Top             =   1680
            Width           =   4215
            Begin VB.PictureBox picBG 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   120
               ScaleHeight     =   15
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   263
               TabIndex        =   45
               ToolTipText     =   "Click here to change"
               Top             =   240
               Width           =   3975
            End
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00000000&
            X1              =   1440
            X2              =   2040
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00000000&
            X1              =   1200
            X2              =   1680
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line8 
            BorderColor     =   &H00000000&
            X1              =   1680
            X2              =   1680
            Y1              =   1320
            Y2              =   1080
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H80000000&
         Caption         =   "Draw Style"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   0
         TabIndex        =   42
         Top             =   5040
         Width           =   4455
         Begin VB.CommandButton cmdStop 
            Caption         =   "Stop"
            Enabled         =   0   'False
            Height          =   330
            Left            =   3240
            TabIndex        =   137
            ToolTipText     =   "Stop Drawing Height-Map"
            Top             =   210
            Width           =   1095
         End
         Begin VB.CommandButton cmdDraw 
            Caption         =   "Draw"
            Enabled         =   0   'False
            Height          =   330
            Left            =   2040
            TabIndex        =   134
            ToolTipText     =   "Draw Height-Map"
            Top             =   210
            Width           =   1095
         End
         Begin VB.ComboBox cboDS 
            Height          =   315
            ItemData        =   "frmMain.frx":0DDD
            Left            =   120
            List            =   "frmMain.frx":0DF3
            TabIndex        =   133
            Text            =   "cboDS"
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame fraOptions 
         BackColor       =   &H80000000&
         Caption         =   "Rotation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   3255
         Index           =   1
         Left            =   0
         TabIndex        =   38
         Top             =   5760
         Visible         =   0   'False
         Width           =   4455
         Begin VB.PictureBox picPlaceHolder 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000000&
            Height          =   2175
            Left            =   3200
            ScaleHeight     =   145
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   80
            TabIndex        =   84
            Top             =   720
            Width           =   1200
            Begin VB.PictureBox picPiSource 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   600
               Left            =   240
               ScaleHeight     =   40
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   48
               TabIndex        =   86
               Top             =   1440
               Width           =   720
            End
            Begin VB.PictureBox picPiDest 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   1200
               Left            =   0
               ScaleHeight     =   78
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   78
               TabIndex        =   85
               Top             =   120
               Width           =   1200
            End
            Begin VB.Shape Shape6 
               BorderColor     =   &H00C00000&
               FillColor       =   &H00C00000&
               Height          =   630
               Left            =   225
               Top             =   1425
               Width           =   750
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H80000000&
            ForeColor       =   &H00000000&
            Height          =   2895
            Left            =   60
            TabIndex        =   41
            Top             =   240
            Width           =   3100
            Begin VB.HScrollBar hsRotation 
               Height          =   255
               Left            =   720
               Max             =   360
               TabIndex        =   81
               Top             =   2520
               Width           =   1815
            End
            Begin VB.OptionButton optDeg 
               BackColor       =   &H80000000&
               Caption         =   "0 "
               ForeColor       =   &H00FF0000&
               Height          =   255
               Index           =   0
               Left            =   2475
               TabIndex        =   80
               Top             =   1080
               Value           =   -1  'True
               Width           =   375
            End
            Begin VB.OptionButton optDeg 
               BackColor       =   &H80000000&
               Caption         =   "45"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   1
               Left            =   2160
               TabIndex        =   79
               Top             =   480
               Width           =   495
            End
            Begin VB.OptionButton optDeg 
               BackColor       =   &H80000000&
               Caption         =   "90"
               ForeColor       =   &H00000000&
               Height          =   245
               Index           =   2
               Left            =   1470
               TabIndex        =   78
               Top             =   240
               Width           =   495
            End
            Begin VB.OptionButton optDeg 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000000&
               Caption         =   "135"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   3
               Left            =   300
               TabIndex        =   77
               Top             =   495
               Width           =   615
            End
            Begin VB.OptionButton optDeg 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000000&
               Caption         =   "180"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   4
               Left            =   40
               TabIndex        =   76
               Top             =   1200
               Width           =   590
            End
            Begin VB.OptionButton optDeg 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000000&
               Caption         =   "225"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   5
               Left            =   300
               TabIndex        =   75
               Top             =   1900
               Width           =   615
            End
            Begin VB.OptionButton optDeg 
               BackColor       =   &H80000000&
               Caption         =   "270"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   6
               Left            =   1470
               TabIndex        =   74
               Top             =   2175
               Width           =   615
            End
            Begin VB.OptionButton optDeg 
               BackColor       =   &H80000000&
               Caption         =   "315"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   7
               Left            =   2160
               TabIndex        =   73
               Top             =   1920
               Width           =   615
            End
            Begin VB.OptionButton optDeg 
               BackColor       =   &H80000000&
               Caption         =   "360"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   8
               Left            =   2475
               TabIndex        =   72
               Top             =   1320
               Width           =   600
            End
            Begin VB.Label lblAngle 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Angle"
               ForeColor       =   &H00000000&
               Height          =   195
               Left            =   120
               TabIndex        =   83
               Top             =   2520
               Width           =   405
            End
            Begin VB.Label lblRotate 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   2640
               TabIndex        =   82
               Top             =   2520
               Width           =   105
            End
            Begin VB.Shape Shape5 
               BackColor       =   &H00FF0000&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00000000&
               FillColor       =   &H0000FF00&
               Height          =   100
               Left            =   1440
               Shape           =   3  'Circle
               Top             =   1270
               Width           =   255
            End
            Begin VB.Line lAngle 
               BorderColor     =   &H00C00000&
               BorderWidth     =   2
               X1              =   1560
               X2              =   2400
               Y1              =   1320
               Y2              =   1320
            End
            Begin VB.Line Line10 
               BorderColor     =   &H00000000&
               X1              =   720
               X2              =   2400
               Y1              =   1320
               Y2              =   1320
            End
            Begin VB.Line Line9 
               BorderColor     =   &H00000000&
               X1              =   1560
               X2              =   1560
               Y1              =   480
               Y2              =   2160
            End
            Begin VB.Shape shapePI 
               BackColor       =   &H80000000&
               BackStyle       =   1  'Opaque
               BorderColor     =   &H00000000&
               BorderWidth     =   2
               FillColor       =   &H00FFFFFF&
               Height          =   1695
               Left            =   700
               Shape           =   3  'Circle
               Top             =   480
               Width           =   1695
            End
         End
         Begin VB.Line Line11 
            X1              =   3360
            X2              =   4200
            Y1              =   700
            Y2              =   700
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Example"
            Height          =   195
            Left            =   3480
            TabIndex        =   87
            Top             =   480
            Width           =   600
         End
      End
      Begin VB.Frame fraOptions 
         BackColor       =   &H80000000&
         Caption         =   "Input Height"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   2175
         Index           =   0
         Left            =   0
         TabIndex        =   33
         Top             =   5760
         Width           =   4455
         Begin VB.ListBox lstRGBH 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            ForeColor       =   &H00FFFFFF&
            Height          =   705
            ItemData        =   "frmMain.frx":0E25
            Left            =   120
            List            =   "frmMain.frx":0E35
            Style           =   1  'Checkbox
            TabIndex        =   68
            Top             =   290
            Width           =   4215
         End
         Begin VB.CommandButton cmdHeightDefault 
            Caption         =   "Default"
            Height          =   315
            Left            =   120
            TabIndex        =   65
            Top             =   1800
            Width           =   4215
         End
         Begin VB.HScrollBar hsHO 
            Height          =   255
            LargeChange     =   40
            Left            =   1200
            Max             =   255
            Min             =   1
            TabIndex        =   35
            Top             =   1080
            Value           =   28
            Width           =   2775
         End
         Begin VB.CheckBox chkInvert 
            BackColor       =   &H80000000&
            Caption         =   "Invert Height"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   3120
            TabIndex        =   34
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lblLP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   3120
            TabIndex        =   98
            Top             =   1440
            Width           =   90
         End
         Begin VB.Label lblHP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   1200
            TabIndex        =   97
            Top             =   1440
            Width           =   90
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lowest Point:"
            Height          =   195
            Left            =   2040
            TabIndex        =   96
            Top             =   1440
            Width           =   960
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Highest Point:"
            Height          =   195
            Left            =   120
            TabIndex        =   95
            Top             =   1440
            Width           =   990
         End
         Begin VB.Label lblHOSV 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "28"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   4080
            TabIndex        =   37
            Top             =   1080
            Width           =   105
         End
         Begin VB.Label lblHOS 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Height offset"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   1080
            Width           =   915
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000000&
         Caption         =   "View Source Image As... / Draw As..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   0
         TabIndex        =   9
         Top             =   4320
         Width           =   4455
         Begin VB.OptionButton optImageStyle 
            BackColor       =   &H80000000&
            Caption         =   "Best-Fit"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton optImageStyle 
            BackColor       =   &H80000000&
            Caption         =   "Normal"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   11
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optImageStyle 
            BackColor       =   &H80000000&
            Caption         =   "Stretch to fit"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame fraTheImage 
         BackColor       =   &H80000000&
         Caption         =   "The Image"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4215
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   4455
         Begin VB.CommandButton cmdCenter 
            Caption         =   "C"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4080
            TabIndex        =   71
            Top             =   3840
            Width           =   255
         End
         Begin VB.VScrollBar vsSrcScroll 
            Enabled         =   0   'False
            Height          =   3300
            Left            =   4080
            Max             =   32000
            TabIndex        =   70
            Top             =   480
            Width           =   255
         End
         Begin VB.HScrollBar hsScrScroll 
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            Max             =   32000
            TabIndex        =   69
            Top             =   3840
            Width           =   3900
         End
         Begin VB.PictureBox picSrcViewArea 
            BackColor       =   &H00000000&
            Height          =   3330
            Left            =   120
            ScaleHeight     =   218
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   258
            TabIndex        =   2
            Top             =   480
            Width           =   3930
            Begin VB.PictureBox picSrcImage 
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
               TabIndex        =   4
               Top             =   0
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.PictureBox picSrcDest 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   975
               Left            =   0
               ScaleHeight     =   65
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   97
               TabIndex        =   3
               Top             =   0
               Visible         =   0   'False
               Width           =   1455
            End
         End
         Begin VB.Label lblH 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Image Height:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2400
            TabIndex        =   8
            Top             =   240
            Width           =   990
         End
         Begin VB.Label lblW 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Image Width:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   945
         End
         Begin VB.Label lblSrcH 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   3480
            TabIndex        =   6
            Top             =   240
            Width           =   90
         End
         Begin VB.Label lblSrcW 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   1200
            TabIndex        =   5
            Top             =   240
            Width           =   90
         End
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOptions 
         Caption         =   "&Open"
         Index           =   1
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "&Save (Export Heighmap)"
         Enabled         =   0   'False
         Index           =   3
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHistory 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistory 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHistory 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuIS 
      Caption         =   "Image Syle"
      Begin VB.Menu mnuVI 
         Caption         =   "&Best-Fit"
         Index           =   0
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuVI 
         Caption         =   "&Normal"
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuVI 
         Caption         =   "Stretch to &Fit"
         Index           =   2
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuZoom 
      Caption         =   "Zoom"
      Begin VB.Menu mnuZ 
         Caption         =   "200%"
         Index           =   0
      End
      Begin VB.Menu mnuZ 
         Caption         =   "175%"
         Index           =   1
      End
      Begin VB.Menu mnuZ 
         Caption         =   "150%"
         Index           =   2
      End
      Begin VB.Menu mnuZ 
         Caption         =   "125%"
         Index           =   3
      End
      Begin VB.Menu mnuZ 
         Caption         =   "100%"
         Index           =   4
      End
      Begin VB.Menu mnuZ 
         Caption         =   "75%"
         Index           =   5
      End
      Begin VB.Menu mnuZ 
         Caption         =   "50%"
         Index           =   6
      End
      Begin VB.Menu mnuZ 
         Caption         =   "25%"
         Index           =   7
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const strExt = "Supported Image Formats|*.jpg;" _
     & "*.bmp;*.gif;*.dib;*.wmf;*.emf;*.ico;*.cur;" _
     & "|.jpg|*.jpg*|.bmp|*.bmp*|.gif (non-anim)|" _
     & "*.gif*|.dib (view only)|*.dib*|.wmf (view only)" _
     & "|*.wmf*|.emf (view only)|*.emf*|.ico (view only)" _
     & "|*.ico*|.cur (view only)|*.cur*" 'Common Dialog box filter
     
Dim dblZoomAmount As Double 'Zoom amount of image
Dim ImageLoaded As Boolean 'Test: is the image loaded
Dim HMDrawn As Boolean 'Test: has is height map draw
Dim blnDraw As Boolean 'Test: stop drawing
Dim CurX 'mouse x position on image
Dim CurY 'mouse y postion on image
Dim strX As String 'X value from string in string array strXYZ()
Dim strY As String 'Y value from string in string array strXYZ()
Dim strZ As String 'Z value from string in string array strXYZ()
'Dim lngGraphColorA() As Long 'Stored colors for graph
Dim intGraphYArray() As Integer 'Storage for calculated graph y values
Dim intGraphXMax As Integer 'Largest value to cycle graph x values
Dim intGraphYMax As Integer 'Largest value to cycle graph y values

Option Explicit

Private Sub cboColorStyle_Click()
    If cboColorStyle.ListIndex = 2 Or cboColorStyle.ListIndex = 0 Then
    optOPC(2).Value = True 'set to default color type
    End If
End Sub

Private Sub cboOptions_Click()
'Show the current frame from "fraOpt" array, using the index from the combo box(cboOptions).
'Hide all other frames from frame array "fraOpt"
    Call CheckUncheck(4, fraOptions, cboOptions.ListIndex, "Visible")
End Sub

Private Sub chkAutoSize_Click()
'Autosize the height map canvass to picture
    Call cmdDefault2_Click
End Sub

Private Sub chkRGB_Click(Index As Integer)
    optOPC(0).Value = True 'check option rgb
End Sub

Private Sub chkShowPos_Click()
'Show the x,y lines that point to the current point the heightmap
    If chkShowPos.Value = 1 Then 'is checked
        If HMDrawn = True And fraColorValue.Enabled = True Then 'has height-map been drawn
            xL.Visible = True 'show the x-horizontal line
            yL.Visible = True 'show the y-vertical line
            shpFollow.Visible = True
        End If
    Else 'is not checked
        xL.Visible = False 'hide the x-horizontal line
        yL.Visible = False 'hide the y-vertical line
        shpFollow.Visible = False
    End If
End Sub

Private Sub cmdDefault2_Click()
'Set default canvass size
    If chkAutoSize.Value = 1 And ImageLoaded = False Then
        picHMDest.Width = picHMViewArea.Width 'Assign width from txtCanSize, on fraOpt(6)
        picHMDest.Height = picHMViewArea.Height 'Assign height from txtCanSize, on fraOpt(6)
        txtCanSize(0).Text = Int(picHMViewArea.Width) 'Display canvass width size, in canvass size frame options
        txtCanSize(1).Text = Int(picHMViewArea.Height) 'Display canvass height Size, in canvass size frame options
        lblCanXY(0).Caption = "X " & Int(picHMViewArea.Width) 'Display canvass width size, in canvass size frame options
        lblCanXY(1).Caption = "Y " & Int(picHMViewArea.Height) 'Display canvass height size, in canvass size frame options
        'Display canvass w/h on status bar, at bottom of form
        SB1.Panels(1).Text = "Canvass Size:(" & val(txtCanSize(0).Text) & "," & val(txtCanSize(1).Text) & ")"
    ElseIf chkAutoSize.Value = 0 Then
        picHMDest.Width = 800 'Assign width from txtCanSize, on fraOpt(6)
        picHMDest.Height = 800 'Assign height from txtCanSize, on fraOpt(6)
        txtCanSize(0).Text = 800
        txtCanSize(1).Text = 800
        lblCanXY(0).Caption = "X " & 800
        lblCanXY(1).Caption = "Y " & 800
        SB1.Panels(1).Text = "Canvass Size:(" & val(txtCanSize(0).Text) & "," & val(txtCanSize(1).Text) & ")" 'Show w/h
    End If
    vsCScroll.Value = 0 'Reset picHMDest vertical scroll value
    hsCScroll.Value = 0 'reset picHMDest horizontal scroll value
    Call UpdateScrolls(picHMViewArea, picHMDest, hsCScroll, vsCScroll, cmdCanvassCenter)

    'UpdateScrolls 'show/hide scrolls vsCScroll/hsCScroll & cmdCanvassCenter. assigns vsCScroll/hsCScroll max
End Sub

Private Sub cmdDraw_Click()
'Initate a heightmap to be draw. Iniate the various routines that need to
'preformed both before and after a height-map is draw
    Dim dblRadious As Double
    If chkAutoSize.Value = 1 Then 'AutoSize Canvass
        dblRadious = Sqr(picSrcDest.Width ^ 2 + picSrcDest.Height ^ 2)
        picHMDest.Width = Int(dblRadious + 255)
        picHMDest.Height = Int(dblRadious + 255)
        txtCanSize(0).Text = Int(picHMDest.Width)
        txtCanSize(1).Text = Int(picHMDest.Height)
        lblCanXY(0).Caption = "X " & picHMDest.Width
        lblCanXY(1).Caption = "Y " & picHMDest.Height
        SB1.Panels(1).Text = "Canvass Size:(" & val(txtCanSize(0).Text) & "," & val(txtCanSize(1).Text) & ")" 'Show w/h
    End If
    Call SetDrawState(True) 'determine if we will allow a drawing
    DoEvents
    'moves controls around, vs, hs, picHMViewArea H/W, cmdCanvassCenter
    Call cmdCanvassCenter_Click 'recenter canvass
    Call MoveControls(Me, picHMViewArea, picHMDest, picTop, picLeft, SB1, hsCScroll, vsCScroll, cmdCanvassCenter) 'Reposition some of the controls
    'Call cmdCanvassCenter_Click 'Center picHMDest to (x,y) cordinates of the picture box picHMViewArea
    modMapInfo.SetHeight CInt(hsHO.Value) 'set the height offset
    Call DrawHeightMap 'choose draw type then draw
    modPrepImage.SampleImage picHMDest, picSample 'draw image in small widow
    mnuFileOptions(3).Enabled = True 'enable the ablilty to save the new heightmap image
    
End Sub

Private Sub DrawHeightMap()
'Choose the appropriate draw routine.
    
    Dim DrawTimer As clsTimer ' timer
    Set DrawTimer = New clsTimer 'set new timer
    
    Call PreNewHeightMap 'Preform various initialisations priar to new hieghtmap
    DrawTimer.StartTimer 'start timer
    modMapInfo.InitialiseHM picSrcDest, picHMDest, -CInt(hsRotation.Value), cboDS.ListIndex 'draw dots
    DrawTimer.StopTimer 'end timer
    Call PostNewHeightMap(DrawTimer.Elasped) 'preform various inits after height map
    Set DrawTimer = Nothing 'clear drawtimer from memory
    
End Sub

Private Sub PreNewHeightMap()
    cmdStop.Enabled = True
    Call UpdateScrolls(picHMViewArea, picHMDest, hsCScroll, vsCScroll, cmdCanvassCenter) 'show/hide scrolls vsCScroll/hsCScroll & cmdCanvassCenter. assigns vsCScroll/hsCScroll max
    Call cmdCanvassCenter_Click
    Shape4.Visible = False
    picSample.Cls
    picSample.Refresh
    picHMDest.Cls 'clear old image
    picHMDest.Refresh 'make sure image was cleared
    pg.Max = picHMDest.Height 'Status Bar max
    Me.Caption = "Drawing..." 'display on form that drawing started
    Me.MousePointer = 11 'hour glass
    txtY.Text = 0
    txtX.Text = 0
    hsX.Value = 0
    hsY.Value = 0
    hsX.Max = 0
    hsY.Max = 0
    intGraphXMax = 0
    intGraphYMax = 0
    lblMax.Caption = "(0 - 0 , 0 - 0)"
    fraColorValue.Enabled = False
End Sub

Private Sub PostNewHeightMap(ByVal strTime As String)
    Me.MousePointer = 0 'default arrow
    picHMDest.Refresh 'make sure image is displayed
    Me.Caption = "Processing Time: " & strTime & " ms" 'Show draw time taken
    HMDrawn = True
    cmdStop.Enabled = False
    Select Case cmdCanvassCenter.Enabled
        Case True
            Shape4.Visible = True
            Call cmdCanvassCenter_Click
        Case False
            Shape4.Visible = False
    End Select
    
    If chkColorMap.Value = 1 Then
        HMDrawn = True
        fraColorValue.Enabled = True
        intGraphXMax = Int(((GetImageScaleW) - 1) / intStep)
        intGraphYMax = Int(((GetImageScaleH) - 1) / intStep)
        lblMax.Caption = "(0 - " & intGraphXMax & " , 0 - " & intGraphYMax & ")"
        lblGraphMax.Caption = intGraphXMax
        lblCGX.Caption = 0
        txtX.Text = 0
        txtY.Text = 0
        hsX.Max = intGraphXMax
        hsY.Max = intGraphYMax
        ReDim intGraphYArray(intGraphXMax)
        'ReDim lngGraphColorA(intGraphXMax, intGraphYMax)
        Call cmdGet_Click
    End If
End Sub

Private Sub cmdGet_Click()
'Scroll the color value array, inputed x and y values
    
    Call ColorArrayValues(val(txtX.Text), val(txtY.Text))
End Sub

Private Sub cmdSetCHW_Click()
'Set Canvass Height and Width, from frame array fraOpt(6), "Canvass Size Width/Height"
    picHMDest.Width = val(txtCanSize(0).Text) 'Assign width from txtCanSize, on fraOpt(6)
    picHMDest.Height = val(txtCanSize(1).Text) 'Assign height from txtCanSize, on fraOpt(6)
    SB1.Panels(1).Text = "Canvass Size:(" & val(txtCanSize(0).Text) & "," & val(txtCanSize(1).Text) & ")" 'Show w/h
    vsCScroll.Value = 0 'Reset picHMDest vertical scroll value
    hsCScroll.Value = 0 'reset picHMDest horizontal scroll value
    Call UpdateScrolls(picHMViewArea, picHMDest, hsCScroll, vsCScroll, cmdCanvassCenter) 'show/hide scrolls vsCScroll/hsCScroll & cmdCanvassCenter. assigns vsCScroll/hsCScroll max
End Sub

Private Sub cmdHeightDefault_Click()
'Reset Default height offset values
    hsHO.Value = 28
    lstRGBH.Selected(lstRGBH.ListCount - 1) = True 'check default
End Sub

Private Sub cmdStop_Click()
    Call SetDrawState(False)
    tmrReset.Enabled = True
End Sub

Public Function StopDrawing() As Boolean
    StopDrawing = blnDraw
End Function

Private Sub Form_Load()
'Start Program
    Call Init 'initialize various variables, etc....
    Call CheckDirectories 'Check if our ini file exists
    Call LoadSettings 'Load any saved settings from the ini file
    Call DrawGraphLines 'Draw the compartive lines, of graph
End Sub

Private Sub Init()
    dblZoomAmount = 1 'Inital zoom amount of image 1=100%
    cboOptions.ListIndex = 0 'display first frame "Input Height"
    lstRGBH.ListIndex = 3 'Start list item with 3rd item on list, "Defualt"
    lstRGBH.Selected(lstRGBH.ListIndex) = True 'Check "Default"
    cboColorStyle.ListIndex = 0 'Start combo box with 1st item, "Default"
    cboRotOptions.ListIndex = 0 'Start combo box with 1st item, "Areial"
    cboGraph.ListIndex = 1 'Start combo box with 1st item, "Dots"
    cboDS.ListIndex = 0 ''Start combo box with 1st item, "Dots"
    HMDrawn = False 'Height-map has not been drawn
    ImageLoaded = False 'An image has not been drawn
    'Display canvass w/h on status bar, at bottom of form
    SB1.Panels(1).Text = "Canvass Size:(" & val(txtCanSize(0).Text) & "," & val(txtCanSize(1).Text) & ")" 'Show w/h
    picSample.BackColor = vbBlack 'Set sample picture background to black
End Sub

Private Sub Form_Resize()
'Resize,move, various controls on the form
    'Reposition some of the controls
    Call MoveControls(Me, picHMViewArea, picHMDest, picTop, picLeft, SB1, hsCScroll, vsCScroll, cmdCanvassCenter)
    Call TopLeftBottomRight 'Display:Top-Left(x,y)/Bottom-Right(x,y) of viewed canvass
    Call cmdDefault2_Click 'Set default canvass size
    Call cmdCanvassCenter_Click
    'Display view area width/height
    lblVAXY.Caption = picHMViewArea.Width & "," & picHMViewArea.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call mnuExit_Click 'clear memory and exit
End Sub


Private Sub hsHO_Change()
'Height-Offset scroll bar from frame array fraOpt(0)
    lblHOSV.Caption = hsHO.Value 'Display height-offset value in label next to scroll
    Call SetHeight(CInt(hsHO.Value))
End Sub

Private Sub hsHO_Scroll()
'Height-Offset scroll bar from frame array fraOpt(0)
    lblHOSV.Caption = hsHO.Value 'Display height-offset value in label next to scroll
End Sub

Private Sub hsRotation_Change()
'Height-Map-Rotation scroll bar from frame array fraOpt(1)
    Call Rotation 'Rotation value
End Sub

Private Sub hsRotation_Scroll()
'Height-Map-Rotation scroll bar from frame array fraOpt(1)
    Call Rotation 'rotation value
End Sub

Private Sub Rotation()
'Height-Map-Rotation scroll bar from frame array fraOpt(1)
    Const PI As Double = 3.14159265358979
    Const Rad As Double = PI / 180 'radians
    Dim newX As Single
    Dim newY As Single
    Dim nIndex As Integer 'array index that holds the red label
    
    newX = (shapePI.Width * 0.5) * Cos(hsRotation.Value * Rad) 'Calculate new X point
    newY = (shapePI.Width * 0.5) * Sin(hsRotation.Value * Rad) 'Calculate new Y point
    
    lAngle.X2 = lAngle.X1 + newX 'assign angled x
    lAngle.Y2 = lAngle.Y1 - newY 'assign angled y
    lblRotate.Caption = hsRotation.Value 'Display height-map rotation value in label next to scroll
    
    For nIndex = 0 To optDeg.Count - 1 'cycle option degrees
        If hsRotation.Value = Left(optDeg(nIndex).Caption, 3) Then 'value same as option caption
            optDeg(nIndex).ForeColor = &HFF0000    'Highlight it red
            optDeg(nIndex).Value = True 'set it's true
        Else
            optDeg(nIndex).ForeColor = vbBlack  'Highlight it black
            optDeg(nIndex).Value = False 'set it's false
        End If
    Next nIndex
    'Draw the example rotated image
    modRotatePIExample.RotatePIPicture picPiSource, picPiDest, -CDbl(hsRotation.Value)
End Sub

Private Sub hsScrScroll_Scroll()
    Call hsScrScroll_Change
End Sub

Private Sub hsX_Change()
'Scroll the color value array, via x values
    txtX.Text = hsX.Value
    Call ColorArrayValues(hsX.Value, hsY.Value)
End Sub

Private Sub hsX_Scroll()
'Scroll the color value array, via x values
    Call hsX_Change 'dispaly array values then graph
End Sub

Private Sub hsY_Change()
'Scroll the color value array, via y values
    txtY.Text = hsY.Value 'display current y value
    Call ColorArrayValues(hsX.Value, hsY.Value) ' dispaly array values then graph
End Sub

Private Sub hsY_Scroll()
'Scroll the color value array, via y values
    Call hsY_Change
End Sub

Private Sub ColorArrayValues(ByVal intRow As Integer, ByVal intCol As Integer)
    Dim intX As Integer 'X index of graph, or strXYZ array row postion
    Dim intY As Integer 'Y index of graph, or strXYZ array colum postion
    Dim lngColor As Long 'Color from color arrays, lngR,lngG,lngB
    intX = val(intRow) 'assign current x valued, from hsX scroll bar
    intY = val(intCol) 'assign current y valued, from hsY scroll bar
    If ImageLoaded = True Then 'if image is loaded then proceed
        If intX >= 0 And intX <= intGraphXMax Then 'x is not less than 0
            If intY >= 0 And intY <= intGraphYMax Then 'y is not less than 0
                lblListValue.Caption = strXYZ(intX, intY) 'display value at x,y from array
                Call ReadArray(strXYZ(), intX, intY) 'Break up info from array
                
                'show colors stored in rgb arrays, for current x,y point
                pColorMap.BackColor = lngColorA(intX, intY)
                picOriginalColor.BackColor = RGB(lngRA(intX, intY), lngGA(intX, intY), lngBA(intX, intY))
                shpFollow.FillColor = pColorMap.BackColor 'set circle fill color
                'lngGraphColorA(intX, intY) = pColorMap.BackColor
                'display colors values stored in rgb arrays, for current x,y point
                lblColorAt.Caption = "(" & lngRA(intX, intY) & "," & lngGA(intX, intY) & "," & lngBA(intX, intY) & ")"
                'set postion of x,y lines that point to current point on heightmap
                Call SetLineLocation(CInt(strX), CInt(strY) - CInt(strZ))
                'Draw Graph points
                Call ColorGraph(intY)
                hsY.Value = val(txtY.Text) 'set current hsY value
                hsX.Value = val(txtX.Text) 'set current hsY value
                lblCGX.Caption = hsX.Value 'Display current x postion on graph
            End If
        End If
    End If
    
End Sub

Private Function ReadArray(ByRef strA() As String, x As Integer, y As Integer)
    Dim strArray() As String 'array filled with color string
    Dim intC As Integer 'comma count
    Dim intI As Integer 'index position of string
    Dim intL As Integer 'string length
    Dim strString As String 'the given string
    Dim strTmp As String 'temp holding string
    
    strString = strA(x, y) 'get the string
    intL = Len(strString) 'get the length of the string
    
    ReDim strArray(0 To intL) 'Redim array to length of string
    
    For intI = 1 To intL 'Cycle letters on string
        strArray(intI) = Mid(strString, intI, 1) 'fill array with string
    Next intI
    
    For intI = 0 To intL 'cycle string array
        strTmp = strTmp & strArray(intI) 'build a string
        If strArray(intI) = "," And intC = 0 Then 'if 1st comma is reached
            strX = Left(strTmp, Len(strTmp) - 1) 'extract red number
            strTmp = "" 'reset string to nothing
        End If
        If strArray(intI) = "," And intC = 1 Then 'if 2nd comma is reached
            strY = Left(strTmp, Len(strTmp) - 1) 'extract green number
            strTmp = "" 'reset string to nothing
        End If
        If strArray(intI) = "," Then intC = 1 'after 1st comma is reached set to 1
    Next intI
    strZ = strTmp 'remainder of cycle string is blue
    
End Function

Private Sub ColorGraph(ByRef y As Integer)
 
    Dim dblSRH As Double
    Dim dblVARW As Double
    Dim dblGraphX As Double
    Dim dblGraphY As Double '
    Dim lngAverage As Long 'average of the colors for point we are working on
    Dim intI As Integer 'Index to start drawing points
    Dim lResult As Long 'Result:did the API succeed, fail=zero, succeed=non-zero
    pGraph.Cls 'clear old graph
    Dim intlx As Long
    Dim intly As Long
    Call DrawGraphLines 'Draw the compartive lines, of graph
    
    dblVARW = (pGraph.Width / intGraphXMax) 'Percent view area width is to source width
    dblSRH = (pGraph.Height / 255) 'percent source height is to view area height

    For intI = 0 To hsX.Value
        
        'Calculate average color, from rgb arrays, at given x and y
        lngAverage = (lngRA(intI, y) + lngGA(intI, y) + lngBA(intI, y)) / 3
        'Calculate and store the height or graph y point, at given x
        intGraphYArray(intI) = pGraph.Height - (lngAverage * dblSRH)
        dblGraphY = intGraphYArray(intI) 'Get the y point at given x
        dblGraphX = intI * dblVARW 'Calculate the x point, scaled to graph
   
        Select Case cboGraph.ListIndex 'Draw the points (x,y) on the graph
            Case 0 'Dots
                lResult = SetPixelV(pGraph.hdc, CLng(dblGraphX), CLng(dblGraphY), vbBlue) '&HFFFF&)
                  Case 1 'Line
                If intlx = 0 Then
                    intlx = CLng(dblGraphX) 'new line needs to start, x's have cycled
                    intly = CLng(dblGraphY) 'new line needs to start, x's have cycled
                End If
                pGraph.Line (intlx, intly)-(CLng(dblGraphX), CLng(dblGraphY)), vbBlue   '&HFFFF&
                intlx = CLng(dblGraphX) 'new line needs to start, x's have cycled
                intly = CLng(dblGraphY)
            Case 2 'Bars
                lResult = SetPixelV(pGraph.hdc, CLng(dblGraphX), CLng(dblGraphY), vbBlue) '&HFFFF&)
                pGraph.Line (CLng(dblGraphX), CLng(dblGraphY + 2))-(dblGraphX, pGraph.Height), lngColorA(intI, y) '&HFFFF&
                'pGraph.Refresh
        End Select
    Next intI
    
End Sub

Private Sub DrawGraphLines()
'Draw the compartive lines of x for the the graph
    pGraph.DrawStyle = 2 'Dots
    'top
    pGraph.Line (0, 0)-(pGraph.Width, 0), vbBlack
    '1/4 from top, or 3/4 from bottom
    pGraph.Line (0, pGraph.Height * 0.25)-(pGraph.Width, pGraph.Height * 0.25), vbBlack
    '3/4's from top, or 1/4 from bottom
    pGraph.Line (0, pGraph.Height * 0.75)-(pGraph.Width, pGraph.Height * 0.75), vbBlack
    'middle
    pGraph.Line (0, pGraph.Height * 0.5)-(pGraph.Width, pGraph.Height * 0.5), vbBlack
    'bottom
    pGraph.Line (0, pGraph.Height - 1)-(pGraph.Width, pGraph.Height - 1), vbBlack
    pGraph.DrawStyle = 0 'Solid
End Sub

Private Sub SetLineLocation(ByVal intX As String, ByVal intY As String)
    If lblListValue.Caption = "0" Then Exit Sub
    yL.X1 = intX
    yL.X2 = intX
    xL.Y1 = intY
    xL.Y2 = intY
    xL.X2 = intX
    yL.Y2 = intY
    shpFollow.Move intX - shpFollow.Width * 0.5, intY - shpFollow.Height * 0.5
    If chkShowPos.Value = 1 Then
        If HMDrawn = True Then
            xL.Visible = True
            yL.Visible = True
            shpFollow.Visible = True
        End If
    End If
End Sub

Private Sub lstRGBH_ItemCheck(Item As Integer)
    Dim intI As Integer 'list index
    
    For intI = 0 To lstRGBH.ListCount - 2 'cycle list
        If lstRGBH.Selected(intI) = True Then 'if red, green or blue is checked
            lstRGBH.Selected(lstRGBH.ListCount - 1) = False 'then uncheck default
        End If
    Next intI
    
    If Item = lstRGBH.ListCount - 1 Then 'if item is Default
        For intI = 0 To lstRGBH.ListCount - 2 'cycle list
            lstRGBH.Selected(intI) = False 'then uncheck all but default
        Next intI
    End If
    
    'If r,g,b are unchecked the check default
    If lstRGBH.Selected(0) = False And lstRGBH.Selected(1) = False And lstRGBH.Selected(2) = False Then
        lstRGBH.Selected(lstRGBH.ListCount - 1) = True 'check default
    End If
    
End Sub

Private Sub mnuDB_Click()
'Send image to microsoft paint editor
    'If ImageLoaded = True Then Call OpenEditor(Me, GetPath)
End Sub

Private Sub mnuAbout_Click()
'Show about form
    frmAbout.Show
End Sub

Private Sub mnuExit_Click()
'Clears memory of all image boxes, and unloades it'self from memory.
'Ends program.
    'MsgBox "Releasing Memory Used.", vbInformation, "Clearing Memory"
    picSrcImage.Cls 'clear source image
    picSrcDest.Cls 'clear the image
    picHMDest.Cls 'clear the image
    picSample.Cls 'clear the image
    picPiSource.Cls 'clear the image
    picPiDest.Cls 'clear the image
    Set picHMDest = Nothing 'clear memory
    Set picSrcImage = Nothing 'clear memory
    Set picSrcDest = Nothing 'clear memory
    Set picSample = Nothing 'clear memory
    Set picPiSource = Nothing 'clear memory
    Set picPiDest = Nothing 'clear memory
    If frmCapture.Visible = True Then Call frmCapture.ExitForm 'clear memory
    Unload Me 'clear form from memory
    End
End Sub

Private Sub mnuFileOptions_Click(Index As Integer)
    Select Case Index
        Case 1 'open
            Call OpenImage
        Case 3 'save
            Call SavePic(picHMDest, CD1)
    End Select
End Sub

Private Sub mnuHistory_Click(Index As Integer)
'Reads the recorderd path that is associated with the history menu item, from the ini file.
'If the image path is not null, it will load the image from the path read from the ini file.
'It will then test if the image loaded succesfully.

Dim strImagePath As String 'Image Path

    Call CheckUncheck(2, mnuHistory, Index, "Checked") 'Check and uncheck menu items
    strImagePath = ReadTheINI("Paths", "History" & Index, "") 'Get stored image path
    
    If strImagePath <> "" Then 'Test: is there an actual image path
        Me.MousePointer = 11 'hour glass
        ImageLoaded = LoadImage(strImagePath, picSrcImage) 'Load the source image/return if it loaded
        If ImageLoaded = True Then Call DrawImage(strImagePath) 'If imaged loaded successfully, draw the image
    End If
    Me.MousePointer = 0 'arrow
End Sub

Private Sub OpenImage()
'Reads the last recorderd path that the common dialog box was at, from the ini file. If no path exsists
'it will assign the local app.path. Opens a common dialog box "CD1", to allow the user
'to choose an image to load. If the image path is not null. Then cancel was not selected
'from the common dialog box. The path to where the imaged was loaded from, then becomes
'the last recordered path of the common dialog box, and is written to the ini file. Then
'it will load the image chosen from the common dialog box. It will then test if the image
'loaded succesfully.

Dim strImagePath As String 'Image Path
Dim CDPath As String 'Common dialog path

    CDPath = ReadTheINI("Paths", "CDPath", CDPath) 'Get last past of ComDial box, default=app path
    strImagePath = GetImage(strExt, CDPath, CD1) 'Get image, returns it's path
    
    If strImagePath <> "" Then 'Cancel was not selected write the path of cd
        Me.MousePointer = 11 'hour glass
        WriteToINI "Paths", "CDPath", GetDir(strImagePath)
        ImageLoaded = LoadImage(strImagePath, picSrcImage) 'Load the source image/return if it loaded
        If ImageLoaded = True Then Call DrawImage(strImagePath) 'If imaged loaded successfully, draw the image
    End If
    Me.MousePointer = 0 'arrow
End Sub

Private Sub DrawImage(ByRef strPath As String)
'Prep and draw image
    modComDialog.SetPath strPath
    Call NewHeightMap(strPath) 'prepare for a new heightmap image
    Call cmdCanvassCenter_Click 'recenter canvass
    Call PrepImage(Me, 1, picSrcImage, picSrcViewArea, picSrcDest) 'Prepare and draw the image
    Call ImageDimenstions 'Display image dimenstions
    Call History(strPath) 'Calculate history menu names and paths
    modPrepImage.SampleImage picSrcImage, picPiSource 'draw pi source image
    Call Rotation 'draw picPiSource image to picPiDest at initial rotation of zero
End Sub

Private Sub NewHeightMap(ByRef strPath As String)
'prepare for a new heightmap image
    cmdDraw.Enabled = True ' enable the ability to draw HM
    picSample.Cls 'clear old heightmap sample image
    picHMDest.Cls 'clear old height map
    pg.Value = 0 'Reset status bar
    frmMain.lblPixelCycles.Caption = "(0,0)"
    frmMain.lblPerc.Caption = "0%" 'percent completed
    fraTheImage.Caption = "The Image: " & ExtractName(strPath) 'Assign image name to frame caption
    Me.Caption = "DDDV"
    chkShowPos.Value = 0
    lblHP.Caption = 0
    lblLP.Caption = 0
End Sub

Private Sub mnuVI_Click(Index As Integer)
'Menu ViewImages as "Best-Fit", "Normal", "FitToWidth".
'Calls optImageStyle to determine, how to view the image when loaded.

    Call optImageStyle_Click(Index)
End Sub

Private Sub mnuZ_Click(Index As Integer)
    Select Case Index
        Case 0 'Zoom 200 percent
            dblZoomAmount = 2
        Case 1 'Zoom 175 percent
            dblZoomAmount = 1.75
        Case 2 'Zoom 150 percent
            dblZoomAmount = 1.5
        Case 3 'Zoom 100 percent
            dblZoomAmount = 1.25
        Case 4
            dblZoomAmount = 1
        Case 5 'Zoom 75 percent
            dblZoomAmount = 0.75
        Case 6 'Zoom 50 percent
            dblZoomAmount = 0.5
        Case 7 'Zoom 25 percent
            dblZoomAmount = 0.25
        Case 8 'normal/stretch
            dblZoomAmount = 1
    End Select
   
    
    If ImageLoaded = True Then 'If image is loaded then prepimage and redraw it
        
        Call PrepImage(Me, dblZoomAmount, picSrcImage, picSrcViewArea, picSrcDest) 'Prepare and draw the image
        Call ImageDimenstions 'Display image dimenstions
    End If
End Sub

Private Sub optDeg_Click(Index As Integer)
    Dim nIndex As Integer 'array index that holds the red label

    hsRotation.Value = CInt(optDeg(Index).Caption)
    optDeg(Index).ForeColor = &HFF0000
    For nIndex = 0 To optDeg.Count - 1
        If optDeg(nIndex).Value <> True Then
            optDeg(nIndex).ForeColor = vbBlack
        End If
    Next nIndex
End Sub

Private Sub optImageStyle_Click(Index As Integer)
'Determines how the loaded image will be shown(Best-Fit, Normal, FitToWidth).
'Check and uncheck which mnuVI(menu View Image) as, (Best-Fit, Normal, FitToWidth)
    Call CheckUncheck(2, optImageStyle, Index, "Value")
    Call CheckUncheck(2, mnuVI, Index, "Checked") 'Check and uncheck menu items
    
    If ImageLoaded = True Then 'If image is loaded then prepimage and redraw it
        mnuZ_Click (4)
    End If
End Sub

Public Sub cmdCanvassCenter_Click()
'Center picHMDest to (x,y) cordinates of the picture box picHMViewArea
    'Call UpdateScrolls
    If picHMDest.Width > picHMViewArea.Width Then hsCScroll.Value = hsCScroll.Max / 2 '(picHMDest.Width - lblDraW2) / 2
    If picHMDest.Height > picHMViewArea.Height Then vsCScroll.Value = vsCScroll.Max / 2 '(picHMDest.Height - lblDrawH2) / 2
End Sub

Public Sub cmdCenter_Click()
'Center picSrcDest to (x,y) (0,0) cordinates of the picture box picViewArea
    DoEvents
    hsScrScroll.Value = 0 'Center image to left to 0
    vsSrcScroll.Value = 0 'Center image to top to 0
End Sub


Private Sub picBG_Click()
'Set background color, for picHMDest from picture box, "picBG"
'located on frame array fraOpt(3), "Output Colors"
On Error GoTo HandleError

    CD1.CancelError = True
    MsgBox "Warning. If there is a Height-Map drawn. It will be cleared if you choose a new color.", vbExclamation
    CD1.ShowColor 'Show color pallete
    picBG.BackColor = CD1.Color 'Assign color chosen
    picHMDest.BackColor = CD1.Color 'Set background color
HandleError:
    If Err.Number = 32755 Then 'cd cancel was selected
        Err.Clear 'clear cancel error
    End If
End Sub

Private Sub picDefineColor_Click()
'Set color of the height-map to be drawn as. For picHMDest from picture box, "picDefineColor"
'located on frame array fraOpt(3), "Output Colors"
On Error GoTo ResolveError
    optOPC(1).Value = True 'check custom colors option
    CD1.CancelError = True
    CD1.ShowColor 'Show color pallete
    picDefineColor.BackColor = CD1.Color 'Assign color chosen
    Exit Sub
ResolveError:
    Exit Sub
End Sub

Private Sub picLeft_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Screen.MousePointer = 0
End Sub


Private Sub picSample_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then 'if left mouse button
        If hsCScroll.Enabled = True Or vsCScroll.Enabled = True Then
            Call MoveSampleSquare(x, y) 'Move sample image sqaure
        End If
    End If
End Sub

Private Sub MoveSampleSquare(ByRef x As Single, ByRef y As Single)
'Move sample image sqaure according to width/height of heightmap and sampleimage x and y cursor position
'Assign hs/vs scroll values
    Dim sglMX As Single 'Square x move
    Dim sglMY As Single 'Square y move
    Dim dblHR As Double 'hor ratio
    Dim dblVR As Double 'vert ratio
    
    sglMX = x - Shape4.Width * 0.5 'calc Square x move
    sglMY = y - Shape4.Height * 0.5 'calc Square y move

    dblVR = vsCScroll.Max / picSample.ScaleHeight 'calc hor ratio
    dblHR = hsCScroll.Max / picSample.ScaleWidth 'calc vert ratio
    
    'Make sure sqaure x left is not less than 0 and not greater than samplepic width
    If sglMX >= 0 And sglMX + Shape4.Width < picSample.ScaleWidth Then Shape4.Left = sglMX
    'Make sure sqaure y top is not less than 0 and not greater than samplepic height
    If sglMY >= 0 And sglMY + Shape4.Height < picSample.ScaleHeight Then Shape4.Top = sglMY
    'Make sure ratio x HorValue is not less than 0 and not greater than hsCScroll.Max
    If Int((x * dblHR)) >= 0 And Int((x * dblHR)) <= hsCScroll.Max Then hsCScroll.Value = Int(x * dblHR)
    'Make sure ratio y HorValue is not less than 0 and not greater than vsCScroll.Max
    If Int((y * dblVR)) >= 0 And Int(y * dblVR) <= vsCScroll.Max Then vsCScroll.Value = Int(y * dblVR)
    
End Sub

Private Sub picViewArea_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Screen.MousePointer = 0 'default arrow mouse pointer
End Sub

Private Sub picSrcDest_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    
    If Button = 1 Then
        If hsScrScroll.Enabled = True Then hsScrScroll.Value = hsScrScroll.Value + CurX - x
        If vsSrcScroll.Enabled = True Then vsSrcScroll.Value = vsSrcScroll.Value + CurY - y
        DoEvents
    End If
    'Theres no sence in constantly writing this while moveing the mouse
    If SB1.Panels(2).Text <> "Tip: Hold left mouse button to move image!" Then
        SB1.Panels(2).Text = "Tip: Hold left mouse button to move image!"
        SB1.Panels(2).Visible = True
    End If
    If HMDrawn = True And chkFollowMode.Value = 1 Then
        txtX.Text = x
        txtY.Text = y
        Call ColorArrayValues(CInt(x), CInt(y))
    End If
End Sub

Private Sub picHMDest_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
    lblMouseCords.Caption = "(" & CInt(x) & "," & CInt(y) & ")" 'Display where the mouse is
    If Button = 1 Then
        If hsCScroll.Enabled = True Then hsCScroll.Value = hsCScroll.Value + CurX - x
        If vsCScroll.Enabled = True Then vsCScroll.Value = vsCScroll.Value + CurY - y
    End If
     
    If Button = vbRightButton Then Call StopShape(x, y)
    
    'Theres no sence in constantly writing this while moveing the mouse
    If SB1.Panels(2).Text <> "Tip: Mouse buttons, (Left to move image) (Right copies a selection)" Then
        SB1.Panels(2).Text = "Tip: Mouse buttons, (Left to move image) (Right copies a selection)"
        SB1.Panels(2).Visible = True
    End If
End Sub
'Mouse Down Events-------------------------------------------------------------
Private Sub picHMDest_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    CurX = x
    CurY = y
    If Button = vbLeftButton And cmdCanvassCenter.Enabled = True Then Screen.MousePointer = 5
    If Button = vbRightButton And HMDrawn = True Then
        Call StartShape(shCapture, x, y, picHMDest)
    End If
    
    'if x>=
End Sub

Private Sub picSrcDest_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    CurX = x
    CurY = y
    If optImageStyle(1).Value = True Then Screen.MousePointer = 5
End Sub
'End Mouse Down Events*********************************************************

'Mouse Up Events---------------------------------------------------------------
Private Sub picHMDest_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Screen.MousePointer = 0
    If Button = vbRightButton Then Call DrawShapeArea(frmCapture.picSource)
    If Button = vbRightButton And ErrorNumber <> 91 Then
        frmCapture.Show
    End If
End Sub

Private Sub picSrcDest_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If optImageStyle(1).Value = True Then Screen.MousePointer = 0
End Sub
'End Mouse Up Events***********************************************************

'Scroll Bar Events-------------------------------------------------------------
Private Sub hs_Scroll()
'Loaded image horizontal scroll move left and right
    'hs_Change
End Sub

Private Sub hsCScroll_Scroll()
'Height-Map image horizontal scroll move left and right
    hsCScroll_Change
End Sub


Private Sub tmrReset_Timer()
    Call PreNewHeightMap
    Me.MousePointer = 0 'arrow
    HMDrawn = False
    Call NewHeightMap(GetPath)
    cmdStop.Enabled = False
    tmrReset.Enabled = False
End Sub

Private Sub txtCanSize_Change(Index As Integer)
'Display the canvass size, on labels located on frame array fraOpt(6)"Canvass Size Width/Height"
    Select Case Index
        Case 0 'Width
            lblCanXY(0).Caption = "X " & val(txtCanSize(0).Text)
        Case 1 'Height
            lblCanXY(1).Caption = "Y " & val(txtCanSize(1).Text)
    End Select
End Sub

Private Sub vs_Scroll()
'Loaded image vertical scroll move up and down
    'vs_Change
End Sub

Private Sub vsCScroll_Scroll()
'Height-Map vertical scroll move up and down
    vsCScroll_Change
End Sub

Private Sub hsScrScroll_Change()
'Loaded image horizontal move left and right
    picSrcDest.Left = 0 - hsScrScroll.Value
End Sub

Private Sub hsCScroll_Change()
'Height-Map image horizontal move left and right
    picHMDest.Left = 0 - hsCScroll.Value
    Call TopLeftBottomRight
    Call MoveSquare
End Sub

Private Sub vsSrcScroll_Change()
'Loaded image vertical move up and down
    picSrcDest.Top = 0 - vsSrcScroll.Value
End Sub

Private Sub vsCScroll_Change()
'Height-Map image vertical move up and down
    picHMDest.Top = 0 - vsCScroll.Value
    Call TopLeftBottomRight
    Call MoveSquare
End Sub

Private Sub MoveSquare()
'Move sample image sqaure according to hor/vert scroll values
    Dim dblHR As Double 'hor x ratio
    Dim dblVR As Double 'vert y ratio
    Dim sglMX As Double 'move x amount
    Dim sglMY As Double 'move y amount
    
    dblVR = vsCScroll.Max / picSample.ScaleHeight 'calc hor x ratio
    dblHR = hsCScroll.Max / picSample.ScaleWidth 'cal vert y ratio
    If dblHR > 0 Then  'prevent division by zero
        sglMX = (hsCScroll.Value / dblHR) - Shape4.Width * 0.5 'move x amount
    End If
    If dblVR > 0 Then
        sglMY = (vsCScroll.Value / dblVR) - Shape4.Height * 0.5 'move y amount
    End If
    
    'Make sure sqaure x left is not less than 0 and not greater than samplepic width
    If sglMX >= 0 And sglMX + Shape4.Width < picSample.ScaleWidth Then Shape4.Left = sglMX
    'Make sure sqaure y top is not less than 0 and not greater than samplepic height
    If sglMY >= 0 And sglMY + Shape4.Height < picSample.ScaleHeight Then Shape4.Top = sglMY

End Sub

Private Function LoadImage(path As String, picBox As PictureBox) As Boolean
'Loades an image to the assigned picture box.
'Receives path = path to image, picBox = passed picture box to load image to.
'Returns true/successfull, false/failed

On Error GoTo HandleError
    picBox.Cls
    Set picBox.Picture = LoadPicture(path) 'load the source image
    LoadImage = True 'the image is loaded, set to true
    HMDrawn = False ' this image has not been drawn to heightmap
HandleError:
    If Err.Number = 481 Then 'Invalid Picture
        LoadImage = False 'image was not loaded, set to false
        picSrcDest.Visible = False 'Hide the image
        MsgBox "Invalid Picture", vbCritical, "Cannot Render Image"
    End If
End Function

Private Sub ImageDimenstions()
'Display source image width/height and source image scaled width/height
    lblSrcW.Caption = Int(GetImageScaleW) 'Get image width, located in modPrepImage
    lblSrcH.Caption = Int(GetImageScaleH) 'Get image height, located in modPrepImage
    
End Sub

Private Sub History(strPath As String)
'Assign names to the mnuHistory menus, from images chosen.
    Dim intI As Integer 'Index of current mnuHistory

    Call CheckUncheck(2, mnuHistory, 0, "Checked") 'check/uncheck history menu items
    'Place history captions, filter out old numbering and add new number
    'If 4th char of menu <> null, then caption = number + previous menu caption.
    'Exp: Caption = "3)" do not add, or Caption = "3) CraterLake.jpg" then add
    If Mid(Mid(mnuHistory(1).Caption, 3), 4, 1) <> vbNullString Then mnuHistory(2).Caption = "3)" & Mid(mnuHistory(1).Caption, 3) 'caption 3 becomes 2
    If Mid(Mid(mnuHistory(0).Caption, 3), 4, 1) <> vbNullString Then mnuHistory(1).Caption = "2)" & Mid(mnuHistory(0).Caption, 3) 'caption 2 becomes 1
    mnuHistory(0).Caption = "1) " & ExtractName(strPath) 'caption 1 becomes new image
    
    'Save and cycle paths to menu history items
    WriteToINI "Paths", "History2", ReadTheINI("Paths", "History1", "") 'path 3 becomes 2
    WriteToINI "Paths", "History1", ReadTheINI("Paths", "History0", "") 'path 2 becomes 1
    WriteToINI "Paths", "History0", strPath 'path 1 becomes new path

    For intI = 0 To 2 'Cycle history items
        If mnuHistory(intI).Caption <> vbNullString Then 'If an item is added
            mnuHistory(intI).Visible = True 'show the item
            mnu2.Visible = True 'show the spacer line between mnuexit and history item
        End If
    Next intI
    
End Sub

Private Sub CheckUncheck(intArrayCount As Integer, varItem As Variant, intIndex As Integer, strDo As String)
'Multipurpose routine to check/uncheck, true/false values, enable/disable, true/false visible. For
'larger than one index control arrays
'Receives intArrayCount = number of items on the array
'Receives varItem = option box, menu, button, etc
'Receives intIndex = index item of varItem that will be true, all others will be false
'Recieves strDo = what will be done, checked, value, enabled, visible.

    Dim ZArray() As Integer 'compare array
    Dim nIndex As Integer 'array index that holds the red label

    ReDim ZArray(intArrayCount)
    ZArray(intIndex) = -1 'assign item that is true
    
    Select Case strDo
        Case "Checked"
            For nIndex = 0 To UBound(ZArray)
                If ZArray(nIndex) <> -1 Then 'uncheck
                    varItem(nIndex).Checked = False 'Uncheck
                Else
                    varItem(nIndex).Checked = True 'check
                End If
            Next nIndex
        Case "Value"
            For nIndex = 0 To UBound(ZArray)
                If ZArray(nIndex) <> -1 Then 'uncheck
                    varItem(nIndex).Value = False 'false
                Else
                    varItem(nIndex).Value = True 'true
                End If
            Next nIndex
        Case "Enabled"
            For nIndex = 0 To UBound(ZArray)
                If ZArray(nIndex) <> -1 Then 'uncheck
                    varItem(nIndex).Value = False 'disabled
                Else
                    varItem(nIndex).Value = True 'enabled
                End If
            Next nIndex
        Case "Visible"
            For nIndex = 0 To UBound(ZArray)
                If ZArray(nIndex) <> -1 Then 'uncheck
                    varItem(nIndex).Visible = False 'Hide
                Else
                    varItem(nIndex).Visible = True 'show
                End If
            Next nIndex
    End Select
End Sub

Private Sub LoadSettings()
'Load various saved settings from the ini file.
    Dim intI As Integer 'Multipurpose index
    Dim strTmpName As String 'extracted name of saved history path
    
    'menu saved settings
    For intI = 0 To 2 'Display any saved history
        strTmpName = (intI + 1) & ") " & ExtractName(ReadTheINI("Paths", "History" & intI, "")) 'extract any saved history names
        'If the name is not null then add it to history
        If Mid(strTmpName, 4, 1) <> vbNullString Then  'If an item is added
            mnuHistory(intI).Caption = strTmpName
            mnuHistory(intI).Visible = True 'show the item
            mnu2.Visible = True 'show the spacer line between mnuexit and history item
        End If
    Next intI
    
    'Check uncheck image style menu items, according to option image styles
    For intI = 0 To optImageStyle.UBound 'Cycle image option styles
        'option is true, check corrosponding mnu item
        If optImageStyle(intI).Value = True Then mnuVI(intI).Checked = True
        'option is false, uncheck corrosponding mnu item
        If optImageStyle(intI).Value = False Then mnuVI(intI).Checked = False
    Next intI
End Sub

Private Function TopLeftBottomRight()
'Display the Top-Left(x,y) locational coordinates of the picHMDest, inside the picHMViewArea
'Display the Bottom-Right (x,y) locational coordinates of the picHMDest, inside the picHMViewArea
    lblTopLeft.Caption = hsCScroll.Value & "," & vsCScroll.Value 'show top-left
    lblBottomRight.Caption = hsCScroll.Value + picHMViewArea.Width & "," & vsCScroll.Value + picHMViewArea.Height 'show bottom right
End Function

Private Sub vsSrcScroll_Scroll()
vsSrcScroll_Change
End Sub
