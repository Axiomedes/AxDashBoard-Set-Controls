VERSION 5.00
Object = "*\AAxFramework.vbp"
Begin VB.Form Form7 
   Caption         =   "Demo: AxDashAnimLabel"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12390
   LinkTopic       =   "Form7"
   ScaleHeight     =   491
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   826
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrShiftDemo 
      Interval        =   800
      Left            =   120
      Top             =   120
   End
   Begin AxDashboardSet.AxDashAnimLabel AnimLabel1 
      Height          =   1875
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3307
      BackColor1      =   16777215
      BackColor2      =   14737632
      BorderColor     =   986895
      EffectFading    =   -1  'True
      Caption1        =   "Sales Trend"
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1Color   =   8388608
      Caption2        =   "$12,450"
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2Color   =   15832166
      Caption3        =   "Last 10 periods"
      BeginProperty Caption3Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption3Color   =   8421504
      GraphLineColor  =   15832166
      GraphBackColor  =   15832166
      GraphPointColor =   32768
      GraphFillOpacity=   30
      GraphGridLines  =   -1  'True
      GraphGridColor  =   8421504
      AnimSpeed       =   50
   End
   Begin AxDashboardSet.AxDashAnimLabel AnimLabel2 
      Height          =   1875
      Left            =   6240
      TabIndex        =   1
      Top             =   240
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   3307
      BackColor1      =   460551
      BackColor2      =   460551
      BorderColor     =   986895
      Caption1        =   "Visitors (Real-Time)"
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2        =   "0"
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2Color   =   10741301
      Caption3        =   "Mode: Shift (live data)"
      BeginProperty Caption3Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption3Color   =   11842760
      GraphLine       =   0
      GraphMatrix     =   "50,50,50,50,50,50,50,50,50,50,50,50"
      GraphLineColor  =   10741301
      GraphBackColor  =   10741301
      GraphFillOpacity=   25
      AnimEnabled     =   0   'False
      AnimMode        =   1
      AnimSpeed       =   800
      AnimStep        =   3
   End
   Begin AxDashboardSet.AxDashAnimLabel AnimLabel3 
      Height          =   1875
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   3307
      BackColor1      =   14737632
      BackColor2      =   14737632
      BorderColor     =   986895
      Caption1        =   "Server Load"
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1Color   =   4210752
      Caption2        =   "Normal"
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2Color   =   6514417
      Caption3        =   "Loop mode - Bars"
      BeginProperty Caption3Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption3Color   =   128
      GraphLine       =   2
      GraphMatrix     =   "10,30,70,45,80,55,25,65,40,90,35,60"
      GraphStyle      =   1
      GraphLineColor  =   6514417
      GraphBackColor  =   6514417
      GraphFillOpacity=   45
      AnimSpeed       =   80
      AnimStep        =   1
   End
   Begin AxDashboardSet.AxDashAnimLabel AnimLabel4 
      Height          =   1875
      Left            =   6240
      TabIndex        =   3
      Top             =   2400
      Width           =   5910
      _ExtentX        =   10425
      _ExtentY        =   3307
      BackColor1      =   460551
      BackColor2      =   460551
      BorderColor     =   986895
      Caption1        =   "Network Rx"
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2        =   "0 Mbps"
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2Color   =   3462041
      Caption3        =   "Shift mode + Boxed icon"
      BeginProperty Caption3Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption3Color   =   11842760
      Boxed           =   -1  'True
      BoxedColor      =   3462041
      GraphMatrix     =   "40,40,40,40,40,40,40,40,40,40,40,40,40,40,40"
      GraphLineColor  =   3462041
      GraphBackColor  =   3462041
      GraphFillOpacity=   30
      AnimEnabled     =   0   'False
      AnimMode        =   1
      AnimSpeed       =   800
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Simula datos en tiempo real para los controles en modo amShift
Private shiftTick As Long

Private Sub Form_Load()
    shiftTick = 0
    ' Precarga valores base para los shift labels
    Dim i As Integer
    For i = 1 To 12
        AnimLabel2.PushValue 50
        AnimLabel4.PushValue 40
    Next i
    AnimLabel2.AnimEnabled = False
    AnimLabel4.AnimEnabled = False
End Sub

Private Sub tmrShiftDemo_Timer()
    ' Simular nuevo dato para controles en modo Shift
    Randomize Timer
    Dim newVal2 As Single, newVal4 As Single
    newVal2 = 30 + Rnd * 70
    newVal4 = 20 + Rnd * 75

    AnimLabel2.PushValue newVal2
    AnimLabel4.PushValue newVal4

    AnimLabel2.Caption2 = Format(newVal2 * 100, "#,##0")
    AnimLabel4.Caption2 = Format(newVal4 * 1.2, "0.0") & " Mbps"

    shiftTick = shiftTick + 1
End Sub
