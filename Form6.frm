VERSION 5.00
Object = "*\AAxFramework.vbp"
Begin VB.Form Form6 
   Caption         =   "Demo: AxDashGaugeLabel"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12390
   LinkTopic       =   "Form6"
   ScaleHeight     =   491
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   826
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrDemo 
      Interval        =   1500
      Left            =   120
      Top             =   120
   End
   Begin AxDashboardSet.AxDashGaugeLabel Gauge1 
      Height          =   3375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5953
      BackColor1      =   460551
      BackColor2      =   460551
      BorderColor     =   986895
      CornerCurve     =   10
      Caption1        =   "CPU Load"
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1Color   =   11842760
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption3        =   "Procesador"
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
      GaugeColorMode  =   2
      GaugeColor1     =   16739880
      GaugeColor2     =   2686974
      GaugeThickness  =   14
      GaugeColorWarning=   570602
   End
   Begin AxDashboardSet.AxDashGaugeLabel Gauge2 
      Height          =   3375
      Left            =   3840
      TabIndex        =   1
      Top             =   360
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5953
      BackColor1      =   460551
      BackColor2      =   460551
      BorderColor     =   986895
      CornerCurve     =   10
      Caption1        =   "Memory"
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1Color   =   11842760
      Caption2        =   "0 GB"
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption3        =   "RAM Usage"
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
      GaugeColorMode  =   1
      GaugeColor1     =   14239468
      GaugeColor2     =   10092390
      GaugeTrackColor =   5259344
      GaugeThickness  =   14
      GaugeNeedle     =   2
      GaugeAnimSpeed  =   6
   End
   Begin AxDashboardSet.AxDashGaugeLabel Gauge3 
      Height          =   3375
      Left            =   7440
      TabIndex        =   2
      Top             =   345
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5953
      BackColor1      =   460551
      BackColor2      =   460551
      BackAngle       =   45
      BorderColor     =   16777215
      BorderWidth     =   2
      CornerCurve     =   10
      Clickable       =   -1  'True
      EffectFading    =   -1  'True
      InitialOpacity  =   10
      Caption1        =   "Disk I/O"
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1Color   =   11842760
      Caption2        =   "0 MB/s"
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption3        =   "Write Speed"
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
      GaugeStyle      =   1
      GaugeColorMode  =   2
      GaugeColor1     =   6514417
      GaugeTrackColor =   4208720
      GaugeThickness  =   16
      GaugeNeedle     =   0
      GaugeAnimSpeed  =   10
   End
   Begin AxDashboardSet.AxDashGaugeLabel Gauge4 
      Height          =   3375
      Left            =   240
      TabIndex        =   3
      Top             =   3960
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5953
      BackColor1      =   460551
      BackColor2      =   460551
      BorderColor     =   986895
      CornerCurve     =   10
      Caption1        =   "Temperature"
      BeginProperty Caption1Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1Color   =   11842760
      Caption2        =   "0 C"
      BeginProperty Caption2Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption3        =   "GPU Temp"
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
      GaugeStyle      =   2
      GaugeColorMode  =   2
      GaugeColor1     =   6514417
      GaugeTrackColor =   4210768
      GaugeThresholdWarning=   50
      GaugeThresholdDanger=   75
      GaugeColorWarning=   570602
      GaugeAnimSpeed  =   7
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Valores simulados para demo
Private demoStep As Long

Private Sub Form_Load()
    demoStep = 0
    ' Inicializar con valores base
    Gauge1.GaugeValue = 45
    Gauge2.GaugeValue = 32
    Gauge3.GaugeValue = 18
    Gauge4.GaugeValue = 55
    UpdateCaptions
End Sub

Private Sub tmrDemo_Timer()
    ' Simular variacion de valores
    Dim v1 As Single, v2 As Single, v3 As Single, v4 As Single
    Randomize Timer
    v1 = 30 + Rnd * 60
    v2 = 20 + Rnd * 70
    v3 = 10 + Rnd * 80
    v4 = 35 + Rnd * 55

    Gauge1.GaugeValue = v1
    Gauge2.GaugeValue = v2
    Gauge3.GaugeValue = v3
    Gauge4.GaugeValue = v4
    UpdateCaptions
End Sub

Private Sub UpdateCaptions()
    Gauge1.Caption2 = Format(Gauge1.GaugeValue, "0") & "%"
    Gauge2.Caption2 = Format(Gauge2.GaugeValue * 0.32, "0.0") & " GB"
    Gauge3.Caption2 = Format(Gauge3.GaugeValue * 5.5, "0") & " MB/s"
    Gauge4.Caption2 = Format(Gauge4.GaugeValue * 1.2 + 30, "0") & " C"
End Sub
