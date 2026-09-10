VERSION 5.00
Begin VB.UserControl AxDashGaugeLabel 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   240
   ToolboxBitmap   =   "AxDashGaugeLabel.ctx":0000
   Begin VB.Timer tmrAnim 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   150
      Top             =   75
   End
   Begin VB.Timer tmrEffect 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   75
      Top             =   75
   End
End
Attribute VB_Name = "AxDashGaugeLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-UC-VB6-----------------------------
'UC Name  : AxDashGaugeLabel
'Version  : 1.00
'Editor   : David Rojas [AxioUK]
'Date     : 10/09/2026
'Desc     : Gauge / Medidor circular animado para Dashboards
'------------------------------------
Option Explicit

'--- Win32 API ---
Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTL) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

'--- GDI+ Core ---
Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal Token As Long)
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal mhDC As Long, ByRef mGraphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
'--- GDI Double Buffer ---
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long

'--- GDI+ Pen ---
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long

'--- GDI+ Brush ---
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal ARGB As Long, ByRef brush As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal brush As Long) As Long
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As Long, ByRef mLineGradient As Long) As Long

'--- GDI+ Path ---
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long

'--- GDI+ Arc (floating point) ---
Private Declare Function GdipDrawArc Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long

'--- GDI+ Ellipse ---
Private Declare Function GdipFillEllipse Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipDrawEllipse Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long

'--- GDI+ Line ---
Private Declare Function GdipDrawLineI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long

'--- GDI+ Text / Font ---
Private Declare Function GdipAddPathString Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFamily As Long, ByVal mStyle As Long, ByVal mEmSize As Single, ByRef mLayoutRect As RECTS, ByVal mFormat As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "GdiPlus.dll" (ByRef mNativeFamily As Long) As Long
Private Declare Function GdipSetStringFormatTrimming Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mTrimming As eStringTrimming) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatFlags Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mFlags As eStringFormatFlags) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As eStringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mAlign As eStringAlignment) As Long
Private Declare Function GdipDeleteStringFormat Lib "GdiPlus.dll" (ByVal mFormat As Long) As Long

'--- GDI+ Transform ---
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "GdiPlus.dll" (ByVal graphics As Long) As Long

'--- Private Types ---
Private Type RECTL
    Left   As Long
    Top    As Long
    Width  As Long
    Height As Long
End Type

Private Type RECTS
    Left   As Single
    Top    As Single
    Width  As Single
    Height As Single
End Type

Private Type POINTL
    X As Long
    Y As Long
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

Private Type GDIPLUS_FONTSTYLE
    FontStyle As Long
End Type

'--- Constants ---
Private Const SmoothingModeAntiAlias As Long = 4
Private Const UnitPixel As Long = 2
Private Const WrapModeTileFlipXY As Long = 3
Private Const FontStyleBold As Long = 1
Private Const FontStyleItalic As Long = 2
Private Const FontStyleUnderline As Long = 4
Private Const FontStyleStrikeout As Long = 8
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const PI As Double = 3.14159265358979

'--- Pi conversion ---
Private Const DEG2RAD As Double = PI / 180#

'--- GDI+ token ---
Private GdipToken As Long
Private hGraphics As Long
Private nScale As Double   ' DPI factor

'--- Private state variables ---
Private m_Enabled As Boolean
Private m_Visible As Boolean
Private m_Clickable As Boolean
Private m_Transparent As Boolean
Private m_Filled As Boolean
Private m_Opacity As Long
Private m_InitialOpacity As Long
Private m_EffectFade As Boolean

'--- Background ---
Private m_Color1 As OLE_COLOR
Private m_Color2 As OLE_COLOR
Private m_Angulo As Single
Private m_BorderColor As OLE_COLOR
Private m_BorderWidth As Long
Private m_CornerCurve As Long

'--- Defaults ---
Private m_def_Color1 As OLE_COLOR
Private m_def_Color2 As OLE_COLOR
Private m_def_Angulo As Single
Private m_def_ForeColor As OLE_COLOR

'--- Caption 1 (titulo) ---
Private m_Caption As String
Private m_Font As StdFont
Private m_ForeColor As OLE_COLOR
Private m_CaptionAlignH As eTextAlignH
Private m_CaptionAlignV As eTextAlignV

'--- Caption 2 (valor central) ---
Private m_Caption2 As String
Private m_Font2 As StdFont
Private m_ForeColor2 As OLE_COLOR
Private m_Caption2AlignH As eTextAlignH
Private m_Caption2AlignV As eTextAlignV

'--- Caption 3 (unidad / info) ---
Private m_Caption3 As String
Private m_Font3 As StdFont
Private m_ForeColor3 As OLE_COLOR
Private m_Caption3AlignH As eTextAlignH
Private m_Caption3AlignV As eTextAlignV

'--- Icon ---
Private m_IconCharCode As String
Private m_IconFont As StdFont
Private m_IconForeColor As OLE_COLOR
Private m_IconAlignH As eTextAlignH
Private m_IconAlignV As eTextAlignV

'--- Gauge properties ---
Private m_GaugeValue As Single       ' 0-100 (normalizado)
Private m_GaugeMin As Single         ' valor real minimo
Private m_GaugeMax As Single         ' valor real maximo
Private m_GaugeStyle As eGaugeStyle  ' estilo del arco
Private m_GaugeColorMode As eGaugeColorMode
Private m_GaugeColor1 As OLE_COLOR
Private m_GaugeColor2 As OLE_COLOR
Private m_GaugeTrackColor As OLE_COLOR
Private m_GaugeThickness As Long
Private m_ThresholdWarning As Single
Private m_ThresholdDanger As Single
Private m_GaugeColorWarning As OLE_COLOR
Private m_GaugeColorDanger As OLE_COLOR
Private m_GaugeNeedle As eGaugeNeedleStyle
Private m_GaugeNeedleColor As OLE_COLOR
Private m_GaugeAnimated As Boolean
Private m_GaugeAnimSpeed As Long

'--- Animation internal state ---
Private m_AnimValue As Single        ' valor visual actual (interpolado)

'--- Tag ---
Private m_Tag As String

'--- Events ---
Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

'===========================================================================
' PUBLIC METHODS
'===========================================================================

Public Sub CopyAmbient()
    UserControl.BackColor = UserControl.Ambient.BackColor
    Refresh
End Sub

Public Sub Refresh()
    Draw
End Sub

'===========================================================================
' CORE DRAWING
'===========================================================================

Private Sub Draw()
    Dim REC As RECTL
    Dim W As Long, H As Long
    Dim CX As Single, CY As Single
    Dim R As Single
    Dim arcStart As Single, arcFull As Single, arcValue As Single
    Dim gX As Single, gY As Single, gW As Single, gH As Single
    Dim gaugeZoneH As Long
    Dim capREC As RECTS, cap2REC As RECTS, cap3REC As RECTS, icoREC As RECTS
    Dim margin As Long

    With UserControl
        '--- Double Buffer: dibujar en DC oculto ---
        Dim memDC As Long, memBmp As Long, oldBmp As Long
        Const SRCCOPY As Long = &HCC0020
        memDC = CreateCompatibleDC(.hdc)
        memBmp = CreateCompatibleBitmap(.hdc, .ScaleWidth, .ScaleHeight)
        oldBmp = SelectObject(memDC, memBmp)
        GdipCreateFromHDC memDC, hGraphics
        GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias

        W = .ScaleWidth
        H = .ScaleHeight
        margin = 6

        '--- Background ---
        REC.Left = 0:     REC.Top = 0
        REC.Width = W - 1: REC.Height = H - 1
        gRoundRect hGraphics, REC, ARGB(m_Color1, 100), ARGB(m_Color2, 100), _
                   m_Angulo, m_BorderWidth, ARGB(m_BorderColor, 100), m_CornerCurve, m_Filled

        '--- Layout: gauge ocupa la zona superior, texto el resto ---
        gaugeZoneH = CLng(H * 0.72)

        '--- Gauge circle sizing ---
        Dim circDiam As Single
        Dim thickHalf As Single
        thickHalf = (m_GaugeThickness * nScale) / 2#

        ' El diametro respeta tanto el ancho como la altura disponible con margen seguro
        Dim availSz As Single
        availSz = IIf(W < gaugeZoneH, W, gaugeZoneH)
        circDiam = availSz - (margin * 4) - (thickHalf * 2)
        If circDiam < 16 Then circDiam = 16
        R = circDiam / 2#
        If R < 8 Then R = 8

        CX = W / 2#
        ' Desplazamos CY hacia abajo con margen superior suficiente para que el grosor del arco no se recorte
        CY = margin * 2 + thickHalf + R

        gX = CX - R: gY = CY - R
        gW = R * 2:  gH = R * 2

        '--- Arc parameters by style ---
        Select Case m_GaugeStyle
            Case gsArcDial     ' 270 grados (default)
                arcStart = 135
                arcFull = 270
            Case gsSemicircle  ' 180 grados
                arcStart = 180
                arcFull = 180
            Case gsFullCircle  ' 360 grados
                arcStart = 90
                arcFull = 360
        End Select

        arcValue = (m_AnimValue / 100) * arcFull

        '--- Draw Track (pista de fondo) ---
        DrawGaugeArc hGraphics, gX, gY, gW, gH, arcStart, arcFull, _
                     ARGB(m_GaugeTrackColor, 45), m_GaugeThickness * nScale

        '--- Draw value arc ---
        If arcValue > 0.1 Then
            DrawValueArc hGraphics, gX, gY, gW, gH, arcStart, arcValue, arcFull

            '--- Needle / indicator at arc tip ---
            DrawNeedle hGraphics, CX, CY, R, arcStart, arcValue
        End If

        '--- Caption areas ---
        ' Icon (arriba del valor)
        icoREC.Left = CX - (R * 0.5)
        icoREC.Top = CY - R * 0.55
        icoREC.Width = R
        icoREC.Height = R * 0.4
        DrawCaption hGraphics, m_IconCharCode, m_IconFont, icoREC, m_IconForeColor, 100, 0, eCenter, eMiddle, True

        ' Caption2 = valor actual (centro del gauge)
        cap2REC.Left = CX - R * 0.7
        cap2REC.Top = CY - R * 0.2
        cap2REC.Width = R * 1.4
        cap2REC.Height = R * 0.45
        DrawCaption hGraphics, m_Caption2, m_Font2, cap2REC, m_ForeColor2, 100, 0, eCenter, eMiddle, False

        ' Caption3 = unidad / subtitulo (debajo del valor)
        cap3REC.Left = CX - R * 0.6
        cap3REC.Top = CY + R * 0.22
        cap3REC.Width = R * 1.2
        cap3REC.Height = R * 0.3
        DrawCaption hGraphics, m_Caption3, m_Font3, cap3REC, m_ForeColor3, 100, 0, eCenter, eTop, False

        ' Caption1 = titulo (parte inferior del control)
        capREC.Left = margin
        capREC.Top = CY + R + thickHalf + 2
        capREC.Width = W - (margin * 2)
        capREC.Height = H - capREC.Top - margin
        If capREC.Height < 12 Then capREC.Height = 12
        DrawCaption hGraphics, m_Caption, m_Font, capREC, m_ForeColor, 100, 0, m_CaptionAlignH, eMiddle, False

        '--- Transparencia ---
        GdipDeleteGraphics hGraphics
        '--- Blit atomico: una sola operacion, sin flicker ---
        BitBlt .hdc, 0, 0, .ScaleWidth, .ScaleHeight, memDC, 0, 0, SRCCOPY
        '--- Liberar recursos del buffer ---
        SelectObject memDC, oldBmp
        DeleteObject memBmp
        DeleteDC memDC
        If m_Transparent Then
            .BackStyle = 0
            .MaskColor = .BackColor
            Set .MaskPicture = .Image
        End If
    End With
End Sub

'--- Dibuja un arco solido (unico color) ---
Private Sub DrawGaugeArc(ByVal hG As Long, ByVal X As Single, ByVal Y As Single, _
                          ByVal W As Single, ByVal H As Single, _
                          ByVal startAngle As Single, ByVal sweepAngle As Single, _
                          ByVal arcColor As Long, ByVal thickness As Single)
    Dim hPen As Long
    If sweepAngle < 0.1 Then Exit Sub
    GdipCreatePen1 arcColor, thickness, UnitPixel, hPen
    GdipDrawArc hG, hPen, X, Y, W, H, startAngle, sweepAngle
    GdipDeletePen hPen
End Sub

'--- Dibuja el arco de valor segun el modo de color ---
Private Sub DrawValueArc(ByVal hG As Long, ByVal X As Single, ByVal Y As Single, _
                          ByVal W As Single, ByVal H As Single, _
                          ByVal arcStart As Single, ByVal arcValue As Single, _
                          ByVal arcFull As Single)
    Dim hPen As Long
    Dim thickness As Single
    thickness = m_GaugeThickness * nScale

    Select Case m_GaugeColorMode

        Case gcmSolid
            DrawGaugeArc hG, X, Y, W, H, arcStart, arcValue, ARGB(m_GaugeColor1, 100), thickness

        Case gcmGradient
            DrawGradientArc hG, X, Y, W, H, arcStart, arcValue, _
                            m_GaugeColor1, m_GaugeColor2, thickness

        Case gcmThreshold
            ' Calculamos sweeps de cada zona sobre el arco completo
            Dim sw1 As Single, sw2 As Single, sw3 As Single
            sw1 = (m_ThresholdWarning / 100) * arcFull
            sw2 = ((m_ThresholdDanger - m_ThresholdWarning) / 100) * arcFull
            sw3 = ((100 - m_ThresholdDanger) / 100) * arcFull

            ' Zona verde (0 a ThresholdWarning)
            If arcValue > 0 Then
                Dim greenSw As Single
                greenSw = IIf(arcValue < sw1, arcValue, sw1)
                If greenSw > 0 Then
                    DrawGaugeArc hG, X, Y, W, H, arcStart, greenSw, ARGB(m_GaugeColor1, 100), thickness
                End If
            End If

            ' Zona amarilla (ThresholdWarning a ThresholdDanger)
            If arcValue > sw1 Then
                Dim yellowSw As Single
                yellowSw = IIf(arcValue - sw1 < sw2, arcValue - sw1, sw2)
                If yellowSw > 0 Then
                    DrawGaugeArc hG, X, Y, W, H, arcStart + sw1, yellowSw, ARGB(m_GaugeColorWarning, 100), thickness
                End If
            End If

            ' Zona roja (ThresholdDanger a 100)
            If arcValue > sw1 + sw2 Then
                Dim redSw As Single
                redSw = arcValue - sw1 - sw2
                If redSw > 0 Then
                    DrawGaugeArc hG, X, Y, W, H, arcStart + sw1 + sw2, redSw, ARGB(m_GaugeColorDanger, 100), thickness
                End If
            End If

    End Select
End Sub

'--- Arco con degradado de color (N segmentos interpolados) ---
Private Sub DrawGradientArc(ByVal hG As Long, ByVal X As Single, ByVal Y As Single, _
                             ByVal W As Single, ByVal H As Single, _
                             ByVal startAngle As Single, ByVal sweepAngle As Single, _
                             ByVal Color1 As OLE_COLOR, ByVal Color2 As OLE_COLOR, _
                             ByVal penW As Single)
    Const STEPS As Long = 24
    Dim i As Long
    Dim t As Single, stepAng As Single
    Dim R1 As Long, G1 As Long, B1 As Long
    Dim R2 As Long, G2 As Long, B2 As Long
    Dim Rc As Long, Gc As Long, Bc As Long
    Dim blendColor As Long
    Dim hPen As Long

    ' Resolver system colors si es necesario
    If (Color1 And &H80000000) Then Color1 = GetSysColor(Color1 And &HFF&)
    If (Color2 And &H80000000) Then Color2 = GetSysColor(Color2 And &HFF&)

    R1 = Color1 And &HFF&
    G1 = (Color1 And &HFF00&) \ &H100
    B1 = (Color1 And &HFF0000) \ &H10000
    R2 = Color2 And &HFF&
    G2 = (Color2 And &HFF00&) \ &H100
    B2 = (Color2 And &HFF0000) \ &H10000

    stepAng = sweepAngle / STEPS

    For i = 0 To STEPS - 1
        t = i / (STEPS - 1)
        Rc = CLng(R1 + (R2 - R1) * t)
        Gc = CLng(G1 + (G2 - G1) * t)
        Bc = CLng(B1 + (B2 - B1) * t)
        blendColor = ARGB(RGB(Rc, Gc, Bc), 100)

        GdipCreatePen1 blendColor, penW, UnitPixel, hPen
        GdipDrawArc hG, hPen, X, Y, W, H, startAngle + i * stepAng, stepAng + 0.3
        GdipDeletePen hPen
    Next i
End Sub

'--- Dibuja aguja / indicador en el extremo del arco ---
Private Sub DrawNeedle(ByVal hG As Long, ByVal CX As Single, ByVal CY As Single, _
                        ByVal R As Single, ByVal arcStart As Single, ByVal arcValue As Single)
    If m_GaugeNeedle = gnNone Then Exit Sub

    Dim endAngleRad As Double
    endAngleRad = (arcStart + arcValue) * DEG2RAD
    Dim tipX As Single, tipY As Single
    tipX = CX + R * Cos(endAngleRad)
    tipY = CY + R * Sin(endAngleRad)

    Dim hBrush As Long, hPen As Long
    Dim dotSz As Single
    dotSz = (m_GaugeThickness * nScale) + 2

    Select Case m_GaugeNeedle
        Case gnDot
            GdipCreateSolidFill ARGB(m_GaugeNeedleColor, 100), hBrush
            GdipFillEllipse hG, hBrush, tipX - dotSz / 2, tipY - dotSz / 2, dotSz, dotSz
            GdipDeleteBrush hBrush
            ' Hub central pequeno
            GdipCreateSolidFill ARGB(m_GaugeNeedleColor, 70), hBrush
            GdipFillEllipse hG, hBrush, CX - 4 * nScale, CY - 4 * nScale, 8 * nScale, 8 * nScale
            GdipDeleteBrush hBrush

        Case gnLine
            Dim lineEndX As Single, lineEndY As Single
            lineEndX = CX + (R - dotSz * 0.5) * Cos(endAngleRad)
            lineEndY = CY + (R - dotSz * 0.5) * Sin(endAngleRad)
            GdipCreatePen1 ARGB(m_GaugeNeedleColor, 100), 2 * nScale, UnitPixel, hPen
            GdipDrawLineI hG, hPen, CLng(CX), CLng(CY), CLng(lineEndX), CLng(lineEndY)
            GdipDeletePen hPen
            ' Hub central
            GdipCreateSolidFill ARGB(m_GaugeNeedleColor, 100), hBrush
            GdipFillEllipse hG, hBrush, CX - 4 * nScale, CY - 4 * nScale, 8 * nScale, 8 * nScale
            GdipDeleteBrush hBrush
    End Select
End Sub

'--- DrawCaption: renderiza texto o icono con GDI+ path ---
Private Function DrawCaption(ByVal hG As Long, sString As Variant, oFont As StdFont, _
                              layoutRect As RECTS, TextColor As OLE_COLOR, _
                              ColorOpacity As Long, mAngle As Single, _
                              HAlign As eTextAlignH, VAlign As eTextAlignV, _
                              Icon As Boolean) As Long
    Dim hPath As Long
    Dim hBrush As Long
    Dim hFontFamily As Long
    Dim hFormat As Long
    Dim lFontSize As Long
    Dim lFontStyle As Long

    If IsNull(sString) Or IsEmpty(sString) Then Exit Function
    If CStr(sString) = vbNullString Or CStr(sString) = "" Then Exit Function
    If oFont Is Nothing Then Exit Function

    GetFontStyleAndSize oFont, lFontStyle, lFontSize
    If lFontSize <= 0 Then Exit Function

    GdipCreateStringFormat 0, 0, hFormat

    Select Case HAlign
        Case eLeft:   GdipSetStringFormatAlign hFormat, StringAlignmentNear
        Case eCenter: GdipSetStringFormatAlign hFormat, StringAlignmentCenter
        Case eRight:  GdipSetStringFormatAlign hFormat, StringAlignmentFar
    End Select

    Select Case VAlign
        Case eTop:    GdipSetStringFormatLineAlign hFormat, StringAlignmentNear
        Case eMiddle: GdipSetStringFormatLineAlign hFormat, StringAlignmentCenter
        Case eBottom: GdipSetStringFormatLineAlign hFormat, StringAlignmentFar
    End Select

    GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap
    GdipSetStringFormatTrimming hFormat, StringTrimmingEllipsisCharacter

    If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) <> 0 Then
        GdipGetGenericFontFamilySansSerif hFontFamily
    End If

    If mAngle <> 0 Then
        GdipTranslateWorldTransform hG, layoutRect.Left + layoutRect.Width / 2, _
                                    layoutRect.Top + layoutRect.Height / 2, 0
        GdipRotateWorldTransform hG, mAngle, 0
        GdipTranslateWorldTransform hG, -(layoutRect.Left + layoutRect.Width / 2), _
                                    -(layoutRect.Top + layoutRect.Height / 2), 0
    End If

    GdipCreatePath 0, hPath

    If Icon Then
        GdipAddPathString hPath, StrPtr(ChrW2(sString)), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
    Else
        GdipAddPathString hPath, StrPtr(CStr(sString)), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
    End If

    GdipDeleteStringFormat hFormat
    GdipCreateSolidFill ARGB(TextColor, ColorOpacity), hBrush
    GdipFillPath hG, hBrush, hPath
    GdipDeleteBrush hBrush
    If mAngle <> 0 Then GdipResetWorldTransform hG

    GdipDeletePath hPath
    GdipDeleteFontFamily hFontFamily
End Function

'--- gRoundRect: dibuja rectangulo con esquinas redondeadas ---
Private Function gRoundRect(ByVal hG As Long, RECT As RECTL, ByVal Color1 As Long, _
                             ByVal Color2 As Long, ByVal Angulo As Single, _
                             ByVal BorderWidth As Long, ByVal BorderColor As Long, _
                             ByVal Round As Long, Filled As Boolean) As Long
    Dim hPen As Long, hBrush As Long, mPath As Long, mRound As Long
    If m_BorderWidth > 0 Then GdipCreatePen1 BorderColor, BorderWidth * nScale, &H2, hPen
    If Filled Then GdipCreateLineBrushFromRectWithAngleI RECT, Color1, Color2, Angulo + 90, 0, WrapModeTileFlipXY, hBrush
    GdipCreatePath &H0, mPath
    With RECT
        mRound = GetSafeRound(Round * nScale, .Width * 2, .Height * 2)
        If mRound < 1 Then mRound = 1
        GdipAddPathArcI mPath, .Left, .Top, mRound, mRound, 180, 90
        GdipAddPathArcI mPath, (.Left + .Width) - mRound, .Top, mRound, mRound, 270, 90
        GdipAddPathArcI mPath, (.Left + .Width) - mRound, (.Top + .Height) - mRound, mRound, mRound, 0, 90
        GdipAddPathArcI mPath, .Left, (.Top + .Height) - mRound, mRound, mRound, 90, 90
    End With
    GdipClosePathFigures mPath
    If Filled Then GdipFillPath hG, hBrush, mPath
    If m_BorderWidth > 0 Then GdipDrawPath hG, hPen, mPath
    GdipDeletePath mPath
    GdipDeleteBrush hBrush
    GdipDeletePen hPen
End Function

'--- Helpers ---
Private Function GetSafeRound(ByVal Round As Long, ByVal MaxW As Long, ByVal MaxH As Long) As Long
    If Round > MaxW Then Round = MaxW
    If Round > MaxH Then Round = MaxH
    GetSafeRound = Round
End Function

Private Function GetFontStyleAndSize(oFont As StdFont, lFontStyle As Long, lFontSize As Long)
    On Error GoTo ErrO
    Dim hdc As Long
    lFontStyle = 0
    If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
    hdc = GetDC(0)
    lFontSize = CLng(oFont.Size * GetDeviceCaps(hdc, LOGPIXELSY) / 72)
    ReleaseDC 0, hdc
ErrO:
End Function

Private Function ARGB(ByVal RGBColor As Long, ByVal Opacity As Long) As Long
    If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
    ARGB = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
    Opacity = CByte((Abs(Opacity) / 100) * 255)
    If Opacity < 128 Then
        If Opacity < 0& Then Opacity = 0&
        ARGB = ARGB Or Opacity * &H1000000
    Else
        If Opacity > 255& Then Opacity = 255&
        ARGB = ARGB Or (Opacity - 128&) * &H1000000 Or &H80000000
    End If
End Function

Private Function ChrW2(ByVal CharCode As Long) As String
    Const POW10 As Long = 2 ^ 10
    If CharCode <= &HFFFF& Then ChrW2 = ChrW$(CharCode) Else _
                                ChrW2 = ChrW$(&HD800& + (CharCode And &HFFFF&) \ POW10) & _
                                        ChrW$(&HDC00& + (CharCode And (POW10 - 1)))
End Function

Private Function GetWindowsDPI() As Double
    Dim hdc As Long, LPX As Double
    hdc = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
    ReleaseDC 0, hdc
    If LPX = 0 Then GetWindowsDPI = 1# Else GetWindowsDPI = LPX / 96#
End Function

Private Function IsMouseOver(hWnd As Long) As Boolean
    Dim Pt As POINTL
    GetCursorPos Pt
    IsMouseOver = (WindowFromPoint(Pt.X, Pt.Y) = hWnd)
End Function

Private Sub InitGDI()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Sub

Private Sub TerminateGDI()
    Call GdiplusShutdown(GdipToken)
End Sub

'===========================================================================
' PROPERTIES
'===========================================================================

'--- Background ---
Public Property Get BackAngle() As Single: BackAngle = m_Angulo: End Property
Public Property Let BackAngle(ByVal v As Single): m_Angulo = v: Refresh: End Property

Public Property Get BackColor1() As OLE_COLOR: BackColor1 = m_Color1: End Property
Public Property Let BackColor1(ByVal v As OLE_COLOR): m_Color1 = v: Refresh: End Property

Public Property Get BackColor2() As OLE_COLOR: BackColor2 = m_Color2: End Property
Public Property Let BackColor2(ByVal v As OLE_COLOR): m_Color2 = v: Refresh: End Property

Public Property Get BorderColor() As OLE_COLOR: BorderColor = m_BorderColor: End Property
Public Property Let BorderColor(ByVal v As OLE_COLOR): m_BorderColor = v: Refresh: End Property

Public Property Get BorderWidth() As Long: BorderWidth = m_BorderWidth: End Property
Public Property Let BorderWidth(ByVal v As Long): m_BorderWidth = v: Refresh: End Property

Public Property Get CornerCurve() As Long: CornerCurve = m_CornerCurve: End Property
Public Property Let CornerCurve(ByVal v As Long): m_CornerCurve = v: Refresh: End Property

Public Property Get Filled() As Boolean: Filled = m_Filled: End Property
Public Property Let Filled(ByVal v As Boolean): m_Filled = v: Refresh: End Property

'--- Caption 1 ---
Public Property Get Caption1() As String: Caption1 = m_Caption: End Property
Public Property Let Caption1(ByVal v As String): m_Caption = v: Refresh: End Property
Public Property Get Caption1Font() As StdFont: Set Caption1Font = m_Font: End Property
Public Property Set Caption1Font(ByVal v As StdFont): Set m_Font = v: Refresh: End Property
Public Property Get Caption1Color() As OLE_COLOR: Caption1Color = m_ForeColor: End Property
Public Property Let Caption1Color(ByVal v As OLE_COLOR): m_ForeColor = v: Refresh: End Property
Public Property Get Caption1AlignH() As eTextAlignH: Caption1AlignH = m_CaptionAlignH: End Property
Public Property Let Caption1AlignH(ByVal v As eTextAlignH): m_CaptionAlignH = v: Refresh: End Property

'--- Caption 2 ---
Public Property Get Caption2() As String: Caption2 = m_Caption2: End Property
Public Property Let Caption2(ByVal v As String): m_Caption2 = v: Refresh: End Property
Public Property Get Caption2Font() As StdFont: Set Caption2Font = m_Font2: End Property
Public Property Set Caption2Font(ByVal v As StdFont): Set m_Font2 = v: Refresh: End Property
Public Property Get Caption2Color() As OLE_COLOR: Caption2Color = m_ForeColor2: End Property
Public Property Let Caption2Color(ByVal v As OLE_COLOR): m_ForeColor2 = v: Refresh: End Property

'--- Caption 3 ---
Public Property Get Caption3() As String: Caption3 = m_Caption3: End Property
Public Property Let Caption3(ByVal v As String): m_Caption3 = v: Refresh: End Property
Public Property Get Caption3Font() As StdFont: Set Caption3Font = m_Font3: End Property
Public Property Set Caption3Font(ByVal v As StdFont): Set m_Font3 = v: Refresh: End Property
Public Property Get Caption3Color() As OLE_COLOR: Caption3Color = m_ForeColor3: End Property
Public Property Let Caption3Color(ByVal v As OLE_COLOR): m_ForeColor3 = v: Refresh: End Property

'--- Icon ---
Public Property Get IconCharCode() As String: IconCharCode = m_IconCharCode: End Property
Public Property Let IconCharCode(ByVal v As String): m_IconCharCode = v: Refresh: End Property
Public Property Get IconFont() As StdFont: Set IconFont = m_IconFont: End Property
Public Property Set IconFont(v As StdFont): Set m_IconFont = v: Refresh: End Property
Public Property Get IconForeColor() As OLE_COLOR: IconForeColor = m_IconForeColor: End Property
Public Property Let IconForeColor(ByVal v As OLE_COLOR): m_IconForeColor = v: Refresh: End Property

'--- Gauge ---
Public Property Get GaugeValue() As Single: GaugeValue = m_GaugeValue: End Property
Public Property Let GaugeValue(ByVal v As Single)
    If v < 0 Then v = 0
    If v > 100 Then v = 100
    m_GaugeValue = v
    If m_GaugeAnimated Then
        tmrAnim.Enabled = True
    Else
        m_AnimValue = v
        Refresh
    End If
End Property

Public Property Get GaugeMin() As Single: GaugeMin = m_GaugeMin: End Property
Public Property Let GaugeMin(ByVal v As Single): m_GaugeMin = v: Refresh: End Property

Public Property Get GaugeMax() As Single: GaugeMax = m_GaugeMax: End Property
Public Property Let GaugeMax(ByVal v As Single): m_GaugeMax = v: Refresh: End Property

Public Property Get GaugeStyle() As eGaugeStyle: GaugeStyle = m_GaugeStyle: End Property
Public Property Let GaugeStyle(ByVal v As eGaugeStyle): m_GaugeStyle = v: Refresh: End Property

Public Property Get GaugeColorMode() As eGaugeColorMode: GaugeColorMode = m_GaugeColorMode: End Property
Public Property Let GaugeColorMode(ByVal v As eGaugeColorMode): m_GaugeColorMode = v: Refresh: End Property

Public Property Get GaugeColor1() As OLE_COLOR: GaugeColor1 = m_GaugeColor1: End Property
Public Property Let GaugeColor1(ByVal v As OLE_COLOR): m_GaugeColor1 = v: Refresh: End Property

Public Property Get GaugeColor2() As OLE_COLOR: GaugeColor2 = m_GaugeColor2: End Property
Public Property Let GaugeColor2(ByVal v As OLE_COLOR): m_GaugeColor2 = v: Refresh: End Property

Public Property Get GaugeTrackColor() As OLE_COLOR: GaugeTrackColor = m_GaugeTrackColor: End Property
Public Property Let GaugeTrackColor(ByVal v As OLE_COLOR): m_GaugeTrackColor = v: Refresh: End Property

Public Property Get GaugeThickness() As Long: GaugeThickness = m_GaugeThickness: End Property
Public Property Let GaugeThickness(ByVal v As Long)
    If v < 2 Then v = 2
    If v > 40 Then v = 40
    m_GaugeThickness = v: Refresh
End Property

Public Property Get GaugeThresholdWarning() As Single: GaugeThresholdWarning = m_ThresholdWarning: End Property
Public Property Let GaugeThresholdWarning(ByVal v As Single)
    If v < 0 Then v = 0
    If v > 100 Then v = 100
    m_ThresholdWarning = v: Refresh
End Property

Public Property Get GaugeThresholdDanger() As Single: GaugeThresholdDanger = m_ThresholdDanger: End Property
Public Property Let GaugeThresholdDanger(ByVal v As Single)
    If v < 0 Then v = 0
    If v > 100 Then v = 100
    m_ThresholdDanger = v: Refresh
End Property

Public Property Get GaugeColorWarning() As OLE_COLOR: GaugeColorWarning = m_GaugeColorWarning: End Property
Public Property Let GaugeColorWarning(ByVal v As OLE_COLOR): m_GaugeColorWarning = v: Refresh: End Property

Public Property Get GaugeColorDanger() As OLE_COLOR: GaugeColorDanger = m_GaugeColorDanger: End Property
Public Property Let GaugeColorDanger(ByVal v As OLE_COLOR): m_GaugeColorDanger = v: Refresh: End Property

Public Property Get GaugeNeedle() As eGaugeNeedleStyle: GaugeNeedle = m_GaugeNeedle: End Property
Public Property Let GaugeNeedle(ByVal v As eGaugeNeedleStyle): m_GaugeNeedle = v: Refresh: End Property

Public Property Get GaugeNeedleColor() As OLE_COLOR: GaugeNeedleColor = m_GaugeNeedleColor: End Property
Public Property Let GaugeNeedleColor(ByVal v As OLE_COLOR): m_GaugeNeedleColor = v: Refresh: End Property

Public Property Get GaugeAnimated() As Boolean: GaugeAnimated = m_GaugeAnimated: End Property
Public Property Let GaugeAnimated(ByVal v As Boolean): m_GaugeAnimated = v: End Property

Public Property Get GaugeAnimSpeed() As Long: GaugeAnimSpeed = m_GaugeAnimSpeed: End Property
Public Property Let GaugeAnimSpeed(ByVal v As Long)
    If v < 1 Then v = 1
    If v > 20 Then v = 20
    m_GaugeAnimSpeed = v
End Property

'--- Misc ---
Public Property Get Enabled() As Boolean: Enabled = m_Enabled: End Property
Public Property Let Enabled(ByVal v As Boolean): m_Enabled = v: Refresh: End Property

Public Property Get Visible() As Boolean: Visible = UserControl.Visible: End Property
Public Property Let Visible(ByVal v As Boolean): UserControl.Visible = v: End Property

Public Property Get Clickable() As Boolean: Clickable = m_Clickable: End Property
Public Property Let Clickable(ByVal v As Boolean): m_Clickable = v: Refresh: End Property

Public Property Get Transparent() As Boolean: Transparent = m_Transparent: End Property
Public Property Let Transparent(ByVal v As Boolean): m_Transparent = v: Refresh: End Property

Public Property Get EffectFading() As Boolean: EffectFading = m_EffectFade: End Property
Public Property Let EffectFading(ByVal v As Boolean): m_EffectFade = v: End Property

Public Property Get InitialOpacity() As Long: InitialOpacity = m_InitialOpacity: End Property
Public Property Let InitialOpacity(ByVal v As Long)
    If v < 0 Then v = 0
    If v > 100 Then v = 100
    m_InitialOpacity = v: m_Opacity = v
End Property

Public Property Get Tag() As String: Tag = m_Tag: End Property
Public Property Let Tag(ByVal v As String): m_Tag = v: End Property

Public Property Get hdc() As Long: hdc = UserControl.hdc: End Property
Public Property Get hWnd() As Long: hWnd = UserControl.hWnd: End Property

Public Property Get Version() As String: Version = "1.00": End Property

'===========================================================================
' USERCONTROL EVENTS
'===========================================================================

Private Sub UserControl_Initialize()
    InitGDI
    nScale = GetWindowsDPI()
End Sub

Private Sub UserControl_InitProperties()
    m_Enabled = True
    m_Visible = True
    m_Clickable = False
    m_Transparent = False
    m_Filled = True
    m_Opacity = 100
    m_InitialOpacity = 100
    m_EffectFade = False

    m_Color1 = RGB(30, 30, 46)
    m_Color2 = RGB(30, 30, 46)
    m_def_Color1 = m_Color1
    m_def_Color2 = m_Color2
    m_Angulo = 0
    m_BorderColor = RGB(60, 60, 90)
    m_BorderWidth = 1
    m_CornerCurve = 8

    Set m_Font = UserControl.Ambient.Font
    Set m_Font2 = UserControl.Ambient.Font
    Set m_Font3 = UserControl.Ambient.Font
    Set m_IconFont = UserControl.Font

    m_ForeColor = vbWhite
    m_ForeColor2 = vbWhite
    m_ForeColor3 = RGB(180, 180, 200)
    m_def_ForeColor = vbWhite

    m_CaptionAlignH = eCenter
    m_CaptionAlignV = eMiddle
    m_Caption = Ambient.DisplayName
    m_Caption2 = "0%"
    m_Caption3 = ""
    m_IconCharCode = ""
    m_IconForeColor = vbWhite

    ' Gauge defaults
    m_GaugeValue = 0
    m_AnimValue = 0
    m_GaugeMin = 0
    m_GaugeMax = 100
    m_GaugeStyle = gsArcDial       ' 270 grados default
    m_GaugeColorMode = gcmSolid
    m_GaugeColor1 = RGB(99, 102, 241)   ' Indigo
    m_GaugeColor2 = RGB(236, 72, 153)   ' Pink
    m_GaugeTrackColor = RGB(60, 60, 80)
    m_GaugeThickness = 12
    m_ThresholdWarning = 60
    m_ThresholdDanger = 80
    m_GaugeColorWarning = RGB(234, 179, 8)    ' Amarillo
    m_GaugeColorDanger = RGB(239, 68, 68)     ' Rojo
    m_GaugeNeedle = gnDot
    m_GaugeNeedleColor = vbWhite
    m_GaugeAnimated = True
    m_GaugeAnimSpeed = 8
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Enabled = .ReadProperty("Enabled", True)
        m_Color1 = .ReadProperty("BackColor1", m_def_Color1)
        m_Color2 = .ReadProperty("BackColor2", m_def_Color2)
        m_Angulo = .ReadProperty("BackAngle", 0)
        m_BorderColor = .ReadProperty("BorderColor", RGB(60, 60, 90))
        m_BorderWidth = .ReadProperty("BorderWidth", 1)
        m_CornerCurve = .ReadProperty("CornerCurve", 8)
        m_Filled = .ReadProperty("Filled", True)
        m_Transparent = .ReadProperty("Transparent", False)
        m_Clickable = .ReadProperty("Clickable", False)
        m_EffectFade = .ReadProperty("EffectFading", False)
        m_InitialOpacity = .ReadProperty("InitialOpacity", 100)
        m_Opacity = m_InitialOpacity

        m_Caption = .ReadProperty("Caption1", "")
        Set m_Font = .ReadProperty("Caption1Font", UserControl.Ambient.Font)
        m_ForeColor = .ReadProperty("Caption1Color", vbWhite)
        m_CaptionAlignH = .ReadProperty("Caption1AlignH", eCenter)
        m_Caption2 = .ReadProperty("Caption2", "0%")
        Set m_Font2 = .ReadProperty("Caption2Font", UserControl.Ambient.Font)
        m_ForeColor2 = .ReadProperty("Caption2Color", vbWhite)
        m_Caption3 = .ReadProperty("Caption3", "")
        Set m_Font3 = .ReadProperty("Caption3Font", UserControl.Ambient.Font)
        m_ForeColor3 = .ReadProperty("Caption3Color", RGB(180, 180, 200))

        m_IconCharCode = .ReadProperty("IconCharCode", "")
        Set m_IconFont = .ReadProperty("IconFont", UserControl.Font)
        m_IconForeColor = .ReadProperty("IconForeColor", vbWhite)

        m_GaugeValue = .ReadProperty("GaugeValue", 0)
        m_AnimValue = m_GaugeValue
        m_GaugeMin = .ReadProperty("GaugeMin", 0)
        m_GaugeMax = .ReadProperty("GaugeMax", 100)
        m_GaugeStyle = .ReadProperty("GaugeStyle", gsArcDial)
        m_GaugeColorMode = .ReadProperty("GaugeColorMode", gcmSolid)
        m_GaugeColor1 = .ReadProperty("GaugeColor1", RGB(99, 102, 241))
        m_GaugeColor2 = .ReadProperty("GaugeColor2", RGB(236, 72, 153))
        m_GaugeTrackColor = .ReadProperty("GaugeTrackColor", RGB(60, 60, 80))
        m_GaugeThickness = .ReadProperty("GaugeThickness", 12)
        m_ThresholdWarning = .ReadProperty("GaugeThresholdWarning", 60)
        m_ThresholdDanger = .ReadProperty("GaugeThresholdDanger", 80)
        m_GaugeColorWarning = .ReadProperty("GaugeColorWarning", RGB(234, 179, 8))
        m_GaugeColorDanger = .ReadProperty("GaugeColorDanger", RGB(239, 68, 68))
        m_GaugeNeedle = .ReadProperty("GaugeNeedle", gnDot)
        m_GaugeNeedleColor = .ReadProperty("GaugeNeedleColor", vbWhite)
        m_GaugeAnimated = .ReadProperty("GaugeAnimated", True)
        m_GaugeAnimSpeed = .ReadProperty("GaugeAnimSpeed", 8)
        m_Tag = .ReadProperty("Tag", "")
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Enabled", m_Enabled, True
        .WriteProperty "BackColor1", m_Color1, m_def_Color1
        .WriteProperty "BackColor2", m_Color2, m_def_Color2
        .WriteProperty "BackAngle", m_Angulo, 0
        .WriteProperty "BorderColor", m_BorderColor, RGB(60, 60, 90)
        .WriteProperty "BorderWidth", m_BorderWidth, 1
        .WriteProperty "CornerCurve", m_CornerCurve, 8
        .WriteProperty "Filled", m_Filled, True
        .WriteProperty "Transparent", m_Transparent, False
        .WriteProperty "Clickable", m_Clickable, False
        .WriteProperty "EffectFading", m_EffectFade, False
        .WriteProperty "InitialOpacity", m_InitialOpacity, 100

        .WriteProperty "Caption1", m_Caption, ""
        .WriteProperty "Caption1Font", m_Font, UserControl.Ambient.Font
        .WriteProperty "Caption1Color", m_ForeColor, vbWhite
        .WriteProperty "Caption1AlignH", m_CaptionAlignH, eCenter
        .WriteProperty "Caption2", m_Caption2, "0%"
        .WriteProperty "Caption2Font", m_Font2, UserControl.Ambient.Font
        .WriteProperty "Caption2Color", m_ForeColor2, vbWhite
        .WriteProperty "Caption3", m_Caption3, ""
        .WriteProperty "Caption3Font", m_Font3, UserControl.Ambient.Font
        .WriteProperty "Caption3Color", m_ForeColor3, RGB(180, 180, 200)

        .WriteProperty "IconCharCode", m_IconCharCode, ""
        .WriteProperty "IconFont", m_IconFont, UserControl.Font
        .WriteProperty "IconForeColor", m_IconForeColor, vbWhite

        .WriteProperty "GaugeValue", m_GaugeValue, 0
        .WriteProperty "GaugeMin", m_GaugeMin, 0
        .WriteProperty "GaugeMax", m_GaugeMax, 100
        .WriteProperty "GaugeStyle", m_GaugeStyle, gsArcDial
        .WriteProperty "GaugeColorMode", m_GaugeColorMode, gcmSolid
        .WriteProperty "GaugeColor1", m_GaugeColor1, RGB(99, 102, 241)
        .WriteProperty "GaugeColor2", m_GaugeColor2, RGB(236, 72, 153)
        .WriteProperty "GaugeTrackColor", m_GaugeTrackColor, RGB(60, 60, 80)
        .WriteProperty "GaugeThickness", m_GaugeThickness, 12
        .WriteProperty "GaugeThresholdWarning", m_ThresholdWarning, 60
        .WriteProperty "GaugeThresholdDanger", m_ThresholdDanger, 80
        .WriteProperty "GaugeColorWarning", m_GaugeColorWarning, RGB(234, 179, 8)
        .WriteProperty "GaugeColorDanger", m_GaugeColorDanger, RGB(239, 68, 68)
        .WriteProperty "GaugeNeedle", m_GaugeNeedle, gnDot
        .WriteProperty "GaugeNeedleColor", m_GaugeNeedleColor, vbWhite
        .WriteProperty "GaugeAnimated", m_GaugeAnimated, True
        .WriteProperty "GaugeAnimSpeed", m_GaugeAnimSpeed, 8
        .WriteProperty "Tag", m_Tag, ""
    End With
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub

Private Sub UserControl_Paint()
    Draw
End Sub

Private Sub UserControl_Terminate()
    TerminateGDI
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    CopyAmbient
End Sub

Private Sub UserControl_Click()
    If m_Clickable Then RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
    If m_EffectFade Then
        m_Opacity = m_InitialOpacity
        tmrEffect.Enabled = True
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'--- Timer: animacion del gauge (easing exponencial) ---
Private Sub tmrAnim_Timer()
    Dim diff As Single
    diff = m_GaugeValue - m_AnimValue
    If Abs(diff) < 0.25 Then
        m_AnimValue = m_GaugeValue
        tmrAnim.Enabled = False
    Else
        m_AnimValue = m_AnimValue + diff * (m_GaugeAnimSpeed / 100)
    End If
    Refresh
End Sub

'--- Timer: efecto hover fade ---
Private Sub tmrEffect_Timer()
    If IsMouseOver(UserControl.hWnd) Then
        If m_Opacity < 100 Then
            m_Opacity = m_Opacity + 2
            Refresh
        Else
            Exit Sub
        End If
    Else
        m_Opacity = m_InitialOpacity
        Refresh
        tmrEffect.Enabled = False
    End If
End Sub
