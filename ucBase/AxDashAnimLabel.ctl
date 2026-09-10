VERSION 5.00
Begin VB.UserControl AxDashAnimLabel 
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6315
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
   ScaleHeight     =   175
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   ToolboxBitmap   =   "AxDashAnimLabel.ctx":0000
   Begin VB.Timer tmrAnim 
      Enabled         =   0   'False
      Interval        =   50
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
Attribute VB_Name = "AxDashAnimLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-UC-VB6-----------------------------
'UC Name  : AxDashAnimLabel
'Version  : 1.00
'Editor   : David Rojas [AxioUK]
'Date     : 10/09/2026
'Desc     : Grafico de linea animado en loop/shift para Dashboards
'------------------------------------
Option Explicit

'--- Win32 API ---
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

'--- GDI+ Pen / Brush ---
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
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
Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long

'--- GDI+ Draw shapes ---
Private Declare Function GdipDrawLineI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipFillEllipse Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal X As Single, ByVal Y As Single, ByVal Width As Single, ByVal Height As Single) As Long
Private Declare Function GdipAddPathCurve Lib "gdiplus" (ByVal path As Long, pPoints As Any, ByVal count As Long) As Long

'--- GDI+ Text ---
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

'--- GDI+ Transform / Clip ---
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "GdiPlus.dll" (ByVal graphics As Long) As Long
Private Declare Function GdipSetClipRectI Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long, ByVal combineMode As Long) As Long
Private Declare Function GdipResetClip Lib "GdiPlus.dll" (ByVal graphics As Long) As Long

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

Private Type POINTS
    X As Integer
    Y As Integer
End Type

Private Type GDIPlusStartupInput
    GdiPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
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

'--- GDI+ token / graphics handle ---
Private GdipToken As Long
Private hGraphics As Long
Private hBrush As Long
Private hPen As Long
Private nScale As Double

'--- Standard properties ---
Private m_Enabled As Boolean
Private m_Visible As Boolean
Private m_Clickable As Boolean
Private m_Transparent As Boolean
Private m_Filled As Boolean
Private m_EffectFade As Boolean
Private m_Opacity As Long
Private m_InitialOpacity As Long
Private m_CornerCurve As Long
Private m_Tag As String

'--- Background ---
Private m_Color1 As OLE_COLOR
Private m_Color2 As OLE_COLOR
Private m_Angulo As Single
Private m_BorderColor As OLE_COLOR
Private m_BorderWidth As Long
Private m_def_Color1 As OLE_COLOR
Private m_def_Color2 As OLE_COLOR
Private m_def_ForeColor As OLE_COLOR

'--- Caption 1 (arriba izquierda) ---
Private m_Caption As String
Private m_Font As StdFont
Private m_ForeColor As OLE_COLOR
Private m_CaptionAlignH As eTextAlignH
Private m_CaptionAlignV As eTextAlignV

'--- Caption 2 (arriba derecha / valor) ---
Private m_Caption2 As String
Private m_Font2 As StdFont
Private m_ForeColor2 As OLE_COLOR
Private m_Caption2AlignH As eTextAlignH

'--- Caption 3 (inferior) ---
Private m_Caption3 As String
Private m_Font3 As StdFont
Private m_ForeColor3 As OLE_COLOR
Private m_Caption3AlignH As eTextAlignH

'--- Icon ---
Private m_IconCharCode As String
Private m_IconFont As StdFont
Private m_IconForeColor As OLE_COLOR
Private m_IconAlignH As eTextAlignH
Private m_IconAlignV As eTextAlignV
Private m_isBoxed As Boolean
Private m_BoxedColor As OLE_COLOR

'--- Graph static properties (heredado de GraphLabel) ---
Private m_Graph As eGraph
Private m_GraphMatrix As String
Private m_GraphMatrixTooltip As String
Private m_GraphStyle As eGraphStyle
Private m_GraphLineColor As OLE_COLOR
Private m_GraphBackColor As OLE_COLOR
Private m_GraphPointColor As OLE_COLOR
Private m_GraphFillOpacity As Long
Private m_GraphGridLines As Boolean
Private m_GraphGridColor As OLE_COLOR

'--- Animation properties ---
Private m_AnimEnabled As Boolean
Private m_AnimMode As eAnimMode
Private m_AnimSpeed As Long        ' intervalo timer en ms
Private m_AnimStep As Long         ' px de desplazamiento por frame (modo visual)
Private m_AnimSmooth As Boolean

'--- Internal animation state ---
Private m_AnimBuffer() As Single   ' buffer de trabajo (puede rotar)
Private m_AnimOffset As Single     ' desplazamiento sub-pixel actual
Private m_AnimBufferCount As Long  ' numero de puntos en buffer
Private m_AnimInitialized As Boolean

'--- Tooltip ---
Private mShowTtp As Boolean
Private sToolTip As String
Private m_PosX As Long, m_PosY As Long
Private gREC() As RECTL
Private iPts() As POINTS

'--- Events ---
Public Event Click(ByVal Serie As Long)
Public Event DblClick()
Public Event AnimCycle()
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

'--- Agregar un valor nuevo al final del buffer (modo amShift) ---
Public Sub PushValue(ByVal NewValue As Single)
    If NewValue < 0 Then NewValue = 0
    If NewValue > 100 Then NewValue = 100
    If m_AnimBufferCount < 2 Then
        InitAnimBuffer
        Exit Sub
    End If
    ' Shift left: descarta el primero, agrega el nuevo al final
    Dim i As Long
    For i = 0 To m_AnimBufferCount - 2
        m_AnimBuffer(i) = m_AnimBuffer(i + 1)
    Next i
    m_AnimBuffer(m_AnimBufferCount - 1) = NewValue
    If m_AnimMode = amShift Then Refresh
End Sub

'--- Resetear y recargar el buffer desde GraphMatrix ---
Public Sub ResetAnimation()
    m_AnimOffset = 0
    InitAnimBuffer
    Refresh
End Sub

'===========================================================================
' ANIMATION BUFFER MANAGEMENT
'===========================================================================

Private Sub InitAnimBuffer()
    Dim pts() As String
    Dim i As Long

    If m_GraphMatrix = "" Or m_GraphMatrix = "0" Then
        m_AnimBufferCount = 2
        ReDim m_AnimBuffer(1)
        m_AnimBuffer(0) = 0: m_AnimBuffer(1) = 0
        m_AnimInitialized = True
        Exit Sub
    End If

    pts = Split(m_GraphMatrix, ",")
    m_AnimBufferCount = UBound(pts) + 1
    ReDim m_AnimBuffer(m_AnimBufferCount - 1)
    For i = 0 To m_AnimBufferCount - 1
        m_AnimBuffer(i) = Val(pts(i))
        If m_AnimBuffer(i) < 0 Then m_AnimBuffer(i) = 0
        If m_AnimBuffer(i) > 100 Then m_AnimBuffer(i) = 100
    Next i
    m_AnimOffset = 0
    m_AnimInitialized = True
End Sub

'--- Rotacion izquierda del buffer (modo Loop) ---
Private Sub RotateBufferLeft()
    If m_AnimBufferCount < 2 Then Exit Sub
    Dim firstVal As Single
    Dim i As Long
    firstVal = m_AnimBuffer(0)
    For i = 0 To m_AnimBufferCount - 2
        m_AnimBuffer(i) = m_AnimBuffer(i + 1)
    Next i
    m_AnimBuffer(m_AnimBufferCount - 1) = firstVal
    RaiseEvent AnimCycle
End Sub

'===========================================================================
' CORE DRAWING
'===========================================================================

Private Sub Draw()
    Dim W As Long, H As Long
    Dim REC As RECTL
    Dim headerH As Long, footerH As Long, graphH As Long
    Dim graphREC As RECTL
    Dim stREC As RECTS, stREC2 As RECTS, stREC3 As RECTS
    Dim IcoBox As RECTS
    Dim BOX As RECTL
    Dim lBorder As Long, mBorder As Long
    Dim margin As Long

    If Not m_AnimInitialized Then InitAnimBuffer

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
        lBorder = m_BorderWidth * 2
        mBorder = lBorder / 2

        '--- Background ---
        REC.Left = 0: REC.Top = 0
        REC.Width = W - 1: REC.Height = H - 1
        gRoundRect hGraphics, REC, ARGB(m_Color1, 100), ARGB(m_Color2, 100), _
                   m_Angulo, m_BorderWidth, ARGB(m_BorderColor, 100), m_CornerCurve, m_Filled

        '--- Layout ---
        headerH = CLng(H * 0.3)   ' 30% header (captions 1 y 2 + icono)
        footerH = CLng(H * 0.15)  ' 15% footer (caption 3)
        graphH = H - headerH - footerH

        ' Icon box (si esta activo)
        BOX.Left = margin
        BOX.Top = margin
        BOX.Width = headerH - (margin * 2)
        BOX.Height = headerH - (margin * 2)

        IcoBox.Left = BOX.Left: IcoBox.Top = BOX.Top
        IcoBox.Width = BOX.Width: IcoBox.Height = BOX.Height

        Dim captionOffsetL As Long
        captionOffsetL = IIf(m_isBoxed, BOX.Left + BOX.Width + margin, margin)

        ' Caption 1 (titulo / izquierda)
        stREC.Left = captionOffsetL
        stREC.Top = mBorder + margin
        stREC.Width = (W \ 2) - captionOffsetL - margin
        stREC.Height = headerH - (margin * 2)

        ' Caption 2 (valor / derecha)
        stREC2.Left = W \ 2
        stREC2.Top = mBorder + margin
        stREC2.Width = (W \ 2) - margin * 2
        stREC2.Height = headerH - (margin * 2)

        ' Caption 3 (inferior)
        stREC3.Left = margin
        stREC3.Top = H - footerH
        stREC3.Width = W - margin * 2
        stREC3.Height = footerH - 2

        '--- Graph region ---
        graphREC.Left = margin
        graphREC.Top = headerH
        graphREC.Width = W - margin * 2
        graphREC.Height = graphH

        '--- Dibujar separador header/grafico ---
        Dim sepBrush As Long
        GdipCreateSolidFill ARGB(m_BorderColor, 30), sepBrush
        GdipFillRectangleI hGraphics, sepBrush, margin, headerH, W - margin * 2, 1
        GdipDeleteBrush sepBrush

        '--- Grafico con clipping para no salirse del area ---
        GdipSetClipRectI hGraphics, graphREC.Left, graphREC.Top, graphREC.Width, graphREC.Height, 0
        DrawAnimatedGraph hGraphics, graphREC
        GdipResetClip hGraphics

        '--- Icon box ---
        If m_isBoxed Then
            gRoundRect hGraphics, BOX, ARGB(m_BoxedColor, 100), ARGB(m_BoxedColor, 100), _
                       m_Angulo, m_BorderWidth, ARGB(m_BoxedColor, 100), m_CornerCurve, m_Filled
            DrawCaption hGraphics, m_IconCharCode, m_IconFont, IcoBox, m_IconForeColor, 100, 0, eCenter, eMiddle, True
        End If

        '--- Captions ---
        DrawCaption hGraphics, m_Caption, m_Font, stREC, m_ForeColor, 100, 0, m_CaptionAlignH, eMiddle, False
        DrawCaption hGraphics, m_Caption2, m_Font2, stREC2, m_ForeColor2, 100, 0, m_Caption2AlignH, eMiddle, False
        DrawCaption hGraphics, m_Caption3, m_Font3, stREC3, m_ForeColor3, 100, 0, m_Caption3AlignH, eMiddle, False

        '--- Grid lines (opcional) ---
        If m_GraphGridLines Then
            DrawGridLines hGraphics, graphREC
        End If

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

'--- Dibuja el grafico con offset de animacion ---
Private Sub DrawAnimatedGraph(ByVal hG As Long, gREC As RECTL)
    If m_AnimBufferCount < 2 Then Exit Sub

    Dim n As Long
    Dim i As Long
    Dim W As Long, H As Long
    Dim stepPx As Single        ' ancho por punto en pixels
    Dim offsetX As Long
    Dim px As Long, py As Long
    Dim pts() As POINTS

    W = gREC.Width
    H = gREC.Height
    n = m_AnimBufferCount - 1
    stepPx = W / n

    ' El offset de animacion desplaza el grafico hacia la izquierda
    offsetX = CLng(m_AnimOffset)

    ' Necesitamos un punto extra a la derecha para continuidad visual en modo loop
    Dim extraCount As Long
    extraCount = m_AnimBufferCount + 1   ' buffer + wrap-around
    ReDim pts(extraCount - 1)

    For i = 0 To m_AnimBufferCount - 1
        pts(i).X = CInt(gREC.Left + stepPx * i - offsetX)
        pts(i).Y = CInt(gREC.Top + (H / 100) * (100 - m_AnimBuffer(i)))
    Next i
    ' Punto extra de wrap-around (primer punto del buffer para continuidad)
    pts(m_AnimBufferCount).X = CInt(gREC.Left + stepPx * m_AnimBufferCount - offsetX)
    pts(m_AnimBufferCount).Y = CInt(gREC.Top + (H / 100) * (100 - m_AnimBuffer(0)))

    ' Fill area bajo la curva
    Dim fillREC As RECTL
    fillREC = gREC

    If m_GraphStyle = gsGradient Then
        GdipCreateLineBrushFromRectWithAngleI fillREC, ARGB(m_GraphBackColor, m_GraphFillOpacity), _
            ARGB(m_Color1, 20), 90, 0, WrapModeTileFlipXY, hBrush
    Else
        GdipCreateSolidFill ARGB(m_GraphBackColor, m_GraphFillOpacity), hBrush
    End If

    GdipCreatePen1 ARGB(m_GraphLineColor, 80), 1.5 * nScale, UnitPixel, hPen

    Select Case m_Graph
        Case egRectLine
            ' Linea recta punto a punto
            Dim mPath As Long
            GdipCreatePath &H0, mPath
            For i = 0 To extraCount - 2
                GdipAddPathLineI mPath, pts(i).X, pts(i).Y, pts(i + 1).X, pts(i + 1).Y
            Next i
            ' Cerrar el path hacia abajo para el relleno
            ' Linea vertical del ultimo punto al borde inferior
            GdipAddPathLineI mPath, pts(extraCount - 1).X, pts(extraCount - 1).Y, pts(extraCount - 1).X, gREC.Top + gREC.Height
            ' Linea horizontal inferior de vuelta al inicio
            GdipAddPathLineI mPath, pts(extraCount - 1).X, gREC.Top + gREC.Height, pts(0).X, gREC.Top + gREC.Height
            GdipClosePathFigures mPath
            GdipFillPath hG, hBrush, mPath
            ' Redibujar solo la linea
            GdipDeletePath mPath
            GdipCreatePath &H0, mPath
            For i = 0 To extraCount - 2
                GdipAddPathLineI mPath, pts(i).X, pts(i).Y, pts(i + 1).X, pts(i + 1).Y
            Next i
            GdipDrawPath hG, hPen, mPath
            GdipDeletePath mPath

        Case egCurvedLine
            ' Curva suavizada
            If extraCount >= 3 Then
                Dim curvePath As Long
                GdipCreatePath &H0, curvePath
                GdipAddPathCurve curvePath, pts(0), extraCount
                ' Extender para relleno: vertical al borde inferior, luego horizontal al inicio
                GdipAddPathLineI curvePath, pts(extraCount - 1).X, pts(extraCount - 1).Y, pts(extraCount - 1).X, gREC.Top + gREC.Height
                GdipAddPathLineI curvePath, pts(extraCount - 1).X, gREC.Top + gREC.Height, pts(0).X, gREC.Top + gREC.Height
                GdipClosePathFigures curvePath
                GdipFillPath hG, hBrush, curvePath
                GdipDeletePath curvePath
                ' Redibujar curva
                GdipCreatePath &H0, curvePath
                GdipAddPathCurve curvePath, pts(0), extraCount
                GdipDrawPath hG, hPen, curvePath
                GdipDeletePath curvePath
            End If

        Case egBars
            Dim barW As Long
            barW = CLng(stepPx * 0.7)
            If barW < 2 Then barW = 2
            For i = 0 To extraCount - 1
                Dim bY As Long, bH As Long
                bY = pts(i).Y
                bH = gREC.Top + gREC.Height - bY
                If bH > 0 Then
                    GdipFillRectangleI hG, hBrush, pts(i).X - barW \ 2, bY, barW, bH
                    GdipDrawRectangleI hG, hPen, pts(i).X - barW \ 2, bY, barW, bH
                End If
            Next i
    End Select

    ' Puntos marcadores
    Dim dotBrush As Long
    GdipCreateSolidFill ARGB(m_GraphPointColor, 90), dotBrush
    For i = 0 To m_AnimBufferCount - 1
        GdipFillEllipse hG, dotBrush, pts(i).X - 3 * nScale, pts(i).Y - 3 * nScale, _
                        6 * nScale, 6 * nScale
    Next i
    GdipDeleteBrush dotBrush
    GdipDeleteBrush hBrush
    GdipDeletePen hPen
End Sub

'--- Grid lines horizontales de referencia ---
Private Sub DrawGridLines(ByVal hG As Long, gREC As RECTL)
    Dim gridPen As Long
    Dim i As Long, lineY As Long
    GdipCreatePen1 ARGB(m_GraphGridColor, 30), 1, UnitPixel, gridPen
    For i = 1 To 3   ' 25%, 50%, 75%
        lineY = gREC.Top + CLng(gREC.Height * (i / 4))
        GdipDrawLineI hG, gridPen, gREC.Left, lineY, gREC.Left + gREC.Width, lineY
    Next i
    GdipDeletePen gridPen
End Sub

'--- DrawCaption: renderiza texto o icono ---
Private Function DrawCaption(ByVal hG As Long, sString As Variant, oFont As StdFont, _
                              layoutRect As RECTS, TextColor As OLE_COLOR, _
                              ColorOpacity As Long, mAngle As Single, _
                              HAlign As eTextAlignH, VAlign As eTextAlignV, _
                              Icon As Boolean) As Long
    Dim hPath As Long
    Dim hCaptBrush As Long
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

    GdipCreatePath 0, hPath
    If Icon Then
        GdipAddPathString hPath, StrPtr(ChrW2(sString)), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
    Else
        GdipAddPathString hPath, StrPtr(CStr(sString)), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
    End If

    GdipDeleteStringFormat hFormat
    GdipCreateSolidFill ARGB(TextColor, ColorOpacity), hCaptBrush
    GdipFillPath hG, hCaptBrush, hPath
    GdipDeleteBrush hCaptBrush
    GdipDeletePath hPath
    GdipDeleteFontFamily hFontFamily
End Function

'--- gRoundRect ---
Private Function gRoundRect(ByVal hG As Long, RECT As RECTL, ByVal Color1 As Long, _
                             ByVal Color2 As Long, ByVal Angulo As Single, _
                             ByVal BorderWidth As Long, ByVal BorderColor As Long, _
                             ByVal Round As Long, Filled As Boolean) As Long
    Dim hPenR As Long, hBrushR As Long, mPath As Long, mRound As Long
    If m_BorderWidth > 0 Then GdipCreatePen1 BorderColor, BorderWidth * nScale, &H2, hPenR
    If Filled Then GdipCreateLineBrushFromRectWithAngleI RECT, Color1, Color2, Angulo + 90, 0, WrapModeTileFlipXY, hBrushR
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
    If Filled Then GdipFillPath hG, hBrushR, mPath
    If m_BorderWidth > 0 Then GdipDrawPath hG, hPenR, mPath
    GdipDeletePath mPath
    GdipDeleteBrush hBrushR
    GdipDeletePen hPenR
End Function

'--- Helpers ---
Private Function GetSafeRound(ByVal Round As Long, ByVal MaxW As Long, ByVal MaxH As Long) As Long
    If Round > MaxW Then Round = MaxW
    If Round > MaxH Then Round = MaxH
    GetSafeRound = Round
End Function

Private Function GetFontStyleAndSize(oFont As StdFont, lFontStyle As Long, lFontSize As Long)
    On Error GoTo ErrO
    Dim hdcF As Long
    lFontStyle = 0
    If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
    hdcF = GetDC(0)
    lFontSize = CLng(oFont.Size * GetDeviceCaps(hdcF, LOGPIXELSY) / 72)
    ReleaseDC 0, hdcF
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
    Dim hdcD As Long, LPX As Double
    hdcD = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hdcD, LOGPIXELSX))
    ReleaseDC 0, hdcD
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

'--- Captions ---
Public Property Get Caption1() As String: Caption1 = m_Caption: End Property
Public Property Let Caption1(ByVal v As String): m_Caption = v: Refresh: End Property
Public Property Get Caption1Font() As StdFont: Set Caption1Font = m_Font: End Property
Public Property Set Caption1Font(ByVal v As StdFont): Set m_Font = v: Refresh: End Property
Public Property Get Caption1Color() As OLE_COLOR: Caption1Color = m_ForeColor: End Property
Public Property Let Caption1Color(ByVal v As OLE_COLOR): m_ForeColor = v: Refresh: End Property
Public Property Get Caption1AlignH() As eTextAlignH: Caption1AlignH = m_CaptionAlignH: End Property
Public Property Let Caption1AlignH(ByVal v As eTextAlignH): m_CaptionAlignH = v: Refresh: End Property

Public Property Get Caption2() As String: Caption2 = m_Caption2: End Property
Public Property Let Caption2(ByVal v As String): m_Caption2 = v: Refresh: End Property
Public Property Get Caption2Font() As StdFont: Set Caption2Font = m_Font2: End Property
Public Property Set Caption2Font(ByVal v As StdFont): Set m_Font2 = v: Refresh: End Property
Public Property Get Caption2Color() As OLE_COLOR: Caption2Color = m_ForeColor2: End Property
Public Property Let Caption2Color(ByVal v As OLE_COLOR): m_ForeColor2 = v: Refresh: End Property
Public Property Get Caption2AlignH() As eTextAlignH: Caption2AlignH = m_Caption2AlignH: End Property
Public Property Let Caption2AlignH(ByVal v As eTextAlignH): m_Caption2AlignH = v: Refresh: End Property

Public Property Get Caption3() As String: Caption3 = m_Caption3: End Property
Public Property Let Caption3(ByVal v As String): m_Caption3 = v: Refresh: End Property
Public Property Get Caption3Font() As StdFont: Set Caption3Font = m_Font3: End Property
Public Property Set Caption3Font(ByVal v As StdFont): Set m_Font3 = v: Refresh: End Property
Public Property Get Caption3Color() As OLE_COLOR: Caption3Color = m_ForeColor3: End Property
Public Property Let Caption3Color(ByVal v As OLE_COLOR): m_ForeColor3 = v: Refresh: End Property
Public Property Get Caption3AlignH() As eTextAlignH: Caption3AlignH = m_Caption3AlignH: End Property
Public Property Let Caption3AlignH(ByVal v As eTextAlignH): m_Caption3AlignH = v: Refresh: End Property

'--- Icon ---
Public Property Get IconCharCode() As String: IconCharCode = m_IconCharCode: End Property
Public Property Let IconCharCode(ByVal v As String): m_IconCharCode = v: Refresh: End Property
Public Property Get IconFont() As StdFont: Set IconFont = m_IconFont: End Property
Public Property Set IconFont(v As StdFont): Set m_IconFont = v: Refresh: End Property
Public Property Get IconForeColor() As OLE_COLOR: IconForeColor = m_IconForeColor: End Property
Public Property Let IconForeColor(ByVal v As OLE_COLOR): m_IconForeColor = v: Refresh: End Property
Public Property Get Boxed() As Boolean: Boxed = m_isBoxed: End Property
Public Property Let Boxed(ByVal v As Boolean): m_isBoxed = v: Refresh: End Property
Public Property Get BoxedColor() As OLE_COLOR: BoxedColor = m_BoxedColor: End Property
Public Property Let BoxedColor(ByVal v As OLE_COLOR): m_BoxedColor = v: Refresh: End Property

'--- Graph ---
Public Property Get GraphLine() As eGraph: GraphLine = m_Graph: End Property
Public Property Let GraphLine(ByVal v As eGraph): m_Graph = v: Refresh: End Property

Public Property Get GraphMatrix() As String: GraphMatrix = m_GraphMatrix: End Property
Public Property Let GraphMatrix(ByVal v As String)
    m_GraphMatrix = v
    m_AnimInitialized = False
    InitAnimBuffer
    Refresh
End Property

Public Property Get GraphMatrixTooltip() As String: GraphMatrixTooltip = m_GraphMatrixTooltip: End Property
Public Property Let GraphMatrixTooltip(ByVal v As String): m_GraphMatrixTooltip = v: End Property

Public Property Get GraphStyle() As eGraphStyle: GraphStyle = m_GraphStyle: End Property
Public Property Let GraphStyle(ByVal v As eGraphStyle): m_GraphStyle = v: Refresh: End Property

Public Property Get GraphLineColor() As OLE_COLOR: GraphLineColor = m_GraphLineColor: End Property
Public Property Let GraphLineColor(ByVal v As OLE_COLOR): m_GraphLineColor = v: Refresh: End Property

Public Property Get GraphBackColor() As OLE_COLOR: GraphBackColor = m_GraphBackColor: End Property
Public Property Let GraphBackColor(ByVal v As OLE_COLOR): m_GraphBackColor = v: Refresh: End Property

Public Property Get GraphPointColor() As OLE_COLOR: GraphPointColor = m_GraphPointColor: End Property
Public Property Let GraphPointColor(ByVal v As OLE_COLOR): m_GraphPointColor = v: Refresh: End Property

Public Property Get GraphFillOpacity() As Long: GraphFillOpacity = m_GraphFillOpacity: End Property
Public Property Let GraphFillOpacity(ByVal v As Long)
    If v < 0 Then v = 0
    If v > 100 Then v = 100
    m_GraphFillOpacity = v: Refresh
End Property

Public Property Get GraphGridLines() As Boolean: GraphGridLines = m_GraphGridLines: End Property
Public Property Let GraphGridLines(ByVal v As Boolean): m_GraphGridLines = v: Refresh: End Property

Public Property Get GraphGridColor() As OLE_COLOR: GraphGridColor = m_GraphGridColor: End Property
Public Property Let GraphGridColor(ByVal v As OLE_COLOR): m_GraphGridColor = v: Refresh: End Property

'--- Animation ---
Public Property Get AnimEnabled() As Boolean: AnimEnabled = m_AnimEnabled: End Property
Public Property Let AnimEnabled(ByVal v As Boolean)
    m_AnimEnabled = v
    tmrAnim.Enabled = v
End Property

Public Property Get AnimMode() As eAnimMode: AnimMode = m_AnimMode: End Property
Public Property Let AnimMode(ByVal v As eAnimMode)
    m_AnimMode = v
    m_AnimOffset = 0
    InitAnimBuffer
End Property

Public Property Get AnimSpeed() As Long: AnimSpeed = m_AnimSpeed: End Property
Public Property Let AnimSpeed(ByVal v As Long)
    If v < 10 Then v = 10
    If v > 1000 Then v = 1000
    m_AnimSpeed = v
    tmrAnim.Interval = v
End Property

Public Property Get AnimStep() As Long: AnimStep = m_AnimStep: End Property
Public Property Let AnimStep(ByVal v As Long)
    If v < 1 Then v = 1
    If v > 20 Then v = 20
    m_AnimStep = v
End Property

Public Property Get AnimSmooth() As Boolean: AnimSmooth = m_AnimSmooth: End Property
Public Property Let AnimSmooth(ByVal v As Boolean): m_AnimSmooth = v: End Property

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
    m_Clickable = False
    m_Transparent = False
    m_Filled = True
    m_Opacity = 100
    m_InitialOpacity = 100
    m_EffectFade = False

    m_Color1 = RGB(20, 20, 35)
    m_Color2 = RGB(20, 20, 35)
    m_def_Color1 = m_Color1
    m_def_Color2 = m_Color2
    m_Angulo = 0
    m_BorderColor = RGB(50, 50, 75)
    m_BorderWidth = 1
    m_CornerCurve = 8

    Set m_Font = UserControl.Ambient.Font
    Set m_Font2 = UserControl.Ambient.Font
    Set m_Font3 = UserControl.Ambient.Font
    Set m_IconFont = UserControl.Font

    m_ForeColor = vbWhite
    m_ForeColor2 = vbWhite
    m_ForeColor3 = RGB(150, 150, 180)
    m_def_ForeColor = vbWhite
    m_CaptionAlignH = eLeft
    m_Caption2AlignH = eRight
    m_Caption3AlignH = eLeft
    m_Caption = Ambient.DisplayName
    m_Caption2 = ""
    m_Caption3 = ""

    m_Graph = egCurvedLine
    m_GraphMatrix = "20,45,30,70,55,80,65,40,75,60"
    m_GraphStyle = gsGradient
    m_GraphLineColor = RGB(99, 102, 241)
    m_GraphBackColor = RGB(99, 102, 241)
    m_GraphPointColor = RGB(255, 255, 255)
    m_GraphFillOpacity = 35
    m_GraphGridLines = False
    m_GraphGridColor = RGB(100, 100, 130)

    m_AnimEnabled = True
    m_AnimMode = amLoop
    m_AnimSpeed = 60
    m_AnimStep = 2
    m_AnimSmooth = True
    m_AnimOffset = 0
    m_AnimInitialized = False

    m_isBoxed = False
    m_BoxedColor = RGB(99, 102, 241)
    m_IconForeColor = vbWhite
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_Enabled = .ReadProperty("Enabled", True)
        m_Color1 = .ReadProperty("BackColor1", m_def_Color1)
        m_Color2 = .ReadProperty("BackColor2", m_def_Color2)
        m_Angulo = .ReadProperty("BackAngle", 0)
        m_BorderColor = .ReadProperty("BorderColor", RGB(50, 50, 75))
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
        m_CaptionAlignH = .ReadProperty("Caption1AlignH", eLeft)
        m_Caption2 = .ReadProperty("Caption2", "")
        Set m_Font2 = .ReadProperty("Caption2Font", UserControl.Ambient.Font)
        m_ForeColor2 = .ReadProperty("Caption2Color", vbWhite)
        m_Caption2AlignH = .ReadProperty("Caption2AlignH", eRight)
        m_Caption3 = .ReadProperty("Caption3", "")
        Set m_Font3 = .ReadProperty("Caption3Font", UserControl.Ambient.Font)
        m_ForeColor3 = .ReadProperty("Caption3Color", RGB(150, 150, 180))
        m_Caption3AlignH = .ReadProperty("Caption3AlignH", eLeft)

        m_IconCharCode = .ReadProperty("IconCharCode", "")
        Set m_IconFont = .ReadProperty("IconFont", UserControl.Font)
        m_IconForeColor = .ReadProperty("IconForeColor", vbWhite)
        m_isBoxed = .ReadProperty("Boxed", False)
        m_BoxedColor = .ReadProperty("BoxedColor", RGB(99, 102, 241))

        m_Graph = .ReadProperty("GraphLine", egCurvedLine)
        m_GraphMatrix = .ReadProperty("GraphMatrix", "20,45,30,70,55,80,65,40,75,60")
        m_GraphMatrixTooltip = .ReadProperty("GraphMatrixTooltip", "")
        m_GraphStyle = .ReadProperty("GraphStyle", gsGradient)
        m_GraphLineColor = .ReadProperty("GraphLineColor", RGB(99, 102, 241))
        m_GraphBackColor = .ReadProperty("GraphBackColor", RGB(99, 102, 241))
        m_GraphPointColor = .ReadProperty("GraphPointColor", vbWhite)
        m_GraphFillOpacity = .ReadProperty("GraphFillOpacity", 35)
        m_GraphGridLines = .ReadProperty("GraphGridLines", False)
        m_GraphGridColor = .ReadProperty("GraphGridColor", RGB(100, 100, 130))

        m_AnimEnabled = .ReadProperty("AnimEnabled", True)
        m_AnimMode = .ReadProperty("AnimMode", amLoop)
        m_AnimSpeed = .ReadProperty("AnimSpeed", 60)
        m_AnimStep = .ReadProperty("AnimStep", 2)
        m_AnimSmooth = .ReadProperty("AnimSmooth", True)
        m_Tag = .ReadProperty("Tag", "")
    End With
    m_AnimOffset = 0
    m_AnimInitialized = False
    InitAnimBuffer
    tmrAnim.Interval = m_AnimSpeed
    tmrAnim.Enabled = m_AnimEnabled
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "Enabled", m_Enabled, True
        .WriteProperty "BackColor1", m_Color1, m_def_Color1
        .WriteProperty "BackColor2", m_Color2, m_def_Color2
        .WriteProperty "BackAngle", m_Angulo, 0
        .WriteProperty "BorderColor", m_BorderColor, RGB(50, 50, 75)
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
        .WriteProperty "Caption1AlignH", m_CaptionAlignH, eLeft
        .WriteProperty "Caption2", m_Caption2, ""
        .WriteProperty "Caption2Font", m_Font2, UserControl.Ambient.Font
        .WriteProperty "Caption2Color", m_ForeColor2, vbWhite
        .WriteProperty "Caption2AlignH", m_Caption2AlignH, eRight
        .WriteProperty "Caption3", m_Caption3, ""
        .WriteProperty "Caption3Font", m_Font3, UserControl.Ambient.Font
        .WriteProperty "Caption3Color", m_ForeColor3, RGB(150, 150, 180)
        .WriteProperty "Caption3AlignH", m_Caption3AlignH, eLeft

        .WriteProperty "IconCharCode", m_IconCharCode, ""
        .WriteProperty "IconFont", m_IconFont, UserControl.Font
        .WriteProperty "IconForeColor", m_IconForeColor, vbWhite
        .WriteProperty "Boxed", m_isBoxed, False
        .WriteProperty "BoxedColor", m_BoxedColor, RGB(99, 102, 241)

        .WriteProperty "GraphLine", m_Graph, egCurvedLine
        .WriteProperty "GraphMatrix", m_GraphMatrix, "20,45,30,70,55,80,65,40,75,60"
        .WriteProperty "GraphMatrixTooltip", m_GraphMatrixTooltip, ""
        .WriteProperty "GraphStyle", m_GraphStyle, gsGradient
        .WriteProperty "GraphLineColor", m_GraphLineColor, RGB(99, 102, 241)
        .WriteProperty "GraphBackColor", m_GraphBackColor, RGB(99, 102, 241)
        .WriteProperty "GraphPointColor", m_GraphPointColor, vbWhite
        .WriteProperty "GraphFillOpacity", m_GraphFillOpacity, 35
        .WriteProperty "GraphGridLines", m_GraphGridLines, False
        .WriteProperty "GraphGridColor", m_GraphGridColor, RGB(100, 100, 130)

        .WriteProperty "AnimEnabled", m_AnimEnabled, True
        .WriteProperty "AnimMode", m_AnimMode, amLoop
        .WriteProperty "AnimSpeed", m_AnimSpeed, 60
        .WriteProperty "AnimStep", m_AnimStep, 2
        .WriteProperty "AnimSmooth", m_AnimSmooth, True
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
    If m_Clickable Then RaiseEvent Click(0)
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

'--- Timer: animacion del grafico ---
Private Sub tmrAnim_Timer()
    If Not m_AnimInitialized Then InitAnimBuffer: Exit Sub
    If m_AnimBufferCount < 2 Then Exit Sub

    Dim stepPx As Single
    stepPx = UserControl.ScaleWidth / (m_AnimBufferCount - 1)

    ' Avanzar offset
    m_AnimOffset = m_AnimOffset + m_AnimStep

    ' Cuando el offset supera el ancho de un paso, rotar el buffer
    If m_AnimOffset >= stepPx Then
        m_AnimOffset = m_AnimOffset - stepPx
        If m_AnimMode = amLoop Then
            RotateBufferLeft
        End If
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
