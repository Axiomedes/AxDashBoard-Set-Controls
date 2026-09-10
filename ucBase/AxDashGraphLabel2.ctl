VERSION 5.00
Begin VB.UserControl AxDGraphLabel2 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H008D4214&
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   75
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   282
   ToolboxBitmap   =   "AxDashGraphLabel2.ctx":0000
   Begin VB.Timer tmrEffect 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   75
      Top             =   75
   End
End
Attribute VB_Name = "AxDGraphLabel2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-UC-VB6-----------------------------
'UC Name  : AxDGraphLabel2
'Version  : 0.01
'Editor   : David Rojas [AxioUK]
'Date     : 17/07/2021
'------------------------------------
Option Explicit

Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function TlsGetValue Lib "kernel32.dll" (ByVal dwTlsIndex As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'-
'Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long

'Private Declare Function ReleaseCapture Lib "User32" () As Long
'Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

'Private Declare Function LoadCursor Lib "User32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
'Private Declare Function DestroyCursor Lib "User32" (ByVal hCursor As Long) As Long
'Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal Token As Long)
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As Long, ByRef mLineGradient As Long) As Long
'Private Declare Function GdipDrawRectangle Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function GdipFillRectangle Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal brush As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As Long
'Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function GdipFillRectangleI Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal brush As Long, ByVal x As Long, ByVal y As Long, ByVal Width As Long, ByVal Height As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal mhDC As Long, ByRef mGraphics As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal brush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
'Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal ARGB As Long, ByRef brush As Long) As Long
Private Declare Function GdipAddPathString Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFamily As Long, ByVal mStyle As Long, ByVal mEmSize As Single, ByRef mLayoutRect As RECTS, ByVal mFormat As Long) As Long
'Private Declare Function GdipMeasureString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTS, ByVal mStringFormat As Long, ByRef mBoundingBox As RECTS, ByRef mCodepointsFitted As Long, ByRef mLinesFilled As Long) As Long
'Private Declare Function GdipCreateFont Lib "GdiPlus.dll" (ByVal mFontFamily As Long, ByVal mEmSize As Single, ByVal mStyle As Long, ByVal mUnit As Long, ByRef mFont As Long) As Long
'Private Declare Function GdipDeleteFont Lib "GdiPlus.dll" (ByVal mFont As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "GdiPlus.dll" (ByRef mNativeFamily As Long) As Long
'Private Declare Function GdipDrawString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTS, ByVal mStringFormat As Long, ByVal mBrush As Long) As Long
Private Declare Function GdipSetStringFormatTrimming Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mTrimming As eStringTrimming) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatFlags Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mFlags As eStringFormatFlags) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As eStringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mAlign As eStringAlignment) As Long
Private Declare Function GdipDeleteStringFormat Lib "GdiPlus.dll" (ByVal mFormat As Long) As Long
Private Declare Function GdipDrawLineI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
'Private Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByRef pPoints As Any, ByVal count As Long) As Long
'Private Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByRef pPoints As Any, ByVal count As Long, ByVal FillMode As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "GdiPlus.dll" (ByVal graphics As Long) As Long
'Private Declare Function GdipResetPath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
'Private Declare Function GdipAddPathCurve Lib "gdiplus" (ByVal path As Long, pPoints As Any, ByVal count As Long) As Long
'Private Declare Function GdipDrawEllipse Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single) As Long
Private Declare Function GdipDrawEllipseI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Long
Private Declare Function GdipFillEllipse Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single) As Long
Private Declare Function GdipFillEllipseI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
'Private Declare Function GdipDrawClosedCurve Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, POINTS As POINTS, ByVal count As Long) As Long
'Private Declare Function GdipFillClosedCurve Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, POINTS As POINTS, ByVal count As Long) As Long
Private Declare Function GdipDrawCurve Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, POINTS As POINTS, ByVal count As Long) As Long
Private Declare Function GdipDrawPie Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipFillPie Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single, ByVal startAngle As Single, ByVal sweepAngle As Single) As Long
Private Declare Function GdipAddPathEllipse Lib "gdiplus" (ByVal path As Long, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single) As Long
Private Declare Function GdipAddPathEllipseI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long) As Long
Private Declare Function GdipSetPathGradientCenterColor Lib "GdiPlus.dll" (ByVal mBrush As Long, ByVal mColors As Long) As Long
Private Declare Function GdipSetPathGradientSurroundColorsWithCount Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mColor As Long, ByRef mCount As Long) As Long
Private Declare Function GdipCreatePathGradientFromPath Lib "GdiPlus.dll" (ByVal mPath As Long, ByRef mPolyGradient As Long) As Long
Private Declare Function GdipSetLinePresetBlend Lib "GdiPlus.dll" (ByVal mBrush As Long, ByRef mBlend As Long, ByRef mPositions As Single, ByVal mCount As Long) As Long
'---
'Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpszProgID As Long, pCLSID As Any) As Long
'Private Declare Function GdipCreateFromHWND Lib "gdiplus" (ByVal hWnd As Long, graphics As Long) As Long
'Private Declare Function GdipCreateBitmapFromGraphics Lib "gdiplus" (ByVal Width As Long, ByVal Height As Long, ByVal graphics As Long, bitmap As Long) As Long
'Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal hImage As Long, ByVal sFileName As String, clsidEncoder As Any, encoderParams As Any) As Long
'---
'Private Declare Function DrawTextW Lib "user32.dll" (ByVal hdc As Long, lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTL) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECTS, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Type RECTL
    Left As Long
    Top As Long
    width As Long
    height As Long
End Type

Private Type RECTS
    Left As Single
    Top As Single
    width As Single
    height As Single
End Type

Private Type POINTS
   x As Single
   y As Single
End Type

Private Type POINTL
    x As Long
    y As Long
End Type

Private Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum

'Private Const LF_FACESIZE = 32
'Private Const SYSTEM_FONT = 13
'Private Const OBJ_FONT As Long = 6&

Private Type GDIPlusStartupInput
    GdiPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

'Private Enum HotkeyPrefix
'    HotkeyPrefixNone = &H0
'    HotkeyPrefixShow = &H1
'    HotkeyPrefixHide = &H2
'End Enum

'Private Type PicBmp
'  Size As Long
'  type As Long
'  hBmp As Long
'  hpal As Long
'  Reserved As Long
'End Type

'Private Enum PenAlignment
'    PenAlignmentCenter = &H0
'    PenAlignmentInset = &H1
'End Enum

'EVENTS------------------------------------
Public Event Click(ByVal Serie As Long)
Public Event DblClick()
'Public Event ChangeValue(ByVal Value As Boolean)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
''-----------------------------------------

'Private Const CombineModeExclude As Long = &H4
Private Const WrapModeTileFlipXY = &H3
'Private Const SmoothingModeHighQuality As Long = &H2
Private Const SmoothingModeAntiAlias As Long = &H4
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const TLS_MINIMUM_AVAILABLE As Long = 64
'Private Const CLR_INVALID = -1
'Private Const WM_NCLBUTTONDOWN = &HA1
'Private Const HTCAPTION = 2
Private Const IDC_HAND As Long = 32649
Private Const UnitPixel As Long = &H2&

'Default Property Values:
Const m_def_ForeColor = 0
Const m_def_Appearance = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_Color1 = &HD59B5B
Const m_def_Color2 = &H6A5444
Const m_def_Angulo = 0

'Property Variables:
Private GdipToken As Long
Private nScale    As Single
Private hCur      As Long
Private hFontCollection As Long
Private hGraphics As Long
Private hPen      As Long
'Private hBrush    As Long

Private m_BorderColor   As OLE_COLOR
Private m_ForeColor     As OLE_COLOR
Private m_ForeColor2    As OLE_COLOR
Private m_Color1        As OLE_COLOR
Private m_Color2        As OLE_COLOR
Private m_BorderWidth   As Long
Private m_Enabled       As Boolean
Private m_Angulo        As Single
Private m_CornerCurve   As Long
'Private m_CaptionAngle  As Single
Private m_Caption       As String
Private m_Caption2      As String
'Private m_Top         As Long
'Private m_Left        As Long
Private m_Opacity     As Long
Private cl_hWnd       As Long

Private m_OldBorderColor As OLE_COLOR
Private m_BoxedColor As OLE_COLOR
Private m_ChangeBorderOnFocus As Boolean
'Private m_OnFocus As Boolean

Private m_Font          As StdFont
Private m_Font2         As StdFont
Private m_IconFont      As StdFont
Private m_IconCharCode  As Long
Private m_IconForeColor As Long
Private m_IconAlignH    As eTextAlignH
Private m_IconAlignV    As eTextAlignV
Private m_PosY          As Single
Private m_PosX          As Single
'Private m_MouseOver     As Boolean
Private m_CaptionAlignV As eTextAlignV
Private m_CaptionAlignH As eTextAlignH
Private m_Caption2AlignV As eTextAlignV
Private m_Caption2AlignH As eTextAlignH

'Private m_StringPosX  As Long
'Private m_StringPosY  As Long
Private m_EffectFade  As Boolean
Private m_InitialOpacity As Long
Private m_Transparent As Boolean
'--- Mejoras Fase 4 ---
Private m_Tag As String
Private m_DropShadow As Boolean
Private m_ShadowColor As OLE_COLOR
Private m_ShadowDepth As Long
Private m_ShadowOpacity As Long
Private m_GlassEffect As Boolean
Private m_PieHoleSize As Single
Private m_PieShowPercent As Boolean
Private m_GraphFillOpacity As Long
Private m_GraphGridLines As Boolean
Private m_GraphGridColor As OLE_COLOR

Private iPts()          As POINTS
Private gREC()          As RECTL
Private m_GraphMatrix   As String
Private m_Graph         As eGraph2
Private m_GraphStyle    As eGraphStyle
Private m_GraphLineColor  As OLE_COLOR
Private m_GraphBackColor  As OLE_COLOR
Private m_GraphPointColor As OLE_COLOR
Private m_ToolTip()       As String
Private sToolTip          As String
Private mSerie            As Long

Private m_isBoxed      As Boolean
Private m_IconBox       As eBoxed
Private m_Clicked      As Boolean
Private m_Clickable    As Boolean
'Private m_OptionButton As Boolean
Private m_Filled       As Boolean

Private m_GraphMatrixTooltip As String
Private mShowTtp As Boolean


Public Sub CopyAmbient()
Dim OPic As StdPicture

On Error GoTo Err

    With UserControl
        Set .Picture = Nothing
        Set OPic = Extender.Container.Image
        .BackColor = Extender.Container.BackColor
        UserControl.PaintPicture OPic, 0, 0, , , Extender.Left, Extender.Top
        Set .Picture = .Image
    End With
Err:
End Sub

Public Sub Refresh()
  UserControl.Cls
  CopyAmbient
  Draw
End Sub

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

Public Function ChrW2(ByVal CharCode As Long) As String
  Const POW10 As Long = 2 ^ 10
  If CharCode <= &HFFFF& Then ChrW2 = ChrW$(CharCode) Else _
                              ChrW2 = ChrW$(&HD800& + (CharCode And &HFFFF&) \ POW10) & _
                                      ChrW$(&HDC00& + (CharCode And (POW10 - 1)))
End Function

Private Sub Draw()
Dim IcoBox As RECTS
Dim tpRec As RECTL
Dim REC As RECTL
Dim Grp As RECTL
Dim BOX As RECTL
Dim stREC As RECTS
Dim stREC2 As RECTS
'Dim stREC3 As RECTS
Dim spREC As RECTS
Dim lBorder As Long, mBorder As Long
Dim strSeccion As Long

With UserControl
    
  GdipCreateFromHDC .hdc, hGraphics
  GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias

  If m_DropShadow Then DrawDropShadow hGraphics, REC, m_ShadowColor, m_ShadowOpacity, m_ShadowDepth, m_CornerCurve
  lBorder = m_BorderWidth * 2
  mBorder = lBorder / 2
  strSeccion = (.ScaleHeight / 5)
    
  REC.Left = 1:     REC.Top = 1
  REC.width = .ScaleWidth - 2
  REC.height = .ScaleHeight - 2
  
  Select Case m_IconBox
    Case ebLeft
      BOX.Left = (.ScaleHeight / 20) + mBorder * nScale
    Case ebRight
      BOX.Left = .ScaleWidth - (.ScaleHeight / 20) - (strSeccion * 2) - mBorder * nScale
  End Select
  BOX.Top = (.ScaleHeight / 20) + mBorder * nScale
  BOX.width = (strSeccion * 2) * nScale
  BOX.height = (strSeccion * 2) * nScale

  stREC.Left = IIf(m_IconBox = ebLeft, BOX.Left + BOX.width + 3, mBorder + 3) * nScale
  stREC.Top = mBorder + 6 * nScale
  stREC.width = .ScaleWidth - (lBorder + BOX.width + 8) * nScale
  stREC.height = strSeccion * nScale
  
  stREC2.Left = IIf(m_IconBox = ebLeft, BOX.Left + BOX.width + 3, mBorder + 3) * nScale
  stREC2.Top = strSeccion + mBorder + 5 * nScale
  stREC2.width = .ScaleWidth - (lBorder + BOX.width + 8) * nScale
  stREC2.height = strSeccion * nScale
    
  Grp.Left = lBorder + 10 * nScale
  Grp.Top = (strSeccion * 2) + mBorder + 10 * nScale
  Grp.width = (.ScaleWidth / 10) * 7 * nScale
  Grp.height = (strSeccion * 3) - (lBorder + 12) * nScale
      
  IcoBox.Left = BOX.Left - 2: IcoBox.Top = BOX.Top - 2
  IcoBox.width = BOX.width + 5: IcoBox.height = BOX.height + 5
      
  SafeRange m_Opacity, 0, 100
  
  '-DRAW Control--------------
  If m_EffectFade Then
    gRoundRect hGraphics, REC, ARGB(m_Color1, m_Opacity), ARGB(m_Color2, m_Opacity), m_Angulo, m_BorderWidth, ARGB(m_BorderColor, m_Opacity), m_CornerCurve, m_Filled
    'Graph
    DrawGraphic hGraphics, m_Graph, m_GraphMatrix, Grp, ARGB(m_GraphBackColor, 50), 0, ARGB(m_GraphLineColor, 50), 2, m_GraphStyle, True, ARGB(m_GraphPointColor, 80)
    'IconBox
    If m_isBoxed Then gRoundRect hGraphics, BOX, ARGB(m_BoxedColor, m_Opacity), ARGB(m_BoxedColor, m_Opacity), m_Angulo, m_BorderWidth, ARGB(m_BoxedColor, m_Opacity), m_CornerCurve, m_Filled
    'Icon
    DrawCaption hGraphics, IconCharCode, IconFont, IcoBox, m_IconForeColor, m_Opacity, 0, m_IconAlignH, m_IconAlignV, True
    'Captions
    DrawCaption hGraphics, m_Caption, m_Font, stREC, m_ForeColor, m_Opacity, 0, m_CaptionAlignH, m_CaptionAlignV, False
    DrawCaption hGraphics, m_Caption2, m_Font2, stREC2, m_ForeColor2, m_Opacity, 0, m_Caption2AlignH, m_Caption2AlignV, False
  Else
  'Back
    gRoundRect hGraphics, REC, ARGB(m_Color1, 100), ARGB(m_Color2, 100), m_Angulo, m_BorderWidth, ARGB(m_BorderColor, 100), m_CornerCurve, m_Filled
    'Graph
    DrawGraphic hGraphics, m_Graph, m_GraphMatrix, Grp, ARGB(m_GraphBackColor, 50), 0, ARGB(m_GraphLineColor, 50), 2, m_GraphStyle, True, ARGB(m_GraphPointColor, 80)
    'IconBox
    If m_isBoxed Then gRoundRect hGraphics, BOX, ARGB(m_BoxedColor, 100), ARGB(m_BoxedColor, 100), m_Angulo, m_BorderWidth, ARGB(m_BoxedColor, 100), m_CornerCurve, m_Filled
    'Icon
    DrawCaption hGraphics, IconCharCode, IconFont, IcoBox, m_IconForeColor, CLng(100), 0, m_IconAlignH, m_IconAlignV, True
    'Captions
    DrawCaption hGraphics, m_Caption, m_Font, stREC, m_ForeColor, 100, 0, m_CaptionAlignH, m_CaptionAlignV, False
    DrawCaption hGraphics, m_Caption2, m_Font2, stREC2, m_ForeColor2, 100, 0, m_Caption2AlignH, m_Caption2AlignV, False
  End If
' '---------------
  'ToolTip
    tpRec.Left = m_PosX
    tpRec.Top = m_PosY - (.TextHeight(sToolTip) + 5)
    tpRec.width = .TextWidth(sToolTip) + 10: tpRec.height = .TextHeight(sToolTip) + 5
    spREC.Left = tpRec.Left:   spREC.Top = tpRec.Top
    spREC.width = tpRec.width: spREC.height = tpRec.height
    
  If mShowTtp Then
    Debug.Print sToolTip
    gRoundRect hGraphics, tpRec, ARGB(vbWhite, 100), ARGB(vbWhite, 100), m_Angulo, 1, ARGB(vbBlack, 100), 3, m_Filled
    DrawCaption hGraphics, sToolTip, .Font, spREC, vbBlack, 100, 0, eCenter, eMiddle, False
  End If
' '---------------
  If m_GlassEffect Then DrawGlassEffect hGraphics, REC, m_CornerCurve
  If m_GraphGridLines Then DrawGridLines hGraphics, REC
  GdipDeleteGraphics hGraphics
  '---------------
  If m_Transparent Then
    .BackStyle = 0
    .MaskColor = .BackColor
    Set .MaskPicture = .Image
  End If
  '---------------
End With

End Sub

Private Function DrawCaption(ByVal hGraphics As Long, sString As Variant, oFont As StdFont, layoutRect As RECTS, _
                             TextColor As OLE_COLOR, ColorOpacity As Long, mAngle As Single, HAlign As eTextAlignH, _
                             VAlign As eTextAlignV, Icon As Boolean) As Long
Dim hPath As Long
'Dim hPen As Long
Dim hBrush As Long
Dim hFontFamily As Long
Dim hFormat As Long
Dim lFontSize As Long
Dim lFontStyle As GDIPLUS_FONTSTYLE
'Dim hFont As Long
Dim newY As Long, newX As Long

On Error Resume Next

    If GdipCreatePath(&H0, hPath) = 0 Then

        If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
            GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap 'bWordWrap
            GdipSetStringFormatTrimming hFormat, StringTrimmingEllipsisWord
            GdipSetStringFormatAlign hFormat, HAlign
            GdipSetStringFormatLineAlign hFormat, VAlign
        End If

        GetFontStyleAndSize oFont, lFontStyle, lFontSize

        If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) Then
            If hFontCollection Then
                If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), hFontCollection, hFontFamily) Then
                    If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
                End If
            Else
                If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
            End If
        End If
'------------------------------------------------------------------------
        If mAngle <> 0 Then
            newY = (layoutRect.height / 2)
            newX = (layoutRect.width / 2)
            Call GdipTranslateWorldTransform(hGraphics, newX, newY, 0)
            Call GdipRotateWorldTransform(hGraphics, mAngle, 0)
            Call GdipTranslateWorldTransform(hGraphics, -newX, -newY, 0)
        End If
'------------------------------------------------------------------------
      If Icon Then
        GdipAddPathString hPath, StrPtr(ChrW2(sString)), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
      Else
        GdipAddPathString hPath, StrPtr(sString), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
      End If
'------------------------------------------------------------------------
        GdipDeleteStringFormat hFormat

        GdipCreateSolidFill ARGB(TextColor, ColorOpacity), hBrush

        GdipFillPath hGraphics, hBrush, hPath
        GdipDeleteBrush hBrush

        If mAngle <> 0 Then GdipResetWorldTransform hGraphics

        GdipDeleteFontFamily hFontFamily
            
        GdipDeletePath hPath
    End If

End Function

Private Sub DrawGraphic(ByVal iGraphics As Long, mShape As eGraph2, GrMatrix As String, Rct As RECTL, _
                      BackColor As Long, Angulo As Single, LineColor As Long, BorderW As Long, _
                      BackStyle As eGraphStyle, POINTS As Boolean, PointColor As Long)
Dim i As Integer
Dim pBrush As Long
Dim hBrush As Long
Dim hPen As Long
Dim x As Long, y As Long
Dim W As Long, H As Long
Dim Pt() As String
Dim p As Long
Dim s As Long

x = Rct.Left:  y = Rct.Top
W = Rct.width: H = Rct.height

If GrMatrix = "0" Or GrMatrix = "" Or GrMatrix = vbNullString Then Exit Sub

GdipCreatePen1 LineColor, BorderW, UnitPixel, hPen

Pt = Split(GrMatrix, ",")
p = UBound(Pt)
s = W / (p + 1)

ReDim iPts(p) As POINTS

Select Case mShape
    Case eRectLine, eCurvedLine
  
    For i = 0 To p
      iPts(i).x = x + (W / p) * i
      iPts(i).y = y + (H / 100) * (100 - CInt(Pt(i)))
    Next i
  
    If mShape = egRectLine Then
          'Rect Graph
        For i = 0 To p - 1
            GdipDrawLineI iGraphics, hPen, iPts(i).x, iPts(i).y, iPts(i + 1).x, iPts(i + 1).y
        Next i
    Else
          'Curved Graph
          GdipDrawCurve iGraphics, hPen, iPts(0), p + 1
    End If
        
    Case eBars
          
    ReDim gREC(p) As RECTL
    
      For i = 0 To p
        gREC(i).Left = (x + s * i) + 3
        gREC(i).Top = y + (H / 100) * (100 - CInt(Pt(i)))
        gREC(i).height = (H - iPts(i).y + 2)
        gREC(i).width = s - 2
        If BackStyle = gsGradient Then
          gRoundRect iGraphics, gREC(i), BackColor, ARGB(vbWhite, 40), Angulo, BorderW, LineColor, 1, True
        Else
          gRoundRect iGraphics, gREC(i), BackColor, BackColor, Angulo, BorderW, LineColor, 1, True
        End If
      Next i
      GoTo zEnd
      
    Case ePie
       Dim nW As Single, nH As Single
       Dim startAngle  As Single
       Dim sweepAngle As Single, UsedAngle As Single
       Dim Ptotal As Single
       Dim mPath As Long
       
       For i = 0 To p
        Ptotal = Ptotal + Val(Pt(i))
       Next i
       Debug.Print Ptotal
              
       'Define the pie.
       x = Rct.Left:    y = Rct.Top
       nW = Rct.width:  nH = Rct.height
       
      If BackStyle = gsGradient Then
        Call GdipCreatePath(&H0, mPath)
        GdipAddPathEllipseI mPath, x, y, nW, nH
        GdipCreatePathGradientFromPath mPath, hBrush
        GdipSetPathGradientCenterColor hBrush, BackColor
        GdipSetPathGradientSurroundColorsWithCount hBrush, PointColor, 1
        'Call GdipFillEllipseI(hGraphics, hBrush, x, y, Width, Height)
      Else
        GdipCreateSolidFill BackColor, hBrush
      End If
      
       For i = 0 To p
          If i = 0 Then
            startAngle = 0
          Else
            UsedAngle = UsedAngle + (360 / Ptotal) * CInt(Pt(i - 1))
            startAngle = UsedAngle
          End If
          sweepAngle = (360 / Ptotal) * CInt(Pt(i))
          Debug.Print "Start:" & startAngle & " | Sweep:" & sweepAngle
          'Draw the pie
          GdipDrawPie iGraphics, hPen, x, y, nW, nH, startAngle, sweepAngle
          GdipFillPie iGraphics, hBrush, x, y, nW, nH, startAngle, sweepAngle
                    'If BackStyle = gsSolid Then GdipFillPie iGraphics, hBrush, x, y, nW, nH, startAngle, sweepAngle
       Next i
       
       '--- Donut hole cutout ---
       If m_PieHoleSize > 0 Then
          Dim holeDiam As Single, holeX As Single, holeY As Single
          Dim hHoleBrush As Long, hHolePen As Long
          holeDiam = (IIf(nW < nH, nW, nH) * m_PieHoleSize) / 100
          holeX = x + (nW - holeDiam) / 2
          holeY = y + (nH - holeDiam) / 2
          GdipCreateSolidFill ARGB(m_Color1, 100), hHoleBrush
          GdipFillEllipseI iGraphics, hHoleBrush, CLng(holeX), CLng(holeY), CLng(holeDiam), CLng(holeDiam)
          GdipDeleteBrush hHoleBrush
          If BorderW > 0 Then
              GdipCreatePen1 LineColor, BorderW, UnitPixel, hHolePen
              GdipDrawEllipseI iGraphics, hHolePen, CLng(holeX), CLng(holeY), CLng(holeDiam), CLng(holeDiam)
              GdipDeletePen hHolePen
          End If
       End If
        
        Call GdipDeleteBrush(hBrush)
        Call GdipDeletePath(mPath)
        GoTo zEnd
End Select


zPoints:
  If POINTS Then
    GdipCreateSolidFill PointColor, pBrush
    For i = 0 To UBound(iPts)
      GdipFillEllipse iGraphics, pBrush, iPts(i).x - 4, iPts(i).y - 4, 8, 8
    Next i
    Call GdipDeleteBrush(pBrush)
  End If

zEnd:
  
  Call GdipDeletePen(hPen)
  
End Sub

Private Function GetFontStyleAndSize(oFont As StdFont, lFontStyle As Long, lFontSize As Long)
On Error GoTo ErrO
    Dim hdc As Long
    lFontStyle = 0
    If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
    
    hdc = GetDC(0&)
    lFontSize = MulDiv(oFont.Size, GetDeviceCaps(hdc, LOGPIXELSY), 72)
    ReleaseDC 0&, hdc
ErrO:
End Function

Private Function GetSafeRound(Angle As Integer, width As Long, height As Long) As Integer
    Dim lRet As Integer
    lRet = Angle
    If lRet * 2 > height Then lRet = height \ 2
    If lRet * 2 > width Then lRet = width \ 2
    GetSafeRound = lRet
End Function

'Private Function GetSystemHandCursor() As Picture
'  Dim Pic As PicBmp
'  Dim IPic As IPicture
'  Dim GUID(0 To 3) As Long
'
'  If hCur Then DestroyCursor hCur: hCur = 0
'
'  hCur = LoadCursor(ByVal 0&, IDC_HAND)
'
'  GUID(0) = &H7BF80980
'  GUID(1) = &H101ABF32
'  GUID(2) = &HAA00BB8B
'  GUID(3) = &HAB0C3000
'
'  With Pic
'    .Size = Len(Pic)
'    .type = vbPicTypeIcon
'    .hBmp = hCur
'    .hpal = 0
'  End With
'
'  Call OleCreatePictureIndirect(Pic, GUID(0), 1, IPic)
'
'  Set GetSystemHandCursor = IPic
'End Function

Private Function GetWindowsDPI() As Double
    Dim hdc As Long, LPX  As Double, LPY As Double
    hdc = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
    LPY = CDbl(GetDeviceCaps(hdc, LOGPIXELSY))
    ReleaseDC 0, hdc

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function


Private Sub DrawDropShadow(ByVal hG As Long, RECT As RECTL, ByVal sColor As Long, ByVal sOpacity As Long, ByVal sDepth As Long, ByVal sRound As Long)
    Dim sREC As RECTL
    Dim hBrushS As Long, mPathS As Long, mRoundS As Long
    sREC.Left = RECT.Left + (sDepth * nScale)
    sREC.Top = RECT.Top + (sDepth * nScale)
    sREC.width = RECT.width
    sREC.height = RECT.height
    GdipCreatePath &H0, mPathS
    With sREC
        mRoundS = GetSafeRound((sRound * nScale), .width * 2, .height * 2)
        If mRoundS = 0 Then mRoundS = 1
        GdipAddPathArcI mPathS, .Left, .Top, mRoundS, mRoundS, 180, 90
        GdipAddPathArcI mPathS, (.Left + .width) - mRoundS, .Top, mRoundS, mRoundS, 270, 90
        GdipAddPathArcI mPathS, (.Left + .width) - mRoundS, (.Top + .height) - mRoundS, mRoundS, mRoundS, 0, 90
        GdipAddPathArcI mPathS, .Left, (.Top + .height) - mRoundS, mRoundS, mRoundS, 90, 90
    End With
    GdipClosePathFigures mPathS
    GdipCreateSolidFill ARGB(sColor, sOpacity), hBrushS
    GdipFillPath hG, hBrushS, mPathS
    GdipDeletePath mPathS
    GdipDeleteBrush hBrushS
End Sub

Private Sub DrawGlassEffect(ByVal hG As Long, RECT As RECTL, ByVal sRound As Long)
    Dim gREC As RECTL
    Dim hBrushG As Long, mPathG As Long, mRoundG As Long
    gREC.Left = RECT.Left + 1
    gREC.Top = RECT.Top + 1
    gREC.width = RECT.width - 2
    gREC.height = CLng(RECT.height * 0.45)
    If gREC.height < 4 Then Exit Sub
    GdipCreatePath &H0, mPathG
    With gREC
        mRoundG = GetSafeRound((sRound * nScale), .width * 2, .height * 2)
        If mRoundG = 0 Then mRoundG = 1
        GdipAddPathArcI mPathG, .Left, .Top, mRoundG, mRoundG, 180, 90
        GdipAddPathArcI mPathG, (.Left + .width) - mRoundG, .Top, mRoundG, mRoundG, 270, 90
        GdipAddPathArcI mPathG, (.Left + .width) - mRoundG, (.Top + .height), 1, 1, 0, 90
        GdipAddPathArcI mPathG, .Left, (.Top + .height), 1, 1, 90, 90
    End With
    GdipClosePathFigures mPathG
    GdipCreateLineBrushFromRectWithAngleI gREC, ARGB(vbWhite, 22), ARGB(vbWhite, 0), 90, 0, WrapModeTileFlipXY, hBrushG
    GdipFillPath hG, hBrushG, mPathG
    GdipDeletePath mPathG
    GdipDeleteBrush hBrushG
End Sub

Private Sub DrawGridLines(ByVal hG As Long, gREC As RECTL)
    Dim gridPen As Long
    Dim i As Long, lineY As Long
    GdipCreatePen1 ARGB(m_GraphGridColor, 30), 1, UnitPixel, gridPen
    For i = 1 To 3
        lineY = gREC.Top + CLng(gREC.height * (i / 4))
        GdipDrawLineI hG, gridPen, gREC.Left, lineY, gREC.Left + gREC.width, lineY
    Next i
    GdipDeletePen gridPen
End Sub
Private Function gRoundRect(ByVal hGraphics As Long, RECT As RECTL, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Angulo As Single, ByVal BorderWidth As Long, ByVal BorderColor As Long, ByVal Round As Long, Filled As Boolean) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mPath As Long
    Dim mRound As Long
    
    If m_BorderWidth > 0 Then GdipCreatePen1 BorderColor, BorderWidth * nScale, &H2, hPen   '&H1 * nScale, &H2, hPen
    If Filled Then GdipCreateLineBrushFromRectWithAngleI RECT, Color1, Color2, Angulo + 90, 0, WrapModeTileFlipXY, hBrush
    GdipCreatePath &H0, mPath   '&H0
    
    With RECT
        mRound = GetSafeRound((Round * nScale), .width * 2, .height * 2)
        If mRound = 0 Then mRound = 1
            GdipAddPathArcI mPath, .Left, .Top, mRound, mRound, 180, 90
            GdipAddPathArcI mPath, (.Left + .width) - mRound, .Top, mRound, mRound, 270, 90
            GdipAddPathArcI mPath, (.Left + .width) - mRound, (.Top + .height) - mRound, mRound, mRound, 0, 90
            GdipAddPathArcI mPath, .Left, (.Top + .height) - mRound, mRound, mRound, 90, 90
    End With
    
    GdipClosePathFigures mPath
    GdipFillPath hGraphics, hBrush, mPath
    GdipDrawPath hGraphics, hPen, mPath
    
    Call GdipDeletePath(mPath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)

    gRoundRect = mPath
End Function

'Inicia GDI+
Private Sub InitGDI()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Sub

Private Function IsMouseOver(hWnd As Long) As Boolean
    Dim Pt As POINTL
    GetCursorPos Pt
    IsMouseOver = (WindowFromPoint(Pt.x, Pt.y) = hWnd)
End Function

'Private Function MousePointerHands(ByVal NewValue As Boolean)
'  If NewValue Then
'    If Ambient.UserMode Then
'      UserControl.MousePointer = vbCustom
'      UserControl.MouseIcon = GetSystemHandCursor
'    End If
'  Else
'    If hCur Then DestroyCursor hCur: hCur = 0
'    UserControl.MousePointer = vbDefault
'    UserControl.MouseIcon = Nothing
'  End If
'
'End Function

Private Function ReadValue(ByVal lProp As Long, Optional Default As Long) As Long
    Dim i       As Long
    For i = 0 To TLS_MINIMUM_AVAILABLE - 1
        If TlsGetValue(i) = lProp Then
            ReadValue = TlsGetValue(i + 1)
            Exit Function
        End If
    Next
    ReadValue = Default
End Function

Private Sub SafeRange(Value, Min, Max)
    If Value < Min Then Value = Min
    If Value > Max Then Value = Max
End Sub

'Termina GDI+
Private Sub TerminateGDI()
    Call GdiplusShutdown(GdipToken)
End Sub

Private Sub tmrEffect_Timer()
If IsMouseOver(UserControl.hWnd) Then
  If m_Opacity < 100 Then
    m_Opacity = m_Opacity + 2
    Refresh
  Else
    Refresh
    Exit Sub
  End If
Else
  m_Opacity = m_InitialOpacity
  mShowTtp = False
  Refresh
  tmrEffect.Enabled = False
End If

End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
  CopyAmbient
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent DblClick
End Sub

Private Sub UserControl_Click()
RaiseEvent Click(mSerie)
End Sub

Private Sub UserControl_Initialize()
    InitGDI
    nScale = GetWindowsDPI
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
hFontCollection = ReadValue(&HFC)
cl_hWnd = UserControl.ContainerHwnd

  m_Clicked = False
  m_Color1 = &HFFC0C0
  m_Color2 = &HFFC0C0
  m_Angulo = 0
  m_BorderColor = &HFF8080
  m_BorderWidth = 1
  m_CornerCurve = 10
  
  m_isBoxed = False
  m_Graph = egRectLine
  
  m_ForeColor = &HFFFFFF
  m_ForeColor2 = &HFFFFFF
  
  Set m_Font = UserControl.Font
  Set m_Font2 = UserControl.Font
  Set m_IconFont = UserControl.Font
  
  m_Caption = Ambient.DisplayName
  m_CaptionAlignV = eMiddle
  m_CaptionAlignH = eCenter
  m_Caption2 = Ambient.DisplayName
  m_Caption2AlignV = eMiddle
  m_Caption2AlignH = eCenter
      
  m_Filled = True
  m_EffectFade = True
  m_Opacity = 50
  m_InitialOpacity = m_Opacity
  
  m_IconCharCode = "&H0"
  m_IconForeColor = &HFFFFFF
  m_IconAlignV = eMiddle
  m_IconAlignH = eCenter
  m_IconBox = ebLeft
  
  m_GraphStyle = gsGradient
  m_GraphLineColor = &H800000
  m_GraphBackColor = &HFF0000
  m_GraphPointColor = &HFF0000
    
  m_GraphMatrix = "20,15,35,30,70,65"
  m_GraphMatrixTooltip = "$20M|$15M|$35M|$30M|$70M|$65M"
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

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer

On Error GoTo hErr
m_ToolTip = Split(GraphMatrixTooltip, "|")

  Select Case m_Graph
    Case egRectLine, egCurvedLine
      For i = 0 To UBound(iPts)
        If x >= iPts(i).x - 3 And x <= iPts(i).x + 3 And y >= iPts(i).y - 3 And y <= iPts(i).y + 3 Then
          m_PosX = x
          m_PosY = y
          mShowTtp = True
          sToolTip = m_ToolTip(i)
          mSerie = i
          Exit For
        Else
          mShowTtp = False
          mSerie = -1
        End If
      Next i
    Case egBars
      For i = 0 To UBound(gREC)
        If x > gREC(i).Left And x < gREC(i).Left + gREC(i).width And y > gREC(i).Top And y < gREC(i).Top + gREC(i).height Then
          m_PosX = x
          m_PosY = y
          mShowTtp = True
          sToolTip = m_ToolTip(i)
          mSerie = i
          Exit For
        Else
          mShowTtp = False
          mSerie = -1
        End If
      Next i
  End Select

tmrEffect.Enabled = True

hErr:
RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

With PropBag
  m_Enabled = .ReadProperty("Enabled", True)
  
  m_Color1 = .ReadProperty("BackColor1", &HFFC0C0)
  m_Color2 = .ReadProperty("BackColor2", &HFFC0C0)
  m_Angulo = .ReadProperty("BackAngle", 0)
  m_BorderColor = .ReadProperty("BorderColor", &HFF8080)
  m_BorderWidth = .ReadProperty("BorderWidth", 1)
  m_CornerCurve = .ReadProperty("CornerCurve", 10)
  m_Filled = .ReadProperty("Filled", True)
  
  m_isBoxed = .ReadProperty("Boxed", False)
  m_Graph = .ReadProperty("GraphLine", egRectLine)
  
  m_ForeColor = .ReadProperty("Caption1Color", &HFFFFFF)
  m_ForeColor2 = .ReadProperty("Caption2Color", &HFFFFFF)
  
  Set m_Font = .ReadProperty("Caption1Font", UserControl.Font)
  m_Caption = .ReadProperty("Caption1", Ambient.DisplayName)
  m_CaptionAlignV = .ReadProperty("Caption1AlignV", 1)
  m_CaptionAlignH = .ReadProperty("Caption1AlignH", 1)
  Set m_Font2 = .ReadProperty("Caption2Font", UserControl.Font)
  m_Caption2 = .ReadProperty("Caption2", Ambient.DisplayName)
  m_Caption2AlignV = .ReadProperty("Caption2AlignV", 1)
  m_Caption2AlignH = .ReadProperty("Caption2AlignH", 1)
  
  m_Transparent = .ReadProperty("Transparent", True)
    
  m_ChangeBorderOnFocus = .ReadProperty("ChangeColorOnFocus", False)
  m_EffectFade = .ReadProperty("EffectFading", True)
  m_InitialOpacity = .ReadProperty("InitialOpacity", 50)
  
  Set m_IconFont = .ReadProperty("IconFont", UserControl.Font)
  m_IconCharCode = .ReadProperty("IconCharCode", "&H0")
  m_IconForeColor = .ReadProperty("IconForeColor", &HFFFFFF)
  m_IconAlignV = .ReadProperty("IconAlignV", 1)
  m_IconAlignH = .ReadProperty("IconAlignH", 1)
  m_IconBox = .ReadProperty("IconBoxSide", ebLeft)
  
  m_GraphStyle = .ReadProperty("GraphStyle", 0)
  m_GraphLineColor = .ReadProperty("GraphLineColor", &H800000)
  m_GraphBackColor = .ReadProperty("GraphBackColor", &HFF0000)
  m_GraphPointColor = .ReadProperty("GraphPointColor", &HFF0000)
    
  m_GraphMatrix = .ReadProperty("GraphMatrix", "20,15,35,30,70,65")
  m_GraphMatrixTooltip = .ReadProperty("GraphMatrixTooltip", "$20M|$15M|$35M|$30M|$70M|$65M")
End With
  
  m_Opacity = m_InitialOpacity
  UserControl.Enabled = m_Enabled
  
End Sub

Private Sub UserControl_Resize()
Refresh
End Sub

Private Sub UserControl_Terminate()
TerminateGDI
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
  Call .WriteProperty("Enabled", m_Enabled)
  Call .WriteProperty("BackColor1", m_Color1, m_def_Color1)
  Call .WriteProperty("BackColor2", m_Color2, m_def_Color2)
  Call .WriteProperty("BackAngle", m_Angulo, m_def_Angulo)
  Call .WriteProperty("BorderColor", m_BorderColor, &HC0&)
  Call .WriteProperty("BorderWidth", m_BorderWidth, 1)
  Call .WriteProperty("CornerCurve", m_CornerCurve, 0)
  Call .WriteProperty("Filled", m_Filled)
  
  Call .WriteProperty("Boxed", m_isBoxed)
  Call .WriteProperty("GraphLine", m_Graph)
  
  Call .WriteProperty("Caption1Color", m_ForeColor, m_def_ForeColor)
  Call .WriteProperty("Caption2Color", m_ForeColor2, m_def_ForeColor)
  
  Call .WriteProperty("Caption1Font", m_Font, UserControl.Ambient.Font)
  Call .WriteProperty("Caption1", m_Caption, Ambient.DisplayName)
  Call .WriteProperty("Caption1AlignV", m_CaptionAlignV, 1)
  Call .WriteProperty("Caption1AlignH", m_CaptionAlignH, 1)
  Call .WriteProperty("Caption2Font", m_Font2, UserControl.Ambient.Font)
  Call .WriteProperty("Caption2", m_Caption2, Ambient.DisplayName)
  Call .WriteProperty("Caption2AlignV", m_Caption2AlignV, 1)
  Call .WriteProperty("Caption2AlignH", m_Caption2AlignH, 1)
  
  Call .WriteProperty("Transparent", m_Transparent, True)
    
  Call .WriteProperty("BoxedColor", m_BoxedColor, vbWhite)
  Call .WriteProperty("ChangeColorOnFocus", m_ChangeBorderOnFocus, False)
  Call .WriteProperty("EffectFading", m_EffectFade, False)
  Call .WriteProperty("InitialOpacity", m_InitialOpacity, 50)
  
  Call .WriteProperty("IconFont", m_IconFont)
  Call .WriteProperty("IconCharCode", m_IconCharCode, 0)
  Call .WriteProperty("IconForeColor", m_IconForeColor, vbButtonText)
  Call .WriteProperty("IconAlignV", m_IconAlignV, 1)
  Call .WriteProperty("IconAlignH", m_IconAlignH, 1)
  Call .WriteProperty("IconBoxSide", m_IconBox)
  
  Call .WriteProperty("GraphStyle", m_GraphStyle, 0)
  Call .WriteProperty("GraphLineColor", m_GraphLineColor)
  Call .WriteProperty("GraphBackColor", m_GraphBackColor)
  Call .WriteProperty("GraphPointColor", m_GraphPointColor)
    
  Call .WriteProperty("GraphMatrix", m_GraphMatrix)
  Call .WriteProperty("GraphMatrixTooltip", m_GraphMatrixTooltip)

End With
  
End Sub

Public Property Get BackAngle() As Single
  BackAngle = m_Angulo
End Property

Public Property Let BackAngle(ByVal New_Angulo As Single)
  m_Angulo = New_Angulo
  PropertyChanged "BackAngle"
  Refresh
End Property

Public Property Get BackColor1() As OLE_COLOR
  BackColor1 = m_Color1
End Property

Public Property Let BackColor1(ByVal New_Color1 As OLE_COLOR)
  m_Color1 = New_Color1
  PropertyChanged "BackColor1"
  Refresh
End Property

Public Property Get BackColor2() As OLE_COLOR
  BackColor2 = m_Color2
End Property

Public Property Let BackColor2(ByVal New_Color2 As OLE_COLOR)
  m_Color2 = New_Color2
  PropertyChanged "BackColor2"
  Refresh
End Property

'Properties-------------------
Public Property Get BorderColor() As OLE_COLOR
  BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)
  m_BorderColor = NewBorderColor
  'm_OldBorderColor = m_BorderColor
  PropertyChanged "BorderColor"
  Refresh
End Property

Public Property Get BoxedColor() As OLE_COLOR
  BoxedColor = m_BoxedColor
End Property

Public Property Let BoxedColor(ByVal NewBoxedColor As OLE_COLOR)
  m_BoxedColor = NewBoxedColor
  PropertyChanged "BoxedColor"
  Refresh
End Property

Public Property Get BorderWidth() As Long
  BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal NewBorderWidth As Long)
  m_BorderWidth = NewBorderWidth
  PropertyChanged "BorderWidth"
  Refresh
End Property

Public Property Get Caption1Font() As StdFont
  Set Caption1Font = m_Font
End Property

Public Property Set Caption1Font(ByVal New_Font As StdFont)
  Set m_Font = New_Font
  PropertyChanged "Caption1Font"
  Refresh
End Property

Public Property Get Caption1Color() As OLE_COLOR
  Caption1Color = m_ForeColor
End Property

Public Property Let Caption1Color(ByVal NewForeColor As OLE_COLOR)
  m_ForeColor = NewForeColor
  PropertyChanged "Caption1Color"
  Refresh
End Property

Public Property Get Caption1() As String
  Caption1 = m_Caption
End Property

Public Property Let Caption1(ByVal NewCaption As String)
  m_Caption = NewCaption
  PropertyChanged "Caption1"
  Refresh
End Property

Public Property Get Caption1AlignH() As eTextAlignH
  Caption1AlignH = m_CaptionAlignH
End Property

Public Property Let Caption1AlignH(ByVal NewCaptionAlignH As eTextAlignH)
  m_CaptionAlignH = NewCaptionAlignH
  PropertyChanged "Caption1AlignH"
  Refresh
End Property

Public Property Get Caption1AlignV() As eTextAlignV
  Caption1AlignV = m_CaptionAlignV
End Property

Public Property Let Caption1AlignV(ByVal NewCaptionAlignV As eTextAlignV)
  m_CaptionAlignV = NewCaptionAlignV
  PropertyChanged "Caption1AlignV"
  Refresh
End Property

Public Property Get Caption2Font() As StdFont
  Set Caption2Font = m_Font2
End Property

Public Property Set Caption2Font(ByVal New_Font As StdFont)
  Set m_Font2 = New_Font
  PropertyChanged "Caption2Font"
  Refresh
End Property

Public Property Get Caption2Color() As OLE_COLOR
  Caption2Color = m_ForeColor2
End Property

Public Property Let Caption2Color(ByVal NewForeColor As OLE_COLOR)
  m_ForeColor2 = NewForeColor
  PropertyChanged "Caption2Color"
  Refresh
End Property
Public Property Get Caption2() As String
  Caption2 = m_Caption2
End Property

Public Property Let Caption2(ByVal NewCaption As String)
  m_Caption2 = NewCaption
  PropertyChanged "Caption2"
  Refresh
End Property

Public Property Get Caption2AlignH() As eTextAlignH
  Caption2AlignH = m_Caption2AlignH
End Property

Public Property Let Caption2AlignH(ByVal NewCaptionAlignH As eTextAlignH)
  m_Caption2AlignH = NewCaptionAlignH
  PropertyChanged "Caption2AlignH"
  Refresh
End Property

Public Property Get Caption2AlignV() As eTextAlignV
  Caption2AlignV = m_Caption2AlignV
End Property

Public Property Let Caption2AlignV(ByVal NewCaptionAlignV As eTextAlignV)
  m_Caption2AlignV = NewCaptionAlignV
  PropertyChanged "Caption2AlignV"
  Refresh
End Property
'----
'----
Public Property Get ChangeBorderOnFocus() As Boolean
  ChangeBorderOnFocus = m_ChangeBorderOnFocus
End Property

Public Property Let ChangeBorderOnFocus(ByVal NewChangeBorderOnFocus As Boolean)
  m_ChangeBorderOnFocus = NewChangeBorderOnFocus
  m_OldBorderColor = m_BorderColor
  PropertyChanged "ChangeBorderOnFocus"
End Property

Public Property Get CornerCurve() As Long
  CornerCurve = m_CornerCurve
End Property

Public Property Let CornerCurve(ByVal NewCornerCurve As Long)
  m_CornerCurve = NewCornerCurve
  PropertyChanged "CornerCurve"
  Refresh
End Property

Public Property Get EffectFading() As Boolean
EffectFading = m_EffectFade
End Property

Public Property Let EffectFading(ByVal vNewValue As Boolean)
m_EffectFade = vNewValue
PropertyChanged "EffectFading"
Refresh
End Property

Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  m_Enabled = New_Enabled
  PropertyChanged "Enabled"
End Property

Public Property Get Filled() As Boolean
  Filled = m_Filled
End Property

Public Property Let Filled(ByVal New_Filled As Boolean)
  m_Filled = New_Filled
  PropertyChanged "Filled"
  Refresh
End Property

Public Property Get GraphLine() As eGraph2
GraphLine = m_Graph
End Property

Public Property Let GraphLine(ByVal eG As eGraph2)
  m_Graph = eG
  PropertyChanged "GraphLine"
  Refresh
End Property

Public Property Get GraphMatrix() As String
  GraphMatrix = m_GraphMatrix
End Property

Public Property Let GraphMatrix(ByVal NewGraphMatrix As String)
  m_GraphMatrix = NewGraphMatrix
  PropertyChanged "GraphMatrix"
  Refresh
End Property

Public Property Get GraphMatrixTooltip() As String
  GraphMatrixTooltip = m_GraphMatrixTooltip
End Property

Public Property Let GraphMatrixTooltip(ByVal NewGraphMatrixTooltip As String)
  m_GraphMatrixTooltip = NewGraphMatrixTooltip
  PropertyChanged "GraphMatrixTooltip"
End Property

Public Property Get GraphStyle() As eGraphStyle
  GraphStyle = m_GraphStyle
End Property

Public Property Let GraphStyle(ByVal NewGraphStyle As eGraphStyle)
  m_GraphStyle = NewGraphStyle
  PropertyChanged "GraphStyle"
  Refresh
End Property

Public Property Get GraphBackColor() As OLE_COLOR
  GraphBackColor = m_GraphBackColor
End Property

Public Property Let GraphBackColor(ByVal NewGraphBackColor As OLE_COLOR)
  m_GraphBackColor = NewGraphBackColor
  PropertyChanged "GraphBackColor"
  Refresh
End Property

Public Property Get GraphLineColor() As OLE_COLOR
  GraphLineColor = m_GraphLineColor
End Property

Public Property Let GraphLineColor(ByVal NewGraphLineColor As OLE_COLOR)
  m_GraphLineColor = NewGraphLineColor
  PropertyChanged "GraphLineColor"
  Refresh
End Property

Public Property Get GraphPointColor() As OLE_COLOR
  GraphPointColor = m_GraphPointColor
End Property

Public Property Let GraphPointColor(ByVal NewGraphPointColor As OLE_COLOR)
  m_GraphPointColor = NewGraphPointColor
  PropertyChanged "GraphPointColor"
  Refresh
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get IconCharCode() As String
    IconCharCode = "&H" & Hex(m_IconCharCode)
End Property

Public Property Let IconCharCode(ByVal New_IconCharCode As String)
    New_IconCharCode = UCase(Replace(New_IconCharCode, Space(1), vbNullString))
    New_IconCharCode = UCase(Replace(New_IconCharCode, "U+", "&H"))
    If Not VBA.Left$(New_IconCharCode, 2) = "&H" And Not IsNumeric(New_IconCharCode) Then
        m_IconCharCode = "&H" & New_IconCharCode
    Else
        m_IconCharCode = New_IconCharCode
    End If
    PropertyChanged "IconCharCode"
    Refresh
End Property

Public Property Get IconFont() As StdFont
    Set IconFont = m_IconFont
End Property

Public Property Set IconFont(New_Font As StdFont)
  Set m_IconFont = New_Font
    PropertyChanged "IconFont"
  Refresh
End Property

Public Property Get IconForeColor() As OLE_COLOR
    IconForeColor = m_IconForeColor
End Property

Public Property Let IconForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_IconForeColor = New_ForeColor
    PropertyChanged "IconForeColor"
    Refresh
End Property

Public Property Get IconAlignH() As eTextAlignH
  IconAlignH = m_IconAlignH
End Property

Public Property Let IconAlignH(ByVal NewIconAlignH As eTextAlignH)
  m_IconAlignH = NewIconAlignH
  PropertyChanged "IconAlignH"
  Refresh
End Property

Public Property Get IconAlignV() As eTextAlignV
  IconAlignV = m_IconAlignV
End Property

Public Property Let IconAlignV(ByVal NewIconAlignV As eTextAlignV)
  m_IconAlignV = NewIconAlignV
  PropertyChanged "IconAlignV"
  Refresh
End Property

Public Property Get InitialOpacity() As Long
  InitialOpacity = m_InitialOpacity
End Property

Public Property Let InitialOpacity(ByVal NewInitialOpacity As Long)
  m_InitialOpacity = NewInitialOpacity
  PropertyChanged "InitialOpacity"
  m_Opacity = m_InitialOpacity
  Refresh
End Property

Public Property Get Transparent() As Boolean
    Transparent = m_Transparent
End Property

Public Property Let Transparent(ByVal NewValue As Boolean)
    m_Transparent = NewValue
    PropertyChanged "Transparent"
    Refresh
End Property

Public Property Get Version() As String
Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Property Get Visible() As Boolean
  Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal NewVisible As Boolean)
  Extender.Visible = NewVisible
End Property

Public Property Get Clickable() As Boolean
   Clickable = m_Clickable
End Property

Public Property Let Clickable(ByVal bClickable As Boolean)
   m_Clickable = bClickable
   PropertyChanged "Clickable"
End Property

Public Property Get IconBoxSide() As eBoxed
  IconBoxSide = m_IconBox
End Property

Public Property Let IconBoxSide(ByVal bBoxed As eBoxed)
  m_IconBox = bBoxed
  PropertyChanged "IconBoxSide"
  Refresh
End Property

Public Property Get Boxed() As Boolean
  Boxed = m_isBoxed
End Property

Public Property Let Boxed(ByVal bBoxed As Boolean)
  m_isBoxed = bBoxed
  PropertyChanged "Boxed"
  Refresh
End Property



'--- Mejoras Fase 4: Propiedades ---
Public Property Get Tag() As String: Tag = m_Tag: End Property
Public Property Let Tag(ByVal v As String): m_Tag = v: End Property

Public Property Get DropShadow() As Boolean: DropShadow = m_DropShadow: End Property
Public Property Let DropShadow(ByVal v As Boolean): m_DropShadow = v: Refresh: End Property

Public Property Get ShadowColor() As OLE_COLOR: ShadowColor = m_ShadowColor: End Property
Public Property Let ShadowColor(ByVal v As OLE_COLOR): m_ShadowColor = v: Refresh: End Property

Public Property Get ShadowDepth() As Long: ShadowDepth = m_ShadowDepth: End Property
Public Property Let ShadowDepth(ByVal v As Long): m_ShadowDepth = v: Refresh: End Property

Public Property Get ShadowOpacity() As Long: ShadowOpacity = m_ShadowOpacity: End Property
Public Property Let ShadowOpacity(ByVal v As Long): m_ShadowOpacity = v: Refresh: End Property

Public Property Get GlassEffect() As Boolean: GlassEffect = m_GlassEffect: End Property
Public Property Let GlassEffect(ByVal v As Boolean): m_GlassEffect = v: Refresh: End Property

Public Property Get PieHoleSize() As Single: PieHoleSize = m_PieHoleSize: End Property
Public Property Let PieHoleSize(ByVal v As Single)
    If v < 0 Then v = 0
    If v > 85 Then v = 85
    m_PieHoleSize = v: Refresh
End Property

Public Property Get PieShowPercent() As Boolean: PieShowPercent = m_PieShowPercent: End Property
Public Property Let PieShowPercent(ByVal v As Boolean): m_PieShowPercent = v: Refresh: End Property

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
