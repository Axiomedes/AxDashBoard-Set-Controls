VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eStringAlignment
    StringAlignmentNear = &H0
    StringAlignmentCenter = &H1
    StringAlignmentFar = &H2
End Enum
  
Public Enum eStringTrimming
    StringTrimmingNone = &H0
    StringTrimmingCharacter = &H1
    StringTrimmingWord = &H2
    StringTrimmingEllipsisCharacter = &H3
    StringTrimmingEllipsisWord = &H4
    StringTrimmingEllipsisPath = &H5
End Enum

Public Enum eStringFormatFlags
    StringFormatFlagsNone = &H0
    StringFormatFlagsDirectionRightToLeft = &H1
    StringFormatFlagsDirectionVertical = &H2
    StringFormatFlagsNoFitBlackBox = &H4
    StringFormatFlagsDisplayFormatControl = &H20
    StringFormatFlagsNoFontFallback = &H400
    StringFormatFlagsMeasureTrailingSpaces = &H800
    StringFormatFlagsNoWrap = &H1000
    StringFormatFlagsLineLimit = &H2000
    StringFormatFlagsNoClip = &H4000
End Enum

Public Enum eTextAlignH
    eLeft
    eCenter
    eRight
End Enum

Public Enum eTextAlignV
    eTop
    eMiddle
    eBottom
End Enum

Public Enum CrossPos
    cTopRight
    cBottomRight
    cTopLeft
    cBottomLeft
End Enum

Public Enum CheckPos
    cLeft
    cRight
    cTop
    cBottom
End Enum

Public Enum LineOr
  loVertical
  loHorizontal
End Enum

Public Enum CaptionPos
    cptoUp
    cptoBottom
    cpTop
    cpBottom
End Enum

Public Enum FadeEffect
    None
    BorderOnly
    ControlsOnly
    BorderAndControls
    AllPanel
End Enum

Public Enum ValPos
    pStartBar
    pValueBar
    pCenterBar
End Enum

Public Enum eFlatSide
  rUp
  rBottom
  rLeft
  rRight
End Enum

Public Enum eBoxed
  ebLeft
  ebRight
End Enum

Public Enum pStyle
  pVertical
  pHorizontal
End Enum

Public Enum eGraph
  egRectLine
  egCurvedLine
  egBars
End Enum

Public Enum eGraph2
  eRectLine
  eCurvedLine
  eBars
  ePie
End Enum

Public Enum eGraphStyle
  gsGradient
  gsSolid
End Enum

Public Enum eTypeLine
  eLine
  eDots
  eBoxs
End Enum
