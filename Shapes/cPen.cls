VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************************'
'                                                                             '
'                               --- CPEN.CLS ---                              '
'                                                                             '
'                             (C) SIMON PRICE 2001                            '
'                                                                             '
'                           HTTP://WWW.VBGAMES.CO.UK                          '
'                                                                             '
'*****************************************************************************'

Option Explicit

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

' 2d coordinate
Private Type POINTAPI
    x As Long
    y As Long
End Type
Public Enum PEN_STYLE
    PS_ALTERNATE = 8
    PS_COSMETIC = &H0
    PS_DASH = 1
    PS_DASHDOT = 3
    PS_DASHDOTDOT = 4
    PS_DOT = 2
    PS_ENDCAP_FLAT = &H200
    PS_ENDCAP_MASK = &HF00
    PS_ENDCAP_ROUND = &H0
    PS_ENDCAP_SQUARE = &H100
    PS_GEOMETRIC = &H10000
    PS_INSIDEFRAME = 6
    PS_JOIN_BEVEL = &H1000
    PS_JOIN_MASK = &HF000
    PS_JOIN_MITER = &H2000
    PS_JOIN_ROUND = &H0
    PS_NULL = 5
    PS_SOLID = 0
    PS_STYLE_MASK = &HF
    PS_TYPE_MASK = &HF0000
    PS_USERSTYLE = 7
End Enum
' pen properties
Private Type LOGPEN
    lopnStyle As Long
    lopnWidth As POINTAPI
    lopnColor As Long
End Type
Private Pen As LOGPEN
Private hPen As Long
Private lpPen As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreatePenIndirect Lib "gdi32" (lpLogPen As LOGPEN) As Long
Public Event Change()



Private Sub Class_Initialize()
On Error Resume Next
    ' create default pen
    With Pen
        .lopnColor = vbBlack
        .lopnStyle = PS_SOLID
        .lopnWidth.x = 1
    End With
    RaiseEvent Change
End Sub



' creates a new pen and assigns to a given device context
Public Sub Use(hDC As Long, hPenOld As Long)
On Error Resume Next
    CleanUp
    hPen = CreatePenIndirect(Pen)
    hPenOld = SelectObject(hDC, hPen)
End Sub



' pen style
Public Property Get Style() As PEN_STYLE
On Error Resume Next
    Style = Pen.lopnStyle
End Property
Public Property Let Style(PS As PEN_STYLE)
On Error Resume Next
    Pen.lopnStyle = PS
    RaiseEvent Change
End Property



' pen width
Public Property Get Width() As Long
On Error Resume Next
    Width = Pen.lopnWidth.x
End Property
Public Property Let Width(w As Long)
On Error Resume Next
    Pen.lopnWidth.x = w
    RaiseEvent Change
End Property



' pen color
Public Property Get Color() As Long
On Error Resume Next
    Color = Pen.lopnColor
End Property
Public Property Let Color(c As Long)
On Error Resume Next
    Pen.lopnColor = c
    RaiseEvent Change
End Property




' deallocate resources
Private Sub CleanUp()
On Error Resume Next
    ' delete the pen object
    If hPen Then ' check for valid handle
        DeleteObject hPen
        hPen = Null
    End If
End Sub



Private Sub Class_Terminate()
On Error Resume Next
    CleanUp
End Sub


