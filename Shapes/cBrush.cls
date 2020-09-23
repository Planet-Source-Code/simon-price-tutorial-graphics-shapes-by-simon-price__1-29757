VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************************'
'                                                                             '
'                              --- CBRUSH.CLS ---                             '
'                                                                             '
'                             (C) SIMON PRICE 2001                            '
'                                                                             '
'                           HTTP://WWW.VBGAMES.CO.UK                          '
'                                                                             '
'*****************************************************************************'

Option Explicit

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

' *** BRUSHES ***
Public Enum BRUSH_STYLE
    BS_DIBPATTERN = 5
    BS_DIBPATTERN8X8 = 8
    BS_DIBPATTERNPT = 6
    BS_NULL = 1
    BS_HOLLOW = BS_NULL
    BS_HATCHED = 2
    BS_INDEXED = 4
    BS_PATTERN = 3
    BS_PATTERN8X8 = 7
    BS_SOLID = 0
End Enum
Public Enum HATCH_STYLE
    HS_BDIAGONAL = 3
    HS_BDIAGONAL1 = 7
    HS_CROSS = 4
    HS_DENSE1 = 9
    HS_DENSE2 = 10
    HS_DENSE3 = 11
    HS_DENSE4 = 12
    HS_DENSE5 = 13
    HS_DENSE6 = 14
    HS_DENSE7 = 15
    HS_DENSE8 = 16
    HS_DIAGCROSS = 5
    HS_DITHEREDBKCLR = 24
    HS_DITHEREDCLR = 20
    HS_DITHEREDTEXTCLR = 22
    HS_FDIAGONAL = 2
    HS_FDIAGONAL1 = 6
    HS_HALFTONE = 18
    HS_HORIZONTAL = 0
    HS_NOSHADE = 17
    HS_SOLID = 8
    HS_SOLIDBKCLR = 23
    HS_SOLIDCLR = 19
    HS_SOLIDTEXTCLR = 21
    HS_VERTICAL = 1
End Enum
' brush properties
Private Type LOGBRUSH
    lbStyle As Long
    lbColor As Long
    lbHatch As Long
End Type
Private Brush As LOGBRUSH
Private hBrush As Long
Private hBrushOld As Long
Private lpBrush As Long
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long

Public Event Change()



Private Sub Class_Initialize()
On Error Resume Next
    ' create default brush
    With Brush
        .lbColor = vbBlack
        .lbStyle = PS_SOLID
        .lbHatch = HS_SOLID
    End With
    RaiseEvent Change
End Sub



' creates a new brush on a given device context
Public Sub Use(hDC As Long, hBrushOld As Long)
On Error Resume Next
    CleanUp
    hBrush = CreateBrushIndirect(Brush)
    hBrushOld = SelectObject(hDC, hBrush)
End Sub



' brush style
Public Property Get Style() As BRUSH_STYLE
On Error Resume Next
    Style = Brush.lbStyle
End Property
Public Property Let Style(s As BRUSH_STYLE)
On Error Resume Next
    Brush.lbStyle = s
    RaiseEvent Change
End Property



' brush color
Public Property Get Color() As Long
On Error Resume Next
    Color = Brush.lbColor
End Property
Public Property Let Color(c As Long)
On Error Resume Next
    Brush.lbColor = c
    RaiseEvent Change
End Property



' brush hatch style
Public Property Get Hatch() As HATCH_STYLE
On Error Resume Next
    Hatch = Brush.lbHatch
End Property
Public Property Let Hatch(h As HATCH_STYLE)
On Error Resume Next
    Brush.lbHatch = h
    RaiseEvent Change
End Property



' deallocate resources
Private Sub CleanUp()
On Error Resume Next
    If hBrush Then ' check handle is valid
        DeleteObject hBrush ' delete brush
        hBrush = Null ' reset handle
    End If
End Sub



Private Sub Class_Terminate()
On Error Resume Next
    CleanUp
End Sub


