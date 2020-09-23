VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDeviceContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************************'
'                                                                             '
'                          --- CDEVICECONTEXT.CLS ---                         '
'                                                                             '
'                             (C) SIMON PRICE 2001                            '
'                                                                             '
'                           HTTP://WWW.VBGAMES.CO.UK                          '
'                                                                             '
'*****************************************************************************'

Option Explicit

' default bitmap size
Private Const DEFAULT_WIDTH = 640
Private Const DEFAULT_HEIGHT = 480

' bitmap properties
Private Type BITMAPINFOHEADER
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type
' 32 bit color
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
' 2d coordinate
Private Type POINTAPI
    x As Long
    y As Long
End Type
' bitmap properties
Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As RGBQUAD
End Type
' rectangle type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
' polygon fill modes
Public Enum FILL_MODE
    FLOODFILLBORDER = 0
    FLOODFILLSURFACE = 1
    POLYFILL_LAST = 2
End Enum
' text alignment modes
Public Enum TEXT_ALIGN
   TA_CENTER = 6
   TA_BASELINE = 24
   TA_BOTTOM = 8
   TA_LEFT = 0
   TA_NOUPDATECP = 0
   TA_RIGHT = 2
   TA_TOP = 0
   TA_UPDATECP = 1
   TA_MASK = (TA_BASELINE + TA_CENTER + TA_UPDATECP)
End Enum
' pattern bit block transfer raster operations
Public Enum PATBLT_RASTEROP
   PR_PATCOPY = &HF00021
   PR_PATINVERT = &H5A0049
   PR_PATPAINT = &HFB0A09
   PR_DSTINVERT = &H550009
   PR_BLACKNESS = &H42
   PR_WHITENESS = &HFF0062
End Enum

' handle to the windows gdi device context
Private hDC As Long
' pointer to the bits of the bitmap
Private pBits As Long
' the bitmap
Private Info As BITMAPINFO
' handle to the bitmap
Private hBMP As Long
' current pen
Private WithEvents pPen As cPen
Attribute pPen.VB_VarHelpID = -1
' old pen
Private hPenOld As Long
' current brush
Private WithEvents pBrush As cBrush
Attribute pBrush.VB_VarHelpID = -1
' old brush
Private hBrushOld As Long

' windows API function declarations
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function SaveDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function RestoreDC Lib "gdi32" (ByVal hDC As Long, ByVal nSavedDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal Handle As Long, ByVal dW As Long) As Long
Private Const DIB_RGB_COLORS = 0
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function Arc Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function ArcTo Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function Pie Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FloodFill Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPolyFillMode Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetPolyFillMode Lib "gdi32" (ByVal hDC As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetTextAlign Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long
Private Declare Function GetTextAlign Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long



' create the device context
Function Create(Optional Width As Long = DEFAULT_WIDTH, Optional Height As Long = DEFAULT_HEIGHT) As Boolean
Dim DisplayDC As Long
On Error GoTo Failed
    CleanUp ' destroy old dc
    ' fill bitmap info structure
    With Info.bmiHeader
        .biBitCount = 24
        .biWidth = Width
        .biHeight = Height
        .biPlanes = 1
        .biSize = Len(Info.bmiHeader)
    End With
    ' get a handle to the display
    DisplayDC = GetDC(0)
    ' create a dc compatible with the display
    hDC = CreateCompatibleDC(DisplayDC)
    If hDC = 0 Then GoTo Failed ' check we received a valid handle
    SaveDC hDC
    ' create a device independant bitmap
    hBMP = CreateDIBSection(hDC, Info, DIB_RGB_COLORS, pBits, 0, 0)
    If hBMP = 0 Then GoTo Failed ' check we received a valid handle
    ' use the bitmap for the dc
    SelectObject hDC, hBMP
    ' clean up the display dc
    ReleaseDC 0, DisplayDC
    Create = True
    Exit Function
Failed:
    CleanUp ' destroy failed dc
    Create = False
End Function



' handle to device context
Public Property Get Handle() As Long
On Error Resume Next
    Handle = hDC
End Property



' bitmap info
Public Property Get BitsPerPixel() As Long
    BitsPerPixel = Info.bmiHeader.biBitCount
End Property
Public Property Get Planes() As Long
    Planes = Info.bmiHeader.biPlanes
End Property
Public Property Get Size() As Long
    Size = Width * Height * BitsPerPixel * 0.125
End Property



' bitmap size
Public Property Get Width() As Long
    Width = Info.bmiHeader.biWidth
End Property
Public Property Get Height() As Long
    Height = Info.bmiHeader.biHeight
End Property
Public Property Let Width(w As Long)
On Error Resume Next
    Resize w, Height
End Property
Public Property Let Height(h As Long)
On Error Resume Next
    Resize Width, h
End Property



' re-creates the device context to a new size
Public Function Resize(Optional Width As Long = DEFAULT_WIDTH, Optional Height As Long = DEFAULT_HEIGHT) As Boolean
On Error Resume Next
    Resize = Create(Width, Height)
End Function



' bit blt function
Public Sub CopyRect(hDestDC As Long, Optional x As Long, Optional y As Long, Optional nWidth As Long, Optional nHeight As Long, Optional xSrc As Long, Optional ySrc As Long, Optional RasterOp As RasterOpConstants = vbSrcCopy)
On Error Resume Next
    If nWidth = 0 Then nWidth = Width
    If nHeight = 0 Then nHeight = Height
    BitBlt hDestDC, x, y, nWidth, nHeight, hDC, xSrc, ySrc, RasterOp
End Sub
' stretch blt function
Public Sub StretchCopyRect(hDestDC As Long, Optional x As Long, Optional y As Long, Optional nWidth As Long, Optional nHeight As Long, Optional xSrc As Long, Optional ySrc As Long, Optional nSrcWidth As Long, Optional nSrcHeight As Long, Optional RasterOp As RasterOpConstants = vbSrcCopy)
On Error Resume Next
    If nWidth = 0 Then nWidth = Width
    If nHeight = 0 Then nHeight = Height
    If nSrcWidth = 0 Then nWidth = Width
    If nSrcHeight = 0 Then nHeight = Height
    StretchBlt hDestDC, x, y, nWidth, nHeight, hDC, xSrc, ySrc, nSrcWidth, nSrcHeight, RasterOp
End Sub



' pen object
Public Property Get Pen() As cPen
On Error Resume Next
    Set Pen = pPen
End Property
Public Property Set Pen(newPen As cPen)
On Error Resume Next
    Set pPen = newPen
    pPen_Change
End Property
' when the pen changes, use it on the dc again
Private Sub pPen_Change()
On Error Resume Next
    pPen.Use hDC, hPenOld
End Sub



' brush object
Public Property Get Brush() As cBrush
On Error Resume Next
    Set Brush = pBrush
End Property
Public Property Set Brush(newBrush As cBrush)
On Error Resume Next
    Set pBrush = newBrush
    pBrush_Change
End Property
' when the brush changes, use it on the dc again
Private Sub pBrush_Change()
On Error Resume Next
    pBrush.Use hDC, hBrushOld
End Sub



' gets an individual pixel color in long format
Public Property Get Pixel(x As Long, y As Long) As Long
On Error Resume Next
    Pixel = GetPixel(Me.hDC, x, y)
End Property
' sets an individual pixel color in long format
Public Property Let Pixel(x As Long, y As Long, Color As Long)
On Error Resume Next
    SetPixelV Me.hDC, x, y, Color
End Property




' draws a line
Public Sub DrawLine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
Dim Point As POINTAPI
On Error Resume Next
    MoveToEx hDC, X1, Y1, Point
    LineTo hDC, X2, Y2
End Sub
' draws a rectangle
Public Sub DrawRectangle(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
On Error Resume Next
    Rectangle hDC, X1, Y1, X2, Y2
End Sub
' draws an ellipse
Public Sub DrawEllipse(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
On Error Resume Next
    Ellipse hDC, X1, Y1, X2, Y2
End Sub
' draws a polygon
Public Sub DrawPolygon(x() As Long, y() As Long)
Dim Point() As POINTAPI
Dim i As Long
On Error Resume Next
    ReDim Point(LBound(x) To UBound(x))
    For i = LBound(x) To UBound(x)
        Point(i).x = x(i)
        Point(i).y = y(i)
    Next
    Polygon hDC, Point(LBound(Point)), UBound(Point) - LBound(Point) + 1
End Sub
' draws a triangle
Public Sub DrawTriangle(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, X3 As Long, Y3 As Long)
Dim Point(1 To 3) As POINTAPI
On Error Resume Next
    Point(1).x = X1
    Point(1).y = Y1
    Point(2).x = X2
    Point(2).y = Y2
    Point(3).x = X3
    Point(3).y = Y3
    Polygon hDC, Point(1), 3
End Sub
' draws an arc
Public Sub DrawArc(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, X3 As Long, Y3 As Long, X4 As Long, Y4 As Long)
On Error Resume Next
    Arc hDC, X1, Y1, X2, Y2, X3, Y3, X4, Y4
End Sub
' draws an arc
Public Sub DrawArcTo(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, X3 As Long, Y3 As Long, X4 As Long, Y4 As Long)
On Error Resume Next
    ArcTo hDC, X1, Y1, X2, Y2, X3, Y3, X4, Y4
End Sub
' draws a pie
Public Sub DrawPie(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, X3 As Long, Y3 As Long, X4 As Long, Y4 As Long)
On Error Resume Next
    Pie hDC, X1, Y1, X2, Y2, X3, Y3, X4, Y4
End Sub
' draws a text string
Public Sub DrawText(str As String, Optional x As Long = 0, Optional y As Long = 0)
On Error Resume Next
    TextOut hDC, x, y, str, Len(str)
End Sub
' fills a rectangle with a pattern
Public Sub DrawPattern(Optional x As Long = 0, Optional y As Long = 0, Optional lWidth As Long = 0, Optional lHeight As Long = 0, Optional RasterOp As PATBLT_RASTEROP = PR_PATCOPY)
On Error Resume Next
    If lWidth = 0 Then lWidth = Width
    If lHeight = 0 Then lHeight = Height
    PatBlt hDC, x, y, lWidth, lHeight, RasterOp
End Sub



' deallocate resources
Sub CleanUp()
On Error Resume Next
    ' replace old pen
    If hPenOld Then
        SelectObject hDC, hPenOld
        hPenOld = Null
    End If
    ' replace old brush
    If hBrushOld Then
        SelectObject hDC, hBrushOld
        hBrushOld = Null
    End If
    ' delete bitmap
    If hBMP Then ' check for valid handle
        DeleteObject hBMP
        hBMP = Null
    End If
    ' delete device context
    If hDC Then ' check for valid handle
        RestoreDC hDC, -1
        DeleteDC hDC
        hDC = Null
    End If
End Sub



' deallocate resources before object is destroyed
Private Sub Class_Terminate()
On Error Resume Next
    CleanUp
End Sub

