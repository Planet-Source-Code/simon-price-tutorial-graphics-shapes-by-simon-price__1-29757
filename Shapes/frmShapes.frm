VERSION 5.00
Begin VB.Form frmShapes 
   Caption         =   "Shapes Tutorial by Simon Price http://www.VBgames.co.uk"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************'
'                                                                             '
'                           --- SHAPES TUTORIAL ---                           '
'                                                                             '
'                               BY SIMON PRICE                                '
'                                                                             '
'       SEE HTTP://WWW.VBGAMES.CO.UK FOR MORE GREAT CODE AND TUTORIALS        '
'                                                                             '
'*****************************************************************************'

Option Explicit

' a device context to play around with and test
Private DC As cDeviceContext
' a pen to play around with and test
Private Pen As cPen
' a brush to play around with and test
Private Brush As cBrush

Private Sub Form_Load()
On Error Resume Next
    ' create a new device context
    Set DC = New cDeviceContext
    DC.Create
    ' create a new pen and assign it to the device context
    Set Pen = New cPen
    Set DC.Pen = Pen
    ' create a new brush and assign it to the device context
    Set Brush = New cBrush
    Set DC.Brush = Brush
End Sub

Private Sub Form_Paint()
On Error Resume Next
    ' set pen and brush properties
    Pen.Color = vbWhite
    Pen.Style = PS_SOLID
    Pen.Width = 1
    Brush.Color = vbWhite
    Brush.Style = BS_SOLID
    Brush.Hatch = HS_SOLID
    ' draw a square
    DC.DrawRectangle 50, 50, 150, 150
    ' draw a rectangle
    DC.DrawRectangle 200, 75, 600, 125
    ' draw a circle
    DC.DrawEllipse 50, 200, 150, 300
    ' draw an ellipse
    DC.DrawEllipse 250, 200, 600, 300
    ' draw a triangle
    DC.DrawTriangle 50, 450, 100, 350, 150, 450
    ' draw some lines
    Pen.Width = 5
    DC.DrawLine 200, 350, 300, 450
    DC.DrawLine 200, 450, 300, 350
    ' draw some text
    DC.DrawText " www.VBgames.co.uk ", 400, 400
    ' copy from the device context onto the window
    DC.CopyRect hDC
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    ' delete the pen
    Set Pen = Nothing
    ' delete the brush
    Set Brush = Nothing
    ' delete the device context
    Set DC = Nothing
End Sub

