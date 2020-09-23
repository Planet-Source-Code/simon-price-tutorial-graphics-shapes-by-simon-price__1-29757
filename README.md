<div align="center">

## Tutorial \- Graphics \- Shapes \- by Simon Price


</div>

### Description

This tutorial is third in a series about Win32 API for graphics. In this tutorial, you will learn about functions that help draw various geometric shapes and how to use them in code, with a working example to download.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-12-11 19:36:28
**By**             |[Simon Price](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/simon-price.md)
**Level**          |Advanced
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Tutorial\_\-4187512132001\.zip](https://github.com/Planet-Source-Code/simon-price-tutorial-graphics-shapes-by-simon-price__1-29757/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
<link rel="stylesheet" type="text/css" href="../../vbgames.css">
<meta name="Microsoft Border" content="t, default">
</head>
<body>
<p> </p>
<h4>About this tutorial:</h4>
<p>This tutorial by Simon Price is part of a series held at <a href="http://www.VBgames.co.uk">http://www.VBgames.co.uk</a>.
It requires a good knowledge of Visual Basic programming. This tutorial come
with an example program with VB6 source code, which can be downloaded from <a href="http://www.VBgames.co.uk/tutorials/gdi/dcs.zip">http://www.VBgames.co.uk/tutorials/gdi/pensbrushes.zip</a>
and possibly from other websites hosting this tutorial (such as PSC).</p>
<h4>Before you begin:</h4>
<p>Have you read the previous tutorial - <i>Device Contexts</i>? If not, please
read that first, because this tutorial builds upon the knowledge and code of the
previous tutorial.</p>
<p><b>Pixels</b></p>
<p>A bitmap has dimensions of width and height measured in pixels. A pixel is
the smallest part of a bitmap which can be changed. It's a little dot. Drawing
anything, whether it's a line, a circle, or a 3D model in the latest 3D shoot-em-up,
at the end of the way, it comes down to drawing pixels. So a pixel is the first
"shape" to learn to draw, since everything else relies upon it.</p>
<h4>32 Bit Color</h4>
<p>All colors in the Windows API graphics functions are given in a 32 bit (4
byte) integer - the Long data type in VB. 1 byte is red, 1 byte green, 1 byte
blue, 1 byte is reserved and currently does nothing (<hint> you can use
the last byte for yourself - maybe store alpha values!? </hint>). VB comes
with the RGB function for creating these 32 bit colors.</p>
<h4>API functions for pixels</h4>
<p>Here are the API functions used for pixels:</p>
<hr>
<p> Declare Function <b> GetPixel</b> Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long</p>
<p><b> GetPixel</b> - returns the color of a pixel in a device context, given
it's x and y coordinates</p>
<hr>
<p>Private Declare Function <b> SetPixelV</b> Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long</p>
<p><b> SetPixelV</b> - Sets the color of a pixel in a device context</p>
<hr>
<h4>Reading and Writing Pixel Colors</h4>
<p>Pixels are read and written with the <i>GetPixel</i> and <i>SetPixelV</i>
functions. Here is the code to read and write pixel colors, it is pretty much
self-explanitory:</p>
<p><font size="1">' gets an individual pixel color in long format<br>
Public Property Get Pixel(x As Long, y As Long) As Long<br>
On Error Resume Next<br>
    Pixel = GetPixel(Me.hDC, x, y)<br>
End Property</font></p>
<p><font size="1"><br>
' sets an individual pixel color in long format<br>
Public Property Let Pixel(x As Long, y As Long, Color As Long)<br>
On Error Resume Next<br>
    SetPixelV Me.hDC, x, y, Color<br>
End Property</font></p>
<h4>Shapes</h4>
<p>It would be possible for us to draw other shapes using the pixel drawing
functions we just learnt. If you want, feel free to go do just that! However,
Windows comes with functions to draw several common shapes, which are easy to
use, and faster than what you could make in your own software VB-coded
implementations of the same functions. I suggest you learn the Windows API
rather than doing it yourself. No need to re-invert the wheel yet.</p>
<h4>API functions for shapes</h4>
<p>Here are the API functions used for shapes:</p>
<hr>
<p>Private Declare Function <b> MoveToEx</b> Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long</p>
<p><b> MoveToEx</b> - set the current cursor of the device context</p>
<hr>
<p>Private Declare Function <b> LineTo</b> Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long</p>
<p><b> LineTo</b> - draws a line from the current cursor position to the
specified point</p>
<hr>
<p>Private Declare Function <b> Rectangle</b> Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long</p>
<p><b> Rectangle</b> - draws a rectangle given two opposing points</p>
<hr>
<p>Private Declare Function <b> Ellipse</b> Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long</p>
<p><b> Ellipse</b> - draws an ellipse, given the opposing points of an imaginary
rectangle what would fit around the ellipse</p>
<hr>
<p>Private Declare Function <b> Polygon</b> Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long</p>
<p><b> Polygon</b> - draws a polygon of any number of sides, given a pointer to
and array of 2D points, and the number or points</p>
<hr>
<p>Private Declare Function <b> Arc</b> Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long</p>
<p><b> Arc</b> - Draws and arc</p>
<hr>
<p>Private Declare Function <b> ArcTo</b> Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long</p>
<p><b> ArcTo</b> - draws an arc</p>
<hr>
<p>Private Declare Function <b> Pie</b> Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long</p>
<p><b> Pie</b> - draws sector of a circle</p>
<hr>
<p>Private Declare Function <b> ExtFloodFill</b> Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long</p>
<p><b> ExtFloodFill</b> - floods an area with a color</p>
<hr>
<p>Private Declare Function <b> FloodFill</b> Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long</p>
<p><b> FloodFill</b> - floods an area with color until a border color is found</p>
<hr>
<p>Private Declare Function <b> FillRect</b> Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long</p>
<p><b> FillRect</b> - fills a rectangle with a pattern from a brush</p>
<hr>
<p>Private Declare Function <b> PatBlt</b> Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long</p>
<p><b> PatBlt</b> - fills a rectangle with a pattern</p>
<hr>
<p>Private Declare Function <b> GetPolyFillMode</b> Lib "gdi32" (ByVal hDC As Long) As Long</p>
<p><b> GetPolyFillMode</b> - returns the current polygon filling mode</p>
<hr>
<p>Private Declare Function <b> SetPolyFillMode</b> Lib "gdi32" (ByVal hDC As Long, ByVal nPolyFillMode As Long) As Long</p>
<p><b> SetPolyFillMode</b> - sets the current polygon filling mode</p>
<hr>
<p>Private Declare Function <b> GetTextColor</b> Lib "gdi32" (ByVal hDC As Long) As Long</p>
<p><b> GetTextColor</b> - returns the current text color</p>
<hr>
<p>Private Declare Function <b> SetTextColor</b> Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long</p>
<p><b> SetTextColor</b> - sets the current text color</p>
<hr>
<p>Private Declare Function <b> GetTextAlign</b> Lib "gdi32" (ByVal hDC As Long) As Long</p>
<p><b> GetTextAlign</b> - returns the current text alignment mode</p>
<hr>
<p>Private Declare Function <b> SetTextAlign</b> Lib "gdi32" (ByVal hDC As Long, ByVal wFlags As Long) As Long</p>
<p><b> SetTextAlign</b> - sets the current text alignment mode</p>
<hr>
<p>Private Declare Function <b> TextOut</b> Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long</p>
<p><b> TextOut</b> - draws a specified string of text on a device context using
the current text color and alignment</p>
<hr>
<h4>Lines</h4>
<p>Lines are drawn with the <i>MoveToEx</i> and <i>LineTo</i> functions. A
device context has a current drawing position - a 2D coordinate, like a cursor.
To draw a line, we must put the "cursor" to one end of the line, then
draw a line between the cursor and the other end of the line. That was a lame
explanation, just look at the code and it's easily understood:</p>
<p><font size="1">' draws a line<br>
Public Sub DrawLine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)<br>
Dim Point As POINTAPI<br>
On Error Resume Next<br>
    MoveToEx hDC, X1, Y1, Point<br>
    LineTo hDC, X2, Y2<br>
End Sub<br>
</font></p>
<h4>Rectangles (and squares)</h4>
<p>A rectangle is drawn with the <i>Rectangle</i> function. A rectangle is
specified by the coordinates of 2 opposing corners. There is no need to specify
the other 2 corners since they can be worked out from the given points. Here is
the code to draw a rectangle:</p>
<p><font size="1">' draws a rectangle<br>
Public Sub DrawRectangle(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)<br>
On Error Resume Next<br>
    Rectangle hDC, X1, Y1, X2, Y2<br>
End Sub</font></p>
<p>Note that squares are drawn in the same way, because a square is just a
special type of rectangle where <i>X2 - X1 = Y2 - Y1</i>.</p>
<h4>Ellipses (and circles)</h4>
<p>An ellipse is drawn with the <i>Ellipse</i> function. An ellipse is specified
by an imaginary rectangle what would fit around the ellipse. Here is the code to
draw an ellipse:</p>
<p><font size="1">' draws an ellipse<br>
Public Sub DrawEllipse(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)<br>
On Error Resume Next<br>
    Ellipse hDC, X1, Y1, X2, Y2<br>
End Sub<br>
</font></p>
<p>Note that circles are drawn in the same way, because a circle is just a
special type of ellipse where <i>X2 - X1 = Y2 - Y1</i>.</p>
<h4>Drawing Polygons</h4>
<p>Polygons are drawn with the <i>Polygon</i> function. Polygons are made from
many points joined up. Examples of polygons include triangles, quadrilaterals,
pentagons, hexagons, heptagons, octagons, nonagons, decagons, dodecagons etc.</p>
<p>Here is the function to draw any polygon:</p>
<p><font size="1">' draws a polygon<br>
Public Sub DrawPolygon(x() As Long, y() As Long)<br>
Dim Point() As POINTAPI<br>
Dim i As Long<br>
On Error Resume Next<br>
    ReDim Point(LBound(x) To UBound(x))<br>
    For i = LBound(x) To UBound(x)<br>
        Point(i).x = x(i)<br>
        Point(i).y = y(i)<br>
    Next<br>
    Polygon hDC, Point(LBound(Point)), UBound(Point) - LBound(Point) + 1<br>
End Sub</font></p>
<p>The most commonly drawn polygon is a triangle, so here is an optimised
version of the function, just for triangles:</p>
<p><font size="1">' draws a triangle<br>
Public Sub DrawTriangle(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, X3 As Long, Y3 As Long)<br>
Dim Point(1 To 3) As POINTAPI<br>
On Error Resume Next<br>
    Point(1).x = X1<br>
    Point(1).y = Y1<br>
    Point(2).x = X2<br>
    Point(2).y = Y2<br>
    Point(3).x = X3<br>
    Point(3).y = Y3<br>
    Polygon hDC, Point(1), 3<br>
End Sub</font></p>
<h4>Drawing Patterns</h4>
<p>A rectangular region can be filled with a common hatched pattern with the <i>PatBlt</i>
function. Here is the code to do that:</p>
<p><font size="1">' fills a rectangle with a pattern<br>
Public Sub DrawPattern(Optional x As Long = 0, Optional y As Long = 0, Optional lWidth As Long = 0, Optional lHeight As Long = 0, Optional RasterOp As PATBLT_RASTEROP = PR_PATCOPY)<br>
On Error Resume Next<br>
    If lWidth = 0 Then lWidth = Width<br>
    If lHeight = 0 Then lHeight = Height<br>
    PatBlt hDC, x, y, lWidth, lHeight, RasterOp<br>
End Sub<br>
</font></p>
<h4>Drawing Text</h4>
<p>A string of text can be draw with the <i>TextOut</i> function. Here is the
code to draw text at a specified position on the device context:</p>
<p><font size="1">' draws a text string<br>
Public Sub DrawText(str As String, Optional x As Long = 0, Optional y As Long = 0)<br>
On Error Resume Next<br>
    TextOut hDC, x, y, str, Len(str)<br>
End Sub<br>
</font></p>
<h4>Example Program</h4>
<p align="center"><img border="0" src="http://www.vbgames.co.uk/tutorials/gdi/shapes.JPG" width="648" height="507"></p>
<p>The example program demonstrates most of what has been learnt in this
tutorial. Download and run the code, you should see several shapes and a text
string naming the coolest site for VB games programming on earth! Have a go at
drawing some more, learn them well, these basic shapes are the basis for all
other shapes.</p>
<h4>Coming soon...</h4>
<p>Watch out for the next tutorial in this series! Next we will learn how to
load from and save to bitmap files for persistent graphics!</body>
</html>

