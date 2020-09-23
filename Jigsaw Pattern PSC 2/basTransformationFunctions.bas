Attribute VB_Name = "basTransformationFunctions"
Option Explicit

'/////////////////////////////////////////////////////////////////////////////////////////////
'/// Transformation code adapted from C++ source code in the book
'/// Direct3D Programming (Kickstart) by Clayton Walnum
'/////////////////////////////////////////////////////////////////////////////////////////////

Public Sub Rotate(ByRef shp As BEZIER_CURVE, ByVal degrees As Integer)
      Dim i As Integer
      Dim RotatedX As Integer, RotatedY As Integer
      Dim radians As Double, c As Double, s As Double
      
      radians = 6.283185308 / (360 / degrees)
      c = Cos(radians)
      s = Sin(radians)
      
      For i = 1 To shp.NumPoints
            RotatedX = shp.Points(i).X * c - shp.Points(i).Y * s
            RotatedY = shp.Points(i).Y * c + shp.Points(i).X * s
            
            shp.Points(i).X = RotatedX
            shp.Points(i).Y = RotatedY
      Next i
End Sub

Public Sub Translate(ByRef shp As BEZIER_CURVE, ByVal TransX As Integer, ByVal TransY As Integer)
      Dim i As Integer
      
      For i = 1 To shp.NumPoints
            shp.Points(i).X = shp.Points(i).X + TransX
            shp.Points(i).Y = shp.Points(i).Y + TransY
      Next i
End Sub
'/////////////////////////////////////////////////////////////////////////////////////////////


'/////////////////////////////////////////////////////////////////////////////////////////////
'/// Initialize bezier curve's points
'/////////////////////////////////////////////////////////////////////////////////////////////
Public Sub InitializeBezierData(ByRef shp As BEZIER_CURVE, ByVal nWidth As Integer, ByVal nHeight As Integer)
      Dim X1 As Integer, Y1 As Integer    'First curve's starting point
      Dim X2 As Integer, Y2 As Integer    'First curve's first control point
      Dim X3 As Integer, Y3 As Integer    'First curve's second control point
      Dim X4 As Integer, Y4 As Integer    'First curve's ending point, second curve's starting point
      Dim X5 As Integer, Y5 As Integer    'Second curve's first control point
      Dim X6 As Integer, Y6 As Integer    'Second curve's second control point
      Dim X7 As Integer, Y7 As Integer    'Second curve's ending point
      
      '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
      '/// First curve (first curve's ending point is (X4, Y4), which is second curve's starting point
      '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
      X1 = 0
      Y1 = 0
      
      X2 = X1 + (nWidth * X2RATIO)
      Y2 = Y1 + (nHeight * Y2RATIO)
      
      X3 = X1 + (nWidth * X3RATIO)
      Y3 = Y1 - (nHeight * Y3RATIO)
      
      X4 = X1 + (nWidth * X4RATIO)
      Y4 = Y1 - (nHeight * Y4RATIO)
      
      '/////////////////////////////////////////////////////////////////////////////////////////////
      '/// Second curve (second curve's starting point is (X4, Y4) )
      '/////////////////////////////////////////////////////////////////////////////////////////////
      X7 = X1 + nWidth
      Y7 = Y1
      
      X6 = X7 - (nWidth * X6RATIO)
      Y6 = Y7 + (nHeight * Y6RATIO)
      
      X5 = X7 - (nWidth * X5RATIO)
      Y5 = Y7 - (nHeight * Y5RATIO)
      
      
      '/////////////////////////////////////////////////////////////////////////////////////////////
      shp.Points(1).X = X1:     shp.Points(1).Y = Y1
      shp.Points(2).X = X2:     shp.Points(2).Y = Y2
      shp.Points(3).X = X3:     shp.Points(3).Y = Y3
      shp.Points(4).X = X4:     shp.Points(4).Y = Y4
      shp.Points(5).X = X5:     shp.Points(5).Y = Y5
      shp.Points(6).X = X6:     shp.Points(6).Y = Y6
      shp.Points(7).X = X7:     shp.Points(7).Y = Y7
      
      shp.NumPoints = 7
End Sub
