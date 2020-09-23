Attribute VB_Name = "basMyDeclarations"
Option Explicit

'NOTE : A point's x and y positions are measured from bottom-left corner of the piece.

'X~RATIO  =     point's x position / piece's width
'Y~RATIO  =     point's y position / piece's height
Public Const X2RATIO As Double = 0.760869565217391
Public Const Y2RATIO As Double = 0.183946488294314

Public Const X3RATIO As Double = 8.02675585284281E-02
Public Const Y3RATIO As Double = 0.150501672240803

Public Const X4RATIO As Double = 0.5
Public Const Y4RATIO As Double = Y2RATIO

Public Const X6RATIO As Double = X2RATIO
Public Const Y6RATIO As Double = Y2RATIO

Public Const X5RATIO As Double = X3RATIO
Public Const Y5RATIO As Double = Y3RATIO
      
Public Const NUM_BEZIER_POINTS As Integer = 7

Public Type BEZIER_CURVE
      Points() As POINTAPI    'Array of point structure
      NumPoints As Integer    'Number of point structure that Points() contains
End Type
