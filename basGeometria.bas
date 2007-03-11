Attribute VB_Name = "basGeometria"
Option Explicit
Type Ponto
   coord(0 To 2) As GLdouble
End Type
Const TANTO = 100
Public Pts(1 To TANTO) As Ponto

Sub Gera()
Dim i As Long, t As GLdouble

For i = 1 To TANTO
 t = i / TANTO - 0.5
 Pts(i).coord(0) = Exp(2 * t) * Cos(2 * PI * t)
 Pts(i).coord(1) = Exp(-2 * t) * Sin(2 * PI * t)
 Pts(i).coord(2) = 2 * t
Next i
End Sub

Sub Des_pontinhos()
 Dim i As Long
 
  glColor3d 0.9, 0.8, 0.1
  glPointSize (2#)
  glBegin bmPoints
  For i = 1 To TANTO
   glVertex3dv Pts(i).coord(0)
  Next i
  glEnd
End Sub
