Attribute VB_Name = "basGeometria"
Option Explicit
Type Ponto
   coord(0 To 2) As GLdouble
End Type
Public Const MAX_PONTOS = 10
Public Pts() As Ponto
Public Qtd_Pts As Long

Sub Gera()
Dim i As Long, t As GLdouble

ReDim Pts(1 To 1)
For i = 1 To MAX_PONTOS
 Qtd_Pts = i
 ReDim Preserve Pts(1 To Qtd_Pts)
 
 t = i / MAX_PONTOS - 0.5
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
  For i = 1 To Qtd_Pts
   glVertex3dv Pts(i).coord(0)
  Next i
  glEnd
End Sub
