Attribute VB_Name = "basGeometria"
Option Explicit
Type Ponto
   coord(0 To 2) As GLdouble
End Type
Public Const MAX_PONTOS = 10000
Public Const TAM_BUFER = MAX_PONTOS
Public Pts() As Ponto
Public Qtd_Pts As Long

