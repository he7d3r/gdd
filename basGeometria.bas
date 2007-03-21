Attribute VB_Name = "basGeometria"
Option Explicit

Type Objeto
   Selecionado As Boolean
   Coord(0 To 2) As GLdouble
End Type
Public Const MAX_OBJETOS = 10000
Public Const TAM_BUFER = MAX_OBJETOS
Public Obj() As Objeto
Public Qtd_Obj As Long

