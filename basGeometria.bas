Attribute VB_Name = "basGeometria"
Option Explicit
Public Enum Tipo_De_Objeto
 PONTO 'PONTO = 0. É o valor padrão de novos objetos.
 'PONTO_SOBRE 'PONTO_DE_INTERSECÇÃO
 SEGMENTO
 'VETOR 'SEMI_RETA 'RETA 'TRIÂNGULO 'POLÍGONO 'POLÍGONO_REGULAR 'CIRCUNFERÊNCIA
 'ARCO 'CÔNICA 'PARALELA 'PERPENDICULAR 'MEDIATRIZ 'PONTO_MÉDIO 'BISSETRIZ_PONTOS
 'BISSETRIZ_RETAS  'COMPASSO 'REFLEXÃO 'SIMETRIA 'TRANSLAÇÃO 'INVERSO_CIRCUNFERÊNCIA
 'TEXTO 'ÂNGULO 'EIXOS
End Enum

Type Objeto
 'Id As Integer 'Identifica exclusivamente o objeto (por tipo??). Como o indice de Obj?
 'Tipo As Tipo_De_Objeto 'Que item será guardado
 'N_Param As Byte 'Número de objetos dos quais este é dependente
 'Cor As Long 'Cor utilizada para desenhar na tela
 'Espessura As Byte 'Raio dos pontos ou a largura de curvas e contornos
 'Traço(1 To 2) As Byte 'Tipo de pontilhado
 Selec As Integer 'Indica que o objeto foi o "selec-ésimo" a ser selecionado
 Mostrar As Boolean ' Aparência 'Indica como o objeto será exibido
 'Nome As String 'Um rótulo para exibição em tela
 'Id_Dep() As Long 'Indices dos parametros (objetos)dos quais este depende
 Coord(0 To 2) As GLdouble 'Coordenadas e angulos livres
End Type
Public ObjApontado As Long

Public Const MAX_OBJETOS = 10000
Public Const TAM_BUFER = MAX_OBJETOS
Public Qtd_Obj As Long 'É sempre = Ubound(Obj)
Public Obj() As Objeto
Public Obj_Sel() As Integer
Public N_Sel As Integer 'É sempre = Ubound(Obj_Sel)
Public P_Aux(0 To 2) As GLdouble 'Ponto auxiliar durante a definição de objetos
Public Posicionando As Boolean 'Indica se está sendo posicionado um ponto no espaço
Public Estado_Teclas As Integer 'Indica se ALT, CTRL e SHIFT estão pressionadas

Public Sub Inicializa()
 ReDim Obj(1 To 1)
 ReDim Obj_Sel(1 To 1)
End Sub

Public Sub Des_Plano(Estado As Integer)
 Const RAIO = 3
 Dim k As GLdouble
 Dim PosX As GLdouble, PosY As GLdouble, PosZ As GLdouble
 
 'If Not Posicionando Then Exit Sub
 
 glColor3f 0.5, 0.5, 0.5
 'glLineWidth (1#)
 glBegin bmLines
  Select Case Estado
  Case 0, vbCtrlMask
    For k = -RAIO To RAIO 'Desenhar "bola" sobre xOy
      PosX = Fix(P_Aux(0) + k): PosY = Fix(P_Aux(1) + k)
      If Abs(PosX - P_Aux(0)) < RAIO Then
      glVertex3d PosX, P_Aux(1) + (RAIO - Abs(PosX - P_Aux(0))), 0#
      glVertex3d PosX, P_Aux(1) - (RAIO - Abs(PosX - P_Aux(0))), 0#
      End If
      If Abs(PosY - P_Aux(1)) < RAIO Then
      glVertex3d P_Aux(0) + (RAIO - Abs(PosY - P_Aux(1))), PosY, 0#
      glVertex3d P_Aux(0) - (RAIO - Abs(PosY - P_Aux(1))), PosY, 0#
      End If
    Next k
  Case vbShiftMask, vbShiftMask + vbCtrlMask
    For k = -RAIO To RAIO 'Desenhar "bola" sobre yOz (lembre que o sistema é negativo)
      PosZ = Fix(P_Aux(2) + k): PosY = Fix(P_Aux(1) + k)
      If Abs(PosZ - P_Aux(2)) < RAIO Then
      glVertex3d 0#, P_Aux(1) + (RAIO - Abs(PosZ - P_Aux(2))), PosZ
      glVertex3d 0#, P_Aux(1) - (RAIO - Abs(PosZ - P_Aux(2))), PosZ
      End If
      If Abs(PosY - P_Aux(1)) < RAIO Then
      glVertex3d 0#, PosY, P_Aux(2) + (RAIO - Abs(PosY - P_Aux(1)))
      glVertex3d 0#, PosY, P_Aux(2) - (RAIO - Abs(PosY - P_Aux(1)))
      End If
    Next k
  Case vbAltMask, vbAltMask + vbCtrlMask
    For k = -RAIO To RAIO 'Desenhar "bola" sobre xOz (lembre que o sistema é negativo)
      PosX = Fix(P_Aux(0) + k): PosZ = Fix(P_Aux(2) + k)
      If Abs(PosX - P_Aux(0)) < RAIO Then
      glVertex3d PosX, 0#, P_Aux(2) + (RAIO - Abs(PosX - P_Aux(0)))
      glVertex3d PosX, 0#, P_Aux(2) - (RAIO - Abs(PosX - P_Aux(0)))
      End If
      If Abs(PosZ - P_Aux(2)) < RAIO Then
      glVertex3d P_Aux(0) + (RAIO - Abs(PosZ - P_Aux(2))), 0#, PosZ
      glVertex3d P_Aux(0) - (RAIO - Abs(PosZ - P_Aux(2))), 0#, PosZ
      End If
    Next k
  End Select
 glEnd
End Sub
Public Sub Des_Eixos()
 Const PONTA = 3
 Const INI_SETA = PONTA - PONTA / 10
 Const AF_SETA = PONTA / 20
 'glLineWidth (2#)
 glBegin bmLines
   glColor3f 1#, 0#, 0#
   glVertex3f 0#, 0#, 0#
   glVertex3f PONTA, 0#, 0#
     glVertex3f PONTA, 0#, 0#: glVertex3f INI_SETA, AF_SETA, 0#
     glVertex3f PONTA, 0#, 0#: glVertex3f INI_SETA, -AF_SETA, 0#
     glVertex3f PONTA, 0#, 0#: glVertex3f INI_SETA, 0#, AF_SETA
     glVertex3f PONTA, 0#, 0#: glVertex3f INI_SETA, 0#, -AF_SETA
     
   glColor3f 0#, 1#, 0#
   glVertex3f 0#, 0#, 0#
   glVertex3f 0#, PONTA, 0#
     glVertex3f 0#, PONTA, 0#: glVertex3f 0#, INI_SETA, AF_SETA
     glVertex3f 0#, PONTA, 0#: glVertex3f 0#, INI_SETA, -AF_SETA
     glVertex3f 0#, PONTA, 0#: glVertex3f AF_SETA, INI_SETA, 0#
     glVertex3f 0#, PONTA, 0#: glVertex3f -AF_SETA, INI_SETA, 0#
     
   glColor3f 0#, 0#, 1#
   glVertex3f 0#, 0#, 0#
   glVertex3f 0#, 0#, PONTA
     glVertex3f 0#, 0#, PONTA: glVertex3f 0#, AF_SETA, INI_SETA
     glVertex3f 0#, 0#, PONTA: glVertex3f 0#, -AF_SETA, INI_SETA
     glVertex3f 0#, 0#, PONTA: glVertex3f AF_SETA, 0#, INI_SETA
     glVertex3f 0#, 0#, PONTA: glVertex3f -AF_SETA, 0#, INI_SETA
 glEnd
End Sub
Public Sub Des_Ponto_Aux(Estado As Integer)
  glColor3d 1#, 0.4, 0.1
  
  glPointSize (3#)
  glBegin bmPoints
    glVertex3dv P_Aux(0)
  glEnd
  'If Not Posicionando Then Exit Sub
  glColor3d 0.7, 0.7, 0.7
  glBegin bmLines
   glVertex3d P_Aux(0), P_Aux(1), P_Aux(2)
   Select Case Estado
   Case 0, vbCtrlMask 'Segmento vertical
     glVertex3d P_Aux(0), P_Aux(1), 0#
   Case vbShiftMask, vbShiftMask + vbCtrlMask 'Segmento fronto-horizontal
     glVertex3d 0#, P_Aux(1), P_Aux(2)
   Case vbAltMask, vbAltMask + vbCtrlMask 'Segmento de topo
     glVertex3d P_Aux(0), 0#, P_Aux(2)
   End Select
  glEnd
End Sub
'Public Sub Des_Figura()
 'glPushMatrix
 ' glTranslatef -0.5, 1.5, 0.5
 ' glColor3d 0, 0, 0
 ' gluCylinder QObj, 1.5, 0.5, 2, 12, 2
 'glPopMatrix
'End Sub
Public Sub Des_LT()
 glBegin GL_LINES
  glColor3d 0.5, 0, 0
  glVertex3f -3, 0, 0
  glVertex3f 3, 0, 0
 glEnd
 glPointSize 3#
 glBegin GL_POINTS
  glColor3d 0.5, 0, 0
  glVertex3f 0, 0, 0
 glEnd
End Sub
Public Sub Des_Objetos(Modo As GLenum, Ferram As String)
 Dim i As Long
 
 'já ocorreu um glPushName 0, inicializando a pilha de nomes arbitrariamente
 
 glColor3d 0#, 0#, 0#
 glPointSize (3#)

 For i = 1 To basGeometria.Qtd_Obj
  If Modo = GL_SELECT Then glLoadName i
  If i = ObjApontado Then
   glColor3d 0.8, 0#, 0.5: glPointSize (5#)
   glBegin bmPoints
     glVertex3dv basGeometria.Obj(i).Coord(0)
   glEnd
   glColor3d 0#, 0#, 0#: glPointSize (3#)
  ElseIf basGeometria.Obj(i).Selec > 0 Then
   glColor3d 0.9, 0.4, 0#: glPointSize (3#)
   glBegin bmPoints
     glVertex3dv basGeometria.Obj(i).Coord(0)
   glEnd
   glColor3d 0#, 0#, 0#: glPointSize (3#)
  Else
   glBegin bmPoints
     glVertex3dv basGeometria.Obj(i).Coord(0)
   glEnd
  End If
 Next i
  
 Select Case UCase(Ferram)
  Case "PONTEIRO"
  
  Case "PONTO"
   If Posicionando Then Des_Ponto_Aux (Estado_Teclas)
  Case "SEGMENTO"
  
 End Select
  
End Sub
Public Function ApontaObjeto(hits As GLint, Buf() As GLuint) As String
 'Static Sel_Ant As GLuint 'Indice do objeto que estava selecionado no movimento anterior
 Dim h As Long, Id As Long
 Dim Qtd_Nomes As GLuint 'Cada nome é composto de 'tantas' coordenadas
 Dim Nome As GLuint 'Indice do objeto que está selecionado ao clicar o mouse
 Dim MinZ As Double
 
 Id = 0
 MinZ = 2121212121 'inicializa minZ para um valor grande
 Nome = -1 'Nada selecionado até agora
 
'Para compreender o laço "FOR NEXT", lembre-se do formato de cada REGISTRO (HIT)...
' Reg1: |    Buf(0)            |  Buf(1)   |  Buf(2)   |     Buf( 3... 3+Qtd_Nomes)      |
'       | Qtd de Nomes em Reg1 | Z mínimo  | Z máximo  | Nomes deste Registro (de 0 a n) |
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Reg2: |  Buf(0 + 3+Qtd_Nomes)| e assim vai...
  
 For h = 1 To hits
  Qtd_Nomes = Buf(Id) 'cada nome é composto de 'tantas' coordenadas
  If (Buf(Id + 1) < MinZ) And (Qtd_Nomes > 0) Then
   MinZ = Buf(Id + 1)
   Nome = Buf(Id + 3)
  End If
  Id = Id + 3 + Qtd_Nomes
 Next h
 
 If Nome > 0 Then
  ApontaObjeto = "Ponto " & Nome
  ObjApontado = Nome
 Else
  ApontaObjeto = ""
  ObjApontado = 0
 End If
End Function
