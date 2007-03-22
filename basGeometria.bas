Attribute VB_Name = "basGeometria"
Option Explicit
Public Enum Tipo_De_Objeto
 PONTO 'PONTO = 0. � o valor padr�o de novos objetos.
 'PONTO_SOBRE 'PONTO_DE_INTERSEC��O
 SEGMENTO
 'VETOR 'SEMI_RETA 'RETA 'TRI�NGULO 'POL�GONO 'POL�GONO_REGULAR 'CIRCUNFER�NCIA
 'ARCO 'C�NICA 'PARALELA 'PERPENDICULAR 'MEDIATRIZ 'PONTO_M�DIO 'BISSETRIZ_PONTOS
 'BISSETRIZ_RETAS  'COMPASSO 'REFLEX�O 'SIMETRIA 'TRANSLA��O 'INVERSO_CIRCUNFER�NCIA
 'TEXTO '�NGULO 'EIXOS
End Enum

Public Enum Apar�ncia
 PADRAO 'O objeto � desenhado normalmente
 SELECIONADO 'O objeto � desenhado com destaque
 OCULTO 'O objeto n�o ser� desenhado
End Enum

Type Objeto
 Id As Integer 'Identifica exclusivamente o objeto (por tipo??). Como o indice de Obj?
 Tipo As Tipo_De_Objeto 'Que item ser� guardado
 N_Param As Byte 'N�mero de objetos dos quais este � dependente
 Cor As Long 'Cor utilizada para desenhar na tela
 Espessura As Byte 'Raio dos pontos ou a largura de curvas e contornos
 Tra�o(1 To 2) As Byte 'Tipo de pontilhado
 Mostrar As Apar�ncia 'Indica como o objeto ser� exibido
 Nome As String 'Um r�tulo para exibi��o em tela
 Id_Dep() As Long 'Indices dos parametros (objetos)dos quais este depende
 Coord(0 To 2) As GLdouble 'Coordenadas e angulos livres
End Type

Public Const MAX_OBJETOS = 10000
Public Const TAM_BUFER = MAX_OBJETOS
Public Qtd_Obj As Long '� sempre = Ubound(Obj)
Public Obj() As Objeto
Public P_Aux(0 To 2) As GLdouble 'Ponto auxiliar durante a defini��o de objetos
Public Posicionando As Boolean 'Indica se est� sendo posicionado um ponto no espa�o
Public Estado_Teclas As Integer 'Indica se ALT, CTRL e SHIFT est�o pressionadas

Public Sub Inicializa()
 ReDim basGeometria.Obj(1 To 1)
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
    For k = -RAIO To RAIO 'Desenhar "bola" sobre yOz (lembre que o sistema � negativo)
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
    For k = -RAIO To RAIO 'Desenhar "bola" sobre xOz (lembre que o sistema � negativo)
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
 'glLineWidth (2#)
 glBegin bmLines
   glColor3f 1#, 0#, 0#
   glVertex3f 0#, 0#, 0#
   glVertex3f 3#, 0#, 0#
   
   glColor3f 0#, 1#, 0#
   glVertex3f 0#, 0#, 0#
   glVertex3f 0#, 3#, 0#
   
   glColor3f 0#, 0#, 1#
   glVertex3f 0#, 0#, 0#
   glVertex3f 0#, 0#, 3#
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
 
 'j� ocorreu um glPushName 0, inicializando a pilha de nomes arbitrariamente
 
 glColor3d 0#, 0#, 0#
 glPointSize (3#)

 For i = 1 To basGeometria.Qtd_Obj
  If Modo = GL_SELECT Then glLoadName i
  If basGeometria.Obj(i).Mostrar = SELECIONADO Then glColor3d 0.9, 0.4, 0#: glPointSize (5#)
  glBegin bmPoints
   glVertex3dv basGeometria.Obj(i).Coord(0)
  glEnd
  If basGeometria.Obj(i).Mostrar = SELECIONADO Then glColor3d 0#, 0#, 0#: glPointSize (3#)
 Next i
 
 Select Case UCase(Ferram)
  Case "PONTEIRO"
  
  Case "PONTO"
   If Posicionando Then Des_Ponto_Aux (Estado_Teclas)
  Case "SEGMENTO"
  
 End Select
  
End Sub
Public Function Avalia_Selecao(hits As GLint, Buf() As GLuint) As String
 Static Sel_Ant As GLuint 'Indice do objeto que estava selecionado no movimento anterior
 Dim h As Long, Id As Long
 Dim Qtd_Nomes As GLuint 'Cada nome � composto de 'tantas' coordenadas
 Dim Nome As GLuint 'Indice do objeto que est� selecionado ao clicar o mouse
 Dim MinZ As Double
 
 Id = 0
 MinZ = 2121212121 'inicializa minZ para um valor grande
 Nome = -1 'Nada selecionado at� agora
 
'Para compreender o la�o "FOR NEXT", lembre-se do formato de cada REGISTRO (HIT)...
' Reg1: |    Buf(0)            |  Buf(1)   |  Buf(2)   |     Buf( 3... 3+Qtd_Nomes)      |
'       | Qtd de Nomes em Reg1 | Z m�nimo  | Z m�ximo  | Nomes deste Registro (de 0 a n) |
' - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
' Reg2: |  Buf(0 + 3+Qtd_Nomes)| e assim vai...
  
 For h = 1 To hits
  Qtd_Nomes = Buf(Id) 'cada nome � composto de 'tantas' coordenadas
  If (Buf(Id + 1) < MinZ) And (Qtd_Nomes > 0) Then
   MinZ = Buf(Id + 1)
   Nome = Buf(Id + 3)
  End If
  Id = Id + 3 + Qtd_Nomes
 Next h
 
 If Nome > 0 Then
  Avalia_Selecao = "Ponto " & Nome
  If Sel_Ant > 0 Then basGeometria.Obj(Sel_Ant).Mostrar = PADRAO
  basGeometria.Obj(Nome).Mostrar = SELECIONADO
  Sel_Ant = Nome
 Else
  Avalia_Selecao = ""
  If Sel_Ant > 0 Then basGeometria.Obj(Sel_Ant).Mostrar = PADRAO
 End If
 'For h = 1 To Qtd_Obj
 '  If Nome <> h Then basGeometria.Obj(h).Mostrar = PADRAO 'terei problema com objetos ocultos
 'Next h

End Function
