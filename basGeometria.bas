Attribute VB_Name = "basGeometria"
Option Explicit

Public Sub Inicializa_Objetos(IdDoc As Integer)
 ReDim Doc(IdDoc).Obj(1 To 1)
 ReDim Doc(IdDoc).Obj_Sel(1 To 1)
End Sub
Public Sub Inverter_Todos(IdDoc As Integer)
   Dim N_Obj As Integer
   Dim i As Integer, s As Integer
   With Doc(IdDoc)
      N_Obj = UBound(.Obj)
      .frm.N_Sel = N_Obj - .frm.N_Sel
      If .frm.N_Sel = 0 Then
         ReDim .Obj_Sel(1 To 1)
      Else
         ReDim .Obj_Sel(1 To .frm.N_Sel)
      End If
      s = 1
      For i = 1 To N_Obj
         If .Obj(i).Selec Then
            .Obj(i).Selec = 0
         Else
            .Obj(i).Selec = s 'vale sempre 's<=i'
            .Obj_Sel(s) = i
            s = s + 1
         End If
      Next i
   End With
End Sub
Public Sub Marcar_Todos(IdDoc As Integer, Selecionar As Boolean)
   Dim N_Obj As Integer
   Dim i As Integer
   
   With Doc(IdDoc)
      N_Obj = UBound(.Obj)
      If Selecionar = True Then
         .frm.N_Sel = N_Obj
         ReDim .Obj_Sel(1 To N_Obj)
         For i = 1 To N_Obj
            .Obj(i).Selec = i
            .Obj_Sel(i) = i
         Next i
      Else
         .frm.N_Sel = 0
         ReDim .Obj_Sel(1 To 1)
         For i = 1 To N_Obj
            .Obj(i).Selec = 0
         Next i
      End If
   End With
End Sub

Public Function Aponta_Objeto(IdDoc As Integer, hits As GLint, Buf() As GLuint) As String
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
  Aponta_Objeto = "Ponto " & Nome
  Doc(IdDoc).frm.ObjApontado = Nome
 Else
  Aponta_Objeto = ""
  Doc(IdDoc).frm.ObjApontado = 0
 End If
End Function

Public Sub Des_Plano(Plano As Tipo_De_Plano, Aux() As GLdouble)
   Const RAIO = 3
   Dim k As GLdouble
   Dim PosX As GLdouble, PosY As GLdouble, PosZ As GLdouble
      
   glColor3f 0.5, 0.5, 0.5
   'glLineWidth (1#)
   glBegin bmLines
   Select Case Plano
   Case PL_HORIZONTAL
      For k = -RAIO To RAIO 'Desenhar "bola" sobre xOy
         PosX = Fix(Aux(0) + k): PosY = Fix(Aux(1) + k)
         If Abs(PosX - Aux(0)) < RAIO Then
            glVertex3d PosX, Aux(1) + (RAIO - Abs(PosX - Aux(0))), 0#
            glVertex3d PosX, Aux(1) - (RAIO - Abs(PosX - Aux(0))), 0#
         End If
         If Abs(PosY - Aux(1)) < RAIO Then
            glVertex3d Aux(0) + (RAIO - Abs(PosY - Aux(1))), PosY, 0#
            glVertex3d Aux(0) - (RAIO - Abs(PosY - Aux(1))), PosY, 0#
         End If
      Next k
   Case PL_PERFIL
      For k = -RAIO To RAIO 'Desenhar "bola" sobre yOz (lembre que o sistema é negativo)
         PosZ = Fix(Aux(2) + k): PosY = Fix(Aux(1) + k)
         If Abs(PosZ - Aux(2)) < RAIO Then
            glVertex3d 0#, Aux(1) + (RAIO - Abs(PosZ - Aux(2))), PosZ
            glVertex3d 0#, Aux(1) - (RAIO - Abs(PosZ - Aux(2))), PosZ
         End If
         If Abs(PosY - Aux(1)) < RAIO Then
            glVertex3d 0#, PosY, Aux(2) + (RAIO - Abs(PosY - Aux(1)))
            glVertex3d 0#, PosY, Aux(2) - (RAIO - Abs(PosY - Aux(1)))
         End If
      Next k
   Case PL_FRONTAL
      For k = -RAIO To RAIO 'Desenhar "bola" sobre xOz (lembre que o sistema é negativo)
         PosX = Fix(Aux(0) + k): PosZ = Fix(Aux(2) + k)
         If Abs(PosX - Aux(0)) < RAIO Then
            glVertex3d PosX, 0#, Aux(2) + (RAIO - Abs(PosX - Aux(0)))
            glVertex3d PosX, 0#, Aux(2) - (RAIO - Abs(PosX - Aux(0)))
         End If
         If Abs(PosZ - Aux(2)) < RAIO Then
            glVertex3d Aux(0) + (RAIO - Abs(PosZ - Aux(2))), 0#, PosZ
            glVertex3d Aux(0) - (RAIO - Abs(PosZ - Aux(2))), 0#, PosZ
         End If
      Next k
   End Select
   glEnd
End Sub
Public Sub Des_Eixos()
 Const PONTA = 3
 Const INI_SETA = PONTA - PONTA / 10
 Const ABERTURA_SETA = PONTA / 20
 'glLineWidth (2#)
 glBegin bmLines
   glColor3f 1#, 0#, 0#
   glVertex3f 0#, 0#, 0#
   glVertex3f PONTA, 0#, 0#
     glVertex3f PONTA, 0#, 0#: glVertex3f INI_SETA, ABERTURA_SETA, 0#
     glVertex3f PONTA, 0#, 0#: glVertex3f INI_SETA, -ABERTURA_SETA, 0#
     glVertex3f PONTA, 0#, 0#: glVertex3f INI_SETA, 0#, ABERTURA_SETA
     glVertex3f PONTA, 0#, 0#: glVertex3f INI_SETA, 0#, -ABERTURA_SETA
     
   glColor3f 0#, 1#, 0#
   glVertex3f 0#, 0#, 0#
   glVertex3f 0#, PONTA, 0#
     glVertex3f 0#, PONTA, 0#: glVertex3f 0#, INI_SETA, ABERTURA_SETA
     glVertex3f 0#, PONTA, 0#: glVertex3f 0#, INI_SETA, -ABERTURA_SETA
     glVertex3f 0#, PONTA, 0#: glVertex3f ABERTURA_SETA, INI_SETA, 0#
     glVertex3f 0#, PONTA, 0#: glVertex3f -ABERTURA_SETA, INI_SETA, 0#
     
   glColor3f 0#, 0#, 1#
   glVertex3f 0#, 0#, 0#
   glVertex3f 0#, 0#, PONTA
     glVertex3f 0#, 0#, PONTA: glVertex3f 0#, ABERTURA_SETA, INI_SETA
     glVertex3f 0#, 0#, PONTA: glVertex3f 0#, -ABERTURA_SETA, INI_SETA
     glVertex3f 0#, 0#, PONTA: glVertex3f ABERTURA_SETA, 0#, INI_SETA
     glVertex3f 0#, 0#, PONTA: glVertex3f -ABERTURA_SETA, 0#, INI_SETA
 glEnd
End Sub
Public Sub Des_Ponto_Aux(Plano As Tipo_De_Plano, Aux() As GLdouble)
   glColor3d 1#, 0.4, 0.1
   
   glPointSize (3#)
   glBegin bmPoints
      glVertex3dv Aux(0)
   glEnd
   'If Not Posicionando Then Exit Sub
   glColor3d 0.7, 0.7, 0.7
   glBegin bmLines
      glVertex3dv Aux(0) ', Aux(1), Aux(2)
      Select Case Plano
      Case PL_HORIZONTAL 'Segmento vertical
         glVertex3d Aux(0), Aux(1), 0#
      Case PL_PERFIL 'Segmento fronto-horizontal
         glVertex3d 0#, Aux(1), Aux(2)
      Case PL_FRONTAL 'Segmento de topo
         glVertex3d Aux(0), 0#, Aux(2)
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
Public Sub Des_Objetos(IdDoc As Integer, Modo As GLenum, Ferram As String)
 Dim i As Long
 Dim N_Obj As Long
 
 'já ocorreu um glPushName 0, inicializando a pilha de nomes arbitrariamente
 
 glColor3d 0#, 0#, 0#
 glPointSize (3#)
 'CAUSOU UM PONTO NA ORIGEM
 N_Obj = UBound(Doc(IdDoc).Obj)
 For i = 1 To N_Obj
  If Modo = GL_SELECT Then glLoadName i
  If i = Doc(IdDoc).frm.ObjApontado Then
   glColor3d 0.8, 0#, 0.5: glPointSize (5#)
   glBegin bmPoints
     glVertex3dv Doc(IdDoc).Obj(i).Coord(0)
   glEnd
   glColor3d 0#, 0#, 0#: glPointSize (3#)
  ElseIf Doc(IdDoc).Obj(i).Selec > 0 Then
   glColor3d 0.9, 0.4, 0#: glPointSize (3#)
   glBegin bmPoints
     glVertex3dv Doc(IdDoc).Obj(i).Coord(0)
   glEnd
   glColor3d 0#, 0#, 0#: glPointSize (3#)
  Else
   glBegin bmPoints
     glVertex3dv Doc(IdDoc).Obj(i).Coord(0)
   glEnd
  End If
 Next i
  
 Select Case UCase(Ferram)
  Case "PONTEIRO"
  
  Case "PONTO"
   If Doc(IdDoc).frm.Posicionando Then Des_Ponto_Aux Sobre_Plano, P_Aux
   
  Case "SEGMENTO"
  
 End Select
  
End Sub
