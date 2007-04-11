Attribute VB_Name = "basGeometria"
Option Explicit

Public Sub Inicializa_Objetos(IdDoc As Integer)
 ReDim Doc(IdDoc).Obj(1 To 1)
 ReDim Doc(IdDoc).Obj_Sel(1 To 1)
 ReDim Obj_Aux(1 To 1)
 Obj_Aux(1).Coord(3) = 1
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

Public Function Aponta_Primeiro_Objeto(ByVal IdDoc As Integer, _
                                       ByVal hits As GLint, _
                                       ByRef Buf() As GLuint) As String
   Dim h As Long, Id As Long
   Dim Qtd_Nomes As GLuint 'Cada nome é composto de 'tantas' coordenadas
   Dim Nome As GLuint 'Indice do objeto que está selecionado ao clicar o mouse
   Dim MinZ As Double
   
   Id = 0
   MinZ = 1E+60 'inicializa minZ para um valor grande
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
      Aponta_Primeiro_Objeto = "Ponto " & Nome
      Doc(IdDoc).frm.ObjApontado = Nome
   Else
      Aponta_Primeiro_Objeto = ""
      Doc(IdDoc).frm.ObjApontado = 0
   End If
   
End Function
Public Sub Des_Planos()
 Const INI_PLANO = -5
 Const FIM_PLANO = 5
 
 Dim k As GLdouble
  
 glColor4f 0.7, 0.7, 0.7, 0.1
 glLineWidth (1#)
 glPolygonMode GL_FRONT_AND_BACK, pgmFILL
 glBegin bmPolygon
   glEdgeFlag GL_FALSE
   glVertex3d INI_PLANO, INI_PLANO, 0
   glEdgeFlag GL_FALSE
   glVertex3d FIM_PLANO, INI_PLANO, 0
   glEdgeFlag GL_FALSE
   glVertex3d FIM_PLANO, FIM_PLANO, 0
   glEdgeFlag GL_FALSE
   glVertex3d INI_PLANO, FIM_PLANO, 0
 glEnd
 Exit Sub
 glBegin bmLines
 For k = INI_PLANO To FIM_PLANO
  glVertex3d k, INI_PLANO, 0 'Plano Horizontal (PI')
  glVertex3d k, FIM_PLANO, 0
  glVertex3d INI_PLANO, k, 0
  glVertex3d FIM_PLANO, k, 0
  
  glVertex3d k, 0, INI_PLANO 'Plano Frontal (PI'')
  glVertex3d k, 0, FIM_PLANO
  glVertex3d INI_PLANO, 0, k
  glVertex3d FIM_PLANO, 0, k
  
  glVertex3d 0, k, INI_PLANO 'Plano de Perfil (PI''')
  glVertex3d 0, k, FIM_PLANO
  glVertex3d 0, INI_PLANO, k
  glVertex3d 0, FIM_PLANO, k
 Next k
 glEnd
 
 Erro = glGetError: If Erro <> glerrNoError Then ErroFatal Erro
End Sub

Public Sub Des_Plano(Plano As Tipo_De_Plano, Aux() As GLdouble)
   Const RAIO = 3
   Dim k As GLdouble
   Dim PosX As GLdouble, PosY As GLdouble, PosZ As GLdouble
   Dim Pt(0 To 2) As GLdouble
   
   Pt(0) = Aux(0) / Aux(3)
   Pt(1) = Aux(1) / Aux(3)
   Pt(2) = Aux(2) / Aux(3)
      
   glColor3f 0.5, 0.5, 0.5
   glLineWidth (1#)
   glBegin bmLines
   Select Case Plano
   Case PL_HORIZONTAL
      For k = -RAIO To RAIO 'Desenhar "bola" sobre xOy
         PosX = Fix(Pt(0) + k): PosY = Fix(Pt(1) + k)
         If Abs(PosX - Pt(0)) < RAIO Then
            glVertex3d PosX, Pt(1) + (RAIO - Abs(PosX - Pt(0))), 0#
            glVertex3d PosX, Pt(1) - (RAIO - Abs(PosX - Pt(0))), 0#
         End If
         If Abs(PosY - Pt(1)) < RAIO Then
            glVertex3d Pt(0) + (RAIO - Abs(PosY - Pt(1))), PosY, 0#
            glVertex3d Pt(0) - (RAIO - Abs(PosY - Pt(1))), PosY, 0#
         End If
      Next k
   Case PL_PERFIL
      For k = -RAIO To RAIO 'Desenhar "bola" sobre yOz (lembre que o sistema é negativo)
         PosZ = Fix(Pt(2) + k): PosY = Fix(Pt(1) + k)
         If Abs(PosZ - Pt(2)) < RAIO Then
            glVertex3d 0#, Pt(1) + (RAIO - Abs(PosZ - Pt(2))), PosZ
            glVertex3d 0#, Pt(1) - (RAIO - Abs(PosZ - Pt(2))), PosZ
         End If
         If Abs(PosY - Pt(1)) < RAIO Then
            glVertex3d 0#, PosY, Pt(2) + (RAIO - Abs(PosY - Pt(1)))
            glVertex3d 0#, PosY, Pt(2) - (RAIO - Abs(PosY - Pt(1)))
         End If
      Next k
   Case PL_FRONTAL
      For k = -RAIO To RAIO 'Desenhar "bola" sobre xOz (lembre que o sistema é negativo)
         PosX = Fix(Pt(0) + k): PosZ = Fix(Pt(2) + k)
         If Abs(PosX - Pt(0)) < RAIO Then
            glVertex3d PosX, 0#, Pt(2) + (RAIO - Abs(PosX - Pt(0)))
            glVertex3d PosX, 0#, Pt(2) - (RAIO - Abs(PosX - Pt(0)))
         End If
         If Abs(PosZ - Pt(2)) < RAIO Then
            glVertex3d Pt(0) + (RAIO - Abs(PosZ - Pt(2))), 0#, PosZ
            glVertex3d Pt(0) - (RAIO - Abs(PosZ - Pt(2))), 0#, PosZ
         End If
      Next k
   End Select
   glEnd
   
   Erro = glGetError: If Erro <> glerrNoError Then ErroFatal Erro
End Sub
Public Sub Des_Eixos()
 Const PONTA = 3
 Const INI_SETA = PONTA - PONTA / 10
 Const ABERTURA_SETA = PONTA / 20
 glLineWidth (1#)
 glBegin bmLines
   glColor4f 1#, 0#, 0#, 1#
   glVertex3f 0#, 0#, 0#
   glVertex3f PONTA, 0#, 0#
     glVertex3f PONTA, 0#, 0#: glVertex3f INI_SETA, ABERTURA_SETA, 0#
     glVertex3f PONTA, 0#, 0#: glVertex3f INI_SETA, -ABERTURA_SETA, 0#
     glVertex3f PONTA, 0#, 0#: glVertex3f INI_SETA, 0#, ABERTURA_SETA
     glVertex3f PONTA, 0#, 0#: glVertex3f INI_SETA, 0#, -ABERTURA_SETA
     
   glColor4f 0#, 1#, 0#, 1#
   glVertex3f 0#, 0#, 0#
   glVertex3f 0#, PONTA, 0#
     glVertex3f 0#, PONTA, 0#: glVertex3f 0#, INI_SETA, ABERTURA_SETA
     glVertex3f 0#, PONTA, 0#: glVertex3f 0#, INI_SETA, -ABERTURA_SETA
     glVertex3f 0#, PONTA, 0#: glVertex3f ABERTURA_SETA, INI_SETA, 0#
     glVertex3f 0#, PONTA, 0#: glVertex3f -ABERTURA_SETA, INI_SETA, 0#
     
   glColor4f 0#, 0#, 1#, 1#
   glVertex3f 0#, 0#, 0#
   glVertex3f 0#, 0#, PONTA
     glVertex3f 0#, 0#, PONTA: glVertex3f 0#, ABERTURA_SETA, INI_SETA
     glVertex3f 0#, 0#, PONTA: glVertex3f 0#, -ABERTURA_SETA, INI_SETA
     glVertex3f 0#, 0#, PONTA: glVertex3f ABERTURA_SETA, 0#, INI_SETA
     glVertex3f 0#, 0#, PONTA: glVertex3f -ABERTURA_SETA, 0#, INI_SETA
   
   'glColor3f 0.8, 0.8, 0.8
   'glVertex4f 1, 0#, 0#, ZERO:   glVertex4f -1#, 0#, 0#, ZERO
   'glVertex4f 0#, 1#, 0#, ZERO:  glVertex4f 0#, -1#, 0#, ZERO
   'glVertex4f 0#, 0#, 1#, ZERO:  glVertex4f 0#, 0#, -1#, ZERO

 glEnd
 
 Erro = glGetError: If Erro <> glerrNoError Then ErroFatal Erro
End Sub
Public Sub Des_Ponto_Aux(Plano As Tipo_De_Plano, Aux() As GLdouble)
   Dim Pt(0 To 2) As GLdouble
   
   Pt(0) = Aux(0) / Aux(3)
   Pt(1) = Aux(1) / Aux(3)
   Pt(2) = Aux(2) / Aux(3)
   
   glColor3d 1#, 0.4, 0.1
   
   glPointSize (3#)
   glBegin bmPoints
      glVertex4dv Aux(0)
   glEnd
   
   glColor3d 0.7, 0.7, 0.7
   glLineWidth (1#)
   glBegin bmLines
      glVertex3dv Pt(0)
      Select Case Plano
      Case PL_HORIZONTAL 'Segmento vertical
         glVertex3d Pt(0), Pt(1), 0#
      Case PL_PERFIL 'Segmento fronto-horizontal
         glVertex3d 0#, Pt(1), Pt(2)
      Case PL_FRONTAL 'Segmento de topo
         glVertex3d Pt(0), 0#, Pt(2)
      End Select
   glEnd
   
   Erro = glGetError: If Erro <> glerrNoError Then ErroFatal Erro
End Sub

Public Sub Des_LT()
   Const Tam = 7
   Const DIST = 0.3
   glColor3d 0.5, 0, 0
   glLineWidth (1#)
   glBegin GL_LINES
      glColor3d 0.5, 0, 0
      glVertex3f -Tam, 0, 0
      glVertex3f Tam, 0, 0
      
      glVertex3f -Tam, DIST, 0
      glVertex3f 1 - Tam, DIST, 0
      glVertex3f Tam, DIST, 0
      glVertex3f Tam - 1, DIST, 0
   glEnd
   
   glPointSize 3#
   glBegin GL_POINTS
      glVertex3f 0, 0, 0
   glEnd
   
   Erro = glGetError: If Erro <> glerrNoError Then ErroFatal Erro
End Sub
Public Sub Des_Objetos(ByVal IdDoc As Integer, ByVal Modo As GLenum, Ob() As Objeto)
   Dim i As Long
   Dim N_Obj As Long
   
   'já ocorreu um glPushName 0, inicializando a pilha de nomes arbitrariamente
   
   glColor3d 0#, 0#, 0#
   glPointSize (3#)
   'CAUSOU UM PONTO NA ORIGEM
   N_Obj = UBound(Ob) '(Doc(IdDoc).Obj)
   For i = 1 To N_Obj
      With Ob(i) 'Doc(IdDoc).Obj(i)
         If i = Doc(IdDoc).frm.ObjApontado Then
            glColor3d 0.8, 0#, 0.5
            glPointSize 5#
            glLineWidth 2#
         ElseIf .Selec > 0 Then
            glColor3d 0.9, 0.4, 0#
            If .Tam > 0 Then glPointSize .Tam
            If .Tam > 0 Then glLineWidth .Tam
         Else
            glColor3dv .Cor(0)
            If .Tam > 0 Then glPointSize .Tam
            If .Tam > 0 Then glLineWidth .Tam
         End If

         Select Case .Tipo
         Case PONTO
            If Modo = GL_SELECT Then glLoadName i
            glBegin bmPoints
               glVertex4dv .Coord(0)
            glEnd
         Case SEGMENTO
            glBegin bmLines
               glVertex4dv Ob(.Id_Dep(1)).Coord(0)
               glVertex4dv Ob(.Id_Dep(2)).Coord(0)
            glEnd
         End Select
      End With
   Next i

   Erro = glGetError: If Erro <> glerrNoError Then ErroFatal Erro
End Sub
Public Sub Des_Objetos_Aux(ByVal IdDoc As Integer, Ob() As Objeto)
   Dim i As Long
   Dim N_Obj As Long
   
   glColor3d 0#, 0#, 0#
   glPointSize (3#)
   N_Obj = UBound(Ob)
   For i = 1 To N_Obj
      With Ob(i)
         glColor3d 0.9, 0.4, 0#
         glPointSize 3#
         glLineWidth 2#
         Select Case .Tipo
         Case PONTO
            With Doc(IdDoc).frm
               If .Posicionando Then
                  If .ObjApontado > 0 Then
                     Des_Ponto_Aux Sobre_Plano, Doc(IdDoc).Obj(Doc(IdDoc).frm.ObjApontado).Coord
                  Else
                     If i < N_Obj - 1 Then
                        glBegin bmPoints
                           glVertex4dv Ob(i).Coord(0)
                        glEnd
                     Else
                        Des_Ponto_Aux Sobre_Plano, Ob(i).Coord
                     End If
                  End If
               End If
            End With

         Case SEGMENTO
            glBegin bmLines
               'Em um segmento aux, o segundo vértice é aux mas o primeiro não
               glVertex4dv Doc(IdDoc).Obj(.Id_Dep(1)).Coord(0) 'Ob(.Id_Dep(1)).Coord(0)
               glVertex4dv Ob(.Id_Dep(2)).Coord(0)
            glEnd
         End Select
      End With
   Next i
   Erro = glGetError: If Erro <> glerrNoError Then ErroFatal Erro
End Sub
