VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exemplo - Integrando Vb e OpenGl"
   ClientHeight    =   3150
   ClientLeft      =   585
   ClientTop       =   1170
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   210
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   413
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picViewTela 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   75
      ScaleHeight     =   198
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   298
      TabIndex        =   0
      Top             =   75
      Width           =   4500
   End
   Begin VB.Label lblTexto 
      AutoSize        =   -1  'True
      Caption         =   "Id do Objeto:"
      Height          =   195
      Left            =   4725
      TabIndex        =   2
      Top             =   585
      Width           =   915
   End
   Begin VB.Label lblHits 
      AutoSize        =   -1  'True
      Caption         =   "Hits:"
      Height          =   195
      Left            =   4725
      TabIndex        =   1
      Top             =   135
      Width           =   315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Nome As GLuint

Private Sub Form_Load()
 hDC1 = Me.picViewTela.hDC 'Identificador da ViewPort1 (embora não use + de uma viewport)
 Call Inicializar_OpenGL(hDC1) 'Ajusta formato dos pixels, iluminação, matrizes de projeção...
 Nome = -1
End Sub
Private Sub Des_Quad()
 glBegin bmQuads
  glVertex3f 0, 0, 0
  glVertex3f 1, 0, 0
  glVertex3f 1, 1, 0
  glVertex3f 0, 1, 0
 glEnd
End Sub
Private Sub Des_Eixos()
 glBegin bmLines
   glColor3f 1#, 0#, 0#
   glVertex3f 0#, 0#, 0#
   glVertex3f 5#, 0#, 0#
   
   glColor3f 0#, 1#, 0#
   glVertex3f 0#, 0#, 0#
   glVertex3f 0#, 5#, 0#
   
   glColor3f 0#, 0#, 1#
   glVertex3f 0#, 0#, 0#
   glVertex3f 0#, 0#, 5#
 glEnd
End Sub
Private Sub Desenha_Todos(Mode As GLenum)
 Dim X, Y As Long
 Dim Cor(1 To 3, 0 To 1) As GLfloat 'TRÊS coordenadas para cada uma das DUAS cores
  
 Cor(1, 0) = 1#: Cor(2, 0) = 0.5: Cor(3, 0) = 0#
 Cor(1, 1) = 0.5: Cor(2, 1) = 0.2: Cor(3, 1) = 0.5
 
 glMatrixMode GL_MODELVIEW
 glLoadIdentity
 gluLookAt 4, 3.5, 3, 0, 0, 0, 0, 0, 1
 
 If Mode = GL_SELECT Then glLoadName 0
 glPushAttrib (amAllAttribBits)
  If Nome = 0 Then glLineWidth (3)
  Des_Eixos
 glPopAttrib
 
 glPushMatrix
 For X = 1 To 3 'para cada X
  glPushMatrix
  If Mode = GL_SELECT Then glLoadName X
  For Y = 1 To 3 'para cada Y
   If Mode = GL_SELECT Then glPushName Y
   glColor3fv Cor(1, (X + Y) Mod 2)
   glPushAttrib (amAllAttribBits)
    If Nome = X + 3 * (Y - 1) Then glColor3f 1, 1, 0.3
    Des_Quad
   glPopAttrib
   glTranslatef 0, 1, 0
   If Mode = GL_SELECT Then glPopName
  Next Y
  glPopMatrix
  glTranslatef 1, 0, 0
 Next X
 glPopMatrix
End Sub

Private Sub picViewTela_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Const BUFSIZE = 512
 Dim h As Long, Id As Long
 Dim Qtd_Nomes As GLuint, MinZ As Double
 Dim SelectBuf(0 To BUFSIZE - 1) As GLuint
 Dim Hits As GLint
 'Dim ViewPort(0 To 3) As GLint 'Declaração foi feita globalmente
 
 If Button <> 1 Then Exit Sub
 glGetIntegerv GL_VIEWPORT, ViewPort(0)
 glSelectBuffer BUFSIZE, SelectBuf(0)
 glRenderMode GL_SELECT
 glInitNames
 glPushName 0
 
 glMatrixMode GL_PROJECTION
 glPushMatrix
 glLoadIdentity
 gluPickMatrix X, ViewPort(3) - Y, 5#, 5#, ViewPort(0)
 gluPerspective 70, fAspect, 1#, 50#
 
 Desenha_Todos GL_SELECT
 
 glMatrixMode GL_PROJECTION
 glPopMatrix
 glFlush
 Hits = glRenderMode(GL_RENDER)
 Me.lblHits = "Hits: " & Hits
 
 'processHits Hits, SelectBuf(0)'Procedimento descrito abaixo...
 Id = 0
 MinZ = 2121212121 'inicializa minZ para um valor grande
 Nome = -1 'Nada selecionado até agora
 'Para compreender o laço "FOR NEXT", lembre-se do formato de cada REGISTRO (HIT)...
  ' Reg1: |    SelectBuf(0)      | SelectBuf(1)  | SelectBuf(2) |  SelectBuf( 3... 3+Qtd_Nomes)   |
  '       | Qtd de Nomes em Reg1 |   Z mínimo    |   Z máximo   | Nomes deste Registro (de 0 a n) |
  '  ...próximo registro é similar!
  ' Reg2: |SelectBuf(0 + 3+Qtd_Nomes)| e assim vai...
  
 For h = 1 To Hits
  Qtd_Nomes = SelectBuf(Id) 'o nome é composto de 'tantas' coordenadas
  If (SelectBuf(Id + 1) < MinZ) And (Qtd_Nomes > 0) Then
   MinZ = SelectBuf(Id + 1)
   Nome = SelectBuf(Id + 3)
   If Qtd_Nomes = 2 Then Nome = Nome + 3 * (-1 + SelectBuf(Id + 4))
  End If
  Id = Id + 3 + Qtd_Nomes
 Next h
 Me.lblTexto = "Id do objeto: " & Nome
 If Nome > 0 Then Me.lblTexto = Me.lblTexto & vbCrLf & "[ " & (Nome - 1) Mod 3 + 1 & ", " & (Nome \ 3) + 1 & "]"
 picViewTela_Paint 'glutPostRedisplay 'GLUT CAUSA ERRO NO VB
 
End Sub

Private Sub picViewTela_Paint()
 glClear clrColorBufferBit Or clrDepthBufferBit
 
 Desenha_Todos GL_RENDER 'GL_SELECT
 
 SwapBuffers hDC1
End Sub
Private Sub Form_Unload(Cancel As Integer)
Call Finalizar_OpenGL
End Sub
