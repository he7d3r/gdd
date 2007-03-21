VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exemplo - Integrando Vb e OpenGl"
   ClientHeight    =   6915
   ClientLeft      =   585
   ClientTop       =   1170
   ClientWidth     =   10395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   461
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   693
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Slider sldDistancia 
      Height          =   330
      Left            =   6840
      TabIndex        =   6
      Top             =   2655
      Width           =   3030
      _ExtentX        =   5345
      _ExtentY        =   582
      _Version        =   393216
      MousePointer    =   9
      SelectRange     =   -1  'True
      TickStyle       =   3
      TextPosition    =   1
   End
   Begin VB.Frame fraVistas 
      Caption         =   "Visualizar"
      Height          =   1590
      Left            =   6930
      TabIndex        =   2
      Top             =   375
      Width           =   1545
      Begin VB.OptionButton optVista 
         Caption         =   "3ª Projeção"
         Height          =   375
         Index           =   3
         Left            =   135
         TabIndex        =   5
         Top             =   1125
         Width           =   1250
      End
      Begin VB.OptionButton optVista 
         Caption         =   "2ª Projeção"
         Height          =   375
         Index           =   2
         Left            =   135
         TabIndex        =   4
         Top             =   675
         Width           =   1250
      End
      Begin VB.OptionButton optVista 
         Caption         =   "1ª Projeção"
         Height          =   375
         Index           =   1
         Left            =   135
         TabIndex        =   3
         Top             =   225
         Value           =   -1  'True
         Width           =   1250
      End
   End
   Begin VB.PictureBox picPerspectiva 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   6030
      Left            =   450
      MouseIcon       =   "Main.frx":0000
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      ToolTipText     =   "Botão direito: Mover camera."
      Top             =   450
      Width           =   6030
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Distante"
      Height          =   195
      Left            =   9180
      TabIndex        =   9
      Top             =   3105
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Próximo"
      Height          =   195
      Left            =   6930
      TabIndex        =   8
      Top             =   3105
      Width           =   555
   End
   Begin VB.Label lblDistancia 
      AutoSize        =   -1  'True
      Caption         =   "Posição do observador:"
      Height          =   195
      Left            =   6930
      TabIndex        =   7
      Top             =   2340
      Width           =   1680
   End
   Begin VB.Label lblPerspectiva 
      Caption         =   "Perspectiva:"
      Height          =   195
      Left            =   465
      TabIndex        =   1
      ToolTipText     =   "Teclas [ + ] e [ - ] alteram a distância da câmera."
      Top             =   195
      Width           =   6015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private X_Ini As Integer, Y_Ini As Integer       'Usado no movimento da camera
Private Phi_Ini As GLfloat, Theta_Ini As GLfloat 'Idem
Private Proj As GLuint 'Número da projeção que será realizada (1, 2 ou 3)

Private Sub Form_Load()
 hDCPerspectiva = Me.picPerspectiva.hDC 'Identificador das ViewPort's
 'hDCObservador = Me.picObservador.hDC
 Call Inicializar_OpenGL 'Ajusta formato dos pixels, iluminação, matrizes de projeção...
 
 sldDistancia.Max = DIST_MAX
 sldDistancia.Min = DIST_MIN
 optVista(2).Value = True

End Sub
Private Sub Form_Unload(Cancel As Integer)
Call Finalizar_OpenGL
End Sub

Private Sub sldDistancia_Change()
 picPerspectiva_Paint
End Sub

Private Sub sldDistancia_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = 0 Then Exit Sub
 
 picPerspectiva_Paint
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

 If Chr(KeyAscii) = "+" Then Ro = Ro - 1
 If Chr(KeyAscii) = "-" Then Ro = Ro + 1
 If Ro < DIST_MIN Then Ro = DIST_MIN
 If Ro > DIST_MAX Then Ro = DIST_MAX
  Cam_X = Ro * Sin(Phi * DEG) * Cos(Theta * DEG)
  Cam_Y = Ro * Sin(Phi * DEG) * Sin(Theta * DEG)
  Cam_Z = Ro * Cos(Phi * DEG)
  
  wglMakeCurrent hDCPerspectiva, hGLRCPerspectiva
  glMatrixMode GL_MODELVIEW
  glLoadIdentity
  gluLookAt Cam_X, Cam_Y, Cam_Z, Centro_X, Centro_Y, Centro_Z, 0, 0, 1
  glMultMatrixf Troca_X_Y(0)
  
  picPerspectiva_Paint
End Sub
Private Sub Des_Planos()
 Const INI_PLANO = -1 '-2 * LADO_PRISMA - 1
 Const FIM_PLANO = 2 * LADO_PRISMA + 1
 
 Dim k As GLdouble
 
 glColor3f 0.7, 0.7, 0.7
 glLineWidth (1#)
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
 
End Sub
Private Sub Des_Eixos(tam As GLfloat)
 
 glLineWidth (2#)
 glBegin bmLines
   glColor3f 1#, 0#, 0#
   glVertex3f 0#, 0#, 0#
   glVertex3f tam, 0#, 0#
   
   glColor3f 0#, 1#, 0#
   glVertex3f 0#, 0#, 0#
   glVertex3f 0#, tam, 0#
   
   glColor3f 0#, 0#, 1#
   glVertex3f 0#, 0#, 0#
   glVertex3f 0#, 0#, tam
 glEnd
End Sub

Private Sub Des_Observador(tam As GLfloat)
Dim r_cabeca As GLfloat
Dim r_membros As GLfloat
Dim tam_membros As GLfloat
 r_cabeca = tam / 7
 r_membros = tam / 28
 tam_membros = 5 / 21 * Sqr(3) * tam
 
 glEnable GL_LIGHTING
 glPushMatrix
  'glTranslatef x, y, tam
  glColor3f 1#, 0.7, 0.5
  gluSphere QObj, r_cabeca, 10, 10
  glTranslatef 0, 0, -9 * tam / 14
  gluCylinder QObj, r_membros, r_membros, 9 * tam / 14, 12, 2
  glRotatef 150, 0, 1, 0
  gluCylinder QObj, r_membros, r_membros, tam_membros, 12, 2
  glRotatef 60, 0, 1, 0
  gluCylinder QObj, r_membros, r_membros, tam_membros, 12, 2
  glRotatef 150, 0, 1, 0
  glTranslatef 0, 0, 3 * tam / 7
  glRotatef 150, 0, 1, 0
  gluCylinder QObj, r_membros, r_membros, tam_membros, 12, 2
  glRotatef 60, 0, 1, 0
  gluCylinder QObj, r_membros, r_membros, tam_membros, 12, 2
  'glColor3d 0, 0, 0
 glPopMatrix
 glDisable GL_LIGHTING

End Sub
Private Sub Des_Prisma(L As GLfloat)
Exit Sub
glEnable GL_LIGHTING
 glColor3f 0.2, 0.4, 0.4
 glBegin bmTriangles
  glNormal3f 1 / Sqr(2), -1 / Sqr(2), 0
  glVertex3f L, 0, 0
  glVertex3f 2 * L, L, 0
  glVertex3f 3 * L / 2, L / 2, L * Sqr(6) / 2
  
  glNormal3f -1 / Sqr(2), 1 / Sqr(2), 0
  glVertex3f 0, L, 0
  glVertex3f L, 2 * L, 0
  glVertex3f L / 2, 3 * L / 2, L * Sqr(6) / 2
 glEnd

 glBegin bmQuads
  glNormal3f 0, 0, -1
  glVertex3f 0, L, 0
  glVertex3f L, 0, 0
  glVertex3f 2 * L, L, 0
  glVertex3f L, 2 * L, 0
  
  glNormal3f Sqr(3 / 14), Sqr(3 / 14), 2 / Sqr(7)
  glVertex3f 2 * L, L, 0
  glVertex3f L, 2 * L, 0
  glVertex3f L / 2, 3 * L / 2, L * Sqr(6) / 2
  glVertex3f 3 * L / 2, L / 2, L * Sqr(6) / 2
  
  glNormal3f -Sqr(3 / 14), -Sqr(3 / 14), 2 / Sqr(7)
  glVertex3f L / 2, 3 * L / 2, L * Sqr(6) / 2
  glVertex3f 3 * L / 2, L / 2, L * Sqr(6) / 2
  glVertex3f L, 0, 0
  glVertex3f 0, L, 0
 glEnd
 glDisable GL_LIGHTING
End Sub

Private Sub Des_Cena(Vista As GLuint)
 Dim i As Long
 
 glLineWidth (1#)
 'glPointSize (3#)
 'glBegin bmPoints
 'glEnd
 Select Case Vista
 Case 1
  Ox = LADO_PRISMA: Oy = LADO_PRISMA: Oz = sldDistancia.Value
  glPushMatrix
   glTranslatef Ox, Oy, Oz
   glRotatef 90, 1, 0, 0
   Des_Observador LADO_PRISMA
  glPopMatrix
  
  glColor3d 0.5, 0.5, 0.7
  glBegin bmLineStrip
   glVertex3f Ox, 0, 0  'P=(Ox, 0, 0)
   glVertex3f Ox, Oy, Oz
   glVertex3f 0, Ox, 0   'P=(0, Ox, 0)
  glEnd
  glBegin bmLineStrip
   glVertex3f 2 * Ox, Ox, 0  'P=(2 * Ox, Ox, 0)
   glVertex3f Ox, Oy, Oz
   glVertex3f Ox, 2 * Ox, 0 'P=(Ox, 2 * Ox, 0)
  glEnd
  glBegin bmLineStrip
   glVertex3f Ox * (Sqr(6) * Ox - 3 * Oz) / (Ox * Sqr(6) - 2 * Oz), Ox * (Sqr(6) * Oy - Oz) / (Ox * Sqr(6) - 2 * Oz), 0  '3 * Ox / 2, Ox / 2, Ox * Sqr(6) / 2
   glVertex3f Ox, Oy, Oz
   glVertex3f Ox * (Sqr(6) * Ox - Oz) / (Ox * Sqr(6) - 2 * Oz), Ox * (Sqr(6) * Oy - 3 * Oz) / (Ox * Sqr(6) - 2 * Oz), 0 'Ox / 2, 3 * Ox / 2, Ox * Sqr(6) / 2
  glEnd
  
  glColor3d 0, 0, 0.2
  glBegin bmPolygon
   glVertex3f Ox, 0, -0.01  'P=(Ox, 0, 0)
   glVertex3f 0, Ox, -0.01   'P=(0, Ox, 0)
   glVertex3f Ox * (Sqr(6) * Ox - Oz) / (Ox * Sqr(6) - 2 * Oz), Ox * (Sqr(6) * Oy - 3 * Oz) / (Ox * Sqr(6) - 2 * Oz), -0.01 'Ox / 2, 3 * Ox / 2, Ox * Sqr(6) / 2
   glVertex3f Ox, 2 * Ox, -0.01 'P=(Ox, 2 * Ox, 0)
   glVertex3f 2 * Ox, Ox, -0.01  'P=(2 * Ox, Ox, 0)
   glVertex3f Ox * (Sqr(6) * Ox - 3 * Oz) / (Ox * Sqr(6) - 2 * Oz), Ox * (Sqr(6) * Oy - Oz) / (Ox * Sqr(6) - 2 * Oz), -0.01  '3 * Ox / 2, Ox / 2, Ox * Sqr(6) / 2
  glEnd
 Case 2
  Ox = LADO_PRISMA: Oy = sldDistancia.Value: Oz = LADO_PRISMA
  glPushMatrix
   glTranslatef Ox, Oy, Oz
   Des_Observador LADO_PRISMA
  glPopMatrix
  
  glColor3d 0.5, 0.7, 0.5
  glBegin bmLineStrip
  glVertex3f Oz, 0, 0
  glVertex3f Ox, Oy, Oz
  glVertex3f -Ox * Oz / (Oy - Oz), 0, Oz * Oz / (Oz - Oy) '0, Oz, 0
  glEnd
  glBegin bmLineStrip
  glVertex3f (2 * Oz * Oy - Ox * Oz) / (Oy - Oz), 0, Oz * Oz / (Oz - Oy) '2 * Oz, Oz, 0
  glVertex3f Ox, Oy, Oz
  glVertex3f (Oz * Oy - Ox * 2 * Oz) / (Oy - 2 * Oz), 0, 2 * Oz * Oz / (2 * Oz - Oy) 'Oz, 2 * Oz, 0
  glEnd
  glBegin bmLineStrip
  glVertex3f Oz * (3 * Oy - Ox) / (2 * Oy - Oz), 0, Oz * (Oy * Sqr(6) - Oz) / (2 * Oy - Oz) '3 * Oz / 2, Oz / 2, Oz * Sqr(6) / 2
  glVertex3f Ox, Oy, Oz
  glVertex3f Oz * (Oy - 3 * Ox) / (2 * Oy - 3 * Oz), 0, Oz * (Sqr(6) * Oy - 3 * Oz) / (2 * Oy - 3 * Oz) 'Oz / 2, 3 * Oz / 2, Oz * Sqr(6) / 2
  glEnd
  
  glColor3d 0#, 0.2, 0#
  glBegin bmPolygon
  'glVertex3f Oz, 0, 0
  glVertex3f -Ox * Oz / (Oy - Oz), 0, Oz * Oz / (Oz - Oy) '0, Oz, 0
  glVertex3f (Oz * Oy - Ox * 2 * Oz) / (Oy - 2 * Oz), 0, 2 * Oz * Oz / (2 * Oz - Oy) 'Oz, 2 * Oz, 0
  glVertex3f (2 * Oz * Oy - Ox * Oz) / (Oy - Oz), 0, Oz * Oz / (Oz - Oy) '2 * Oz, Oz, 0
  glVertex3f Oz * (3 * Oy - Ox) / (2 * Oy - Oz), 0, Oz * (Oy * Sqr(6) - Oz) / (2 * Oy - Oz) '3 * Oz / 2, Oz / 2, Oz * Sqr(6) / 2
  glVertex3f Oz * (Oy - 3 * Ox) / (2 * Oy - 3 * Oz), 0, Oz * (Sqr(6) * Oy - 3 * Oz) / (2 * Oy - 3 * Oz) 'Oz / 2, 3 * Oz / 2, Oz * Sqr(6) / 2
  glEnd
 Case 3
  Ox = sldDistancia.Value: Oy = LADO_PRISMA: Oz = LADO_PRISMA
  glPushMatrix
   glTranslatef Ox, Oy, Oz
   glRotatef 90, 0, 1, 0
   glRotatef 90, 1, 0, 0
   Des_Observador LADO_PRISMA
  glPopMatrix
  
  glColor3d 0.7, 0.5, 0.5
  glBegin bmLineStrip
  glVertex3f 0, Oz, 0
  glVertex3f Ox, Oy, Oz
  glVertex3f 0, -Oy * Oz / (Ox - Oz), Oz * Oz / (Oz - Ox) '0, Oz, 0
  glEnd
  glBegin bmLineStrip
  glVertex3f 0, (2 * Oz * Ox - Oy * Oz) / (Ox - Oz), Oz * Oz / (Oz - Ox) '2 * Oz, Oz, 0
  glVertex3f Ox, Oy, Oz
  glVertex3f 0, (Oz * Ox - Oy * 2 * Oz) / (Ox - 2 * Oz), 2 * Oz * Oz / (2 * Oz - Ox) 'Oz, 2 * Oz, 0
  glEnd
  glBegin bmLineStrip
  glVertex3f 0, Oz * (3 * Ox - Oy) / (2 * Ox - Oz), Oz * (Ox * Sqr(6) - Oz) / (2 * Ox - Oz) '3 * Oz / 2, Oz / 2, Oz * Sqr(6) / 2
  glVertex3f Ox, Oy, Oz
  glVertex3f 0, Oz * (Ox - 3 * Oy) / (2 * Ox - 3 * Oz), Oz * (Sqr(6) * Ox - 3 * Oz) / (2 * Ox - 3 * Oz) 'Oz / 2, 3 * Oz / 2, Oz * Sqr(6) / 2
  glEnd
  
  glMatrixMode GL_MODELVIEW
  glPushMatrix
  'glLoadIdentity
  glMultMatrixf M_proj(0)
  'glMultMatrixf Troca_X_Y(0)
  glBegin bmTriangles
    glColor3d 0, 0, 0.3
   glVertex3f LADO_PRISMA, 0, 0
   glVertex3f 2 * LADO_PRISMA, LADO_PRISMA, 0
   glVertex3f 3 * LADO_PRISMA / 2, LADO_PRISMA / 2, LADO_PRISMA * Sqr(6) / 2
    glColor3d 0, 0.3, 0
   glVertex3f 0, LADO_PRISMA, 0
   glVertex3f LADO_PRISMA, 2 * LADO_PRISMA, 0
   glVertex3f LADO_PRISMA / 2, 3 * LADO_PRISMA / 2, LADO_PRISMA * Sqr(6) / 2
  glEnd
   
  glPopMatrix
  
  
  
  'glBegin bmPolygon
   'glVertex3f 0, Oz, 0
  ' glVertex3f 0, -Oy * Oz / (Ox - Oz), Oz * Oz / (Oz - Ox) '0, Oz, 0
  ' glVertex3f 0, (Oz * Ox - Oy * 2 * Oz) / (Ox - 2 * Oz), 2 * Oz * Oz / (2 * Oz - Ox) 'Oz, 2 * Oz, 0
  ' glVertex3f 0, (2 * Oz * Ox - Oy * Oz) / (Ox - Oz), Oz * Oz / (Oz - Ox) '2 * Oz, Oz, 0
  ' glVertex3f 0, Oz * (3 * Ox - Oy) / (2 * Ox - Oz), Oz * (Ox * Sqr(6) - Oz) / (2 * Ox - Oz) '3 * Oz / 2, Oz / 2, Oz * Sqr(6) / 2
  ' glVertex3f 0, Oz * (Ox - 3 * Oy) / (2 * Ox - 3 * Oz), Oz * (Sqr(6) * Ox - 3 * Oz) / (2 * Ox - 3 * Oz) 'Oz / 2, 3 * Oz / 2, Oz * Sqr(6) / 2
  'glEnd
 End Select
End Sub

Private Sub optVista_Click(Index As Integer)
Proj = Index
picPerspectiva_Paint
'picObservador_Paint
End Sub

Private Sub picPerspectiva_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Const VELOCIDADE = 0.5
 Dim dx As Integer, dy As Integer
 
 dx = VELOCIDADE * (x - X_Ini)
 dy = VELOCIDADE * (y - Y_Ini)
 Select Case Button
 Case 1

 Case 2 'botao direito = Mover camera
  Phi = Phi_Ini - dy
  Theta = Theta_Ini - dx
  Phi = IIf(Phi <= 0, 0.0001, Phi): Phi = IIf(Phi > 180, 180, Phi)
  Theta = IIf(Theta <= -180, Theta + 360, Theta): Theta = IIf(Theta > 180, Theta - 360, Theta)
  
  Cam_X = Ro * Sin(Phi * DEG) * Cos(Theta * DEG)
  Cam_Y = Ro * Sin(Phi * DEG) * Sin(Theta * DEG)
  Cam_Z = Ro * Cos(Phi * DEG)
  
  wglMakeCurrent hDCPerspectiva, hGLRCPerspectiva
  glMatrixMode GL_MODELVIEW
  glLoadIdentity
  gluLookAt Cam_X, Cam_Y, Cam_Z, Centro_X, Centro_Y, Centro_Z, 0, 0, 1
  glMultMatrixf Troca_X_Y(0)
  
  picPerspectiva_Paint
 End Select
End Sub
Private Sub picPerspectiva_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 X_Ini = x: Y_Ini = y
 Phi_Ini = Phi:  Theta_Ini = Theta
 
 If Button = 2 Then picPerspectiva.MousePointer = 99
End Sub
Private Sub picPerspectiva_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button = 2 Then picPerspectiva.MousePointer = 0
End Sub
Private Sub picPerspectiva_Paint()

 wglMakeCurrent hDCPerspectiva, hGLRCPerspectiva
 glClear clrColorBufferBit Or clrDepthBufferBit
 
 Des_Planos
 Des_Eixos (2 * LADO_PRISMA + 2)
 Des_Prisma (LADO_PRISMA)
 Des_Cena Proj
 
 SwapBuffers hDCPerspectiva
End Sub
'Private Sub picObservador_Paint()

 'wglMakeCurrent hDCObservador, hGLRCObservador
 'glClear clrColorBufferBit Or clrDepthBufferBit
 
 'Des_Planos
 'Des_Eixos (3#)
 'Des_Cena Proj
 'Des_Prisma (LADO_PRISMA)
 
 'SwapBuffers hDCObservador
'End Sub


