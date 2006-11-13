Attribute VB_Name = "basVisualização"
Option Explicit
Public ModelViewMatrix(0 To 15) As GLfloat
Public ProjectionMatrix(0 To 15) As GLfloat
Public ViewPort(0 To 3) As GLint
Public MostrarMatrizes As Boolean

Public LightPos(3) As GLfloat
Public SpecRef(3) As GLfloat
Public Diffuse(3) As GLfloat
Public lmodel_ambient(3) As GLfloat

Public Centro_X As Single, Centro_Y As Single
Public Larg As GLsizei, Alt As GLsizei
Public fAspect As GLfloat

Public Sub Inicializar_Luz()
 
 LightPos(0) = 0
 LightPos(1) = 0
 LightPos(2) = 100
 LightPos(3) = 1
 
 SpecRef(0) = 0#
 SpecRef(1) = 0#
 SpecRef(2) = 0#
 SpecRef(3) = 1#
 
 glMatrixMode GL_MODELVIEW
  glLoadIdentity
  glEnable GL_LIGHTING
  glEnable GL_LIGHT0
  'glShadeModel GL_SMOOTH
  'glFrontFace GL_CCW
  
  basVisualização.lmodel_ambient(0) = 0.5
  basVisualização.lmodel_ambient(1) = 0.5
  basVisualização.lmodel_ambient(2) = 0.5
  basVisualização.lmodel_ambient(3) = 1#
  
  glLightModelfv GL_LIGHT_MODEL_Ambient, basVisualização.lmodel_ambient(0)
  
  'glMaterialfv GL_FRONT, GL_SPECULAR, SpecRef(0)
  'glMateriali GL_FRONT, GL_SHININESS, 50
 
End Sub
Public Sub Inicializar_OpenGL(ByVal hDC As Long)
 Dim hGLRC As Long
  
 'Ajusta um contexto OpenGl para operar com o objeto com o hDC passado
 basVbOpenGl.SetupPixelFormat hDC
 hGLRC = wglCreateContext(hDC)
 wglMakeCurrent hDC, hGLRC
 
 Larg = frmMain.picViewTela.ScaleWidth
 Alt = frmMain.picViewTela.ScaleHeight
 If Alt > 0 Then
  fAspect = Larg / Alt
 Else
  fAspect = 0
 End If
 
 'Começa a configurar a biblioteca OpenGl
 glEnable GL_DEPTH_TEST
 'glClearDepth 1#  'padrão=1
 glEnable glcColorMaterial
 glClearColor 0.8, 0.8, 1#, 0
      
 Call basVisualização.Ajusta_ViewPort(0, 0, Larg, Alt)
 'Call basVisualização.Inicializar_Luz

End Sub
Sub Ajusta_ViewPort(X_esq As GLint, Y_inf As GLint, Larg As GLsizei, Alt As GLsizei)
 'Ajusta viewport e matriz de projeção
 glViewport X_esq, Y_inf, Larg, Alt
 glMatrixMode GL_PROJECTION
   glLoadIdentity
   If Alt > 0 Then
    fAspect = Larg / Alt
   Else
    fAspect = 0
   End If
   gluPerspective 70, fAspect, 1#, 50#
   'gluOrtho2D -5, 5, -5, 5
   'gluOrtho2D Centro_X - Visivel_X / 2, Centro_X + Visivel_X / 2, _
   '          Centro_Y - Visivel_Y / 2, Centro_Y + Visivel_Y / 2
   
 glMatrixMode GL_MODELVIEW
End Sub
