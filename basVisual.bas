Attribute VB_Name = "basVisual"
Option Explicit
  Public ModelViewMatrix(0 To 15) As GLfloat
  Public ProjectionMatrix(0 To 15) As GLfloat
  Public Viewport(0 To 3) As GLint

Public LightPos(3) As GLfloat
Public SpecRef(3) As GLfloat
Public Diffuse(3) As GLfloat
Public lmodel_ambient(3) As GLfloat

Public m_Grid As Integer
Public m_Translate_X As Integer
Public m_Translate_Y As Integer
Public m_Translate_Z As Integer
Public m_camera_radsFromEast As GLfloat
Public m_translationUnit As Double
Public m_camera_direction_y As Integer

Public Sub Inicializar_Luz()
 'm_Translate_X = 10
 'm_Translate_Z = 10
 'm_translationUnit = 1
 'm_camera_direction_y = 0
 'm_camera_radsFromEast = 1.56
 
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
  glFrontFace GL_CCW
  
  basVisual.lmodel_ambient(0) = 0.5
  basVisual.lmodel_ambient(1) = 0.5
  basVisual.lmodel_ambient(2) = 0.5
  basVisual.lmodel_ambient(3) = 1#
  
  glLightModelfv GL_LIGHT_MODEL_Ambient, basVisual.lmodel_ambient(0)
  
  'glMaterialfv GL_FRONT, GL_SPECULAR, SpecRef(0)
  'glMateriali GL_FRONT, GL_SHININESS, 50
 
End Sub
Public Sub Inicializar_OpenGL(ByVal hDC As Long)
 Dim hGLRC As Long
 Dim fAspect As GLfloat
 
 'Ajusta o contexto OpenGl para operar com o form do Visual Basic
 basVbOpenGl.SetupPixelFormat hDC
 hGLRC = wglCreateContext(hDC)
 wglMakeCurrent hDC, hGLRC
 
 'Começa a configurar a biblioteca OpenGl
 'glEnable GL_DEPTH_TEST
 'glEnable GL_DITHER
 'glDepthFunc GL_LESS
 'glClearDepth 1
 glEnable glcColorMaterial
 glClearColor 0.96, 0.96, 1#, 0  'Fundo quase branco
 'glClearColor 0, 0, 0, 0 'Fundo preto
 'glClearColor 0.1, 0.1, 0.1, 0
 
 'Ajusta matriz de projeção e viewport
 glMatrixMode GL_PROJECTION
   glLoadIdentity
   If frmMain.ScaleHeight > 0 Then
    fAspect = frmMain.ScaleWidth / frmMain.ScaleHeight
   Else
    fAspect = 0
   End If
   
   'gluPerspective 60, fAspect, 1, 2000
   'gluOrtho2D -5, 5, -5, 5
   
' Call Ajusta_ViewPort(0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight)
 
 Call basVisual.Inicializar_Luz

End Sub
Sub Ajusta_ViewPort(X_esq As GLint, Y_inf As GLint, Larg As GLsizei, Alt As GLsizei)
 glMatrixMode GL_PROJECTION
   glLoadIdentity
   glViewport X_esq, Y_inf, Larg, Alt
   'gluOrtho2D -5, 5, -5, 5
   gluOrtho2D Centro_X - Visivel_X / 2, Centro_X + Visivel_X / 2, Centro_Y - Visivel_Y / 2, Centro_Y + Visivel_Y / 2
 glMatrixMode GL_MODELVIEW
 frmMatriz.AtualizaMatrizes
End Sub
Public Sub normalize(out() As GLfloat)
Dim D As GLfloat
    D = Sqr(out(0) * out(0) + out(1) * out(1) + out(2) * out(2))
    If (D = 0) Then
        Exit Sub
    End If
    out(0) = out(0) / D
    out(1) = out(1) / D
    out(2) = out(2) / D
End Sub

Public Sub normcrossprod(v() As GLfloat, w() As GLfloat, out() As GLfloat)
 '[Vx Vy Vz] X [Wx Wy Wz] =[(Vy*Wz-Wy*Vz),(Wx*Vz-Vx*Wz),(Vx*Wy-Wx*Vy)]
 out(0) = v(1) * w(2) - w(1) * v(2)
 out(1) = w(0) * v(2) - v(0) * w(2)
 out(2) = v(0) * w(1) - w(0) * v(1)
 Call normalize(out)
 
End Sub
