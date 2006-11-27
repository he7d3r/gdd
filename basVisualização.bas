Attribute VB_Name = "basVisualização"
Option Explicit
Public Troca_X_Y(0 To 15) As GLfloat
Public ModelViewMatrix(0 To 15) As GLfloat
Public ProjectionMatrix(0 To 15) As GLfloat
Public ViewPort(0 To 3) As GLint

Public Const PI = 3.14159265
Public Const DEG = PI / 180
Public Cam_X As Single, Cam_Y As Single, Cam_Z As Single
Public Phi As GLfloat, Theta As GLfloat, Ro As GLfloat
Public Larg As GLsizei, Alt As GLsizei
Public fAspect As GLfloat

Public Sub Inicializar_OpenGL()
Dim Pfd As PIXELFORMATDESCRIPTOR
Dim Result As Long
 'Ajusta um contexto OpenGl para operar com o objeto com o hDC passado
 Result = basVbOpenGl.SetupPixelFormat(hDCPerspectiva)
 'Define o formato dos pixels conforme descrito em 'pfd'
 SetPixelFormat hDCPerspectiva, Result, Pfd
 hGLRCPerspectiva = wglCreateContext(hDCPerspectiva)
 
 SetPixelFormat hDCFrontal, Result, Pfd
 hGLRCFrontal = wglCreateContext(hDCFrontal)
 
 SetPixelFormat hDCLateral, Result, Pfd
 hGLRCLateral = wglCreateContext(hDCLateral)
 
 SetPixelFormat hDCSuperior, Result, Pfd
 hGLRCSuperior = wglCreateContext(hDCSuperior)
 
 SetPixelFormat hDCEpura, Result, Pfd
 hGLRCEpura = wglCreateContext(hDCEpura)
 
 'inicialize AQUI algumas matrizes de iluminação, e outras...
 QObj = gluNewQuadric()
 
 Phi = 60: Theta = 45: Ro = 7
 Cam_X = Ro * Sin(Phi * DEG) * Cos(Theta * DEG)
 Cam_Y = Ro * Sin(Phi * DEG) * Sin(Theta * DEG)
 Cam_Z = Ro * Cos(Phi * DEG)
  
 'Configurações específicas da PERSPECTIVA
 wglMakeCurrent hDCPerspectiva, hGLRCPerspectiva
  glClearColor 0.9, 1#, 1#, 1#       ' clear view with black color
  glShadeModel (GL_SMOOTH) 'Interpolate colors
  'glEnable GL_CULL_FACE    'Do not calculate BackFace of polys
  'glFrontFace GL_CCW
  glEnable GL_DEPTH_TEST
  'glClearDepth 1
  'glDepthFunc cfLEqual
  'glBlendFunc GL_SRC_ALPHA, GL_DST_ALPHA
  glPolygonMode GL_FRONT_AND_BACK, GL_LINE
  'glEnable GL_CULL_FACE
  glEnable GL_COLOR_MATERIAL
  'glEnable GL_LIGHTING
  'glLightfv GL_LIGHT0, GL_AMBIENT, Amb_Dif_Light(0)
  'glLightfv GL_LIGHT0, GL_DIFFUSE, Amb_Dif_Light(0)
  'glLightfv GL_LIGHT0, GL_POSITION, Light0Pos(0)
  'glEnable GL_LIGHT0
  
  With frmMain.picPerspectiva
   Larg = .ScaleWidth
   Alt = .ScaleHeight
  End With
  If Alt > 0 Then
   fAspect = Larg / Alt
  Else
   fAspect = 0
  End If
  glViewport 0, 0, Larg, Alt
  glMatrixMode GL_PROJECTION
  glLoadIdentity
  'glMultMatrixf Troca_X_Y(0)
  gluPerspective 35!, fAspect, 1!, 100!
  
  glMatrixMode GL_MODELVIEW
  glLoadIdentity
  glPushMatrix
   glRotatef -90, 0#, 0#, 1#
   glScalef -1, 1, 1
   glGetFloatv GL_MODELVIEW_MATRIX, Troca_X_Y(0)
  glPopMatrix
  gluLookAt Cam_X, Cam_Y, Cam_Z, 0, 0, 0, 0, 0, 1
  glMultMatrixf Troca_X_Y(0)
  
  'Configurações específicas da VISTA FRONTAL
  wglMakeCurrent hDCFrontal, hGLRCFrontal
  glClearColor 1#, 1#, 1#, 1#
  glEnable GL_DEPTH_TEST
  glPolygonMode GL_FRONT_AND_BACK, GL_LINE
  With frmMain.picFrontal
   Larg = .ScaleWidth: Alt = .ScaleHeight
  End With
  glViewport 0, 0, Larg, Alt
  glMatrixMode GL_PROJECTION
  glLoadIdentity
  glOrtho -5, 5, -5, 5, -5, 5
  glMatrixMode GL_MODELVIEW
  glLoadIdentity
  glMultMatrixf Troca_X_Y(0)
  glRotatef 90, 0#, 0#, 1#
  glRotatef 90, 1#, 0#, 0#
  
  'Configurações específicas da VISTA LATERAL
  wglMakeCurrent hDCLateral, hGLRCLateral
  glClearColor 1#, 1#, 1#, 1#
  glEnable GL_DEPTH_TEST
  glPolygonMode GL_FRONT_AND_BACK, GL_LINE
  With frmMain.picLateral
   Larg = .ScaleWidth: Alt = .ScaleHeight
  End With

  glViewport 0, 0, Larg, Alt
  glMatrixMode GL_PROJECTION
  glLoadIdentity
  glOrtho -5, 5, -5, 5, -5, 5
  glMatrixMode GL_MODELVIEW
  glLoadIdentity
  glMultMatrixf Troca_X_Y(0)
  glRotatef 180, 0#, 0#, 1#
  glRotatef -90, 0#, 1#, 0#
  
  'Configurações específicas da VISTA SUPERIOR
  wglMakeCurrent hDCSuperior, hGLRCSuperior
  glClearColor 1#, 1#, 1#, 1#
  glEnable GL_DEPTH_TEST
  glPolygonMode GL_FRONT_AND_BACK, GL_LINE
  'glEnable GL_COLOR_MATERIAL
  'glEnable GL_LIGHTING
  'glEnable GL_LIGHT0
  With frmMain.picSuperior
   Larg = .ScaleWidth: Alt = .ScaleHeight
  End With
  glViewport 0, 0, Larg, Alt
  glMatrixMode GL_PROJECTION
  glLoadIdentity
  glOrtho -5, 5, -5, 5, -5, 5
  glMatrixMode GL_MODELVIEW
  glLoadIdentity
  glMultMatrixf Troca_X_Y(0)
  glRotatef 90, 0#, 0#, 1#
  
  'Configurações específicas da ÉPURA:
  wglMakeCurrent hDCEpura, hGLRCEpura
  glClearColor 0.8, 0.95, 0.9, 1#
  glEnable GL_DEPTH_TEST
  glPolygonMode GL_FRONT_AND_BACK, GL_LINE
  With frmMain.picSuperior
   Larg = .ScaleWidth: Alt = .ScaleHeight
  End With
  glViewport 0, 0, Larg, Alt
  glMatrixMode GL_PROJECTION
  glLoadIdentity
  glOrtho -5, 5, -5, 5, -5, 5
  glMatrixMode GL_MODELVIEW
  glLoadIdentity
  glMultMatrixf Troca_X_Y(0)
  
End Sub
