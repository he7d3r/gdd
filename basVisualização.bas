Attribute VB_Name = "basVisualização"
Option Explicit
Public Troca_X_Y(0 To 15) As GLfloat
Public M_proj(0 To 15) As GLfloat

Public ModelViewMatrix(0 To 15) As GLfloat
Public ProjectionMatrix(0 To 15) As GLfloat
Public ViewPort(0 To 3) As GLint

Public Const PI = 3.14159265
Public Const DEG = PI / 180
Public Cam_X As Single, Cam_Y As Single, Cam_Z As Single
Public Centro_X As Single, Centro_Y As Single, Centro_Z As Single
Public Ox As GLfloat, Oy As GLfloat, Oz As GLfloat 'posição do observador
Public Phi As GLfloat, Theta As GLfloat, Ro As GLfloat

Public Const LADO_PRISMA = 2
Public Const DIST_MIN = 3 * LADO_PRISMA
Public Const DIST_MAX = 30 * LADO_PRISMA

Public Larg As GLsizei, Alt As GLsizei
Public fAspect As GLfloat
Public Amb_Light(3) As GLfloat
Public Dif_Light(3) As GLfloat
Public Spec_Light(3) As GLfloat
Public Light0Pos(3) As GLfloat


Public Sub Inicializar_OpenGL()
Dim pfd As PIXELFORMATDESCRIPTOR
Dim Result As Long
Const EP = 0.1


 'Ajusta um contexto OpenGl para operar com o objeto com o hDC passado
 Result = basVbOpenGl.SetupPixelFormat(hDCPerspectiva)
 'Define o formato dos pixels conforme descrito em 'pfd'
 SetPixelFormat hDCPerspectiva, Result, pfd
 hGLRCPerspectiva = wglCreateContext(hDCPerspectiva)
 
 SetPixelFormat hDCObservador, Result, pfd
 hGLRCObservador = wglCreateContext(hDCObservador)
 
 'inicialize AQUI algumas matrizes de iluminação, e outras...
 QObj = gluNewQuadric()
 
 Phi = 60: Theta = 30: Ro = 20
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
  glPolygonMode GL_FRONT_AND_BACK, GL_FILL 'GL_LINE
  'glEnable GL_CULL_FACE
  glEnable GL_COLOR_MATERIAL
  
  Amb_Light(0) = 0.5
  Amb_Light(1) = 0.5
  Amb_Light(2) = 0.5
  Amb_Light(3) = 1
     
  Dif_Light(0) = 0.5
  Dif_Light(1) = 0.5
  Dif_Light(2) = 0.5
  Dif_Light(3) = 1
  
  Spec_Light(0) = 1
  Spec_Light(1) = 1
  Spec_Light(2) = 1
  Spec_Light(3) = 1
  
  Light0Pos(0) = 0
  Light0Pos(1) = 0
  Light0Pos(2) = 5
  Light0Pos(3) = 1
    
  'ajusta LIGHT0
  glLightfv GL_LIGHT0, GL_AMBIENT, Amb_Light(0)
  glLightfv GL_LIGHT0, GL_DIFFUSE, Dif_Light(0)
  glLightfv GL_LIGHT0, GL_POSITION, Light0Pos(0)
  glLightfv GL_LIGHT0, GL_SPECULAR, Spec_Light(0)
  glEnable GL_LIGHT0
  
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
  gluPerspective 35!, fAspect, 1!, 150!
  
  glMatrixMode GL_MODELVIEW
  glLoadIdentity
  'Define a matriz Troca_X_Y para inverter o sistema de coordenadas
  glPushMatrix
   glRotatef -90, 0#, 0#, 1#
   glScalef -1, 1, 1
   glGetFloatv GL_MODELVIEW_MATRIX, Troca_X_Y(0)
  glPopMatrix
  
  gluLookAt Cam_X, Cam_Y, Cam_Z, Centro_X, Centro_Y, Centro_Z, 0, 0, 1
  glMultMatrixf Troca_X_Y(0)
  
  'M_proj(0) = 0:   M_proj(1) = 0:  M_proj(2) = 0:   M_proj(3) = 0
  'M_proj(4) = -Oy: M_proj(5) = Ox: M_proj(6) = 0:   M_proj(7) = 0
  'M_proj(8) = -Oz: M_proj(9) = 0:  M_proj(10) = Ox: M_proj(11) = 0
  'M_proj(12) = -1: M_proj(13) = 0: M_proj(14) = 0:  M_proj(15) = Ox
  
  M_proj(0) = 0:  M_proj(1) = -Oy: M_proj(2) = -Oz: M_proj(3) = -1
  M_proj(4) = 0:  M_proj(5) = Ox:  M_proj(6) = 0:   M_proj(7) = 0
  M_proj(8) = 0:  M_proj(9) = 0:   M_proj(10) = Ox: M_proj(11) = 0
  M_proj(12) = 0: M_proj(13) = 0:  M_proj(14) = 0:  M_proj(15) = Ox
  
  M_proj(0) = 1:  M_proj(1) = -Oy: M_proj(2) = -Oz: M_proj(3) = -1
  M_proj(4) = 0:  M_proj(5) = Ox:  M_proj(6) = 0:   M_proj(7) = 0
  M_proj(8) = 0:  M_proj(9) = 0:   M_proj(10) = Ox: M_proj(11) = 0
  M_proj(12) = 0: M_proj(13) = 0:  M_proj(14) = 0:  M_proj(15) = 1
  
  'Configurações específicas da VISTA OBSERVADOR
  'wglMakeCurrent hDCObservador, hGLRCObservador
  'glClearColor 1#, 1#, 1#, 1#
  'glEnable GL_DEPTH_TEST
  'glPolygonMode GL_FRONT_AND_BACK, GL_LINE
  'glPolygonMode GL_FRONT_AND_BACK, GL_FILL  'GL_LINE
  'With frmMain.picObservador
  ' Larg = .ScaleWidth: Alt = .ScaleHeight
  'End With
  'glViewport 0, 0, Larg, Alt
  'glMatrixMode GL_PROJECTION
  'glOrtho -5, 5, -5, 5, -5, 5
  'glLoadIdentity
  'glFrustum 0, 2 * LADO_PRISMA + 1, 0, 2 * LADO_PRISMA + 1, 0, -10
  'glMatrixMode GL_MODELVIEW
  'glLoadIdentity
  'gluLookAt Ox, Oy, Oz, 0, 1, 0, 0, 0, 1
  'glMultMatrixf Troca_X_Y(0)
  'glRotatef 90, 0#, 0#, 1#
  'glRotatef 90, 1#, 0#, 0#
  
End Sub
