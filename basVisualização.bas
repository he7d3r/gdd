Attribute VB_Name = "basVisualização"
Option Explicit
Public Troca_X_Y(0 To 15) As GLfloat

Public Sub Inicializa_OpenGL(IdDoc As Integer)
   Dim pfd As PIXELFORMATDESCRIPTOR
   Dim j As Vista
   Dim Result As Long
   Dim Larg As GLsizei, Alt As GLsizei

   If Doc(IdDoc).Deletado Then
      ErroFatal "Documento " & IdDoc & " não existe. Não pode ser inicializado!"
      Exit Sub
   End If
   
   'Ajusta um pfd adequado para operar com o Doc usando o hDC passado
   Result = basVbOpenGl.Ajusta_FormatoPixel(Doc(IdDoc).frm.hDC_Vista(PERSPECTIVA))
   For j = PERSPECTIVA To EPURA
      'Define o formato dos pixels conforme descrito em 'pfd'
      SetPixelFormat Doc(IdDoc).frm.hDC_Vista(j), Result, pfd
      Doc(IdDoc).frm.hGLRC_Vista(j) = wglCreateContext(Doc(IdDoc).frm.hDC_Vista(j))
   Next j
 
   'Inicializa parâmetros, algumas matrizes de iluminação, etc...
     
   'QObj = gluNewQuadric()
   With Doc(IdDoc).frm
      .Phi = 60: .Theta = 45: .Ro = 10
      .Cam_X = .Ro * Sin(.Phi * DEG) * Cos(.Theta * DEG)
      .Cam_Y = .Ro * Sin(.Phi * DEG) * Sin(.Theta * DEG)
      .Cam_Z = .Ro * Cos(.Phi * DEG)
   
      'Configurações específicas da PERSPECTIVA
      wglMakeCurrent Doc(IdDoc).frm.hDC_Vista(PERSPECTIVA), Doc(IdDoc).frm.hGLRC_Vista(PERSPECTIVA)
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
      
      With .picPerspectiva
         Larg = .ScaleWidth: Alt = .ScaleHeight
      End With
      If Alt > 0 Then
         .fAspect = Larg / Alt
      Else
         .fAspect = 0
      End If
      glViewport 0, 0, Larg, Alt
      glMatrixMode GL_PROJECTION
      glLoadIdentity
      'glMultMatrixf Troca_X_Y(0)
      gluPerspective 35!, .fAspect, 1!, 100!
      
      glMatrixMode GL_MODELVIEW
      glLoadIdentity
      'Define a matriz Troca_X_Y para inverter o sistema de coordenadas
      glPushMatrix
         glRotatef -90, 0#, 0#, 1#
         glScalef -1, 1, 1
         glGetFloatv GL_MODELVIEW_MATRIX, Troca_X_Y(0)
      glPopMatrix
      
      gluLookAt .Cam_X, .Cam_Y, .Cam_Z, 0, 0, 0, 0, 0, 1
      glMultMatrixf Troca_X_Y(0)
   
  
      'Configurações específicas da VISTA FRONTAL
      wglMakeCurrent Doc(IdDoc).frm.hDC_Vista(FRONTAL), Doc(IdDoc).frm.hGLRC_Vista(FRONTAL)
      
      glClearColor 1#, 1#, 1#, 1#
      glEnable GL_DEPTH_TEST
      glPolygonMode GL_FRONT_AND_BACK, GL_LINE
      With .picFrontal
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
      wglMakeCurrent Doc(IdDoc).frm.hDC_Vista(LATERAL), Doc(IdDoc).frm.hGLRC_Vista(LATERAL)
      
      glClearColor 1#, 1#, 1#, 1#
      glEnable GL_DEPTH_TEST
      glPolygonMode GL_FRONT_AND_BACK, GL_LINE
      With .picLateral
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
      wglMakeCurrent Doc(IdDoc).frm.hDC_Vista(SUPERIOR), Doc(IdDoc).frm.hGLRC_Vista(SUPERIOR)
      
      glClearColor 1#, 1#, 1#, 1#
      glEnable GL_DEPTH_TEST
      glPolygonMode GL_FRONT_AND_BACK, GL_LINE
      'glEnable GL_COLOR_MATERIAL
      'glEnable GL_LIGHTING
      'glEnable GL_LIGHT0
      With .picSuperior
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
      wglMakeCurrent Doc(IdDoc).frm.hDC_Vista(EPURA), Doc(IdDoc).frm.hGLRC_Vista(EPURA)
      
      glClearColor 0.8, 0.95, 0.9, 1#
      glEnable GL_DEPTH_TEST
      glPolygonMode GL_FRONT_AND_BACK, GL_LINE
      With .picEpura
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
  End With
End Sub
