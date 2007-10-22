Attribute VB_Name = "basVisualização"
Option Explicit
Public Troca_X_Y(0 To 15) As GLfloat

Public Sub Inicializa_OpenGL(IdDoc As Integer)
   Dim pfd As PIXELFORMATDESCRIPTOR
   Dim v As Vista
   Dim Result As Long
   Dim Larg As GLsizei, Alt As GLsizei
   Dim Cor_Neblina(0 To 3) As GLfloat

   If Doc(IdDoc).Deletado Then
      ErroFatal "Documento " & IdDoc & " não existe. Não pode ser inicializado!"
      Exit Sub
   End If
   
   'Ajusta um pfd adequado para operar com o Doc usando o hDC passado
   Result = basVbOpenGl.Ajusta_FormatoPixel(Doc(IdDoc).frm.hDC_Vista(PERSPECTIVA))
   For v = PERSPECTIVA To EPURA
      'Define o formato dos pixels conforme descrito em 'pfd'
      SetPixelFormat Doc(IdDoc).frm.hDC_Vista(v), Result, pfd
      Doc(IdDoc).frm.hGLRC_Vista(v) = wglCreateContext(Doc(IdDoc).frm.hDC_Vista(v))
   Next v
 
   'Inicializa parâmetros, algumas matrizes de iluminação, etc...
     
   'QObj = gluNewQuadric()
   With Doc(IdDoc).frm
      .Phi = 70: .Ro = 15: .Theta = 15
      .Cam_X = .Ro * Sin(.Phi * DEG) * Cos(.Theta * DEG)
      .Cam_Y = .Ro * Sin(.Phi * DEG) * Sin(.Theta * DEG)
      .Cam_Z = .Ro * Cos(.Phi * DEG)
      frmMDIGDD.staInfo.Panels(2).Text = "CÂMERA:  ( " _
                                             & Format(.Cam_X, "0.0") & " ;  " _
                                             & Format(.Cam_Y, "0.0") & " ;  " _
                                             & Format(.Cam_Z, "0.0") & ")cart     ( " _
                                             & Format(.Phi, "0") & " ;  " _
                                             & Format(.Theta, "#0") & " ;  " _
                                             & Format(.Ro, "#0") & ")esf"
      For v = PERSPECTIVA To EPURA
         wglMakeCurrent .hDC_Vista(v), .hGLRC_Vista(v)
         
         'Todas as telas têm teste de profundidade e desenham apenas o contorno dos polígonos
         glEnable GL_DEPTH_TEST
         glPolygonMode GL_FRONT_AND_BACK, GL_LINE
         
         With .picVista(v)
            Larg = .ScaleWidth: Alt = .ScaleHeight
         End With
         glViewport 0, 0, Larg, Alt
         glMatrixMode GL_PROJECTION
         glLoadIdentity
         
         If v = PERSPECTIVA Then
            If Alt > 0 Then
               .fAspect = Larg / Alt
            Else
               .fAspect = 0
            End If
            gluPerspective 35!, .fAspect, DIST_MIN_CENA, DIST_MAX_CENA
            'Define a matriz Troca_X_Y para inverter o sistema de coordenadas
            glMatrixMode GL_MODELVIEW
            glLoadIdentity
            glPushMatrix
               glRotatef -90, 0#, 0#, 1#
               glScalef -1, 1, 1
               glGetFloatv GL_MODELVIEW_MATRIX, Troca_X_Y(0)
            glPopMatrix
            
            gluLookAt .Cam_X, .Cam_Y, .Cam_Z, 0, 0, 0, 0, 0, 1
         Else
            glOrtho -5, 5, -5, 5, -5, 5 '-10, 10, -10, 10, -10, 10
            glMatrixMode GL_MODELVIEW
            glLoadIdentity
         End If
         glMultMatrixf Troca_X_Y(0)
         
         Select Case v
         Case PERSPECTIVA
            Cor_Neblina(0) = 0.9: Cor_Neblina(1) = 1#: Cor_Neblina(2) = 1#: Cor_Neblina(3) = 1#
            'glClearColor 0.9, 1#, 1#, 1#
            glClearColor Cor_Neblina(0), Cor_Neblina(1), Cor_Neblina(2), Cor_Neblina(3)
            glEnable GL_FOG
               glFogi GL_FOG_MODE, GL_LINEAR
               'glFogi GL_FOG_MODE, GL_EXP2
               'glFogi GL_FOG_MODE, GL_EXP '=default
               glFogfv GL_FOG_COLOR, Cor_Neblina(0)
               'glFogf GL_FOG_DENSITY, 0.35
               glHint GL_FOG_HINT, GL_DONT_CARE
               glFogf GL_FOG_START, .Ro
               glFogf GL_FOG_END, DIST_MAX_CENA
            glShadeModel (GL_SMOOTH)
            glEnable GL_COLOR_MATERIAL
            
            'glEnable GL_CULL_FACE  'glFrontFace GL_CCW
            'glClearDepth 1         'glDepthFunc cfLEqual
            'glBlendFunc GL_SRC_ALPHA, GL_DST_ALPHA
            'glEnable GL_CULL_FACE  'glEnable GL_LIGHTING
            'glLightfv GL_LIGHT0, GL_AMBIENT, Amb_Dif_Light(0)
            'glLightfv GL_LIGHT0, GL_DIFFUSE, Amb_Dif_Light(0)
            'glLightfv GL_LIGHT0, GL_POSITION, Light0Pos(0)
            'glEnable GL_LIGHT0

         Case FRONTAL
            glClearColor 1#, 1#, 1#, 1#
            glRotatef 90, 0#, 0#, 1#
            glRotatef 90, 1#, 0#, 0#
            
         Case LATERAL
            glClearColor 1#, 1#, 1#, 1#
            glRotatef 180, 0#, 0#, 1#
            glRotatef -90, 0#, 1#, 0#
            
         Case SUPERIOR
            glClearColor 1#, 1#, 1#, 1#
            glRotatef 90, 0#, 0#, 1#
            
         Case EPURA
            glClearColor 0.8, 0.95, 0.9, 1#
            glRotatef 90, 0#, 0#, 1# 'Igual à vista SUPERIOR
            'Na rotina Paint:
            'Desenha uma vez/posiciona vista frontal/desenha de novo/posiciona vista superior
            
         End Select
      Next v
  End With
End Sub
