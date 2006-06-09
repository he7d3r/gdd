Attribute VB_Name = "basObjOpenGl"
Option Explicit
Public Sub MontaEixos()
 glPushMatrix
   m_Grid = glGenLists(1)
   glNewList m_Grid, GL_COMPILE
     glBegin GL_LINES
       glColor3f 1#, 0#, 0#
         glVertex3f 0#, 0#, 0#: glVertex3f 4#, 0#, 0#
       glColor3f 0#, 1#, 0#
         glVertex3f 0#, 0#, 0#: glVertex3f 0#, 4#, 0#
       glColor3f 0#, 0#, 1#
         glVertex3f 0#, 0#, 0#: glVertex3f 0#, 0#, 4#
     glEnd
   glEndList
 glPopMatrix
End Sub
Public Sub Desenha_esfera()
 Dim quadObj As GLUquadric
 
 'glLoadIdentity
 
 'gluLookAt 5, 4, 5, _
 '0#, 0#, 0#, _
 '0#, 0#, 1#
 
 glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
 glLightfv GL_LIGHT0, GL_POSITION, basVisual.LightPos(0)
 glPushMatrix
  quadObj = gluNewQuadric()
  gluQuadricDrawStyle quadObj, GLU_FILL
  gluQuadricNormals quadObj, GLU_SMOOTH
  gluQuadricOrientation quadObj, GLU_OUTSIDE 'GLU_INSIDE
  
 basVisual.Diffuse(0) = 0.5
 basVisual.Diffuse(1) = 0#
 basVisual.Diffuse(2) = 0.5
 basVisual.Diffuse(3) = 1
  glMaterialfv GL_FRONT, GL_AMBIENT_AND_DIFFUSE, basVisual.Diffuse(0)
  'glScalef 2, 2, 2
  gluSphere quadObj, 1, 16, 16
 glPopMatrix

End Sub
Public Sub MostraEixos()

 glPushAttrib GL_LIGHTING
  glDisable GL_LIGHTING

  'glPushMatrix'USAR EM CASO DE TRANSLAÇÃO DOS EIXOS...
   glCallList m_Grid
  'glPopMatrix
 glPopAttrib
 'glEnable GL_LIGHTING
 
End Sub
