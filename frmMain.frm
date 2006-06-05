VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Main"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   377
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   424
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   1111
      ButtonWidth     =   1164
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "CLIQUE"
            Key             =   "Click"
            Object.ToolTipText     =   "Teste este botão..."
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
 Unload Me
 End
End If
End Sub

'Dim xAngle As GLfloat
'Dim yAngle As GLfloat
'Dim zAngle As GLfloat

Private Sub Form_Load()
    Dim hGLRC As Long
    Dim fAspect As GLfloat
    Call basVisual.InitializeArrays
    
    'xAngle = 0
    'yAngle = 0
    'zAngle = 0

    basVisual.SetupPixelFormat hDC
    
    hGLRC = wglCreateContext(hDC)
    wglMakeCurrent hDC, hGLRC
    
    glEnable GL_DEPTH_TEST
    glEnable GL_DITHER
    glDepthFunc GL_LESS
    glClearDepth 1
    glClearColor 0, 0, 0, 0
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    If frmMain.ScaleHeight > 0 Then
    fAspect = frmMain.ScaleWidth / frmMain.ScaleHeight
    Else
    fAspect = 0
    End If
    
    'gluPerspective 60, fAspect, 1, 2000
    gluOrtho2D -5, 5, -5, 5
    glViewport 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight
    

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

    MontaEixos
    
    Form_Paint

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
 Dim x1, y1 As GLdouble
 Const TAM = 5
 Dim cx, cy As GLdouble

 
If Button <> 1 Then Exit Sub
x1 = X / Me.ScaleWidth
y1 = y / Me.ScaleHeight
 cx = 1 - 2 * x1
 cy = 2 * y1 - 1
 
 glMatrixMode GL_PROJECTION
  glLoadIdentity
  gluOrtho2D TAM * (cx - 1), TAM * (cx + 1), TAM * (cy - 1), TAM * (cy + 1)
 
 glMatrixMode GL_MODELVIEW
 Form_Paint
 SwapBuffers hDC
End Sub

Private Sub Form_Paint()
    Dim I As Integer
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim quadObj As GLUquadric
    
    glLoadIdentity

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
    
    'Grid
    'glPushMatrix
    'glTranslatef 0, -2, 0
    MostraEixos
    'glPopMatrix
    SwapBuffers hDC
        
End Sub
Private Sub Form_Resize()

    glViewport 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight
    Form_Paint
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If basVisual.hGLRC <> 0 Then
        wglMakeCurrent 0, 0
        wglDeleteContext basVisual.hGLRC
    End If
    
    'If hPalette <> 0 Then
        'DeleteObject hPalette
    'End If

End Sub
Sub MontaEixos()

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
Sub MostraEixos()

    glPushAttrib GL_LIGHTING
    glDisable GL_LIGHTING
    
    glPushMatrix
        glColor3ub 0, 255, 0
        glCallList m_Grid
    glPopMatrix
    glPopAttrib
    glEnable GL_LIGHTING

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
 Dim imgX As ListImage
 Dim btnButton As Button
 Static n As Integer
 
 n = n + 1
 If Button.Key = "Click" And Toolbar1.Buttons.Count > 1 Then
  Set btnButton = Toolbar1.Buttons(1)
  Set Toolbar1.ImageList = Nothing
  ImageList1.ListImages.Clear
  
  Toolbar1.Buttons.Clear
  Toolbar1.Buttons.Add 1, btnButton.Key, btnButton.Caption, btnButton.Style, btnButton.Image
  n = 0
  Exit Sub
 End If
 
 If Not Dir(App.Path & "\IMG\" & Format(n, "00") & ".bmp") <> "" Then n = n - 1: Exit Sub
 
 Set imgX = ImageList1.ListImages. _
 Add(, , LoadPicture(App.Path & "\IMG\" & Format(n, "00") & ".bmp"))
 imgX.Key = "img" & n ' Use the new reference to assign Key.
 
 ImageList1.MaskColor = vbWhite
 'ImageList1.MaskColor = vbRed
 'ImageList1.MaskColor = vbBlue
 
 Toolbar1.ImageList = ImageList1
 Set btnButton = Toolbar1.Buttons.Add(, "b" & n, "Botão " & n, tbrDefault, n)
btnButton.ToolTipText = "Botão nº " & n
End Sub
