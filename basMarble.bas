Attribute VB_Name = "basMarble"
Option Explicit
'Private Type PALETTEENTRY
'    peRed As Byte
'    peGreen As Byte
'    peBlue As Byte
'    peFlags As Byte
'End Type
'Private Type LOGPALETTE
'    palVersion As Integer
'    palNumEntries As Integer
'    palPalEntry(0 To 255) As PALETTEENTRY
'End Type
Private Type PIXELFORMATDESCRIPTOR
    nSize As Integer
    nVersion As Integer
    dwFlags As Long
    iPixelType As Byte
    cColorBits As Byte
    cRedBits As Byte
    cRedShift As Byte
    cGreenBits As Byte
    cGreenShift As Byte
    cBlueBits As Byte
    cBlueShift As Byte
    cAlphaBits As Byte
    cAlphaShift As Byte
    cAccumBits As Byte
    cAccumRedBits As Byte
    cAccumGreenBits As Byte
    cAccumBlueBits As Byte
    cAccumAlpgaBits As Byte
    cDepthBits As Byte
    cStencilBits As Byte
    cAuxBuffers As Byte
    iLayerType As Byte
    bReserved As Byte
    dwLayerMask As Long
    dwVisibleMask As Long
    dwDamageMask As Long
End Type

Const PFD_TYPE_RGBA = 0
'Const PFD_TYPE_COLORINDEX = 1
Const PFD_MAIN_PLANE = 0
Const PFD_DOUBLEBUFFER = 1
Const PFD_DRAW_TO_WINDOW = &H4
Const PFD_SUPPORT_OPENGL = &H20
'Const PFD_NEED_PALETTE = &H80

Private Declare Function ChoosePixelFormat Lib "gdi32" (ByVal hDC As Long, pfd As PIXELFORMATDESCRIPTOR) As Long
'Private Declare Function CreatePalette Lib "gdi32" (pPal As LOGPALETTE) As Long
'Private Declare Sub DeleteObject Lib "gdi32" (hObject As Long)
'Private Declare Sub DescribePixelFormat Lib "gdi32" (ByVal hDC As Long, ByVal PixelFormat As Long, ByVal nBytes As Long, pfd As PIXELFORMATDESCRIPTOR)
'Private Declare Function GetDC Lib "gdi32" (ByVal hWnd As Long) As Long
'Private Declare Function GetPixelFormat Lib "gdi32" (ByVal hDC As Long) As Long
'Private Declare Sub GetSystemPaletteEntries Lib "gdi32" (ByVal hDC As Long, ByVal start As Long, ByVal entries As Long, ByVal ptrEntries As Long)
'Private Declare Sub RealizePalette Lib "gdi32" (ByVal hPalette As Long)
'Private Declare Sub SelectPalette Lib "gdi32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bln As Long)
Private Declare Function SetPixelFormat Lib "gdi32" (ByVal hDC As Long, ByVal i As Long, pfd As PIXELFORMATDESCRIPTOR) As Boolean
Private Declare Sub SwapBuffers Lib "gdi32" (ByVal hDC As Long)
Private Declare Function wglCreateContext Lib "OpenGL32" (ByVal hDC As Long) As Long
Private Declare Sub wglDeleteContext Lib "OpenGL32" (ByVal hContext As Long)
Private Declare Sub wglMakeCurrent Lib "OpenGL32" (ByVal l1 As Long, ByVal l2 As Long)


'Type MARBLE_STRUCT
    'x As GLint
    'Y As GLint
    'z As GLint
    'i As GLint
    'j As GLint
'End Type
'
'Public hPalette As Long
Public hGLRC As Long

'Public Marble_Coord As MARBLE_STRUCT
Public LightPos(3) As GLfloat
Public SpecRef(3) As GLfloat
Public Diffuse(3) As GLfloat
Public lmodel_ambient(3) As GLfloat
Public Const TRIANGLE_COUNT = 11
Public vData(23, 2) As GLfloat
Public Index(TRIANGLE_COUNT, 2) As GLfloat

Public m_Grid As Integer
Public m_Cube As Integer
Public m_Translate_X As Integer
Public m_Translate_Y As Integer
Public m_Translate_Z As Integer
Public m_camera_radsFromEast As GLfloat
Public m_translationUnit As Double
Public m_camera_direction_y As Integer

Public vMap(10, 10) As GLint

Sub FatalError(ByVal strMessage As String)
    MsgBox "Fatal Error: " & strMessage, vbCritical + vbApplicationModal + vbOKOnly + vbDefaultButton1, "Fatal Error In " & App.Title
    Unload frmMain
    Set frmMain = Nothing
    End
End Sub
Sub SetupPixelFormat(ByVal hDC As Long)
    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim PixelFormat As Integer
    pfd.nSize = Len(pfd)
    pfd.nVersion = 1
    pfd.dwFlags = PFD_SUPPORT_OPENGL Or PFD_DRAW_TO_WINDOW Or PFD_DOUBLEBUFFER Or PFD_TYPE_RGBA
    pfd.iPixelType = PFD_TYPE_RGBA
    pfd.cColorBits = 24
    pfd.cDepthBits = 24
    pfd.iLayerType = PFD_MAIN_PLANE
    PixelFormat = ChoosePixelFormat(hDC, pfd)
    If PixelFormat = 0 Then FatalError "Could not retrieve pixel format!"
    SetPixelFormat hDC, PixelFormat, pfd
End Sub
'Sub SetupPalette(ByVal lhDC As Long)
    'Dim PixelFormat As Long
    'Dim pfd As PIXELFORMATDESCRIPTOR
    'Dim pPal As LOGPALETTE
    'Dim PaletteSize As Long
    'PixelFormat = GetPixelFormat(lhDC)
    'DescribePixelFormat lhDC, PixelFormat, Len(pfd), pfd
    'If (pfd.dwFlags And PFD_NEED_PALETTE) <> 0 Then
        'PaletteSize = 2 ^ pfd.cColorBits
    'Else
        'Exit Sub
    'End If
'
    'pPal.palVersion = &H300
    'pPal.palNumEntries = PaletteSize
    'Dim redMask As Long
    'Dim GreenMask As Long
    'Dim BlueMask As Long
    'Dim i As Long
    'redMask = 2 ^ pfd.cRedBits - 1
    'GreenMask = 2 ^ pfd.cGreenBits - 1
    'BlueMask = 2 ^ pfd.cBlueBits - 1
    'For i = 0 To PaletteSize - 1
        'With pPal.palPalEntry(i)
            '.peRed = i
            '.peGreen = i
            '.peBlue = i
            '.peFlags = 0
        'End With
    'Next
    'GetSystemPaletteEntries frmMain.hDC, 0, 256, VarPtr(pPal.palPalEntry(0))
    'hPalette = CreatePalette(pPal)
    'If hPalette <> 0 Then
        'SelectPalette lhDC, hPalette, False
        'RealizePalette lhDC
    'End If
'End Sub

Public Sub InitializeArrays()
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
    
    
    'Marble_Coord.x = -5#
    'Marble_Coord.Y = 0
    'Marble_Coord.z = -25
   '
    'Marble_Coord.i = 2
    'Marble_Coord.j = 4
    
    
    'm_camera_direction_y = 0#
    'm_Translate_Y = 40
    'm_Translate_Z = 40
    
    'Row 1
'    vMap(0, 0) = 1
    'vMap(0, 1) = 0
    'vMap(0, 2) = 0
    'vMap(0, 3) = 0
    'vMap(0, 4) = 1
    'vMap(0, 5) = 1
    'vMap(0, 6) = 0
    'vMap(0, 7) = 1
    'vMap(0, 8) = 0
    'vMap(0, 9) = 1
    ''Row 2
    'vMap(1, 0) = 1
    'vMap(1, 1) = 0
    'vMap(1, 2) = 1
    'vMap(1, 3) = 1
    'vMap(1, 4) = 1
    'vMap(1, 5) = 1
    'vMap(1, 6) = 0
    'vMap(1, 7) = 0
    'vMap(1, 8) = 0
    'vMap(1, 9) = 1
    ''Row 3
'    vMap(2, 0) = 1
    'vMap(2, 1) = 0
    'vMap(2, 2) = 0
    'vMap(2, 3) = 0
    'vMap(2, 4) = 0
    'vMap(2, 5) = 1
    'vMap(2, 6) = 1
    'vMap(2, 7) = 1
    'vMap(2, 8) = 1
    'vMap(2, 9) = 1
    ''Row 4
    'vMap(3, 0) = 1
    'vMap(3, 1) = 0
    'vMap(3, 2) = 1
    'vMap(3, 3) = 1
    'vMap(3, 4) = 0
    'vMap(3, 5) = 1
    'vMap(3, 6) = 1
    'vMap(3, 7) = 1
    'vMap(3, 8) = 0
    'vMap(3, 9) = 1
    ''Row 5
    'vMap(4, 0) = 1
    'vMap(4, 1) = 0
    'vMap(4, 2) = 1
    'vMap(4, 3) = 1
    'vMap(4, 4) = 0
    'vMap(4, 5) = 1
    'vMap(4, 6) = 1
    'vMap(4, 7) = 1
    'vMap(4, 8) = 0
    'vMap(4, 9) = 1
    ''Row 6
    'vMap(5, 0) = 1
    'vMap(5, 1) = 0
    'vMap(5, 2) = 1
    'vMap(5, 3) = 0
    'vMap(5, 4) = 0
    'vMap(5, 5) = 0
    'vMap(5, 6) = 0
    'vMap(5, 7) = 0
    'vMap(5, 8) = 0
    'vMap(5, 9) = 1
    ''Row 7
    'vMap(6, 0) = 1
    'vMap(6, 1) = 0
    'vMap(6, 2) = 1
    'vMap(6, 3) = 0
    'vMap(6, 4) = 1
    'vMap(6, 5) = 1
    'vMap(6, 6) = 1
    'vMap(6, 7) = 1
    'vMap(6, 8) = 1
'    vMap(6, 9) = 1
    'Row 8
'    vMap(7, 0) = 1
    'vMap(7, 1) = 0
    'vMap(7, 2) = 1
    'vMap(7, 3) = 0
    'vMap(7, 4) = 0
    'vMap(7, 5) = 0
    'vMap(7, 6) = 0
    'vMap(7, 7) = 1
    'vMap(7, 8) = 1
    'vMap(7, 9) = 1
    ''Row 9
    'vMap(8, 0) = 1
    'vMap(8, 1) = 0
    'vMap(8, 2) = 1
    'vMap(8, 3) = 1
    'vMap(8, 4) = 1
    'vMap(8, 5) = 1
    'vMap(8, 6) = 0
    'vMap(8, 7) = 0
    'vMap(8, 8) = 0
    'vMap(8, 9) = 1
    ''Row 10
    'vMap(9, 0) = 1
    'vMap(9, 1) = 1
    'vMap(9, 2) = 1
    'vMap(9, 3) = 1
    'vMap(9, 4) = 1
    'vMap(9, 5) = 1
    'vMap(9, 6) = 1
    'vMap(9, 7) = 1
    'vMap(9, 8) = 1
    'vMap(9, 9) = 1
'
       '
    ''Front (0-3)
    'vData(0, 0) = 1
    'vData(0, 1) = 1
    'vData(0, 2) = 1
    'vData(1, 0) = 1
    'vData(1, 1) = -1
    'vData(1, 2) = 1
    'vData(2, 0) = -1
    'vData(2, 1) = -1
    'vData(2, 2) = 1
    'vData(3, 0) = -1
    'vData(3, 1) = 1
    'vData(3, 2) = 1
'
    ''back (4-7)
    'vData(4, 0) = 1#
    'vData(4, 1) = 1#
    'vData(4, 2) = -1#
    'vData(5, 0) = 1#
    'vData(5, 1) = -1#
    'vData(5, 2) = -1#
    'vData(6, 0) = -1#
    'vData(6, 1) = -1#
    'vData(6, 2) = -1#
    'vData(7, 0) = -1#
    'vData(7, 1) = 1#
    'vData(7, 2) = -1#
    ''right (8-11)
    'vData(8, 0) = 1#
    'vData(8, 1) = 1#
    'vData(8, 2) = 1#
    'vData(9, 0) = 1#
    'vData(9, 1) = 1#
    'vData(9, 2) = -1#
    'vData(10, 0) = 1#
    'vData(10, 1) = -1#
    'vData(10, 2) = -1#
    'vData(11, 0) = 1#
    'vData(11, 1) = -1#
    'vData(11, 2) = 1#
    ''left (12-15)
    'vData(12, 0) = -1#
    'vData(12, 1) = 1#
    'vData(12, 2) = 1#
    'vData(13, 0) = -1#
    'vData(13, 1) = 1#
    'vData(13, 2) = -1#
    'vData(14, 0) = -1#
    'vData(14, 1) = -1#
    'vData(14, 2) = -1#
    'vData(15, 0) = -1#
    'vData(15, 1) = -1#
    'vData(15, 2) = 1#
'
    ''Top (16-20)
    'vData(16, 0) = 1#
    'vData(16, 1) = 1#
    'vData(16, 2) = 1#
    'vData(17, 0) = 1#
    'vData(17, 1) = 1#
    'vData(17, 2) = -1#
    'vData(18, 0) = -1#
    'vData(18, 1) = 1#
    'vData(18, 2) = -1#
    'vData(19, 0) = -1#
    'vData(19, 1) = 1#
    'vData(19, 2) = 1#
'
    ''Botton
    'vData(20, 0) = 1#
    'vData(20, 1) = -1#
    'vData(20, 2) = 1#
    'vData(21, 0) = 1#
    'vData(21, 1) = -1#
    'vData(21, 2) = -1#
    'vData(22, 0) = -1#
    'vData(22, 1) = -1#
    'vData(22, 2) = -1#
    'vData(23, 0) = -1#
    'vData(23, 1) = -1#
    'vData(23, 2) = 1#
'
'
    ''Index
    ''front
    'Index(0, 0) = 0
    'Index(0, 1) = 1
    'Index(0, 2) = 2
    'Index(1, 0) = 0
    'Index(1, 1) = 2
    'Index(1, 2) = 3
    ''Back
    'Index(2, 0) = 4
    'Index(2, 1) = 6
    'Index(2, 2) = 5
    'index(3, 0) = 4
    'Index(3, 1) = 7
    'Index(3, 2) = 6
    ''Right
    'Index(4, 0) = 8
    'Index(4, 1) = 9
    'Index(4, 2) = 10
    'Index(5, 0) = 8
    'Index(5, 1) = 10
    'Index(5, 2) = 11
    ''Left
    'Index(6, 0) = 12
    'Index(6, 1) = 14
    'Index(6, 2) = 13
    'Index(7, 0) = 12
    'Index(7, 1) = 15
    'Index(7, 2) = 14
    ''Top
    'Index(8, 0) = 16
    'Index(8, 1) = 18
    'Index(8, 2) = 17
    'Index(9, 0) = 16
    'Index(9, 1) = 19
    'Index(9, 2) = 18
    ''Bottom
    'Index(10, 0) = 20
    'Index(10, 1) = 21
    'Index(10, 2) = 22
    'Index(11, 0) = 20
    'Index(11, 1) = 22
    'Index(11, 2) = 23
'
    
End Sub

Public Sub RenderTriangle(a As Integer, b As Integer, c As Integer)
    Dim x1 As GLfloat
    Dim y1 As GLfloat
    Dim z1 As GLfloat
    Dim lRC As Long
    
    Dim x2 As GLfloat
    Dim y2 As GLfloat
    Dim z2 As GLfloat
    
    Dim x3 As GLfloat
    Dim y3 As GLfloat
    Dim z3 As GLfloat
    
    Dim v1(3) As GLfloat
    Dim v2(3) As GLfloat
    Dim v3(3) As GLfloat

    Dim v(3) As GLfloat
    Dim w(3) As GLfloat
    Dim out(3) As GLfloat
    
  
    '-----------------------------Get Vertex Data-----------------------
    v1(0) = vData(a, 0)
    v1(1) = vData(a, 1)
    v1(2) = vData(a, 2)
    
    x1 = vData(a, 0)
    y1 = vData(a, 1)
    z1 = vData(a, 2)
    
    v2(0) = vData(b, 0)
    v2(1) = vData(b, 1)
    v2(2) = vData(b, 2)
    
    x2 = vData(b, 0)
    y2 = vData(b, 1)
    z2 = vData(b, 2)
    
    v3(0) = vData(c, 0)
    v3(1) = vData(c, 1)
    v3(2) = vData(c, 2)
    
    x3 = vData(c, 0)
    y3 = vData(c, 1)
    z3 = vData(c, 2)
    '--------------------------------------------------------------------
    
    
    v(0) = x2 - x1
    v(1) = y2 - y1
    v(2) = z2 - z1
    
    w(0) = x3 - x1
    w(1) = y3 - y1
    w(2) = z3 - z1
    
    Call normcrossprod(v, w, out)
    
    Diffuse(0) = 1
    Diffuse(1) = 0
    Diffuse(2) = 0
    Diffuse(3) = 0
    
    glBegin (GL_TRIANGLES)
    glMaterialfv GL_FRONT, GL_AMBIENT_AND_DIFFUSE, Diffuse(0)
    'Flip the normal
    glNormal3f -1 * out(0), -1 * out(1), -1 * out(2)
    glVertex3f v1(0), v1(1), v1(2)
    glVertex3f v2(0), v2(1), v2(2)
    glVertex3f v3(0), v3(1), v3(2)
    glEnd
    
    
End Sub

Public Sub normalize(out() As GLfloat)
Dim d As GLfloat
    d = Sqr(out(0) * out(0) + out(1) * out(1) + out(2) * out(2))
    If (d = 0) Then
        Exit Sub
    End If
    out(0) = out(0) / d
    out(1) = out(1) / d
    out(2) = out(2) / d
    
End Sub

Public Sub normcrossprod(v() As GLfloat, w() As GLfloat, out() As GLfloat)
 '[Vx Vy Vz] X [Wx Wy Wz] =[(Vy*Wz-Wy*Vz),(Wx*Vz-Vx*Wz),(Vx*Wy-Wx*Vy)]
 out(0) = v(1) * w(2) - w(1) * v(2)
 out(1) = w(0) * v(2) - v(0) * w(2)
 out(2) = v(0) * w(1) - w(0) * v(1)
 Call normalize(out)
 
End Sub

'frmMarble
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
 Unload Me
 End
End If
End Sub

'Const WALL_TYPE = 1
'Const MARBLE_BACK = 1
'Const MARBLE_FORWARD = 2
'Const MARBLE_LEFT = 3
'Const MARBLE_RIGHT = 4

'Dim xAngle As GLfloat
'Dim yAngle As GLfloat
'Dim zAngle As GLfloat

'Private Sub cmdBack_Click()
    'CameraBack
'End Sub
'
'Private Sub cmdDeElevate_Click()
    'm_Translate_Y = m_Translate_Y - 1
    'Form_Paint
'End Sub
'
'Private Sub cmdDown_Click()
    'Tilt_Camera_Down
'End Sub

'Private Sub cmdElevate_Click()
'    m_Translate_Y = m_Translate_Y + 1
    'Form_Paint
'End Sub
'
'Private Sub cmdForward_Click()
    'CameraForward
'End Sub
'
'Private Sub cmdLeft_Click()
    'CameraLeft
'End Sub



'Private Sub cmdMarbleBackward_Click()
    'If CheckForWall(MARBLE_BACK) = False Then
    'If Marble_Coord.z < 45 Then
        'Marble_Coord.z = Marble_Coord.z + 10#
        'Form_Paint
        'Marble_Coord.i = Marble_Coord.i + 1
    'End If
    'End If
'End Sub

'Private Sub cmdMarbleForward_Click()
    'If CheckForWall(MARBLE_FORWARD) = False Then
    'If Marble_Coord.z > -45 Then
        'Marble_Coord.z = Marble_Coord.z - 10#
        'Form_Paint
        'Marble_Coord.i = Marble_Coord.i - 1
    'End If
    'End If
'End Sub

'Private Sub cmdMarbleLeft_Click()
    'If CheckForWall(MARBLE_LEFT) = False Then
    'If Marble_Coord.x > -45 Then
        'Marble_Coord.x = Marble_Coord.x - 10#
        'Form_Paint
        'Marble_Coord.j = Marble_Coord.j - 1
    'End If
    'End If
'End Sub

'Private Sub cmdMarbleRight_Click()
    'If CheckForWall(MARBLE_RIGHT) = False Then
    'If Marble_Coord.x < 45 Then
        'Marble_Coord.x = Marble_Coord.x + 10#
        'Form_Paint
        'Marble_Coord.j = Marble_Coord.j + 1
    'End If
    'End If
'End Sub

'Private Sub cmdRight_Click()
    'CameraRight
'End Sub
'
'Private Sub cmdUp_Click()
    'Tilt_Camera_Up
'End Sub

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


'    BuildCube
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
    Dim i As Integer
    Dim a As Integer
    Dim b As Integer
    Dim c As Integer
    Dim quadObj As GLUquadric
    
    glLoadIdentity
    'gluLookAt m_Translate_X, m_Translate_Y, m_Translate_Z, _
    'm_Translate_X + (100# * (Cos(m_camera_radsFromEast))), _
    'm_Translate_Y + m_camera_direction_y, _
    'm_Translate_Z - (100# * Sin(m_camera_radsFromEast)), _
    '0#, 1#, 0#
    'gluLookAt 5, 4, 5, _
    '0#, 0#, 0#, _
    '0#, 0#, 1#
    
    glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
    
    glLightfv GL_LIGHT0, GL_POSITION, basVisual.LightPos(0)
    
    'Marble
    glPushMatrix
        'glTranslatef Marble_Coord.x, Marble_Coord.Y, Marble_Coord.z
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
        gluSphere quadObj, 1, 6, 6
    glPopMatrix
    
    'Grid
    'glPushMatrix
    'glTranslatef 0, -2, 0
    MostraEixos
    'glPopMatrix
    
    'glPushMatrix
    'DisplayWalls
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
    'Dim r As Integer
    'Dim c As Integer
    'Dim nStep As Integer
'
    'nStep = 10
    glPushMatrix
    m_Grid = glGenLists(1)
    glNewList m_Grid, GL_COMPILE
    'glBegin (GL_LINES)
        'For r = -50 To 50 Step nStep
            'glVertex3f r, 0#, -50#
            'glVertex3f r, 0#, 50#
        'Next
        'For c = -50 To 50 Step nStep
            'glVertex3f 50#, 0#, c
            'glVertex3f -50#, 0#, c
        'Next
    'glEnd
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
        'Bottom
        glCallList m_Grid
        '/*Back*/
        'glPushMatrix
            'glTranslatef 0#, 50#, -50#
            'glPushMatrix
                'glRotatef 90#, 1#, 0#, 0#
                'glCallList m_Grid
            'glPopMatrix
        'glPopMatrix
        '/*Front*/
        'glPushMatrix
            'glTranslatef 0#, 50#, 50#
            'glPushMatrix
                'glRotatef 90#, 1#, 0#, 0#
                'glCallList m_Grid
            'glPopMatrix
        'glPopMatrix
        '/*Left Side*/
        'glPushMatrix
            'glTranslatef -50#, 50#, 0#
            'glPushMatrix
                'glRotatef 90, 0#, 0#, 1#
                'glCallList m_Grid
            'glPopMatrix
        'glPopMatrix
        '/*Right Side*/
        'glPushMatrix
            'glTranslatef 50#, 50#, 0#
            'glPushMatrix
                'glRotatef 90, 0#, 0#, 1#
                'glCallList m_Grid
            'glPopMatrix
        'glPopMatrix
    glPopMatrix
    glPopAttrib
    glEnable GL_LIGHTING

End Sub
'
'
'Sub MoveCamera(dStep As Double)
'
    'Dim xChange As Double
    'Dim zChange As Double
'
    'xChange = dStep * Cos(m_camera_radsFromEast)
    'zChange = -dStep * Sin(m_camera_radsFromEast)
'
    'If ((m_Translate_X < 40 + xChange) And (m_Translate_X + xChange) > -40) Then
        'm_Translate_X = m_Translate_X + xChange
    'Else
        'm_Translate_X = m_Translate_X - xChange
    'End If
'
    'If ((m_Translate_Z + zChange) < 40 And (m_Translate_Z + zChange) > -40) Then
        'm_Translate_Z = m_Translate_Z + zChange
    'Else
        'm_Translate_Z = m_Translate_Z - zChange
    'End If
'
'End Sub
'
'Private Sub CameraLeft()
    'm_camera_radsFromEast = m_camera_radsFromEast + (10# / 180# * 3.142)
    'If (m_camera_radsFromEast > (6.28)) Then
    '    m_camera_radsFromEast = 0#
    'End If
    'Form_Paint
'End Sub
'
'Private Sub CameraRight()
'
    'm_camera_radsFromEast = m_camera_radsFromEast - (10# / 180# * 3.142)
    'If (m_camera_radsFromEast < 0#) Then
       'm_camera_radsFromEast = 6.28
    'End If
    'Form_Paint
'End Sub
'
'Private Sub CameraForward()
    'Call MoveCamera(m_translationUnit)
    'Form_Paint
'End Sub
'Private Sub CameraBack()
    'Call MoveCamera(-1# * m_translationUnit)
    'Form_Paint
'End Sub
'
'Private Sub Tilt_Camera_Up()
    'm_camera_direction_y = m_camera_direction_y + 10
    'Form_Paint
'End Sub
'Sub Tilt_Camera_Down()
    'm_camera_direction_y = m_camera_direction_y - 10
    'Form_Paint
'End Sub

'Sub BuildCube()
    'Dim i As Integer
    'Dim a As Integer
    'Dim b As Integer
    'Dim c As Integer
'
    'm_Cube = glGenLists(1)
    'glNewList m_Cube, GL_COMPILE_AND_EXECUTE
     '
    'F'or i = 0 To TRIANGLE_COUNT - 1
        'a = Index(i, 0)
        'b = Index(i, 1)
        'c = Index(i, 2)
        'Call RenderTriangle(a, b, c)
    'Next
        '
    'glEnd
    'glEndList
'
'End Sub

'Private Function CheckForWall(iType As Integer)
    'Dim i As Integer
    'Dim j As Integer
    'Dim iRC As Integer
'
    'Select Case iType
    'Case MARBLE_BACK
        'i = Marble_Coord.i + 1
        'j = Marble_Coord.j
    'Case MARBLE_FORWARD
        'i = Marble_Coord.i - 1
        'j = Marble_Coord.j
    'Case MARBLE_LEFT
        'i = Marble_Coord.i
        'j = Marble_Coord.j - 1
    'Case MARBLE_RIGHT
        'i = Marble_Coord.i
        'j = Marble_Coord.j + 1
    'End Select
'
    '
    'If i = -1 Or j = -1 Then
     '   iRC = False
    'Else
        'If vMap(i, j) = WALL_TYPE Then
            'iRC = True
        'Else
            'iRC = False
        'End If
    'End If
'
    '
    'CheckForWall = iRC
'
'End Function


'Private Sub DisplayWalls()
'Dim i As Integer
'Dim j As Integer
'Dim x As GLfloat
'Dim z As GLfloat
'
    'z = -45
    'For i = 0 To 9
        'x = -45
        'For j = 0 To 9
            'If vMap(i, j) = WALL_TYPE Then
               'glPushMatrix
               'glTranslatef x, 2.5, z
               'glScalef 5, 5, 5
               'glCallList m_Cube
               'glPopMatrix
            'End If
        'x = x + 10
        'Next
        'z = z + 10
   'Next
'
'End Sub






