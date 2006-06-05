Attribute VB_Name = "basVisual"
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
Const PFD_TYPE_COLORINDEX = 1
Const PFD_MAIN_PLANE = 0
Const PFD_DOUBLEBUFFER = 1
Const PFD_DRAW_TO_WINDOW = &H4
Const PFD_SUPPORT_OPENGL = &H20
Const PFD_NEED_PALETTE = &H80

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

'Public hPalette As Long
Public hGLRC As Long


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


