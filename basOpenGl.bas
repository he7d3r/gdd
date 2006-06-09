Attribute VB_Name = "basVbOpenGl"
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
Private Declare Function SetPixelFormat Lib "gdi32" (ByVal hDC As Long, ByVal I As Long, pfd As PIXELFORMATDESCRIPTOR) As Boolean
Private Declare Sub SwapBuffers Lib "gdi32" (ByVal hDC As Long)
Private Declare Function wglCreateContext Lib "OpenGL32" (ByVal hDC As Long) As Long
Private Declare Sub wglDeleteContext Lib "OpenGL32" (ByVal hContext As Long)
Private Declare Sub wglMakeCurrent Lib "OpenGL32" (ByVal l1 As Long, ByVal l2 As Long)

'Public hPalette As Long
Public hGLRC As Long
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
Public Sub Finalizar_OpenGL() 'ByVal hDC As Long)
 If basVbOpenGl.hGLRC <> 0 Then
  wglMakeCurrent 0, 0
  wglDeleteContext basVbOpenGl.hGLRC
 End If
 'If hPalette <> 0 Then
   'DeleteObject hPalette
 'End If
End Sub

