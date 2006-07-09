Attribute VB_Name = "basVbOpenGl"
Option Explicit
'*************************************************************************************
'   Os tipos de dados "PALETTEENTRY", "LOGPALETTE" e "PIXELFORMATDESCRIPTOR"
'são definidos na biblioteca "VBOpenGL", não precisamos redeclará-los.
'As constantes relacionadas abaixo também pertencem a "VBOpenGL".
'
'- Classe GDI (Graphics Device Interface):
' PFD_MAIN_PLANE
' PFD_OVERLAY_PLANE
' PFD_TYPE_COLORINDEX
' PFD_TYPE_RGBA
' PFD_UNDERLAY_PLANE
'
'- Classe WGL:
' PFD_DEPTH_DONTCARE
' PFD_DOUBLEBUFFER
' PFD_DOUBLEBUFFER_DONTCARE
' PFD_DRAW_TO_BITMAP
' PFD_DRAW_TO_WINDOW
' PFD_GENERIC_ACCELERATED
' PFD_GENERIC_FORMAT
' PFD_NEED_PALETTE
' PFD_NEED_SYSTEM_PALETTE
' PFD_STEREO
' PFD_STEREO_DONTCARE
' PFD_SUPPORT_DIRECTDRAW
' PFD_SUPPORT_GDI
' PFD_SUPPORT_OPENGL
' PFD_SWAP_COPY
' PFD_SWAP_EXCHANGE
' PFD_SWAP_LAYER_BUFFERS
'*************************************************************************************

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
Global hDC1 As Long, hDC2 As Long 'HDC's das ViewPorts
Sub FatalError(ByVal msgErro As String)
    MsgBox "ERRO FATAL: " & msgErro, vbCritical + vbApplicationModal + vbOKOnly + vbDefaultButton1, "Erro fatal com """ & App.Title & """"
    Unload frmMain
    Set frmMain = Nothing
    End
End Sub
Sub SetupPixelFormat(ByVal hDC As Long)
    Dim pfd As PIXELFORMATDESCRIPTOR
    Dim PixelFormat As Integer
    
    With pfd
    .nSize = Len(pfd)
    .nVersion = 1
    .dwFlags = PFD_SUPPORT_OPENGL Or _
               PFD_DRAW_TO_WINDOW Or _
               PFD_DOUBLEBUFFER Or _
               PFD_TYPE_RGBA
    .iPixelType = PFD_TYPE_RGBA
    .cColorBits = 24
    .cDepthBits = 24
    .iLayerType = PFD_MAIN_PLANE
    End With
    
    PixelFormat = ChoosePixelFormat(hDC, pfd)
    If PixelFormat = 0 Then FatalError "Não foi possível obter um formato adequado para os pixels!"
    SetPixelFormat hDC, PixelFormat, pfd
End Sub
Public Sub Finalizar_OpenGL() 'ByVal hDC As Long)
 If basVbOpenGl.hGLRC <> 0 Then
  wglMakeCurrent 0, 0 'NULL, NULL
  wglDeleteContext basVbOpenGl.hGLRC
 End If
 'If hPalette <> 0 Then
   'DeleteObject hPalette
 'End If
End Sub

