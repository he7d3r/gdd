Attribute VB_Name = "basVbOpenGl"
Option Explicit
'*************************************************************************************
'   Os tipos de dados "PALETTEENTRY", "LOGPALETTE" e "PIXELFORMATDESCRIPTOR"
'são definidos na biblioteca "VBOpenGL", não precisamos redeclará-los.
'As constantes relacionadas abaixo também pertencem a "VBOpenGL".'
'- Classe GDI (Graphics Device Interface):
' PFD_MAIN_PLANE; PFD_OVERLAY_PLANE; PFD_TYPE_COLORINDEX;
' PFD_TYPE_RGBA; PFD_UNDERLAY_PLANE
'- Classe WGL:
' PFD_DEPTH_DONTCARE; PFD_DOUBLEBUFFER; PFD_DOUBLEBUFFER_DONTCARE;
' PFD_DRAW_TO_BITMAP; PFD_DRAW_TO_WINDOW; PFD_GENERIC_ACCELERATED;
' PFD_GENERIC_FORMAT; PFD_NEED_PALETTE; PFD_NEED_SYSTEM_PALETTE;
' PFD_STEREO; PFD_STEREO_DONTCARE; PFD_SUPPORT_DIRECTDRAW; PFD_SUPPORT_GDI;
' PFD_SUPPORT_OPENGL; PFD_SWAP_COPY; PFD_SWAP_EXCHANGE; PFD_SWAP_LAYER_BUFFERS
'*************************************************************************************
'DECLARE: Usado para referenciar procedimentos externos em DLL's...
Private Declare Function ChoosePixelFormat Lib "gdi32" (ByVal hDC As Long, pfd As PIXELFORMATDESCRIPTOR) As Long
Private Declare Function SetPixelFormat Lib "gdi32" (ByVal hDC As Long, ByVal i As Long, pfd As PIXELFORMATDESCRIPTOR) As Boolean
Private Declare Sub SwapBuffers Lib "gdi32" (ByVal hDC As Long)
Private Declare Function wglCreateContext Lib "OpenGL32" (ByVal hDC As Long) As Long
Private Declare Sub wglDeleteContext Lib "OpenGL32" (ByVal hContext As Long)
Private Declare Sub wglMakeCurrent Lib "OpenGL32" (ByVal l1 As Long, ByVal l2 As Long)

Public Sub ErroFatal(ByVal msgErro As String)
   MsgBox "ERRO FATAL: " & msgErro, _
      vbCritical + vbApplicationModal + vbOKOnly + vbDefaultButton1, _
      "Erro fatal com """ & App.Title & """"
   Unload frmMDIGeo3d
   Set frmMDIGeo3d = Nothing
   End
End Sub
Public Function Ajusta_FormatoPixel(ByVal hDC As Long) As Long
   Dim pfd As PIXELFORMATDESCRIPTOR
   'PIXEL FORMAT: Sempre há 24 tipos básicos disponíveis.
   'Surgem outros se existir uma placa 3d no computador...
   Dim PixelFormat As Integer
   Const ERR_PIX_FORMAT = "Erro na rotina 'ChoosePixelFormat'!"
   
   'Define em 'pfd' as propriedades requeridas para os pixels...
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
   
   'Verifica se está disponível o formato 'pfd'...
   PixelFormat = ChoosePixelFormat(hDC, pfd)
   If PixelFormat = 0 Then ErroFatal ERR_PIX_FORMAT
   Ajusta_FormatoPixel = PixelFormat
End Function
Public Sub Finalizar_OpenGL(IdDoc As Integer)
   Dim j As Vista ' Cada vista no documento
   Dim Qtd As Integer  ' Total de documentos
   
   If Not Doc(IdDoc).Deletado Then
      For j = PERSPECTIVA To EPURA
         If Doc(IdDoc).frm.hGLRC_Vista(j) <> 0 Then
            wglMakeCurrent 0, 0 'NULL, NULL
            wglDeleteContext Doc(IdDoc).frm.hGLRC_Vista(j)
         End If
      Next j
   End If
End Sub
