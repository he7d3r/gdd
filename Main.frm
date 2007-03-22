VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exemplo - Integrando Vb e OpenGl"
   ClientHeight    =   6915
   ClientLeft      =   585
   ClientTop       =   1170
   ClientWidth     =   10980
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   461
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   732
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ilsFerramentas 
      Left            =   180
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   33
      ImageHeight     =   33
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0D36
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrFerramentas 
      Align           =   3  'Align Left
      Height          =   6915
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   12197
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      MousePointer    =   99
      MouseIcon       =   "Main.frx":1A6C
   End
   Begin VB.PictureBox picEpura 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   915
      ScaleHeight     =   198
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   9
      Top             =   3705
      Width           =   3000
   End
   Begin VB.PictureBox picLateral 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   4395
      ScaleHeight     =   198
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   7
      Top             =   3705
      Width           =   3000
   End
   Begin VB.PictureBox picSuperior 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   7755
      ScaleHeight     =   198
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   6
      Top             =   3705
      Width           =   3000
   End
   Begin VB.PictureBox picFrontal 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   7755
      ScaleHeight     =   198
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   1
      Top             =   345
      Width           =   3000
   End
   Begin VB.PictureBox picPerspectiva 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   915
      MouseIcon       =   "Main.frx":1D86
      ScaleHeight     =   198
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   431
      TabIndex        =   0
      ToolTipText     =   "Botão direito: Mover camera."
      Top             =   345
      Width           =   6495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Épura (1ª e 2ª Proj.):"
      Height          =   195
      Left            =   915
      TabIndex        =   8
      Top             =   3465
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Vista Frontal (2ª Proj.):"
      Height          =   195
      Left            =   7755
      TabIndex        =   5
      Top             =   105
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Vista Lateral (3ª Proj.):"
      Height          =   195
      Left            =   4395
      TabIndex        =   4
      Top             =   3465
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vista Superior (1ª Proj.):"
      Height          =   195
      Left            =   7755
      TabIndex        =   3
      Top             =   3465
      Width           =   1665
   End
   Begin VB.Label lblPerspectiva 
      Caption         =   "Perspectiva:"
      Height          =   195
      Left            =   915
      TabIndex        =   2
      ToolTipText     =   "Teclas [ + ] e [ - ] alteram a distância da câmera."
      Top             =   105
      Width           =   6465
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Magnetismo As Boolean 'Indica se os pontos devem "grudar" na malha quadriculada
Private X_Ini As Integer, Y_Ini As Integer       'Usado no movimento da camera
Private Phi_Ini As GLfloat, Theta_Ini As GLfloat 'Idem
'Private Px As GLdouble, Py As GLdouble, Pz As GLdouble '(Px,Py,Pz)=Unico objeto até agora

Private Type Ferramenta
 IdImg As Integer
 Key As String
 TipText As String
End Type

Private Sub Form_Load()
 hDCPerspectiva = Me.picPerspectiva.hDC 'Identificador das ViewPort's
 hDCFrontal = Me.picFrontal.hDC
 hDCLateral = Me.picLateral.hDC
 hDCSuperior = Me.picSuperior.hDC
 hDCEpura = Me.picEpura.hDC
 
 Carrega_Ferramentas
 tbrFerramentas.Tag = tbrFerramentas.Buttons.Item(1).Key
 basGeometria.Inicializa
 
 Magnetismo = True 'habilita o magnetismo entre "ponto" e "grade"
 
 Call Inicializar_OpenGL 'Ajusta formato dos pixels, iluminação, matrizes de projeção...
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Posicionando = False
  picPerspectiva_Paint
  picEpura_Paint
  picSuperior_Paint
  picFrontal_Paint
  picLateral_Paint
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
 If Chr(KeyAscii) = "m" Or Chr(KeyAscii) = "M" Then Magnetismo = Not Magnetismo
 If Chr(KeyAscii) = "+" Then Ro = Ro - 1
 If Chr(KeyAscii) = "-" Then Ro = Ro + 1
 If KeyAscii = vbKeyEscape Then tbrFerramentas.Buttons(1).Value = tbrPressed: tbrFerramentas.Tag = "PONTEIRO"
 If Chr(KeyAscii) = "r" Or Chr(KeyAscii) = "R" Then
  picPerspectiva_Paint
  picEpura_Paint
  picSuperior_Paint
  picFrontal_Paint
  picLateral_Paint
 End If
 If Ro < 3 Then Ro = 3
 If Ro > 20 Then Ro = 20
  Cam_X = Ro * Sin(Phi * DEG) * Cos(Theta * DEG)
  Cam_Y = Ro * Sin(Phi * DEG) * Sin(Theta * DEG)
  Cam_Z = Ro * Cos(Phi * DEG)
  
  wglMakeCurrent hDCPerspectiva, hGLRCPerspectiva
  glMatrixMode GL_MODELVIEW
  glLoadIdentity
  gluLookAt Cam_X, Cam_Y, Cam_Z, 0, 0, 0, 0, 0, 1
  glMultMatrixf Troca_X_Y(0)
  
  picPerspectiva_Paint
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Call Finalizar_OpenGL
End Sub

Sub Carrega_Ferramentas()
 Const Arq_INI = "Tabela.ini"
 Dim imgX As ListImage
 Dim btnButton As Button
 
 Dim Qtd As Integer
 Dim FileNumber As Integer
 Dim N As Integer
 Dim F() As Ferramenta ' IdImg, Key e TipText
 
 FileNumber = FreeFile
 On Error GoTo ERRO
  Open App.Path & "\" & Arq_INI For Input As #FileNumber
 On Error GoTo 0
  
 N = 0
 'ReDim F(1 To N)
 
 With ilsFerramentas
   .ListImages.Clear
   .MaskColor = vbWhite
   Do
    N = N + 1
    ReDim Preserve F(1 To N)
    Input #FileNumber, F(N).IdImg, F(N).Key, F(N).TipText
    Set imgX = .ListImages. _
    Add(N, F(N).Key, LoadPicture(App.Path & "\IMG\" & Format(N, "00") & ".bmp"))
   Loop While Not EOF(FileNumber)
   
   Close #FileNumber
   Qtd = .ListImages.Count '=N-1
 End With
 
 With tbrFerramentas
   .Buttons.Clear
   .ImageList = ilsFerramentas
   For N = 1 To Qtd
    Set btnButton = .Buttons.Add(N, F(N).Key, "", tbrDefault, N)
    btnButton.ToolTipText = F(N).TipText
    btnButton.Style = tbrButtonGroup
   'If N > 3 Then btnButton.Enabled = False
   Next N
   .Buttons(1).Value = tbrPressed
 End With
 
 Exit Sub
ERRO:
 'If Err.Number = 53 Then
  'Err.Clear
  'Recup_Arquivo
  'Inicializa
 'Else
  Err.Raise Err.Number
 'End If
End Sub

Private Sub picPerspectiva_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Const VELOCIDADE = 0.5
 Dim dx As Integer, dy As Integer
 Dim winX As GLdouble, winY As GLdouble, winZ As GLdouble
  
 Dim M_ViewPort(0 To 3) As GLint
 Dim M_ModelView(0 To 15) As GLdouble
 Dim M_Projection(0 To 15) As GLdouble
 Dim Pos As GLdouble
 Dim Y_real As GLint
 Dim x1 As GLdouble, y1 As GLdouble, z1 As GLdouble
 Dim x0 As GLdouble, y0 As GLdouble, z0 As GLdouble
 Dim vx As GLdouble, vy As GLdouble, vz As GLdouble
 Dim px1 As GLdouble, py1 As GLdouble, pz1 As GLdouble
 Dim px2 As GLdouble, py2 As GLdouble, pz2 As GLdouble
 
 Dim Buf_Selec(0 To TAM_BUFER - 1) As GLuint
 Dim N_Hits As GLint
 
 Select Case Button
 Case 0, 1  '=Button: Mover ou definir objetos, conforme tbrFerramentas.Tag
  Select Case tbrFerramentas.Tag
  Case "PONTEIRO"
   'Obtem cópia da matriz de ViewPort, define qual será o Buffer e inicia modo de seleção
   wglMakeCurrent hDCPerspectiva, hGLRCPerspectiva
   glGetIntegerv GL_VIEWPORT, M_ViewPort(0)
   glSelectBuffer TAM_BUFER, Buf_Selec(0)
   glRenderMode GL_SELECT
   glInitNames
   glPushName 0 'valor arbitrário para iniciar a pilha
   
   'Define uma matriz para desenhar apenas próximo do mouse
   glMatrixMode GL_PROJECTION
   glPushMatrix
    glLoadIdentity
    gluPickMatrix X, M_ViewPort(3) - Y, 5#, 5#, M_ViewPort(0)
    gluPerspective 35!, fAspect, 1!, 100!
    
    glClear clrDepthBufferBit Or clrColorBufferBit
    basGeometria.Des_Objetos GL_SELECT, tbrFerramentas.Tag 'GL_RENDER
    glMatrixMode GL_PROJECTION 'As rotinas de desenho mudam para GL_MODELVIEW
   glPopMatrix
   glFlush
   'Envia dados sobre selecao para o basGeometria
   N_Hits = glRenderMode(GL_RENDER)
   
   'conferir se PRECISA INDICE (?) ao passar matriz
   
   picPerspectiva.ToolTipText = basGeometria.Avalia_Selecao(N_Hits, Buf_Selec)
   
  Case "PONTO"
   Estado_Teclas = Shift
   Posicionando = True
   wglMakeCurrent hDCPerspectiva, hGLRCPerspectiva
   glGetIntegerv GL_VIEWPORT, M_ViewPort(0)
   glGetDoublev GL_MODELVIEW_MATRIX, M_ModelView(0)
   glGetDoublev GL_PROJECTION_MATRIX, M_Projection(0)
   
   Y_real = M_ViewPort(3) - Y - 1
   gluUnProject X, Y_real, 0#, M_ModelView(0), M_Projection(0), M_ViewPort(0), x0, y0, z0
   gluUnProject X, Y_real, 1#, M_ModelView(0), M_Projection(0), M_ViewPort(0), x1, y1, z1
   vx = x1 - x0
   vy = y1 - y0
   vz = z1 - z0
   Select Case Estado_Teclas
   Case 0 'Mover sobre um PLANO HORIZONTAL
    If vz = 0 Then vz = z0 - P_Aux(2): MsgBox "vz=0"
    Pos = (P_Aux(2) - z0) / vz
   Case vbShiftMask 'Mover sobre um PLANO DE PERFIL
    If vx = 0 Then vx = x0 - P_Aux(0): MsgBox "vx=0"
    Pos = (P_Aux(0) - x0) / vx
   Case vbAltMask 'Mover sobre um PLANO FRONTAL
    If vy = 0 Then vy = y0 - P_Aux(1): MsgBox "vy=0"
    Pos = (P_Aux(1) - y0) / vy
    
   Case vbCtrlMask + 0 'Mover sobre uma RETA VERTICAL
     If vx = 0 Then vx = x0: MsgBox "vx=0"
     'P1 sobre um PLANO DE PERFIL
     Pos = (P_Aux(0) - x0) / vx
     If (Pos < 0 Or 1 < Pos) Then Exit Sub
     px1 = P_Aux(0): py1 = y0 + Pos * vy: pz1 = z0 + Pos * vz
     
     If X < 5 Then X = 5
     gluUnProject X - 5, Y_real, 0#, M_ModelView(0), M_Projection(0), M_ViewPort(0), x0, y0, z0
     gluUnProject X - 5, Y_real, 1#, M_ModelView(0), M_Projection(0), M_ViewPort(0), x1, y1, z1
     vx = x1 - x0
     vy = y1 - y0
     vz = z1 - z0
     If vx = 0 Then vx = x0: MsgBox "vx=0"
     Pos = (P_Aux(0) - x0) / vx
     If (Pos < 0 Or 1 < Pos) Then Exit Sub
     px2 = P_Aux(0):  py2 = y0 + Pos * vy:  pz2 = z0 + Pos * vz
     'P_Aux(0) = P_Aux(0)
     'P_Aux(1) = P_Aux(1)
     If py2 <> py1 Then P_Aux(2) = pz1 + (P_Aux(1) - py1) * (pz2 - pz1) / (py2 - py1)
   Case vbCtrlMask + vbShiftMask 'Mover sobre uma RETA FRONTO-HORIZONTAL
     If vy = 0 Then vy = y0 - P_Aux(1): MsgBox "vy=0"
     Pos = (P_Aux(1) - y0) / vy 'P1 sobre um PLANO FRONTAL
     If (Pos < 0 Or 1 < Pos) Then Exit Sub
     px1 = x0 + Pos * vx:   py1 = P_Aux(1):   pz1 = z0 + Pos * vz
     
     If Y_real < 5 Then Y_real = 5
     gluUnProject X, Y_real - 5, 0#, M_ModelView(0), M_Projection(0), M_ViewPort(0), x0, y0, z0
     gluUnProject X, Y_real - 5, 1#, M_ModelView(0), M_Projection(0), M_ViewPort(0), x1, y1, z1
     vx = x1 - x0
     vy = y1 - y0
     vz = z1 - z0
     If vy = 0 Then vy = y0 - P_Aux(1): MsgBox "vy=0"
     Pos = (P_Aux(1) - y0) / vy 'P1 sobre um PLANO FRONTAL
     If (Pos < 0 Or 1 < Pos) Then Exit Sub
     px2 = x0 + Pos * vx:   py2 = P_Aux(1):   pz2 = z0 + Pos * vz
     If pz2 <> pz1 Then P_Aux(0) = px1 + (P_Aux(2) - pz1) * (px2 - px1) / (pz2 - pz1)
     'P_Aux(1) = P_Aux(1)
     'P_Aux(2) = P_Aux(2)
   Case vbCtrlMask + vbAltMask 'Mover sobre uma RETA DE TOPO
     If vx = 0 Then vx = x0: MsgBox "vx=0"
     Pos = (P_Aux(0) - x0) / vx 'P1 sobre um PLANO DE PERFIL
     If (Pos < 0 Or 1 < Pos) Then Exit Sub
     px1 = P_Aux(0): py1 = y0 + Pos * vy: pz1 = z0 + Pos * vz
     
     If Y_real < 5 Then Y_real = 5
     gluUnProject X, Y_real - 5, 0#, M_ModelView(0), M_Projection(0), M_ViewPort(0), x0, y0, z0
     gluUnProject X, Y_real - 5, 1#, M_ModelView(0), M_Projection(0), M_ViewPort(0), x1, y1, z1
     vx = x1 - x0
     vy = y1 - y0
     vz = z1 - z0
     If vx = 0 Then vx = x0: MsgBox "vx=0"
     Pos = (P_Aux(0) - x0) / vx 'P1 sobre um PLANO DE PERFIL
     If (Pos < 0 Or 1 < Pos) Then Exit Sub
     px2 = P_Aux(0):  py2 = y0 + Pos * vy:  pz2 = z0 + Pos * vz
     'P_Aux(0) = P_Aux(0)
     If pz2 <> pz1 Then P_Aux(1) = py1 + (P_Aux(2) - pz1) * (py2 - py1) / (pz2 - pz1)
     'P_Aux(2) = P_Aux(2)
   End Select
   If (Estado_Teclas = 0 Or Estado_Teclas = vbShiftMask Or Estado_Teclas = vbAltMask) Then
     If (Pos < 0 Or 1 < Pos) Then Exit Sub
     'Calcula a interseção do raio projetante com o plano escolhido
     P_Aux(0) = x0 + Pos * vx
     P_Aux(1) = y0 + Pos * vy
     P_Aux(2) = z0 + Pos * vz
   End If
   
   If Magnetismo Then
    P_Aux(0) = Round(P_Aux(0))
    P_Aux(1) = Round(P_Aux(1))
    P_Aux(2) = Round(P_Aux(2))
   End If

  Case "SEGMENTO"
  
  End Select
  picPerspectiva_Paint
  picEpura_Paint
  picSuperior_Paint
  picFrontal_Paint
  picLateral_Paint
  
 Case 2 '=Button: Mover camera (botao direito)
  dx = VELOCIDADE * (X - X_Ini)
  dy = VELOCIDADE * (Y - Y_Ini)

  Phi = Phi_Ini - dy
  Theta = Theta_Ini - dx
  Phi = IIf(Phi <= 0, 0.0001, Phi): Phi = IIf(Phi > 180, 180, Phi)
  Theta = IIf(Theta <= -180, Theta + 360, Theta): Theta = IIf(Theta > 180, Theta - 360, Theta)
  
  Cam_X = Ro * Sin(Phi * DEG) * Cos(Theta * DEG)
  Cam_Y = Ro * Sin(Phi * DEG) * Sin(Theta * DEG)
  Cam_Z = Ro * Cos(Phi * DEG)
  
  wglMakeCurrent hDCPerspectiva, hGLRCPerspectiva
  glMatrixMode GL_MODELVIEW
  glLoadIdentity
  gluLookAt Cam_X, Cam_Y, Cam_Z, 0, 0, 0, 0, 0, 1
  glMultMatrixf Troca_X_Y(0)
  
  picPerspectiva_Paint
 End Select
End Sub
Private Sub picPerspectiva_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 If Button = 2 Then picPerspectiva.MousePointer = 99
 Select Case UCase(tbrFerramentas.Tag)
  Case "PONTEIRO"
   
  Case "PONTO"
   
  'Case "SEGMENTO"
 End Select
 
 X_Ini = X: Y_Ini = Y
 Phi_Ini = Phi:  Theta_Ini = Theta
End Sub
Private Sub picPerspectiva_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
 Case 1
  Estado_Teclas = Shift
  If tbrFerramentas.Tag = "PONTO" And basGeometria.Qtd_Obj < basGeometria.MAX_OBJETOS Then
   basGeometria.Qtd_Obj = basGeometria.Qtd_Obj + 1
   ReDim Preserve basGeometria.Obj(1 To basGeometria.Qtd_Obj)
   basGeometria.Obj(Qtd_Obj).Coord(0) = P_Aux(0)
   basGeometria.Obj(Qtd_Obj).Coord(1) = P_Aux(1)
   basGeometria.Obj(Qtd_Obj).Coord(2) = P_Aux(2)
   'P_Aux(0) = 0: P_Aux(1) = 0: P_Aux(2) = 0
  End If
  
  picPerspectiva_Paint
  picEpura_Paint
  picSuperior_Paint
  picFrontal_Paint
  picLateral_Paint
 Case 2
  If Button = 2 Then picPerspectiva.MousePointer = 0
 End Select
End Sub
Private Sub picPerspectiva_Paint()

 wglMakeCurrent hDCPerspectiva, hGLRCPerspectiva
 glClear clrColorBufferBit Or clrDepthBufferBit
 
 basGeometria.Des_Eixos
 basGeometria.Des_Objetos GL_RENDER, tbrFerramentas.Tag 'GL_SELECT
 
 Select Case UCase(tbrFerramentas.Tag)
  'Case "PONTEIRO"
  Case "PONTO"
   If Posicionando Then Des_Plano (Estado_Teclas)
  'Case "SEGMENTO"
 End Select
 SwapBuffers hDCPerspectiva
End Sub
Private Sub picSuperior_Paint()

 wglMakeCurrent hDCSuperior, hGLRCSuperior
 glClear clrColorBufferBit Or clrDepthBufferBit
 Des_Eixos
 basGeometria.Des_Objetos GL_RENDER, tbrFerramentas.Tag 'GL_SELECT
 SwapBuffers hDCSuperior
End Sub
Private Sub picFrontal_Paint()

 wglMakeCurrent hDCFrontal, hGLRCFrontal
 glClear clrColorBufferBit Or clrDepthBufferBit
 Des_Eixos
 basGeometria.Des_Objetos GL_RENDER, tbrFerramentas.Tag 'GL_SELECT
 SwapBuffers hDCFrontal
End Sub
Private Sub picLateral_Paint()

 wglMakeCurrent hDCLateral, hGLRCLateral
 glClear clrColorBufferBit Or clrDepthBufferBit
 Des_Eixos
 basGeometria.Des_Objetos GL_RENDER, tbrFerramentas.Tag 'GL_SELECT
 SwapBuffers hDCLateral
End Sub
Private Sub picEpura_Paint()

 wglMakeCurrent hDCEpura, hGLRCEpura
 glClear clrColorBufferBit Or clrDepthBufferBit
 
 Des_LT
 
 'Posicionado por padrão para vista superior
 'Desenha 1ª projeção
 basGeometria.Des_Objetos GL_RENDER, tbrFerramentas.Tag 'GL_SELECT
 
 'Reposiciona para vista frontal
 glMatrixMode GL_MODELVIEW
  glLoadIdentity
  glMultMatrixf Troca_X_Y(0)
  glRotatef 90, 0#, 0#, 1#
  glRotatef 90, 1#, 0#, 0#
 'Desenha 2ª projeção
 basGeometria.Des_Objetos GL_RENDER, tbrFerramentas.Tag 'GL_SELECT
 
 SwapBuffers hDCEpura
End Sub
Private Sub tbrFerramentas_ButtonClick(ByVal Button As MSComctlLib.Button)
'Convenção: Tag guarda um nome igual aos da enumeração pública de tipos dos objetos
 tbrFerramentas.Tag = Button.Key
 'If tbrFerramentas.Tag <> "PONTO" Then MsgBox "NAO SELECIONANDO!": Posicionando = False
 
End Sub
