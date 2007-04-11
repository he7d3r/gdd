VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Exemplo - Integrando Vb e OpenGl"
   ClientHeight    =   4050
   ClientLeft      =   255
   ClientTop       =   540
   ClientWidth     =   6540
   Icon            =   "Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   270
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   436
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ilsFerramentas 
      Left            =   5445
      Top             =   900
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
            Picture         =   "Main.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":1040
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrFerramentas 
      Align           =   3  'Align Left
      Height          =   4050
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   7144
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
      MousePointer    =   99
      MouseIcon       =   "Main.frx":1D76
   End
   Begin VB.PictureBox picVista 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   1500
      Index           =   5
      Left            =   915
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   98
      TabIndex        =   9
      Top             =   2400
      Width           =   1500
   End
   Begin VB.PictureBox picVista 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1500
      Index           =   3
      Left            =   2610
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   98
      TabIndex        =   7
      Top             =   2400
      Width           =   1500
   End
   Begin VB.PictureBox picVista 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1500
      Index           =   4
      Left            =   4380
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   98
      TabIndex        =   6
      Top             =   2400
      Width           =   1500
   End
   Begin VB.PictureBox picVista 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   1500
      Index           =   2
      Left            =   3555
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   98
      TabIndex        =   1
      Top             =   450
      Width           =   1500
   End
   Begin VB.PictureBox picVista 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1500
      Index           =   1
      Left            =   915
      MouseIcon       =   "Main.frx":2090
      ScaleHeight     =   98
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   98
      TabIndex        =   0
      ToolTipText     =   "Botão direito: Mover camera."
      Top             =   450
      Width           =   1500
   End
   Begin VB.Label lblVista 
      AutoSize        =   -1  'True
      Caption         =   "Épura (1ª e 2ª Proj.):"
      Height          =   195
      Index           =   5
      Left            =   915
      TabIndex        =   8
      Top             =   2175
      Width           =   1440
   End
   Begin VB.Label lblVista 
      AutoSize        =   -1  'True
      Caption         =   "Vista Frontal (2ª Proj.):"
      Height          =   195
      Index           =   2
      Left            =   3510
      TabIndex        =   5
      Top             =   150
      Width           =   1560
   End
   Begin VB.Label lblVista 
      AutoSize        =   -1  'True
      Caption         =   "Vista Lateral (3ª Proj.):"
      Height          =   195
      Index           =   3
      Left            =   2640
      TabIndex        =   4
      Top             =   2175
      Width           =   1560
   End
   Begin VB.Label lblVista 
      AutoSize        =   -1  'True
      Caption         =   "Vista Superior (1ª Proj.):"
      Height          =   195
      Index           =   4
      Left            =   4380
      TabIndex        =   3
      Top             =   2175
      Width           =   1665
   End
   Begin VB.Label lblVista 
      AutoSize        =   -1  'True
      Caption         =   "Perspectiva:"
      Height          =   195
      Index           =   1
      Left            =   1035
      TabIndex        =   2
      ToolTipText     =   "Teclas [ + ] e [ - ] alteram a distância da câmera."
      Top             =   150
      Width           =   885
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private hDC_V(1 To 5) As Long                      'Device Contexts de cada tela no Doc
Private hGLRC_V(1 To 5) As Long                    'GL Rendering Context de cada tela no Doc
Private M_ViewPort(0 To 3) As GLint
Private M_ModelView(0 To 15) As GLdouble
Private M_Projection(0 To 15) As GLdouble

Public X_Ini As Integer, Y_Ini As Integer                'Usado no movimento da camera
Public Cam_X As Single, Cam_Y As Single, Cam_Z As Single 'Coord. cartesianas da câmera

Public Phi_Ini As GLfloat, Theta_Ini As GLfloat          'Usado no movimento da camera
Public Phi As GLfloat, Theta As GLfloat, Ro As GLfloat   'Coord. esféricas da câmera

Public N_Sel As Integer             'Em geral = Ubound(Obj_Sel)
Public ObjApontado As Long          'Indica o índice do objeto sob o mouse
Public Posicionando As Boolean      'Indica se está sendo posicionado um ponto no espaço
Private Clicou_Sobre_Objeto As Boolean          'Indica se um objeto já definido será reposicionado
Public fAspect As GLfloat           'Proporção entre os lados da picPerspectiva

Property Get hDC_Vista(ByVal Index As Vista) As Long
   If Index > UBound(hDC_V) Then ErroFatal "Não existe uma Vista com índice " & Index & "!"
   hDC_Vista = hDC_V(Index)
End Property
Property Get hGLRC_Vista(ByVal Index As Vista) As Long
   If Index > UBound(hGLRC_V) Then ErroFatal "Não existe uma Vista com índice " & Index & "!"
   hGLRC_Vista = hGLRC_V(Index)
End Property
Property Let hGLRC_Vista(ByVal Index As Vista, ByVal v As Long)
   If Index > UBound(hGLRC_V) Then ErroFatal "Não existe uma Vista com índice " & Index & "!"
   hGLRC_V(Index) = v
End Property

Public Sub Redesenhar_Todos()
   Dim v As Vista
   
   For v = PERSPECTIVA To EPURA
      picVista_Paint (v)
   Next v
End Sub

Private Sub Form_Load()
   Dim v As Vista
   
   For v = PERSPECTIVA To EPURA
      hDC_V(v) = Me.picVista(v).hDC   'Identificador das ViewPort's
   Next v
   N_Sel = 0
   'P_Aux(3) = 1
   
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Posicionando = False
   Redesenhar_Todos
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Const MAX_RO = 50
   Const MIN_RO = 3
   If Chr(KeyAscii) = "+" Then Ro = Ro - 1
   If Chr(KeyAscii) = "-" Then Ro = Ro + 1
   If KeyAscii = vbKeyEscape Then
      tbrFerramentas.Buttons(1).Value = tbrPressed: tbrFerramentas.Tag = "PONTEIRO"
      ReDim Obj_Aux(1 To 1)
      Obj_Aux(1).Coord(3) = 1
      Posicionando = False
   End If
   'O ALT NÃO ESTÁ COM PROBLEMA.
   'VOCÊ SELECIONOU O MENU!
   If Chr(KeyAscii) = "r" Or Chr(KeyAscii) = "R" Then
    Redesenhar_Todos
   End If
   
   If Ro < MIN_RO Then Ro = MIN_RO
   If Ro > MAX_RO Then Ro = MAX_RO
   Cam_X = Ro * Sin(Phi * DEG) * Cos(Theta * DEG)
   Cam_Y = Ro * Sin(Phi * DEG) * Sin(Theta * DEG)
   Cam_Z = Ro * Cos(Phi * DEG)
   
   wglMakeCurrent hDC_Vista(PERSPECTIVA), hGLRC_Vista(PERSPECTIVA)
   glFogf GL_FOG_START, Ro
   glMatrixMode GL_MODELVIEW
    glLoadIdentity
    gluLookAt Cam_X, Cam_Y, Cam_Z, 0, 0, 0, 0, 0, 1
    glMultMatrixf Troca_X_Y(0)
  
   picVista_Paint (PERSPECTIVA)
   Erro = glGetError: If Erro <> glerrNoError Then ErroFatal Erro
End Sub

Private Sub Form_Resize()
   Const ESP = 25
   Dim Tam As Single
   Dim Barra As Single
   Dim Linha(1 To 2) As Single
   Dim Coluna(1 To 3) As Single
   Dim l As Single, a As Single
   Dim v As Vista
    
   Barra = tbrFerramentas.Width
   a = (Me.ScaleHeight - 3 * ESP) / 2
   l = (Me.ScaleWidth - 4 * ESP - Barra) / 3
   Tam = IIf(a < l, a, l)
   
   If Tam <= 0 Then Exit Sub
   Linha(1) = ESP:   Linha(2) = 2 * ESP + Tam
   Coluna(1) = Barra + ESP: Coluna(2) = Barra + 2 * ESP + Tam: Coluna(3) = Barra + 3 * ESP + 2 * Tam
   picVista(PERSPECTIVA).Move Coluna(1), Linha(1), 2 * Tam + ESP, Tam
       picVista(FRONTAL).Move Coluna(3), Linha(1), Tam, Tam
         picVista(EPURA).Move Coluna(1), Linha(2), Tam, Tam
       picVista(LATERAL).Move Coluna(2), Linha(2), Tam, Tam
      picVista(SUPERIOR).Move Coluna(3), Linha(2), Tam, Tam
   
   For v = PERSPECTIVA To EPURA
      lblVista(v).Move picVista(v).Left, picVista(v).Top - 15
   Next v
   
   For v = PERSPECTIVA To EPURA
      wglMakeCurrent hDC_Vista(v), hGLRC_Vista(v)
      With picVista(v)
         l = .ScaleWidth: a = .ScaleHeight
      End With
      glViewport 0, 0, l, a
      If v = PERSPECTIVA Then
         If a > 0 Then
          fAspect = l / a
         Else
          fAspect = 0
         End If
         glMatrixMode GL_PROJECTION
          glLoadIdentity
          gluPerspective 35!, fAspect, DIST_MIN_CENA, DIST_MAX_CENA
         glMatrixMode GL_MODELVIEW
      End If
   Next v
   Redesenhar_Todos
   Erro = glGetError: If Erro <> glerrNoError Then ErroFatal Erro
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Doc(Me.Tag).Deletado = True
 Call Finalizar_OpenGL(Me.Tag)
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If Doc(Me.Tag).Alterado Then MsgBox "O documento foi alterado, mas não foi salvo."
End Sub
Private Sub Pos_Ponto(ByVal v As Vista, _
                     ByVal X As Single, ByVal Y As Single, _
                     ByVal Perpendicular As Boolean, _
                     ByRef Aux() As GLdouble)

   Dim Pos As GLdouble
   Dim Y_real As GLint
   Dim x1 As GLdouble, y1 As GLdouble, z1 As GLdouble
   Dim x0 As GLdouble, y0 As GLdouble, z0 As GLdouble
   Dim vx As GLdouble, vy As GLdouble, vz As GLdouble
   Dim px1 As GLdouble, py1 As GLdouble, pz1 As GLdouble
   Dim px2 As GLdouble, py2 As GLdouble, pz2 As GLdouble
   Dim Pt(0 To 3) As GLdouble
      
   Pt(0) = Aux(0) / Aux(3)
   Pt(1) = Aux(1) / Aux(3)
   Pt(2) = Aux(2) / Aux(3)

   wglMakeCurrent hDC_Vista(v), hGLRC_Vista(v)
   glGetIntegerv GL_VIEWPORT, M_ViewPort(0)
   glGetDoublev GL_MODELVIEW_MATRIX, M_ModelView(0)
   glGetDoublev GL_PROJECTION_MATRIX, M_Projection(0)
   
   Y_real = M_ViewPort(3) - Y - 1
   gluUnProject X, Y_real, 0#, M_ModelView(0), M_Projection(0), M_ViewPort(0), x0, y0, z0
   gluUnProject X, Y_real, 1#, M_ModelView(0), M_Projection(0), M_ViewPort(0), x1, y1, z1
   vx = x1 - x0
   vy = y1 - y0
   vz = z1 - z0
   
   If Not Perpendicular Then
      Select Case Sobre_Plano
      Case PL_HORIZONTAL
         If vz = 0 Then vz = z0 - Pt(2): MsgBox "vz=0"
         Pos = (Pt(2) - z0) / vz
      Case PL_PERFIL
         If vx = 0 Then vx = x0 - Pt(0): MsgBox "vx=0"
         Pos = (Pt(0) - x0) / vx
      Case PL_FRONTAL
         If vy = 0 Then vy = y0 - Pt(1): MsgBox "vy=0"
         Pos = (Pt(1) - y0) / vy
      End Select
      If (Pos < 0 Or 1 < Pos) Then
         Posicionando = False
         frmMDIGeo3d.staInfo.Panels(1).Text = ""
         Redesenhar_Todos
         Exit Sub
      Else
         'Calcula a interseção do raio projetante com o plano escolhido
         Pt(0) = x0 + Pos * vx
         Pt(1) = y0 + Pos * vy
         Pt(2) = z0 + Pos * vz
      End If
   Else
      Select Case Sobre_Plano
      Case PL_HORIZONTAL 'Mover sobre uma RETA VERTICAL
         If vx = 0 Then vx = x0: MsgBox "vx=0"
         'P1 sobre um PLANO DE PERFIL
         Pos = (Pt(0) - x0) / vx
         'If (Pos < 0 Or 1 < Pos) Then Exit Sub
         px1 = Pt(0): py1 = y0 + Pos * vy: pz1 = z0 + Pos * vz
         
         If X < 5 Then X = 5
         gluUnProject X - 5, Y_real, 0#, M_ModelView(0), M_Projection(0), M_ViewPort(0), x0, y0, z0
         gluUnProject X - 5, Y_real, 1#, M_ModelView(0), M_Projection(0), M_ViewPort(0), x1, y1, z1
         vx = x1 - x0
         vy = y1 - y0
         vz = z1 - z0
         If vx = 0 Then vx = x0: MsgBox "vx=0"
         Pos = (Pt(0) - x0) / vx
         'If (Pos < 0 Or 1 < Pos) Then Exit Sub
         px2 = Pt(0):  py2 = y0 + Pos * vy:  pz2 = z0 + Pos * vz
         'Pt(0) = Pt(0)
         'Pt(1) = Pt(1)
         If py2 <> py1 Then Pt(2) = pz1 + (Pt(1) - py1) * (pz2 - pz1) / (py2 - py1)
      Case PL_PERFIL 'Mover sobre uma RETA FRONTO-HORIZONTAL
         If vy = 0 Then vy = y0 - Pt(1): MsgBox "vy=0"
         Pos = (Pt(1) - y0) / vy 'P1 sobre um PLANO FRONTAL
         If (Pos < 0 Or 1 < Pos) Then Exit Sub
         px1 = x0 + Pos * vx:   py1 = Pt(1):   pz1 = z0 + Pos * vz
         
         If Y_real < 5 Then Y_real = 5
         gluUnProject X, Y_real - 5, 0#, M_ModelView(0), M_Projection(0), M_ViewPort(0), x0, y0, z0
         gluUnProject X, Y_real - 5, 1#, M_ModelView(0), M_Projection(0), M_ViewPort(0), x1, y1, z1
         vx = x1 - x0
         vy = y1 - y0
         vz = z1 - z0
         If vy = 0 Then vy = y0 - Pt(1): MsgBox "vy=0"
         Pos = (Pt(1) - y0) / vy 'P1 sobre um PLANO FRONTAL
         If (Pos < 0 Or 1 < Pos) Then Exit Sub
         px2 = x0 + Pos * vx:   py2 = Pt(1):   pz2 = z0 + Pos * vz
         If pz2 <> pz1 Then Pt(0) = px1 + (Pt(2) - pz1) * (px2 - px1) / (pz2 - pz1)
         'Pt(1) = Pt(1)
         'Pt(2) = Pt(2)
      Case PL_FRONTAL 'Mover sobre uma RETA DE TOPO
         If vx = 0 Then vx = x0: MsgBox "vx=0"
         Pos = (Pt(0) - x0) / vx 'P1 sobre um PLANO DE PERFIL
         If (Pos < 0 Or 1 < Pos) Then Exit Sub
         px1 = Pt(0): py1 = y0 + Pos * vy: pz1 = z0 + Pos * vz
         
         If Y_real < 5 Then Y_real = 5
         gluUnProject X, Y_real - 5, 0#, M_ModelView(0), M_Projection(0), M_ViewPort(0), x0, y0, z0
         gluUnProject X, Y_real - 5, 1#, M_ModelView(0), M_Projection(0), M_ViewPort(0), x1, y1, z1
         vx = x1 - x0
         vy = y1 - y0
         vz = z1 - z0
         If vx = 0 Then vx = x0: MsgBox "vx=0"
         Pos = (Pt(0) - x0) / vx 'P1 sobre um PLANO DE PERFIL
         If (Pos < 0 Or 1 < Pos) Then Exit Sub
         px2 = Pt(0):  py2 = y0 + Pos * vy:  pz2 = z0 + Pos * vz
         'Pt(0) = Pt(0)
         If pz2 <> pz1 Then Pt(1) = py1 + (Pt(2) - pz1) * (py2 - py1) / (pz2 - pz1)
         'Pt(2) = Pt(2)
      End Select
   End If
  
   If frmMDIGeo3d.mnuEditarMagnetismo.Checked Then
      Aux(0) = Round(Pt(0))
      Aux(1) = Round(Pt(1))
      Aux(2) = Round(Pt(2))
   Else
      Aux(0) = Pt(0)
      Aux(1) = Pt(1)
      Aux(2) = Pt(2)
   End If
   Aux(3) = 1
   
   frmMDIGeo3d.staInfo.Panels(1).Text = "Posição atual: [ " _
                                       & Format(Aux(0), "0.0") & " ;  " _
                                       & Format(Aux(1), "0.0") & " ;  " _
                                       & Format(Aux(2), "0.0") & "]"
   
   Erro = glGetError: If Erro <> glerrNoError Then ErroFatal Erro
End Sub
Function Listar_Objetos_Sob(ByVal v As Vista, _
                           ByVal X As Single, ByVal Y As Single, _
                           ByRef B() As GLuint) As GLint
   
   'Obtem cópia da matriz de ViewPort, define qual será o Buffer e inicia modo de seleção
   wglMakeCurrent hDC_Vista(v), hGLRC_Vista(v)
   glGetIntegerv GL_VIEWPORT, M_ViewPort(0)
   glSelectBuffer TAM_BUFER, B(0)
   glRenderMode GL_SELECT
   glInitNames
   glPushName 0 'valor arbitrário para iniciar a pilha
   
   'Define uma matriz para desenhar apenas próximo do mouse
   glMatrixMode GL_PROJECTION
   glPushMatrix
    glLoadIdentity
    gluPickMatrix X, M_ViewPort(3) - Y, PROX, PROX, M_ViewPort(0)
    gluPerspective 35!, fAspect, DIST_MIN_CENA, DIST_MAX_CENA
    
    glClear clrDepthBufferBit Or clrColorBufferBit
    Des_Objetos Me.Tag, GL_SELECT, Doc(Me.Tag).Obj 'GL_RENDER
    glMatrixMode GL_PROJECTION 'As rotinas de desenho mudam para GL_MODELVIEW
   glPopMatrix
   glFlush
   
   Listar_Objetos_Sob = glRenderMode(GL_RENDER)
   
   Erro = glGetError: If Erro <> glerrNoError Then ErroFatal Erro
End Function
Private Sub picVista_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Select Case Index
   Case PERSPECTIVA
      'Select Case UCase(tbrFerramentas.Tag)
      'Case "PONTEIRO"
      'Case "PONTO"
      'Case "SEGMENTO"
      'End Select
      X_Ini = X: Y_Ini = Y
      Select Case Button
      Case vbRightButton
         'picVista(PERSPECTIVA).MousePointer = 99
         Phi_Ini = Phi:  Theta_Ini = Theta
      Case vbLeftButton
         Clicou_Sobre_Objeto = (ObjApontado > 0)
'         Me.Caption = Clicou_Sobre_Objeto
      End Select
   End Select
End Sub
Private Sub picVista_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Const VELOCIDADE = 0.5
   Dim dx As Integer, dy As Integer
   Dim Buf_Selec(0 To TAM_BUFER - 1) As GLuint
   Dim N_Hits As GLint
   Dim vt As Vista
   
   vt = Index
   Select Case Index
   Case PERSPECTIVA
      If ((Button And vbRightButton) = vbRightButton) Then
         picVista(PERSPECTIVA).MousePointer = 99
         Posicionando = False
         ObjApontado = 0
         dx = VELOCIDADE * (X - X_Ini)
         dy = VELOCIDADE * (Y - Y_Ini)
         
         Phi = Phi_Ini - dy
         Theta = Theta_Ini - dx
         Phi = IIf(Phi <= 0, ZERO, Phi): Phi = IIf(Phi > 180, 180, Phi)
         Theta = IIf(Theta <= -180, Theta + 360, Theta): Theta = IIf(Theta > 180, Theta - 360, Theta)
         
         Cam_X = Ro * Sin(Phi * DEG) * Cos(Theta * DEG)
         Cam_Y = Ro * Sin(Phi * DEG) * Sin(Theta * DEG)
         Cam_Z = Ro * Cos(Phi * DEG)
         
         frmMDIGeo3d.staInfo.Panels(2).Text = "CÂMERA:  ( " _
                                             & Format(Cam_X, "0.0") & " ;  " _
                                             & Format(Cam_Y, "0.0") & " ;  " _
                                             & Format(Cam_Z, "0.0") & ")cart     ( " _
                                             & Format(Phi, "0") & " ;  " _
                                             & Format(Theta, "#0") & " ;  " _
                                             & Format(Ro, "#0") & ")esf"
         
         wglMakeCurrent hDC_Vista(PERSPECTIVA), hGLRC_Vista(PERSPECTIVA)
         glMatrixMode GL_MODELVIEW
         glLoadIdentity
         gluLookAt Cam_X, Cam_Y, Cam_Z, 0, 0, 0, 0, 0, 1
         glMultMatrixf Troca_X_Y(0)
         
         picVista_Paint (PERSPECTIVA)
         Exit Sub
      End If

      Select Case tbrFerramentas.Tag
      Case "PONTEIRO"
         Select Case Button
         Case 0
            'aponte
            N_Hits = Listar_Objetos_Sob(vt, X, Y, Buf_Selec)
            picVista(PERSPECTIVA).ToolTipText = Aponta_Primeiro_Objeto(Me.Tag, N_Hits, Buf_Selec)
         Case vbLeftButton
            If Clicou_Sobre_Objeto Then 'O mouse estava sobre 'ObjApontado' ao clicar
               'mova ObjApontado
               Posicionando = True
               Pos_Ponto vt, X, Y, Shift = vbCtrlMask, Doc(Me.Tag).Obj(ObjApontado).Coord
            Else
               'desenhe retângulo de seleção
            End If
         End Select
      Case "PONTO"
         Select Case Button
         Case 0
            'aponte
            N_Hits = Listar_Objetos_Sob(vt, X, Y, Buf_Selec)
            picVista(PERSPECTIVA).ToolTipText = Aponta_Primeiro_Objeto(Me.Tag, N_Hits, Buf_Selec)
            'posicione p_aux
            Posicionando = True
            If ObjApontado <= 0 Then
               Pos_Ponto vt, X, Y, Shift = vbCtrlMask, Obj_Aux(1).Coord 'P_Aux
            End If
         Case vbLeftButton
            If Clicou_Sobre_Objeto Then
               'mova ObjApontado
               Posicionando = True
               Pos_Ponto vt, X, Y, Shift = vbCtrlMask, Doc(Me.Tag).Obj(ObjApontado).Coord
            Else
               'aponte
               N_Hits = Listar_Objetos_Sob(vt, X, Y, Buf_Selec)
               picVista(PERSPECTIVA).ToolTipText = Aponta_Primeiro_Objeto(Me.Tag, N_Hits, Buf_Selec)
               'arraste p_aux
               Posicionando = True
               Pos_Ponto vt, X, Y, Shift = vbCtrlMask, Obj_Aux(1).Coord 'P_Aux
            End If
         End Select
      Case "SEGMENTO"
         Select Case Button
         Case 0
            'aponte
            N_Hits = Listar_Objetos_Sob(vt, X, Y, Buf_Selec)
            picVista(PERSPECTIVA).ToolTipText = Aponta_Primeiro_Objeto(Me.Tag, N_Hits, Buf_Selec)
            'posicione p_aux
            Posicionando = True
            If ObjApontado <= 0 Then
               If UBound(Obj_Aux) = 1 Then
                  Pos_Ponto vt, X, Y, Shift = vbCtrlMask, Obj_Aux(1).Coord
               Else
                  Pos_Ponto vt, X, Y, Shift = vbCtrlMask, Obj_Aux(2).Coord
               End If
            Else
               If UBound(Obj_Aux) > 1 Then
                  With Doc(Me.Tag).Obj(ObjApontado)
                     Obj_Aux(2).Coord(0) = .Coord(0)
                     Obj_Aux(2).Coord(1) = .Coord(1)
                     Obj_Aux(2).Coord(2) = .Coord(2)
                     Obj_Aux(2).Coord(3) = .Coord(3)
                  End With
               End If
            End If
         Case vbLeftButton
            If Clicou_Sobre_Objeto Then
               'mova ObjApontado
               If Abs(X - X_Ini) > PROX Or Abs(Y - Y_Ini) > PROX Then
                  Posicionando = True
                  Pos_Ponto vt, X, Y, Shift = vbCtrlMask, Doc(Me.Tag).Obj(ObjApontado).Coord
                  If UBound(Obj_Aux) > 1 Then
                     With Doc(Me.Tag).Obj(ObjApontado)
                        Obj_Aux(2).Coord(0) = .Coord(0)
                        Obj_Aux(2).Coord(1) = .Coord(1)
                        Obj_Aux(2).Coord(2) = .Coord(2)
                        Obj_Aux(2).Coord(3) = .Coord(3)
                     End With
                  End If
               End If
            Else
               'aponte
               N_Hits = Listar_Objetos_Sob(vt, X, Y, Buf_Selec)
               picVista(PERSPECTIVA).ToolTipText = Aponta_Primeiro_Objeto(Me.Tag, N_Hits, Buf_Selec)
               'arraste p_aux
               Posicionando = True
               If UBound(Obj_Aux) = 1 Then
                  Pos_Ponto vt, X, Y, Shift = vbCtrlMask, Obj_Aux(1).Coord 'P_Aux
               Else
                  Pos_Ponto vt, X, Y, Shift = vbCtrlMask, Obj_Aux(2).Coord 'P_Aux
               End If
            End If
         End Select
      End Select
      
      Redesenhar_Todos
      
   Case FRONTAL
   Case LATERAL
   Case EPURA
   Case SUPERIOR
   
   End Select
   
   Erro = glGetError: If Erro <> glerrNoError Then ErroFatal Erro
End Sub

Private Sub picVista_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim i As Integer
   Dim Id As Long
   'Dim N_Obj As Integer 'Em geral = Ubound(Doc(me.tag).Obj)
         
   Select Case Index
   Case PERSPECTIVA
      
      'N_Obj = UBound(Doc(Me.Tag).Obj)
      
      Select Case Button
      Case vbLeftButton
         Select Case tbrFerramentas.Tag
         Case "PONTEIRO"
            Posicionando = False
            If ObjApontado <= 0 Then
               If Shift = 0 Then Marcar_Todos Me.Tag, False
            Else
               With Doc(Me.Tag)
               'N_Sel = UBound(.Obj_Sel)
               If Shift = 0 Then
                  If N_Sel = 0 Then 'Ainda não há obj. selecionado
                     N_Sel = 1
                     .Obj(ObjApontado).Selec = 1
                     .Obj_Sel(1) = ObjApontado
                  Else
                     For i = 1 To N_Sel
                      .Obj(.Obj_Sel(i)).Selec = 0 'duplicado
                     Next i
                     If N_Sel > 1 Or .Obj(ObjApontado).Selec = 0 Then _
                                     .Obj(ObjApontado).Selec = 1
                     N_Sel = IIf((.Obj(ObjApontado).Selec = 0) And (N_Sel = 1), 0, 1)
                     
                     ReDim .Obj_Sel(1 To 1)
                     If N_Sel > 0 Then .Obj_Sel(1) = ObjApontado
                  End If
                
               Else 'not Shift = 0
                  If .Obj(ObjApontado).Selec > 0 Then
                     N_Sel = N_Sel - 1
                     For i = .Obj(ObjApontado).Selec To N_Sel
                      .Obj_Sel(i) = .Obj_Sel(i + 1)
                      .Obj(.Obj_Sel(i + 1)).Selec = i
                     Next i
                     If N_Sel > 0 Then ReDim Preserve .Obj_Sel(1 To N_Sel)
                     .Obj(ObjApontado).Selec = 0
                   
                  Else 'o obj apontado nao estava selecionado
                     N_Sel = N_Sel + 1
                     ReDim Preserve .Obj_Sel(1 To N_Sel)
                     .Obj_Sel(N_Sel) = ObjApontado
                     .Obj(ObjApontado).Selec = N_Sel
                  End If '.Obj(ObjApontado).Selec > 0
                
               End If 'Shift = 0
               End With
            End If 'ObjApontado > 0
         Case "PONTO"
            'não define um ponto sobre outro
            'não cria novo ponto ao mover um existente
            If ObjApontado > 0 Then Exit Sub
            
            Def_Ponto Obj_Aux(1).Coord
            
         Case "SEGMENTO"
            If Clicou_Sobre_Objeto Then
               If Posicionando Then
                  If Abs(X - X_Ini) > PROX Then Posicionando = False: Exit Sub
                  If Abs(Y - Y_Ini) > PROX Then Posicionando = False: Exit Sub
               End If
               Id = ObjApontado
            Else
               If ObjApontado > 0 Then
                  Id = ObjApontado
                  If UBound(Obj_Aux) = 1 Then
                     With Doc(Me.Tag).Obj(Id)
                        Obj_Aux(1).Coord(0) = .Coord(0)
                        Obj_Aux(1).Coord(1) = .Coord(1)
                        Obj_Aux(1).Coord(2) = .Coord(2)
                        Obj_Aux(1).Coord(3) = .Coord(3)
                        Obj_Aux(1).Tam = 3
                     End With
                  End If
               Else
                  If UBound(Obj_Aux) = 1 Then
                     Def_Ponto Obj_Aux(1).Coord
                  Else
                     Def_Ponto Obj_Aux(2).Coord
                  End If
                  Id = UBound(Doc(Me.Tag).Obj)
               End If
            End If
            
            If UBound(Obj_Aux) = 1 Then
               ReDim Preserve Obj_Aux(1 To 3) 'dois pontos e um segmento
               If Clicou_Sobre_Objeto Then
                  With Doc(Me.Tag).Obj(Id)
                     Obj_Aux(1).Coord(0) = .Coord(0)
                     Obj_Aux(1).Coord(1) = .Coord(1)
                     Obj_Aux(1).Coord(2) = .Coord(2)
                     Obj_Aux(1).Coord(3) = .Coord(3)
                     Obj_Aux(1).Tam = 3
                  End With
               End If
               With Obj_Aux(1)
                  Obj_Aux(2).Coord(0) = .Coord(0)
                  Obj_Aux(2).Coord(1) = .Coord(1)
                  Obj_Aux(2).Coord(2) = .Coord(2)
                  Obj_Aux(2).Coord(3) = .Coord(3)
                  .Tam = 3
               End With
               With Obj_Aux(3)
                  .Tipo = SEGMENTO
                  .N_Dep = 2
                  .Cor(0) = 1#: .Cor(1) = 0#: .Cor(2) = 0# 'vermelho
                  .Tam = 1
                  ReDim .Id_Dep(1 To .N_Dep)
                  .Id_Dep(1) = Id
                  .Id_Dep(2) = 2
               End With
            Else
               Def_Seg Obj_Aux(3).Id_Dep(1), Id
               
               ReDim Obj_Aux(1 To 1)
               Obj_Aux(1).Coord(3) = 1
            End If
         
         End Select 'tbrFerramentas.Tag
         
         Clicou_Sobre_Objeto = False
         Redesenhar_Todos
         
      Case vbRightButton
         picVista(PERSPECTIVA).MousePointer = 0
         If Abs(X_Ini - X) < PROX And Abs(Y_Ini - Y) < PROX And Not ObjApontado Then PopupMenu frmMDIGeo3d.mnuEditar
      End Select 'Button
   End Select 'Index da Vista
End Sub

Private Sub picVista_Paint(Index As Integer)
   Dim v As Vista
         
   v = Index
   wglMakeCurrent hDC_Vista(v), hGLRC_Vista(v)
   glClear clrColorBufferBit Or clrDepthBufferBit
   If Index <> EPURA Then
      If Index = PERSPECTIVA Then
         If Posicionando Then 'And UCase(tbrFerramentas.Tag) = "PONTO"
            If ObjApontado > 0 Then
               Des_Plano Sobre_Plano, Doc(Me.Tag).Obj(ObjApontado).Coord
            Else
               If UBound(Obj_Aux) > 1 Then
                  Des_Plano Sobre_Plano, Obj_Aux(UBound(Obj_Aux) - 1).Coord 'P_Aux
               Else
                  Des_Plano Sobre_Plano, Obj_Aux(1).Coord 'P_Aux
               End If
            End If
         End If
      End If
      Des_Eixos
      'Des_Planos
      Des_Objetos Me.Tag, GL_RENDER, Doc(Me.Tag).Obj 'GL_SELECT
      Des_Objetos_Aux Me.Tag, Obj_Aux
   Else
      'Desenha usando vista superior
      Des_LT
      Des_Objetos Me.Tag, GL_RENDER, Doc(Me.Tag).Obj 'GL_SELECT
      Des_Objetos_Aux Me.Tag, Obj_Aux
      
      'Redesenha usando vista frontal
      glMatrixMode GL_MODELVIEW
         glLoadIdentity
         glMultMatrixf Troca_X_Y(0)
         glRotatef 90, 0#, 0#, 1#
         glRotatef 90, 1#, 0#, 0#
      Des_Objetos Me.Tag, GL_RENDER, Doc(Me.Tag).Obj 'GL_SELECT
      Des_Objetos_Aux Me.Tag, Obj_Aux
      
      'Reposiciona para a vista superior usada no próximo evento
      glMatrixMode GL_MODELVIEW
         glLoadIdentity
         glMultMatrixf Troca_X_Y(0)
         glRotatef 90, 0#, 0#, 1# 'Igual à vista SUPERIOR
         
   End If
   v = Index
   SwapBuffers hDC_Vista(v)
   Erro = glGetError: If Erro <> glerrNoError Then ErroFatal Erro
End Sub

Private Sub tbrFerramentas_ButtonClick(ByVal Button As MSComctlLib.Button)
'Convenção: Tag guarda um nome igual aos da enumeração pública de tipos dos objetos
 tbrFerramentas.Tag = Button.Key
 'If tbrFerramentas.Tag <> "PONTO" Then MsgBox "NAO SELECIONANDO!": Posicionando = False
 
End Sub

Private Sub Def_Ponto(Pos() As GLdouble)
   Dim N_Obj As Integer 'Em geral = Ubound(Doc(me.tag).Obj)
   
   N_Obj = UBound(Doc(Me.Tag).Obj)

   If N_Obj < MAX_OBJETOS Then
      N_Obj = N_Obj + 1
      ReDim Preserve Doc(Me.Tag).Obj(1 To N_Obj)
      With Doc(Me.Tag).Obj(N_Obj)
      .Coord(0) = Pos(0)
      .Coord(1) = Pos(1)
      .Coord(2) = Pos(2)
      .Coord(3) = Pos(3)
      
      .Cor(0) = 0#: .Cor(1) = 0#: .Cor(2) = 0# 'preto
      .N_Dep = 0
      .Selec = 0
      .Tam = 3
      .Tipo = PONTO
      End With
   Else
      ErroFatal "Você excedeu o número máximo de objetos nesta construção!"
   End If
End Sub
Private Sub Def_Seg(De As Long, Para As Long)
   Dim N_Obj As Integer 'Em geral = Ubound(Doc(me.tag).Obj)
   
   N_Obj = UBound(Doc(Me.Tag).Obj)

   If N_Obj < MAX_OBJETOS Then
      N_Obj = N_Obj + 1
      ReDim Preserve Doc(Me.Tag).Obj(1 To N_Obj)
      With Doc(Me.Tag).Obj(N_Obj)
      ReDim .Id_Dep(1 To 2)
      .Id_Dep(1) = De
      .Id_Dep(2) = Para
      .Cor(0) = 1#: .Cor(1) = 0.8: .Cor(2) = 0#  'laranja
      .N_Dep = 0
      .Selec = 0
      .Tam = 1
      .Tipo = SEGMENTO
      End With
   Else
      ErroFatal "Você excedeu o número máximo de objetos nesta construção!"
   End If
End Sub

