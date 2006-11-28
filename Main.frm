VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exemplo - Integrando Vb e OpenGl"
   ClientHeight    =   7200
   ClientLeft      =   585
   ClientTop       =   1170
   ClientWidth     =   10290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   686
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picEpura 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   240
      ScaleHeight     =   198
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   9
      Top             =   3840
      Width           =   3000
   End
   Begin VB.PictureBox picLateral 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   3720
      ScaleHeight     =   198
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   7
      Top             =   3840
      Width           =   3000
   End
   Begin VB.PictureBox picSuperior 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   7080
      ScaleHeight     =   198
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   6
      Top             =   3840
      Width           =   3000
   End
   Begin VB.PictureBox picFrontal 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   7080
      ScaleHeight     =   198
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   198
      TabIndex        =   1
      Top             =   480
      Width           =   3000
   End
   Begin VB.PictureBox picPerspectiva 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   240
      ScaleHeight     =   198
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   431
      TabIndex        =   0
      Top             =   480
      Width           =   6495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Épura (1ª e 2ª Proj.):"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   3600
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Vista Frontal (2ª Proj.):"
      Height          =   195
      Left            =   7080
      TabIndex        =   5
      Top             =   240
      Width           =   1560
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Vista Lateral (3ª Proj.):"
      Height          =   195
      Left            =   3720
      TabIndex        =   4
      Top             =   3600
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Vista Superior (1ª Proj.):"
      Height          =   195
      Left            =   7080
      TabIndex        =   3
      Top             =   3600
      Width           =   1665
   End
   Begin VB.Label lblPerspectiva 
      AutoSize        =   -1  'True
      Caption         =   "Perspectiva:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   885
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Magnetismo
Private X_Ini As Integer, Y_Ini As Integer
Private Phi_Ini As GLfloat, Theta_Ini As GLfloat
Private Px As GLdouble, Py As GLdouble, Pz As GLdouble

Private Sub Form_KeyPress(KeyAscii As Integer)
 If Chr(KeyAscii) = "m" Or Chr(KeyAscii) = "M" Then Magnetismo = Not Magnetismo
 If Chr(KeyAscii) = "+" Then Ro = Ro - 1
 If Chr(KeyAscii) = "-" Then Ro = Ro + 1
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

Private Sub Form_Load()

 hDCPerspectiva = Me.picPerspectiva.hDC 'Identificador das ViewPort's
 hDCFrontal = Me.picFrontal.hDC
 hDCLateral = Me.picLateral.hDC
 hDCSuperior = Me.picSuperior.hDC
 hDCEpura = Me.picEpura.hDC
 Px = 0: Py = 0: Pz = 0
 Magnetismo = True
 Call Inicializar_OpenGL 'Ajusta formato dos pixels, iluminação, matrizes de projeção...
End Sub
Private Sub Des_Plano()
 Const RAIO = 3
 Dim K As GLdouble
 Dim PosX As GLdouble, PosY As GLdouble
 
 glColor3f 0.5, 0.5, 0.5
 'glLineWidth (1#)
 glBegin bmLines
  For K = -RAIO To RAIO
    PosX = Fix(Px + K): PosY = Fix(Py + K)
    If Abs(PosX - Px) < RAIO Then
    glVertex3d PosX, Py + (RAIO - Abs(PosX - Px)), 0#
    glVertex3d PosX, Py - (RAIO - Abs(PosX - Px)), 0#
    End If
    If Abs(PosY - Py) < RAIO Then
    glVertex3d Px + (RAIO - Abs(PosY - Py)), PosY, 0#
    glVertex3d Px - (RAIO - Abs(PosY - Py)), PosY, 0#
    End If
  Next K
 glEnd
End Sub
Private Sub Des_Eixos()
 'glLineWidth (2#)
 glBegin bmLines
   glColor3f 1#, 0#, 0#
   glVertex3f 0#, 0#, 0#
   glVertex3f 3#, 0#, 0#
   
   glColor3f 0#, 1#, 0#
   glVertex3f 0#, 0#, 0#
   glVertex3f 0#, 3#, 0#
   
   glColor3f 0#, 0#, 1#
   glVertex3f 0#, 0#, 0#
   glVertex3f 0#, 0#, 3#
 glEnd
End Sub
Private Sub Des_Ponto()
  glColor3d 0.1, 0.1, 0.1
  glPointSize (3#)
  glBegin bmPoints
   If Magnetismo Then
    glVertex3d Round(Px), Round(Py), Round(Pz)
   Else
    glVertex3d Px, Py, Pz
   End If
  glEnd
  glColor3d 0.7, 0.7, 0.7
  glBegin bmLines
  ' glVertex3d 0#, Py, Pz
  ' glVertex3d Px, Py, Pz
  If Magnetismo Then
   glVertex3d Round(Px), Round(Py), 0#
   glVertex3d Round(Px), Round(Py), Round(Pz)
  Else
   glVertex3d Px, Py, 0#
   glVertex3d Px, Py, Pz
  End If
  ' glVertex3d Px, 0#, Pz
  ' glVertex3d Px, Py, Pz
  
  glEnd
End Sub
Private Sub Des_Figura()
 glPushMatrix
  glTranslatef -0.5, 1.5, 0.5
  glColor3d 0, 0, 0
  gluCylinder QObj, 1.5, 0.5, 2, 12, 2
 glPopMatrix
End Sub
Private Sub Des_LT()
 glBegin GL_LINES
  glColor3d 0.5, 0, 0
  glVertex3f -3, 0, 0
  glVertex3f 3, 0, 0
 glEnd
 glPointSize 3#
 glBegin GL_POINTS
  glColor3d 0.5, 0, 0
  glVertex3f 0, 0, 0
 glEnd
End Sub
Private Sub Desenha_Todos()
 Des_Figura
 Des_Eixos
 Des_Ponto
End Sub
Private Sub Form_Unload(Cancel As Integer)
Call Finalizar_OpenGL
End Sub

Private Sub picPerspectiva_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 X_Ini = X: Y_Ini = Y
 Phi_Ini = Phi:  Theta_Ini = Theta
End Sub

Private Sub picPerspectiva_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Const VELOCIDADE = 0.5
 Dim dx As Integer, dy As Integer
 Dim winX As GLdouble, winY As GLdouble, winZ As GLdouble
 
 Dim Pos As GLdouble
 Dim ViewPort(0 To 3) As GLint
 Dim mvmatrix(0 To 15) As GLdouble, projmatrix(0 To 15) As GLdouble
 Dim realy As GLint
 Dim x1 As GLdouble, y1 As GLdouble, z1 As GLdouble
 Dim x0 As GLdouble, y0 As GLdouble, z0 As GLdouble
 Dim vx As GLdouble, vy As GLdouble, vz As GLdouble
 Dim px1 As GLdouble, py1 As GLdouble, pz1 As GLdouble
 Dim px2 As GLdouble, py2 As GLdouble, pz2 As GLdouble
 
 Select Case Button
 Case 1
  wglMakeCurrent hDCPerspectiva, hGLRCPerspectiva
  glGetIntegerv GL_VIEWPORT, ViewPort(0)
  glGetDoublev GL_MODELVIEW_MATRIX, mvmatrix(0)
  glGetDoublev GL_PROJECTION_MATRIX, projmatrix(0)
  realy = ViewPort(3) - Y - 1
  gluUnProject X, realy, 0#, mvmatrix(0), projmatrix(0), ViewPort(0), x0, y0, z0
  gluUnProject X, realy, 1#, mvmatrix(0), projmatrix(0), ViewPort(0), x1, y1, z1
  vx = x1 - x0
  vy = y1 - y0
  vz = z1 - z0
  Select Case Shift
  Case 0
   If vz = 0 Then vz = z0: MsgBox "vz=0"
   Pos = (Pz - z0) / vz
   If (Pos < 0 Or 1 < Pos) Then Exit Sub
    Px = x0 + Pos * vx
    Py = y0 + Pos * vy
    Pz = z0 + Pos * vz
   
  Case vbShiftMask
  
  Case vbAltMask
  
  Case vbAltMask + vbShiftMask
  
  Case vbCtrlMask
    If vx = 0 Then vx = x0: MsgBox "vx=0"
    Pos = (Px - x0) / vx
    If (Pos < 0 Or 1 < Pos) Then Exit Sub
    px1 = Px: py1 = y0 + Pos * vy: pz1 = z0 + Pos * vz
    
    If X < 5 Then X = 5
    gluUnProject X - 5, realy, 0#, mvmatrix(0), projmatrix(0), ViewPort(0), x0, y0, z0
    gluUnProject X - 5, realy, 1#, mvmatrix(0), projmatrix(0), ViewPort(0), x1, y1, z1
    If vx = 0 Then vx = x0: MsgBox "vx=0"
    Pos = (Px - x0) / vx
    If (Pos < 0 Or 1 < Pos) Then Exit Sub
    px2 = Px:  py2 = y0 + Pos * vy:  pz2 = z0 + Pos * vz
    'px = px
    'Py = Py
    If py2 <> py1 Then Pz = pz1 + (Py - py1) * (pz2 - pz1) / (py2 - py1)
  Case vbCtrlMask + vbShiftMask
  
  Case vbCtrlMask + vbAltMask
  
  Case vbCtrlMask + vbAltMask + vbShiftMask
  
  End Select
  picPerspectiva_Paint
  picEpura_Paint
  picSuperior_Paint
  picFrontal_Paint
  picLateral_Paint
 Case 2 'botao direito
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

Private Sub picPerspectiva_Paint()

 wglMakeCurrent hDCPerspectiva, hGLRCPerspectiva
 glClear clrColorBufferBit Or clrDepthBufferBit

 Desenha_Todos
 Des_Plano
 
 SwapBuffers hDCPerspectiva
End Sub
Private Sub picSuperior_Paint()

 wglMakeCurrent hDCSuperior, hGLRCSuperior
 glClear clrColorBufferBit Or clrDepthBufferBit
 Desenha_Todos
 SwapBuffers hDCSuperior
End Sub
Private Sub picEpura_Paint()

 wglMakeCurrent hDCEpura, hGLRCEpura
 glClear clrColorBufferBit Or clrDepthBufferBit
 glMatrixMode GL_MODELVIEW
  glLoadIdentity
  glMultMatrixf Troca_X_Y(0)
  glRotatef 90, 0#, 0#, 1#
  glRotatef 90, 1#, 0#, 0#
 
 Des_Figura
 Des_Ponto
 
 glMatrixMode GL_MODELVIEW
  glLoadIdentity
  glMultMatrixf Troca_X_Y(0)
  glRotatef 90, 0#, 0#, 1#
 
 Des_Figura
 Des_Ponto
 Des_LT
 
 SwapBuffers hDCEpura
End Sub
Private Sub picFrontal_Paint()

 wglMakeCurrent hDCFrontal, hGLRCFrontal
 glClear clrColorBufferBit Or clrDepthBufferBit
 Desenha_Todos
 SwapBuffers hDCFrontal
End Sub
Private Sub picLateral_Paint()

 wglMakeCurrent hDCLateral, hGLRCLateral
 glClear clrColorBufferBit Or clrDepthBufferBit
 Desenha_Todos
 SwapBuffers hDCLateral
End Sub
