VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Exemplo - Integrando Vb e OpenGl"
   ClientHeight    =   3360
   ClientLeft      =   600
   ClientTop       =   1185
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   688
   Begin VB.TextBox txtLineStipple 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Text            =   "3"
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtLineWidth 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1800
      TabIndex        =   7
      Text            =   "1"
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtPointSize 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Text            =   "3"
      Top             =   1800
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Opções"
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   7200
      TabIndex        =   1
      Top             =   360
      Width           =   2895
      Begin VB.CheckBox chkQuad 
         Appearance      =   0  'Flat
         Caption         =   "Exibir quadrado"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox chkMarcas 
         Appearance      =   0  'Flat
         Caption         =   "Exibir marcas sobre os eixos"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chkEixos 
         Appearance      =   0  'Flat
         Caption         =   "Exibir eixos X e Y"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2175
      End
   End
   Begin VB.PictureBox picViewTela 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3000
      Left            =   2520
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   120
      Width           =   4500
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Estilo de pontilhado:"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   2760
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Espessura das linhas:"
      Height          =   195
      Left            =   255
      TabIndex        =   8
      Top             =   2280
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tamanho dos pontos:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   1545
   End
   Begin VB.Image imgRight 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   1530
      Picture         =   "Main.frx":0000
      Top             =   660
      Width           =   510
   End
   Begin VB.Image ImgLeft 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   510
      Picture         =   "Main.frx":0442
      Top             =   660
      Width           =   510
   End
   Begin VB.Image ImgDown 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   1020
      Picture         =   "Main.frx":0884
      Top             =   1170
      Width           =   510
   End
   Begin VB.Image imgUp 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   1020
      Picture         =   "Main.frx":0CC6
      Top             =   150
      Width           =   510
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkEixos_Click()
picViewTela_Paint
End Sub

Private Sub chkMarcas_Click()
picViewTela_Paint
End Sub

Private Sub chkQuad_Click()
picViewTela_Paint
End Sub

Private Sub Form_Load()

 hDC1 = Me.picViewTela.hDC 'Identificador da ViewPort1 (embora não use + de uma viewport)
 
 Larg = frmMain.picViewTela.ScaleWidth
 Alt = frmMain.picViewTela.ScaleHeight
 Centro_X = 0
 Centro_Y = 0
 Visivel_X = 10: Visivel_Y = Visivel_X * Alt / Larg
 Call Inicializar_OpenGL(hDC1) 'Ajusta formato dos pixels, iluminação, matrizes de projeção...

End Sub
Private Sub picViewTela_DblClick()
If Not frmMatriz.Visible Then frmMatriz.Show    'vbModal
End Sub

Private Sub picViewTela_Paint()
Dim D As GLfloat, Ini As GLfloat, Fim As GLfloat
 glClear clrColorBufferBit Or clrDepthBufferBit
 
 If chkQuad Then
  glLineWidth 2 * CInt(Val(txtLineWidth))
  glLineStipple 1, CInt(Val(txtLineStipple))
  glEnable glcLineStipple
  glColor3f 0#, 1#, 0#
  glBegin bmLineLoop
   glVertex2f 2, 0
   glVertex2f 0, 2
   glVertex2f -2, 0
   glVertex2f 0, -2
  glEnd
  glDisable glcLineStipple
 End If
 
 glLineWidth CInt(Val(txtLineWidth))
 glPointSize CInt(Val(txtPointSize))

 Ini = Centro_X - (Visivel_X / 2)
 Fim = Ini + Visivel_X
 If chkEixos Then
  glColor3f 0.5, 0.5, 0.5
  glBegin bmLines
   glVertex2f Ini, 0
   glVertex2f Fim, 0
  glEnd
 End If
 If chkMarcas Then
  glColor3f 0.2, 0.2, 0.2
  glBegin bmPoints
   For D = CInt(Ini) To CInt(Fim)
     glVertex2f D, 0
   Next D
  glEnd
 End If
 
 Ini = Centro_Y - (Visivel_Y / 2)
 Fim = Ini + Visivel_Y
 If chkEixos Then
  glColor3f 0.5, 0.5, 0.5
  glBegin bmLines
   glVertex2f 0, Ini
   glVertex2f 0, Fim
  glEnd
 End If
 If chkMarcas Then
  glColor3f 0.2, 0.2, 0.2
  glBegin bmPoints
   For D = CInt(Ini) To CInt(Fim)
     glVertex2f 0, D
   Next D
  glEnd
 End If
 
 SwapBuffers hDC1
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload frmMatriz
Call Finalizar_OpenGL
End Sub

Private Sub ImgDown_Click()
Centro_Y = Centro_Y + 0.15 * Visivel_Y
Call basVisualização.Ajusta_ViewPort(0, 0, Larg, Alt)
Call picViewTela_Paint
End Sub

Private Sub ImgLeft_Click()
Centro_X = Centro_X + 0.15 * Visivel_X
Call basVisualização.Ajusta_ViewPort(0, 0, Larg, Alt)
Call picViewTela_Paint
End Sub

Private Sub imgRight_Click()
Centro_X = Centro_X - 0.15 * Visivel_X
Call basVisualização.Ajusta_ViewPort(0, 0, Larg, Alt)
Call picViewTela_Paint
End Sub

Private Sub imgUp_Click()
Centro_Y = Centro_Y - 0.15 * Visivel_Y
Call basVisualização.Ajusta_ViewPort(0, 0, Larg, Alt)
Call picViewTela_Paint
End Sub

Private Sub txtLineStipple_LostFocus()
Call picViewTela_Paint
End Sub

Private Sub txtLineStipple_Validate(Cancel As Boolean)
If Val(txtLineStipple) > 500 Then txtLineStipple = "3": MsgBox "O valor deve ser menor que 500": Cancel = True
End Sub

Private Sub txtLineWidth_LostFocus()
Call picViewTela_Paint
End Sub

Private Sub txtLineWidth_Validate(Cancel As Boolean)
If Val(txtLineWidth) > 30 Then txtLineWidth = "2": MsgBox "O valor deve ser menor que 30": Cancel = True
End Sub

Private Sub txtPointSize_LostFocus()
Call picViewTela_Paint
End Sub

Private Sub txtPointSize_Validate(Cancel As Boolean)
If Val(txtPointSize) > 30 Then txtPointSize = "5": MsgBox "O valor deve ser menor que 30": Cancel = True
End Sub
