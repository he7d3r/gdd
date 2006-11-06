VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exemplo - Integrando Vb e OpenGl"
   ClientHeight    =   3045
   ClientLeft      =   585
   ClientTop       =   1170
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   203
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   596
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLimpar 
      Cancel          =   -1  'True
      Caption         =   "Limpar"
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtLineWidth 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1155
      TabIndex        =   1
      Text            =   "2"
      Top             =   195
      Width           =   735
   End
   Begin VB.TextBox txtFactor 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1155
      TabIndex        =   2
      Text            =   "1"
      Top             =   915
      Width           =   735
   End
   Begin VB.PictureBox picViewTela 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1560
      Left            =   3480
      ScaleHeight     =   104
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   340
      TabIndex        =   0
      Top             =   240
      Width           =   5100
   End
   Begin VB.Label lblConta 
      Height          =   330
      Left            =   225
      TabIndex        =   11
      Top             =   2520
      Width           =   8115
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   ";"
      Height          =   195
      Left            =   1935
      TabIndex        =   10
      Top             =   1035
      Width           =   45
   End
   Begin VB.Label lblPattern 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2070
      TabIndex        =   9
      Top             =   915
      Width           =   690
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   ")"
      Height          =   195
      Left            =   2835
      TabIndex        =   8
      Top             =   960
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   ")"
      Height          =   195
      Left            =   1935
      TabIndex        =   7
      Top             =   240
      Width           =   45
   End
   Begin VB.Image imgNaoSim 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   3120
      Picture         =   "Main.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgNaoSim 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   0
      Left            =   2760
      Picture         =   "Main.frx":0442
      Stretch         =   -1  'True
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgStipple 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   255
      Index           =   15
      Left            =   6360
      Picture         =   "Main.frx":0884
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgStipple 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   14
      Left            =   6000
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgStipple 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   13
      Left            =   5640
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgStipple 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   12
      Left            =   5280
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgStipple 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   11
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgStipple 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   10
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgStipple 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   9
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgStipple 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   8
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgStipple 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   7
      Left            =   3000
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgStipple 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   6
      Left            =   2640
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgStipple 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   5
      Left            =   2280
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgStipple 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   4
      Left            =   1920
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgStipple 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   3
      Left            =   1320
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgStipple 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   2
      Left            =   960
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgStipple 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   1
      Left            =   600
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Image imgStipple 
      Appearance      =   0  'Flat
      Height          =   255
      Index           =   0
      Left            =   240
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Pattern (Estilo de pontilhado):"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   2070
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "glLineWidth ("
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   240
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "glLineStipple("
      Height          =   195
      Left            =   165
      TabIndex        =   4
      Top             =   960
      Width           =   945
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Stipple(0 To 14) As Boolean
Private Pattern As GLushort
Const SOMA_UM = "Soma = 1 + 0 + 0 + 0 + 0 + 0 + 0 + 0 + 0 + 0 + 0 + 0 + 0 + 0 + 0 + 0 = 1"

Private Sub cmdLimpar_Click()
Dim i As Long
 Stipple(0) = True
 imgStipple(0).Picture = imgNaoSim(1)
 For i = 1 To 14
  Stipple(i) = False
  imgStipple(i).Picture = imgNaoSim(0)
 Next i
 txtLineWidth = 2
 txtFactor = 1
 Pattern = 1
 lblPattern = Pattern
 lblConta = SOMA_UM
 picViewTela_Paint
End Sub

Private Sub Form_Load()
 'Dim i As Integer
 hDC1 = Me.picViewTela.hDC 'Identificador da ViewPort1 (embora não use + de uma viewport)
 
 Larg = frmMain.picViewTela.ScaleWidth
 Alt = frmMain.picViewTela.ScaleHeight
 Centro_X = 0
 Centro_Y = 0
 Visivel_X = 10: Visivel_Y = Visivel_X * Alt / Larg
 Call Inicializar_OpenGL(hDC1) 'Ajusta formato dos pixels, iluminação, matrizes de projeção...
 
 cmdLimpar_Click
End Sub

Private Sub imgStipple_Click(Index As Integer)
Dim i As Long
Stipple(Index) = Not Stipple(Index)
imgStipple(Index).Picture = imgNaoSim(IIf(Stipple(Index), 1, 0))

Pattern = 0: lblConta = "Soma = "
For i = 0 To 14
 lblConta = lblConta & CStr(IIf(Stipple(i), 1, 0) * 2 ^ i) & " + "
 Pattern = Pattern Or (IIf(Stipple(i), 1, 0) * 2 ^ i)
Next i
lblConta = Left(lblConta, Len(lblConta) - 3) & " = " & Pattern
lblPattern = Pattern
picViewTela_Paint
End Sub

Private Sub picViewTela_Paint()
Dim D As GLfloat, Ini As GLfloat, Fim As GLfloat

 glClear clrColorBufferBit Or clrDepthBufferBit
 
 glLineWidth CInt(Val(txtLineWidth))
 glEnable glcLineStipple
 glLineStipple CInt(Val(txtFactor)), CInt(Val(lblPattern))
 Ini = Centro_X - (Visivel_X / 3)
 Fim = Ini + 2 * Visivel_X / 3
 glColor3f 0#, 0#, 0#
 glBegin bmLines
   glVertex2f Ini, 0
   glVertex2f Fim, 0
 glEnd
 'glDisable glcLineStipple
 SwapBuffers hDC1
End Sub
Private Sub Form_Unload(Cancel As Integer)
Call Finalizar_OpenGL
End Sub
Private Sub txtFactor_LostFocus()
Call picViewTela_Paint
End Sub
Private Sub txtLineWidth_LostFocus()
Call picViewTela_Paint
End Sub
Private Sub txtLineWidth_Validate(Cancel As Boolean)
If Val(txtLineWidth) > 30 Then txtLineWidth = "2": MsgBox "O valor deve ser menor que 30": Cancel = True
End Sub

