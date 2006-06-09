VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   Caption         =   "Main"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   639
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1320
      Top             =   2040
   End
   Begin MSComctlLib.StatusBar staInfo 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   9210
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar fsbVertical 
      Height          =   2535
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   4471
      _Version        =   393216
      MousePointer    =   1
      Appearance      =   2
      Orientation     =   1245184
   End
   Begin VB.PictureBox picCanto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   720
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   1440
      Width           =   495
   End
   Begin MSComctlLib.ImageList ilsFerramentas 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar tlbObjetos 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   1111
      ButtonWidth     =   1164
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin MSComCtl2.FlatScrollBar fsbHorizontal 
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   720
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1085
      _Version        =   393216
      MousePointer    =   1
      Appearance      =   2
      Arrows          =   65536
      Orientation     =   1245185
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type Ferramenta
 IdImg As Integer
 Key As String
 TipText As String
End Type

Private multi_sel As Boolean

Private Const DICA = " Dica... "
Private Const TAM_BARRA = 20
Private Const DIST_MIN = 8 'pixels
Private Altura_linha As Single
Private Xant, Yant As Single

Private Sub GeraBarraStatus()
   ' Delete the first Panel object, which is
   ' created automatically.
   'StatusBar1.Panels.Remove 1
   Dim I As Integer

   ' The fourth argument of the Add method
   ' sets the Style property.
   For I = 0 To 6
      staInfo.Panels.Add , , , I
   Next I
   staInfo.Panels(1).AutoSize = sbrSpring
   staInfo.Panels(1).MinWidth = 140
End Sub
Private Sub Form_Initialize()
  Me.Caption = "Geometria dinâmica"
  GeraBarraStatus
  With Me
  '.BackColor = vbWhite
  fsbHorizontal.Move .ScaleLeft, .ScaleTop + .ScaleHeight - TAM_BARRA - staInfo.Height, .ScaleWidth - TAM_BARRA, TAM_BARRA
  fsbVertical.Move .ScaleLeft + .ScaleWidth - TAM_BARRA, .ScaleTop + tlbObjetos.Height, TAM_BARRA, .ScaleHeight - tlbObjetos.Height - TAM_BARRA - staInfo.Height
  fsbHorizontal.Min = -MAX_X: fsbHorizontal.Max = MAX_X
  fsbVertical.Min = MAX_Y:    fsbVertical.Max = -MAX_Y
  picCanto.Move fsbVertical.Left, fsbHorizontal.Top, TAM_BARRA, TAM_BARRA
 End With
 With tlbObjetos.Buttons
  tlbObjetos.Tag = .Item(1).Key
 End With
 'With lstObjetos
 ' .Clear
 ' .AddItem "Medindo linha..."
 ' .height = 0
 ' Altura_linha = .height
 ' .Clear
 ' Me.lblAux.Font.Size = .Font.Size
 '
 ' lblDica.BackColor = vbRed
 ' lblDica.ForeColor = vbBlack
 ' lblDica.Caption = DICA
 ' lblDica.Move 10, Me.ScaleTop + tbrObjetos.height + 10
' End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 
 If KeyCode = 27 Then
  Unload Me
  End
 End If

End Sub

Private Sub Form_Load()
 Call Inicializar_OpenGL(Me.hDC)
 MontaEixos
 
 Form_Paint
 carrega_ferramentas
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim x1, y1 As GLdouble
 Const TAM = 5
 Dim cx, cy As GLdouble

 
 If Button <> 1 Then Exit Sub
 x1 = X / Me.ScaleWidth
 y1 = Y / Me.ScaleHeight
 cx = 1 - 2 * x1
 cy = 2 * y1 - 1
 
 glMatrixMode GL_PROJECTION
  glLoadIdentity
  gluOrtho2D TAM * (cx - 1), TAM * (cx + 1), TAM * (cy - 1), TAM * (cy + 1)
 glMatrixMode GL_MODELVIEW
 
 Form_Paint
 
End Sub

Private Sub Form_Paint()
 
 Desenha_esfera
 MostraEixos
 
 SwapBuffers hDC
 
End Sub
Private Sub Form_Resize()
 Dim Visivel_antes_X As Single, Visivel_antes_Y As Single

 Call Ajusta_ViewPort(0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight)
 'glViewport 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight
 Form_Paint
 
 Visivel_antes_X = Visivel_X 'Guarda a medida da largura visível atualmente
 Visivel_antes_Y = Visivel_Y 'Guarda a medida da altura visível atualmente
 With Me
  'Mede a largura e a altura da área de desenho em "pixels"
  Visivel_X = .ScaleWidth - .fsbVertical.Width
  Visivel_Y = .ScaleHeight - tlbObjetos.Height - fsbHorizontal.Height - staInfo.Height
  'Converte a largura e a altura de PIXELS para CENTÍMETROS
  Visivel_X = Visivel_X * TwipsPerPixelX_INICIAL / Twips_por_Cm
  Visivel_Y = Visivel_Y * TwipsPerPixelY_INICIAL / Twips_por_Cm
  'Atualiza as coordenadas que correspondem ao centro do form.
  Centro_X = Centro_X + (Visivel_X - Visivel_antes_X) / 2
  Centro_Y = Centro_Y - (Visivel_Y - Visivel_antes_Y) / 2
  
  
  On Error Resume Next
  'Como evitar que desapareça a área de desenho ao diminuir MUITO a largura e altura?
  'Reposiciona barras e botões
  fsbHorizontal.Move .ScaleLeft, .ScaleTop + .ScaleHeight - TAM_BARRA - staInfo.Height, .ScaleWidth - TAM_BARRA, TAM_BARRA
  fsbVertical.Move .ScaleLeft + .ScaleWidth - TAM_BARRA, .ScaleTop + tlbObjetos.Height, TAM_BARRA, .ScaleHeight - tlbObjetos.Height - TAM_BARRA - staInfo.Height
  picCanto.Move fsbVertical.Left, fsbHorizontal.Top
  On Error GoTo 0
  Timer1.Enabled = True
  
  '.Cls
  '.Refresh
  'lblDica.Move 10, Me.ScaleTop + tbrObjetos.height + 10
 End With
 
End Sub
Private Sub Form_Unload(Cancel As Integer)

 Call Finalizar_OpenGL '(Me.hDC)'será necessário?
 
End Sub

Sub carrega_ferramentas()
 Const Arq_INI = "Tabela.ini"
 Dim imgX As ListImage
 Dim btnButton As Button
 
 Dim Qtd As Integer
 Dim FileNumber As Variant, Q As Variant
 Dim N As Integer
 Dim F() As Ferramenta
 
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
 
 With tlbObjetos
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
 If Err.Number = 53 Then
  Err.Clear
  'Recup_Arquivo
  'Inicializa
 Else
  Err.Raise Err.Number
 End If
End Sub

Private Sub Timer1_Timer()
  'Esse timer só existe para contornar um defeito na rotina Resize...
  'O programa nao conhece a altura real da barra no instante do redimensionamento, só depois
  With Me
   On Error Resume Next
   fsbVertical.Move .ScaleLeft + .ScaleWidth - TAM_BARRA, .ScaleTop + tlbObjetos.Height, TAM_BARRA, .ScaleHeight - tlbObjetos.Height - TAM_BARRA - staInfo.Height
   On Error GoTo 0
   Timer1.Enabled = False
  End With
End Sub
