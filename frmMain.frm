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
   Begin VB.PictureBox picViewTela 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   2400
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   5
      ToolTipText     =   "Incluir tips conforme posição (ou não!)"
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1320
      Top             =   2040
   End
   Begin MSComctlLib.StatusBar staInfo 
      Align           =   2  'Align Bottom
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   8730
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   1508
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
Private Const TAM_TELA = 30 'Cms
Private Const N_INC = 100
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
   staInfo.Height = TAM_BARRA
   For I = 0 To 6
      staInfo.Panels.Add , , , I
   Next I
   staInfo.Panels(1).AutoSize = sbrSpring
   staInfo.Panels(1).MinWidth = 140
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 
 If KeyCode = 27 Then
  Unload Me
  End
 End If

End Sub

Private Sub Form_Load()
 Call Inicializar_OpenGL(Me.picViewTela.hDC)
 Call Carrega_Ferramentas
 Call GeraBarraStatus
 
 Call MontaEixos
 Call Inicializar_Objetos
 
 Me.Caption = "Geometria dinâmica"
 With Me
  '.BackColor = vbWhite
  fsbHorizontal.Move .ScaleLeft, .ScaleTop + .ScaleHeight - TAM_BARRA - staInfo.Height, .ScaleWidth - TAM_BARRA, TAM_BARRA
  fsbVertical.Move .ScaleLeft + .ScaleWidth - TAM_BARRA, .ScaleTop + tlbObjetos.Height, TAM_BARRA, .ScaleHeight - tlbObjetos.Height - TAM_BARRA - staInfo.Height
  
  fsbVertical.Value = fsbVertical.Max \ 2
  fsbHorizontal.Value = fsbHorizontal.Max \ 2
  
  fsbVertical.SmallChange = fsbVertical.Max \ N_INC
  fsbHorizontal.SmallChange = fsbHorizontal.Max \ N_INC
  
  fsbVertical.LargeChange = fsbVertical.Max \ (N_INC \ 5)
  fsbHorizontal.LargeChange = fsbHorizontal.Max \ (N_INC \ 5)
  
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim x1, y1 As GLdouble
 Dim cx, cy As GLdouble
 
 If Button <> -1 Then Exit Sub
 
 x1 = X / Me.ScaleWidth
 y1 = Y / Me.ScaleHeight
 cx = 1 - 2 * x1
 cy = 2 * y1 - 1
 
 Call Ajusta_ViewPort(0, 2 * TAM_BARRA, Visivel_X_pix, Visivel_Y_pix)
 
 Form_Paint
 
End Sub

Private Sub Form_Paint()
 
 Desenha_esfera
 MostraEixos
 
 SwapBuffers hDC
 
End Sub

Private Sub Form_Resize()
 Dim Visivel_antes_X As Single, Visivel_antes_Y As Single
 Dim AltBarra As Single
 
 With Me.tlbObjetos
 AltBarra = 6 + .ButtonHeight * ((.Buttons.Count - 1) \ Int(Me.ScaleWidth / .ButtonWidth) + 1)
 End With
 
 Visivel_antes_X = Visivel_X 'Guarda a medida da largura visível atualmente
 Visivel_antes_Y = Visivel_Y 'Guarda a medida da altura visível atualmente
 With Me.picViewTela
  .Move Me.ScaleLeft, _
                    Me.ScaleTop + AltBarra, _
                    Me.ScaleWidth - TAM_BARRA, _
                    Me.ScaleHeight - AltBarra - 2 * TAM_BARRA
  'Mede a largura e a altura da área de desenho em "pixels"
  Visivel_X_pix = .Width
  Visivel_Y_pix = .Height
  'Converte a largura e a altura de PIXELS para CENTÍMETROS
  Visivel_X = Visivel_X_pix * Cm_por_Pixel_X
  Visivel_Y = Visivel_Y_pix * Cm_por_Pixel_Y
  'Atualiza as coordenadas que correspondem ao centro do form.
  Centro_X = Centro_X + (Visivel_X - Visivel_antes_X) / 2
  Centro_Y = Centro_Y - (Visivel_Y - Visivel_antes_Y) / 2
  
  On Error Resume Next
  'Como evitar que desapareça a área de desenho ao diminuir MUITO a largura e altura?
  'Use API's do windows
  '
  'Reposiciona barras e botões
  fsbHorizontal.Move .Left, .Top + .Height, .Width, TAM_BARRA
  fsbVertical.Move .Left + .Width, .Top, TAM_BARRA, .Height
  picCanto.Move fsbVertical.Left, fsbHorizontal.Top
  On Error GoTo 0
  'Timer1.Enabled = True
  
 Call Ajusta_ViewPort(0, 0, 100, 100) '.Width, .Height)
 'glViewport 0, 0, frmMain.ScaleWidth, frmMain.ScaleHeight
 Form_Paint
 
  'lblDica.Move 10, Me.ScaleTop + tbrObjetos.height + 10
 End With
 
End Sub
Private Sub Form_Unload(Cancel As Integer)

 Call Finalizar_OpenGL '(Me.hDC)'será necessário?
 
End Sub

Sub Carrega_Ferramentas()
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
Private Sub fsbHorizontal_Change()
'Tratar o bug que ocorre ao redimensionar o form:
'a rotina resize muda para decimais o valor dos centros X e Y,
'mas as scroll's tm value inteiro
 Centro_X = TAM_TELA * (fsbHorizontal.Value / fsbHorizontal.Max - 0.5)
 Call Ajusta_ViewPort(0, 2 * TAM_BARRA, Visivel_X_pix, Visivel_Y_pix)
 Form_Paint
End Sub
Private Sub fsbHorizontal_Scroll()
'Incluir mais valores entre os extremos de MAX_X e -MAX_X permitirá um scroll mais suave
'Conferir compatibilidade com o Zoom
 Centro_X = TAM_TELA * (fsbHorizontal.Value / fsbHorizontal.Max - 0.5)
 Call Ajusta_ViewPort(0, 2 * TAM_BARRA, Visivel_X_pix, Visivel_Y_pix)
 Form_Paint
End Sub
Private Sub fsbVertical_Change()
 Centro_Y = -TAM_TELA * (fsbVertical.Value / fsbVertical.Max - 0.5)
 Call Ajusta_ViewPort(0, 2 * TAM_BARRA, Visivel_X_pix, Visivel_Y_pix)
 Form_Paint
End Sub
Private Sub fsbVertical_Scroll()
'Incluir mais valores entre os extremos de MAX_Y e -MAX_Y permitirá um scroll mais suave
'Conferir compatibilidade com o Zoom
 Centro_Y = -TAM_TELA * (fsbVertical.Value / fsbVertical.Max - 0.5)
 Call Ajusta_ViewPort(0, 2 * TAM_BARRA, Visivel_X_pix, Visivel_Y_pix)
 Form_Paint
End Sub
