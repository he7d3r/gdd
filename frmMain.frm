VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0FFC0&
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
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   1200
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   137
      TabIndex        =   5
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   720
   End
   Begin MSComctlLib.StatusBar staInfo 
      Height          =   435
      Left            =   1320
      TabIndex        =   4
      Top             =   1200
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar fsbVertical 
      Height          =   1935
      Left            =   3360
      TabIndex        =   2
      Top             =   1800
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   3413
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
      Left            =   3360
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   3840
      Width           =   495
   End
   Begin MSComctlLib.ImageList ilsFerramentas 
      Left            =   600
      Top             =   1200
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
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   3840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
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

Private Const vA = 5 'Entre a barra de ferramentas e a área de desenho
Private Const vB = 5 'Entre a margem esquerda do form e a área de desenho
Private Const vC = 5 'Entre a área de desenho e a barra de rolagem vertical
Private Const vD = 5 'Entre a área de desenho e a barra de rolagem horizontal
Private Const vE = 5 'Entre a borda direita do form e a barra de rolagem vertical
Private Const vF = 5 'Entre a borda inferior do form e a barra de rolagem horizontal
Private Const vG = 0
Private Const vH = 0
Private Const vI = 0
Private Const vJ = 0

Private multi_sel As Boolean

Private Const DICA = " Dica... "
Private Const TAM_BARRA = 20
Private Const TAM_FOLHA_DESENHO = 30 '100  'Cms
Private Const N_INC = 100
Private Const DIST_MIN = 8 'pixels
Private Altura_linha As Single
Private Xant, Yant As Single
Private Redesenhar As Boolean 'redesenhar tela no evento change das scroll's?

Private Sub Form_Load()
 hDC1 = Me.picViewTela.hDC 'Identificador da ViewPort1 (embora não use + de uma viewport)
 Call Inicializar_OpenGL(hDC1) 'Ajusta formato dos pixels, iluminação, matrizes de projeção...
 Call Carrega_Ferramentas 'Carrega imagens e nomes dos botões da barra de ferramentas
 Call GeraBarraStatus 'Configura uma barra de status padrão, com altura TAM_BARRA.
 
 Call MontaEixos 'Gera as displaylists para os eixos
 Call Inicializar_Objetos 'Carrega matriz de objetos geométricos
 
    tlbObjetos.Tag = tlbObjetos.Buttons.Item(1).Key
    basObjGeometria.inc_Mov = 0.05
    basObjGeometria.inc_Trans = 1
    basObjGeometria.Zoom = 100#
    Redesenhar = True
    
    basObjGeometria.Centro_X = 0#
    basObjGeometria.Centro_Y = 0#
    
    basObjGeometria.TwipsPerPixelX_INICIAL = Screen.TwipsPerPixelX
    basObjGeometria.TwipsPerPixelY_INICIAL = Screen.TwipsPerPixelY
    'O form mede N twips, cada cm contém M twips, logo o form mede N/M cm's
    basObjGeometria.Cm_por_Pixel_X = TwipsPerPixelX_INICIAL / Twips_por_Cm
    basObjGeometria.Cm_por_Pixel_Y = TwipsPerPixelY_INICIAL / Twips_por_Cm
    
  With frmMain
   '.BackColor = vbWhite
    .Caption = "Geometria dinâmica"
   'Mede a largura e a altura da área de desenho em "pixels"
    basObjGeometria.Visivel_X = .ScaleWidth - (TAM_BARRA + vB + vC + vE)
    basObjGeometria.Visivel_Y = .ScaleHeight - (2 * TAM_BARRA + tlbObjetos.Height + vA + vD + vF)
    picCanto.Move fsbVertical.Left, fsbHorizontal.Top, TAM_BARRA, TAM_BARRA
   'Converte a largura e a altura da área de desenho para "centímetros"
    basObjGeometria.Visivel_X = basObjGeometria.Visivel_X * Cm_por_Pixel_X
    basObjGeometria.Visivel_Y = basObjGeometria.Visivel_Y * Cm_por_Pixel_Y
 End With
 Redesenhar = False
  With fsbVertical
   .SmallChange = .Max \ N_INC
   .LargeChange = .Max \ (N_INC \ 5)
   .Value = .Max \ 2
  End With
  With fsbHorizontal
   .SmallChange = .Max \ N_INC
   .LargeChange = .Max \ (N_INC \ 5)
   .Value = .Max \ 2
  End With
 Redesenhar = True
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

Private Sub Form_Resize()
 Dim Visivel_antes_X As Single, Visivel_antes_Y As Single
 Dim AltBarra As Single
 'Estudar a possibilidade de usar Top/Left em lugar de Centro_Y/Centro_X
 
 If Me.ScaleWidth <= 0 Then Exit Sub
 
 'Calcula o número de botões que cabem em cada linha,
 'Determina o número de linhas necessárias para exibí-los,
 'Multiplica pela altura de cada botão e soma espessura das bordas
 With Me.tlbObjetos
  AltBarra = 6 + .ButtonHeight * ((.Buttons.Count - 1) \ Int(Me.ScaleWidth / .ButtonWidth) + 1)
 End With
 'Guarda a largura e a altura visíveis atualmente
 Visivel_antes_X = basObjGeometria.Visivel_X
 Visivel_antes_Y = basObjGeometria.Visivel_Y
 With Me
  'Mede a largura e a altura da área de desenho em "pixels"
   Visivel_X_pix = .ScaleWidth - (TAM_BARRA + vB + vC + vE)
   Visivel_Y_pix = .ScaleHeight - (2 * TAM_BARRA + AltBarra + vA + vD + vF)
  'Posiciona a área de desenho na tela
   .picViewTela.Move .ScaleLeft + vB, .ScaleTop + (AltBarra + vA), Visivel_X_pix, Visivel_Y_pix
  'POsiciona barras de rolagem
   fsbHorizontal.Move picViewTela.Left, vD + picViewTela.Top + Visivel_Y_pix, picViewTela.Width, TAM_BARRA
   fsbVertical.Move vC + picViewTela.Left + Visivel_X_pix, picViewTela.Top, TAM_BARRA, picViewTela.Height
  'Posiciona canto
   picCanto.Move fsbVertical.Left, fsbHorizontal.Top, TAM_BARRA, TAM_BARRA
  'Converte a largura e a altura da área de desenho para "centímetros"
   basObjGeometria.Visivel_X = Visivel_X_pix * Cm_por_Pixel_X
   basObjGeometria.Visivel_Y = Visivel_Y_pix * Cm_por_Pixel_Y
  'Atualiza as coordenadas que correspondem ao centro do form.
   Centro_X = Centro_X + (Visivel_X - Visivel_antes_X) / 2
   Centro_Y = Centro_Y - (Visivel_Y - Visivel_antes_Y) / 2
  'Isto poderia fazer parte de um evento Change_Centro para a futura classe Folha
   If 2 * Centro_X > TAM_FOLHA_DESENHO Then Centro_X = TAM_FOLHA_DESENHO / 2
   If 2 * Centro_X < -TAM_FOLHA_DESENHO Then Centro_X = -TAM_FOLHA_DESENHO / 2
   If 2 * Centro_Y > TAM_FOLHA_DESENHO Then Centro_X = TAM_FOLHA_DESENHO / 2
   If 2 * Centro_Y < -TAM_FOLHA_DESENHO Then Centro_Y = -TAM_FOLHA_DESENHO / 2
  'Funcção inversa de: Centro_X = TAM_FOLHA_DESENHO * (fsbHorizontal.Value / fsbHorizontal.Max - 0.5)
   Redesenhar = False
    fsbHorizontal.Value = CInt((Centro_X / TAM_FOLHA_DESENHO + 0.5) * fsbHorizontal.Max)
    fsbVertical.Value = CInt((Centro_Y / TAM_FOLHA_DESENHO + 0.5) * fsbVertical.Max)
   Redesenhar = True
  'Como evitar que desapareça a área de desenho ao diminuir MUITO a largura e altura?
  'Use API's do windows
  'Timer1.Enabled = True
  'lblDica.Move 10, Me.ScaleTop + tbrObjetos.height + 10
 End With
 
 Call Ajusta_ViewPort(0, 0, Visivel_X_pix, Visivel_Y_pix)
 Call picViewTela_Paint
 
End Sub

Private Sub picViewTela_Click()
If Not frmMatriz.Visible Then frmMatriz.Show    'vbModal
End Sub

Private Sub picViewTela_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Const FORMATO = "0.0"
 Dim X0 As Single, Y0 As Single
 Dim X1 As Single, Y1 As Single
 Dim X2 As Single, Y2 As Single
 Dim s As String
 
 s = "VB: [" & X & "; " & Y & "],  ViewPort: "
 
 X0 = X
 Y0 = Viewport(3) - Y - 1
 s = s & "[" & X0 & "; " & Y0 & "],  Normalizado: "
 
 X1 = 2 * X0 / Viewport(2) - 1
 Y1 = 2 * Y0 / Viewport(3) - 1
 s = s & "[" & Format(X1, FORMATO) & "; " & Format(Y1, FORMATO) & "],  Cm's: "
 
 X2 = (X1 - ProjectionMatrix(12)) / ProjectionMatrix(0)
 Y2 = (Y1 - ProjectionMatrix(13)) / ProjectionMatrix(5)
 s = s & "[" & Format(X2, FORMATO) & "; " & Format(Y2, FORMATO) & "]"

Me.Caption = s


End Sub

Private Sub picViewTela_Paint()
 Dim N As Integer, N_Obj As Integer
 Dim D As GLfloat, Ini As GLfloat, Fim As GLfloat
 Dim Espec As GLfloat
 'Tornar publico esse valor, atualizando quando adicionar ou remover objetos
 
 glClear clrColorBufferBit 'Or clrDepthBufferBit
 N_Obj = UBound(Obj)
 
 For N = 1 To N_Obj
  With Obj(N)
   If .Mostrar <> OCULTO Then
    Select Case .Tipo
    Case PONTO
     glPointSize (.Espessura)
     If .P_rep(3) <> 0 Then 'Ponto finito
      glGetFloatv glgPointSize, Espec
      glPointSize 2
      'If .Mostrar = SELECIONADO Then Me.Circle (Pixel_X(.P_rep(1) / .P_rep(3)), Pixel_Y(.P_rep(2) / .P_rep(3))), 3, vbRed
      glPointSize Espec
      glColor3f 1#, 0#, 0#
      glBegin bmPoints
       glVertex2f .P_rep(1) / .P_rep(3), .P_rep(2) / .P_rep(3)
      glEnd
     End If
    Case PONTO_SOBRE
    
    Case PONTO_DE_INTERSECÇÃO
    
    Case SEGMENTO
    
    Case VETOR
    
    Case RETA
    
    Case SEMI_RETA
    
    Case TRIÂNGULO
    
    Case POLÍGONO
    
    Case POLÍGONO_REGULAR
    
    Case CIRCUNFERÊNCIA
    
    Case ARCO
    
    Case CÔNICA
    
    Case PERPENDICULAR
    
    Case PARALELA
    
    Case PONTO_MÉDIO
    
    Case BISSETRIZ_PONTOS
    
    Case BISSETRIZ_RETAS
    
    Case EIXOS
     glLineWidth .Espessura
     Ini = Centro_X - (Visivel_X / 2)
     Fim = Centro_X + (Visivel_X / 2)
     glColor3f 0, 0, 0 '0.5, 0.5, 0.5
     'glLineStipple 1, 127
     'glEnable glcLineStipple
     glBegin bmLines
      glVertex2f Ini, 0
      glVertex2f Fim, 0
     glEnd
     'glDisable glcLineStipple
     glColor3f 0.2, 0.2, 0.2
     glPointSize .Espessura * 3
     glBegin bmPoints
      For D = CInt(Ini) To CInt(Fim)
        glVertex2f D, 0
      Next D
     glEnd
     
     glLineWidth (.Espessura)
     Ini = Centro_Y - (Visivel_Y / 2)
     Fim = Centro_Y + (Visivel_Y / 2)
     glColor3f 0.5, 0.5, 0.5
     glBegin bmLines
      glVertex2f 0, Ini
      glVertex2f 0, Fim
     glEnd
     
     glColor3f 0.2, 0.2, 0.2
     glPointSize (.Espessura * 3)
     glBegin bmPoints
      For D = CInt(Ini) To CInt(Fim)
        glVertex2f 0, D
      Next D
     glEnd
    Case COMPASSO
    
    Case REFLEXÃO
    
    Case SIMETRIA
    
    Case TRANSLAÇÃO
    
    Case INVERSO_CIRCUNFERÊNCIA
    
    Case TEXTO
    
    Case ÂNGULO
 
    Case Else
     
    End Select
   End If
  End With
 Next N
  
 'Me.PSet (Pixel_X(Visivel_X / 2), Pixel_Y(Visivel_Y / 2)), vbGreen
  'Desenha_esfera
 'MostraEixos
 SwapBuffers hDC1
End Sub
Private Sub GeraBarraStatus()
   Dim i As Integer
   
   With staInfo
    .Height = TAM_BARRA
    '.Width = Me.Width 'Ajustado automaticamente ao definir Align=2
    .Align = vbAlignBottom
    For i = 0 To 6
       .Panels.Add , , , i
    Next i
    .Panels(1).AutoSize = sbrSpring
    .Panels(1).MinWidth = 140
   End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 
 If KeyCode = 27 Then
  Unload Me
  End
 End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Dim x1, y1 As GLdouble
' Dim cx, cy As GLdouble
'
' If Button <> -1 Then Exit Sub
'
' x1 = X / Me.ScaleWidth
' y1 = Y / Me.ScaleHeight
' cx = 1 - 2 * x1
' cy = 2 * y1 - 1
'
 'Call Ajusta_ViewPort(0, 0, Visivel_X_pix, Visivel_Y_pix)
 
 'picViewTela_Paint
 
End Sub


Private Sub Form_Unload(Cancel As Integer)

 Call Finalizar_OpenGL '(hDC1)'será necessário?
 Unload frmMatriz
End Sub

Sub Carrega_Ferramentas()
 Const Arq_INI = "Tabela.ini"
 Dim imgX As ListImage
 Dim btnButton As Button
 
 Dim Qtd As Integer
 Dim FileNumber As Integer ' Variant
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

Private Sub picCanto_Click()
  fsbVertical.Value = fsbVertical.Max \ 2
  fsbHorizontal.Value = fsbHorizontal.Max \ 2
End Sub

Private Sub Timer1_Timer()
'picViewTela_Paint
'Exit Sub
  
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
 If Redesenhar Then
  Centro_X = TAM_FOLHA_DESENHO * (fsbHorizontal.Value / fsbHorizontal.Max - 0.5)
  Call Ajusta_ViewPort(0, 0, Visivel_X_pix, Visivel_Y_pix)
  
  Call picViewTela_Paint
 End If
End Sub
Private Sub fsbHorizontal_Scroll()
'Conferir compatibilidade com o Zoom
 Centro_X = TAM_FOLHA_DESENHO * (fsbHorizontal.Value / fsbHorizontal.Max - 0.5)
 Call Ajusta_ViewPort(0, 0, Visivel_X_pix, Visivel_Y_pix)
 
 Call picViewTela_Paint
End Sub
Private Sub fsbVertical_Change()
 If Redesenhar Then
  Centro_Y = -TAM_FOLHA_DESENHO * (fsbVertical.Value / fsbVertical.Max - 0.5)
  Call Ajusta_ViewPort(0, 0, Visivel_X_pix, Visivel_Y_pix)
  
  Call picViewTela_Paint
 End If
End Sub
Private Sub fsbVertical_Scroll()
'Conferir compatibilidade com o Zoom
 Centro_Y = -TAM_FOLHA_DESENHO * (fsbVertical.Value / fsbVertical.Max - 0.5)
 Call Ajusta_ViewPort(0, 0, Visivel_X_pix, Visivel_Y_pix)
 
 Call picViewTela_Paint
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Exit Sub
 Select Case MsgBox("Deseja salvar as alterações?", vbQuestion + vbYesNoCancel, "Finalizando o aplicativo...")
 Case vbCancel
  Cancel = True
 Case vbNo
  Cancel = False
 Case vbYes
  Cancel = True
  'SalvarArquivo
  'Unload Me 'fechar definitivamente
 End Select
End Sub
