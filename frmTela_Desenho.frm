VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTela_Desenho 
   BackColor       =   &H00FFFFFF&
   Caption         =   "GDin"
   ClientHeight    =   7155
   ClientLeft      =   165
   ClientTop       =   375
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmTela_Desenho.frx":0000
   MousePointer    =   2  'Cross
   ScaleHeight     =   477
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   760
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstAjuda 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   225
      IntegralHeight  =   0   'False
      ItemData        =   "frmTela_Desenho.frx":030A
      Left            =   600
      List            =   "frmTela_Desenho.frx":0311
      MousePointer    =   5  'Size
      Style           =   1  'Checkbox
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   300
      LargeChange     =   10
      Left            =   360
      Max             =   50
      Min             =   -50
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2280
      Width           =   2265
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   6780
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstCursor 
      Left            =   1080
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483624
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":031B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":0635
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":094F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":0C69
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":0F83
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":129D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":15B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":18D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":1BEB
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":1F05
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":221F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":2381
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":269B
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":27FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":2A2B
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":2B8D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCanto 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   11040
      MousePointer    =   1  'Arrow
      ScaleHeight     =   300
      ScaleWidth      =   300
      TabIndex        =   3
      Top             =   6840
      Width           =   300
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   960
      Top             =   3240
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2265
      LargeChange     =   10
      Left            =   0
      Max             =   -50
      Min             =   50
      MousePointer    =   1  'Arrow
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2640
      Width           =   300
   End
   Begin MSComctlLib.ImageList imglstFerramenta 
      Left            =   375
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   33
      ImageHeight     =   33
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":2CEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":3A25
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":475B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":5491
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":61C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":6EFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":7C33
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":8969
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":969F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":A3D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":B10B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":BE41
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":CB77
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":D8AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":E5E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":F319
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":1004F
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":10D85
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":11ABB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrObjetos 
      Align           =   1  'Align Top
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1191
      ButtonWidth     =   1058
      ButtonHeight    =   1032
      Appearance      =   1
      ImageList       =   "imglstFerramenta"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PONTEIRO"
            Object.ToolTipText     =   "Seleção de objetos"
            ImageIndex      =   1
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PONTO"
            Object.ToolTipText     =   "Ponto livre"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PONTO_SOBRE"
            Object.ToolTipText     =   "Ponto sobre objeto"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PONTO_DE_INTERSECÇÃO"
            Object.ToolTipText     =   "Pontos de Interseção"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SEGMENTO"
            Object.ToolTipText     =   "Segmento de reta"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "VETOR"
            Object.ToolTipText     =   "Vetor"
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SEMI_RETA"
            Object.ToolTipText     =   "Semi-reta"
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RETA"
            Object.ToolTipText     =   "Vetor"
            ImageIndex      =   8
            Style           =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TRIÂNGULO"
            Object.ToolTipText     =   "Triângulo"
            ImageIndex      =   9
            Style           =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "POLÍGONO"
            Object.ToolTipText     =   "Polígono"
            ImageIndex      =   10
            Style           =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "POLÍGONO_REGULAR"
            Object.ToolTipText     =   "Polígono Regular"
            ImageIndex      =   11
            Style           =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CIRCUNFERÊNCIA"
            Object.ToolTipText     =   "Circunferência"
            ImageIndex      =   12
            Style           =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ARCO"
            Object.ToolTipText     =   "Arco de circunferência"
            ImageIndex      =   13
            Style           =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CÔNICA"
            Object.ToolTipText     =   "Cônica por 5 pontos"
            ImageIndex      =   14
            Style           =   2
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PARALELA"
            Object.ToolTipText     =   "Reta paralela "
            ImageIndex      =   15
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PERPENDICULAR"
            Object.ToolTipText     =   "Reta perpendicular"
            ImageIndex      =   16
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MEDIATRIZ"
            Object.ToolTipText     =   "Medriatriz"
            ImageIndex      =   17
            Style           =   2
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PONTO_MÉDIO"
            Object.ToolTipText     =   "Ponto médio"
            ImageIndex      =   18
            Style           =   2
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "BISSETRIZ"
            Object.ToolTipText     =   "Bissetriz"
            ImageIndex      =   19
            Style           =   2
         EndProperty
      EndProperty
      MousePointer    =   1
   End
   Begin VB.ListBox lstObjetos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      Height          =   225
      IntegralHeight  =   0   'False
      ItemData        =   "frmTela_Desenho.frx":127F1
      Left            =   2760
      List            =   "frmTela_Desenho.frx":127F8
      MousePointer    =   1  'Arrow
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblAux 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<<  AUX  >>"
      ForeColor       =   &H80000017&
      Height          =   225
      Left            =   2760
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblDica 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Dica... "
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   600
      MouseIcon       =   "frmTela_Desenho.frx":12802
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Menu mnuArq 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuArqSair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "A&juda"
      Begin VB.Menu mnuAjudaSobre 
         Caption         =   "&Sobre"
      End
   End
End
Attribute VB_Name = "frmTela_Desenho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private multi_sel As Boolean

Private Const DICA = " Dica... "
Private Const TAM_BARRA = 20
Private Const DIST_MIN = 8 'pixels
Private Altura_linha As Single
Private Xant, Yant As Single
'Private Const
'Private Const

Private Sub Form_Initialize()
 Dim i As Integer
 Call Inicializar_Parametros
 Me.Caption = "Geometria dinâmica"
 MakeEight
  With Me
  .BackColor = vbWhite
  HScroll1.Move .ScaleLeft, .ScaleTop + .ScaleHeight - TAM_BARRA - StatusBar1.Height, .ScaleWidth - TAM_BARRA, TAM_BARRA
  VScroll1.Move .ScaleLeft + .ScaleWidth - TAM_BARRA, .ScaleTop + tbrObjetos.Height, TAM_BARRA, .ScaleHeight - tbrObjetos.Height - TAM_BARRA
  HScroll1.Min = -MAX_X: HScroll1.Max = MAX_X
  VScroll1.Min = MAX_Y: VScroll1.Max = -MAX_Y
 End With
 With tbrObjetos.Buttons
  tbrObjetos.Tag = .Item(1).Key
  For i = 3 To .Count
   .Item(i).Enabled = False
  Next i
 End With
 With lstObjetos
  .Clear
  .AddItem "Medindo linha..."
  .Height = 0
  Altura_linha = .Height
  .Clear
  Me.lblAux.Font.Size = .Font.Size
  
  lblDica.BackColor = vbRed
  lblDica.ForeColor = vbBlack
  lblDica.Caption = DICA
  lblDica.Move 10, Me.ScaleTop + tbrObjetos.Height + 10
 End With
 
End Sub

Private Sub Aponta_Objeto(ByVal X As Single, ByVal Y As Single, ByRef Loc() As Long)
 
 Dim N, N_Obj As Integer
 Dim Cor_Ponto_XY As Long
 Dim X_real As Single, Y_real As Single
 
 'Se NUNCA for pintado um objeto de BRANCO, pode-se (?) incluir isso:
 'If Cor_Ponto_XY = Me.BackColor Then Exit Sub
 
 X_real = Cm_X(X): Y_real = Cm_Y(Y)
 
 'Ocorre erro se não houver objetos.
 'Mas Obj() tem (ao menos) um PONTO e um par de EIXOS
 N_Obj = UBound(Obj)
 ReDim Objeto_Localizado(1 To 1)
 Objeto_Localizado(1) = NENHUM
 
 'Avalia cada objeto da lista Obj()
 'Se um item está visível e perto do mouse:
 ' é guardado em Objeto_Localizado()
 ' é alocado espaço para um novo item
 '
 For N = 1 To N_Obj
  With Obj(N)
   If .Mostrar <> OCULTO Then
    Select Case .Tipo
     Case PONTO, PONTO_SOBRE, PONTO_DE_INTERSECÇÃO, PONTO_MÉDIO
      'Me.MousePointer = vbSizeAll
      'If Abs(X_real - .P_rep(1)) < DIST_MIN And _
      '   Abs(Y_real - .P_rep(2)) < DIST_MIN Then Loc = N: Exit Sub
         
      If Abs(X - Pixel_X(.P_rep(1))) < DIST_MIN And _
         Abs(Y - Pixel_Y(.P_rep(2))) < DIST_MIN Then
         Loc(UBound(Loc)) = N 'último objeto localizado tem id=N
         ReDim Preserve Loc(1 To UBound(Loc) + 1) 'Aloque espaço para guardar mais um objeto
         Objeto_Localizado(UBound(Loc)) = NENHUM 'Este objeto ainda não foi encontrado
         'Exit Sub
      End If
     Case SEGMENTO, VETOR
     
     Case SEMI_RETA
     
     Case RETA, PARALELA, PERPENDICULAR, MEDIATRIZ, BISSETRIZ_PONTOS, BISSETRIZ_RETAS
     
    End Select
   End If
  End With
 Next N
 'Loc(1) = NENHUM
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'se (botao = seta)
'  se (aponta objeto)
'   destaque objeto da lista. //colorir ou animar.
'   guarda: Ainda NÃO foi usada a seta provosória. (isso está no lugar certo?)
'   se (shift pressionado)
'    adiciona à seleção atual
'    [ seleção é uma matriz ou uma PROPRIEDADE dos objetos? ]
'   senão
'    seleciona apenas este objeto
'  senao
'   inicia seleção múltipla //guarde x,y
'senão
'  cancele todas as seleções atuais
'  [nao faça nada até ser movido o mouse]

 '-Permitir seleção personalizada, com caixas de verificação para tipos de objetos...
 '-Seleção avançada é parte de um outro botão
 
 Dim i As Long
 
 Select Case UCase(tbrObjetos.Tag)
  Case "PONTEIRO"
   'Atualiza seleção atual
   If Objeto_Localizado(1) <> NENHUM Then
    If Shift And vbShiftMask Then
     Obj(Objeto_Localizado(1)).Mostrar = (SELECIONADO + PADRAO) - Obj(Objeto_Localizado(1)).Mostrar
    Else
     Obj(Objeto_Localizado(1)).Mostrar = SELECIONADO
     'Dica: Tente trocar "UBound(Obj)" por uma variável pública
     'Cancele todas as demais seleções
     For i = 1 To UBound(Obj)
      If (Obj(i).Mostrar = SELECIONADO) And (i <> Objeto_Localizado(1)) Then Obj(i).Mostrar = PADRAO
     Next i
    End If
   Else
    'multi_sel = True
   End If
     'Isto deve estar no MouseMOVE ou aqui???
  Case "SEGMENTO", "VETOR", "SEMI_RETA", "RETA", _
       "TRIÂNGULO", "POLÍGONO", "POLÍGONO_REGULAR", _
       "CIRCUNFERÊNCIA", "ARCO", "CÔNICA"
   'Estes são os objetos que exibem prévia de sua posição enquanto são criados
   'Pode ser problema no "Padrão Cabri/cinderella": botao-->parametros
   'Solução simples no KSeg: Parâmetros-->Botao
  Case Else
  'Cancele todas as seleções
   For i = 1 To UBound(Obj)
    If (Obj(i).Mostrar = SELECIONADO) Then Obj(i).Mostrar = PADRAO
   Next i
 End Select
  
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'se (não está clicado)
'  se (aponta ALGUM objeto)
'   se (botao = seta)
'    exibe lista de possíveis objetos para seleção
'   senao
'    se (existe ponto provisório) mova-o
'    exibe ajuda (chkbox?) para a construção (qual o objeto necessário?)
'senão, se (iniciada seleção múltipla)
'  desenhe retângulo pontilhado
'senão se (selecionado apenas um objeto)
'  se (botao = seta)
'     se (é livre)
'      mova objeto e calcule a nova posição de todos os que dependem dele
'     senão
'      Kseg: Move os pontos DEPENDENTES do atual e os LIVRES dos quais depende
'      Cabri: Não faz nada, nem avisa
'      Opção: Mover toda a construção OU indicar movimento proibido
'  senao
'   seleciona seta enquanto estiver clicado
'senão
'    Transladar os pontos LIVRES que definem os objetos (indiretamente ou não)
'
 Dim i As Integer
 Dim Bkp() As Long
 Dim Mudou As Boolean
 Dim Larg As Single
 Dim Xpos, Ypos As Single
 
 'usar o comando abaixo na ajuda para uma construção:
 'lblAux.Move Xpos, Ypos - lblAux.Height - 2, .Width
 
 Xpos = X + 1 + 2 * DIST_MIN: Ypos = Y + 1 + 2 * DIST_MIN
 
 Bkp = Objeto_Localizado
 With lstObjetos
  If Button = NENHUM Then
   Call Aponta_Objeto(X, Y, Objeto_Localizado)
   If Objeto_Localizado(1) <> NENHUM Then
    For i = 1 To UBound(Objeto_Localizado) - 1
     If Bkp(i) <> Objeto_Localizado(i) Then Mudou = True: Exit For
    Next i
    If Mudou Then
     .Visible = False
     .Clear
     lblAux = ""
     For i = 1 To UBound(Objeto_Localizado) - 1
      .AddItem "Objeto " & Format(Objeto_Localizado(i), "00") & _
             " (" & Nome(Obj(Objeto_Localizado(i)).Tipo) & ")"
      If Len(lblAux) < Len(.List(i - 1)) Then
       lblAux = .List(i - 1)
       Larg = lblAux.Width
      End If
     Next i
    End If
   End If
  End If
  If Objeto_Localizado(1) = NENHUM Then
   .Visible = False
   lblAux.Visible = False
   Me.MousePointer = ccCross
  Else
   If Mudou Then
    lblAux = " " & .List(0) & " "
    Select Case UBound(Objeto_Localizado)
    Case 2 'Um objeto localizado
     lblAux.Move Xpos, Ypos
     .Visible = False
     Me.MousePointer = 99
    Case Else 'Mais de um objeto localizado
     lblAux = "Um destes objetos..."
     lblAux.Move Xpos, Ypos
     Me.MousePointer = ccArrowQuestion
     .Move Xpos, Ypos, Larg + 18, Altura_linha * (.ListCount) + 2
     'lblAux.Visible = False
     '.Visible = True
    End Select
    lblAux.Visible = True
   End If
  End If
 End With
 'Trocar este item por um label que aparece ao ser clicado longe de objetos
 StatusBar1.Panels(1).Text = "Posição atual: [ " & Format(Cm_X(X), "0.0") & " ;  " & Format(Cm_Y(Y), "0.0") & "]"


   'Se for executado, saberemos quando muda o TwipsPerPixel da tela!
   If Screen.TwipsPerPixelX <> TwipsPerPixelX_INICIAL Then
    MsgBox "TwipsPerPixelX mudou de " & TwipsPerPixelX_INICIAL & "para " & Screen.TwipsPerPixelX
   End If
   If Screen.TwipsPerPixelY <> TwipsPerPixelY_INICIAL Then
    MsgBox "TwipsPerPixelY mudou de " & TwipsPerPixelY_INICIAL & "para " & Screen.TwipsPerPixelY
   End If
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Se (iniciada seleção múltipla)
' Efetiva seleção (destaca todos os objetos
' com ALGUM ponto dentro do retângulo de seleção)
'Senão se (botao=seta provisoria)
' retorna ao botao anterior
' guarda: Foi liberada a seta provosória.
' (apenas foram movidos alguns objetos. Nada + para fazer)
'senao //Cria pontos e objetos para as diversas ferramentas
' Se (Aponta_Objeto)
'  Se (botão=construção)
'   Se (constução depende deste tipo de objeto)
'    Use objeto atual na construção
'   Senão se (construção precisa de + algum ponto)
'    Se (objeto atual não é um ponto)
'     se(é possível por ponto sobre este objeto)
'      Cria (ponto_sobre_objeto)
'     senão
'      Cria (ponto comum)
'   Senão
'    ignora mouse_up (sai sem fazer nada e aguarda o próximo)
'  senão se (construção precisa de + algum ponto)
'   cria (Ponto_livre)
'  senão
'   aguarda o procimo mouse up
 
  
If (Not multi_sel) Then
  Select Case UCase(tbrObjetos.Tag)
    Case "PONTEIRO"
     If UBound(Objeto_Localizado) > 2 Then
      lblAux.Visible = False
      Me.lstObjetos.Visible = True
     End If
    Case "PONTO"
    'Param=1
      ReDim Preserve Obj(1 To 1 + UBound(Obj))
      With Obj(UBound(Obj))
        '.Tipo = PONTO
        '.Nome = ""
        '.N_Param = 0
        ReDim .P_rep(1 To 3)
        .P_rep(1) = Cm_X(X)
        .P_rep(2) = Cm_Y(Y)
        .P_rep(3) = 1# 'para coordenadas homogêneas
        .Espessura = 4#
        .Mostrar = PADRAO
        
        '.Traço(0) = 1: .Traço(2) = 1'Irrelevante para pontos.
        ' Usar para indicar forma, como X,O, . ou + ...
      End With
    Case "PONTO_SOBRE"
     'Param=2
    Case "PONTO_DE_INTERSECÇÃO"
    'Param=2
    Case "SEGMENTO", "VETOR"
    'Param=2
      With Obj(UBound(Obj))
        .Tipo = tbrObjetos.Buttons(tbrObjetos.Tag) 'SEGMENTO ou VETOR
        .Nome = ""
        .N_Param = 2
        ReDim .P_dep(1 To .N_Param)
        'decidir como acessar parâmetros. Serão guardados id's em P()??
        .P_dep(1) = P(1): .P_dep(2) = P(2) '(x,y) do ponto  inicial
        .P_dep(3) = P(3): .P_dep(4) = P(4) '(x,y) do ponto  final
        
        .Traço(1) = 1: .Traço(2) = 1
        .Cor = 0
        .Espessura = 4#
        .Mostrar = PADRAO
      End With
    Case "RETA"
    'Param=2
    Case "SEMI_RETA"
    'Param=2
    Case "TRIÂNGULO"
    'Param=3
    Case "POLÍGONO"
    'Param=N
    Case "POLÍGONO_REGULAR"
    'Param=3
    Case "EIXOS"
    'Param=3
    Case "CIRCUNFERÊNCIA"
    'Param=2
    Case "ARCO"
    'Param=3
    Case "CÔNICA"
    'Param=5
    Case "PERPENDICULAR"
    'Param=2
    Case "PARALELA"
    'Param=2
    Case "PONTO_MÉDIO"
    'Param=1 ou 2
    Case "BISSETRIZ_PONTOS"
    'Param=3
    Case "BISSETRIZ_RETAS"
    'Param=2
    Case "COMPASSO"
    'Param=2 ou 3
    Case "REFLEXÃO"
    'Param=2
    Case "SIMETRIA"
    'Param=2
    Case "TRANSLAÇÃO"
    'Param=2
    Case "INVERSO_CIRCUNFERÊNCIA"
    'Param=2
    Case "TEXTO"
    'Param=1 + texto
    Case "ÂNGULO"
    'Param=2 ou 3
    Case Else
  End Select
  End If
  Me.Refresh
End Sub

Private Sub MakeEight()
   ' Delete the first Panel object, which is
   ' created automatically.
   'StatusBar1.Panels.Remove 1
   Dim i As Integer

   ' The fourth argument of the Add method
   ' sets the Style property.
   For i = 0 To 6
      StatusBar1.Panels.Add , , , i
   Next i
   StatusBar1.Panels(1).AutoSize = sbrSpring
   StatusBar1.Panels(1).MinWidth = 140
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
 With Source
  Select Case .Name
  Case "lstAjuda"
   .Move X - 1 - Xant / TwipsPerPixelX_INICIAL, Y - 1 - Yant / TwipsPerPixelY_INICIAL
   lblDica.Move .Left, .Top - lblDica.Height - 5
  Case "lblDica"
   If .BackColor = vbRed Then
    .MousePointer = ccArrowQuestion
   Else
    .MousePointer = ccSize
   End If
   .Move X - Xant / TwipsPerPixelX_INICIAL, Y - Yant / TwipsPerPixelY_INICIAL
   lstAjuda.Move .Left, .Top + .Height + 5
  End Select
 End With
End Sub

Private Sub lblDica_Click()
 With lblDica
  If .BackColor = vbRed Then
   .BackColor = vbActiveTitleBar
   .ForeColor = vbActiveTitleBarText
   .Caption = "Itens necessários..."
   .MousePointer = ccSize
   lstAjuda.DragMode = 1
   lstAjuda.Move lblDica.Left, lblDica.Top + lblDica.Height + 5
   lstAjuda.Visible = True
  Else
   .BackColor = vbRed
   .ForeColor = vbBlack
   .Caption = DICA
   .MousePointer = ccArrowQuestion
   lstAjuda.DragMode = 0
   lstAjuda.Visible = False
  End If
 End With
End Sub

Private Sub lblDica_DragDrop(Source As Control, X As Single, Y As Single)
 With Source
  Select Case .Name
  Case "lblDica"
   If .BackColor = vbRed Then
    .MousePointer = ccArrowQuestion
   Else
    .MousePointer = ccSize
   End If
   If (X = Xant) And (Y = Yant) Then
    lblDica_Click
   Else
     .Move .Left + (X - Xant) / TwipsPerPixelX_INICIAL, _
            .Top + (Y - Yant) / TwipsPerPixelY_INICIAL
     lstAjuda.Move .Left, .Top + .Height + 5
   End If
  Case "lstAjuda"
   With lblDica
    lstAjuda.Move (.Left - 1) + (X - Xant) / TwipsPerPixelX_INICIAL, _
                   (.Top - 1) + (Y - Yant) / TwipsPerPixelY_INICIAL
    .Move lstAjuda.Left, lstAjuda.Top - .Height - 5
    End With
  End Select
 End With
End Sub

Private Sub lblDica_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 lblDica.Drag vbBeginDrag
 Xant = X: Yant = Y 'medido em twips. Exigirá conversão para pixels ao soltar.
 lblDica.MousePointer = 15
End Sub
Private Sub lstAjuda_DragDrop(Source As Control, X As Single, Y As Single)
 With Source
  Select Case .Name
  Case "lblDica"
   If .BackColor = vbRed Then
    .MousePointer = ccArrowQuestion
   Else
    .MousePointer = ccSize
   End If
   lblDica.Move (lstAjuda.Left + 1) + (X - Xant) / TwipsPerPixelX_INICIAL, _
                 (lstAjuda.Top + 1) + (Y - Yant) / TwipsPerPixelY_INICIAL
   lstAjuda.Move .Left, .Top + .Height + 5
  Case "lstAjuda"
    .Move .Left + (X - Xant) / TwipsPerPixelX_INICIAL, _
           .Top + (Y - Yant) / TwipsPerPixelY_INICIAL
    lblDica.Move .Left, .Top - lblDica.Height - 5
  End Select
 End With
End Sub

Private Sub lstAjuda_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Xant = X: Yant = Y 'medido em twips. Exigirá conversão para pixels ao soltar.
End Sub

Private Sub mnuAjudaSobre_Click()
 MsgBox "Este software para geometria está em fase de construção." & vbCrLf & "Use por SUA conta e risco...", vbCritical + vbSystemModal, "Aviso do Helder"
End Sub

Private Sub mnuArqSair_Click()
 Unload Me
End Sub

Private Sub tbrObjetos_ButtonClick(ByVal Button As MSComctlLib.Button)
 'Se (existem objetos selecionados)
 '  Se (determinam a construção do botão atual)
 '    Cria objeto a partir dos atuais (aqui, no mouseUP, ou em ambos???)
 '  Senão
 '    Desfaça a seleção atual
 'Senão
 '  Inicia construção
 '  (pode exibir CheckList que indique os objetos necessários)
 
 'Convenção: Tag guarda um nome igual aos da enumeração pública de tipos dos objetos
 tbrObjetos.Tag = Button.Key
 
 If tbrObjetos.Tag = "PONTEIRO" Then
  lblDica.Visible = False
  lstAjuda.Visible = False
 Else
  lblDica = DICA
  lblDica.Move 10, Me.ScaleTop + tbrObjetos.Height + 10
  lblDica.Visible = True
  lblDica.BackColor = vbRed
  lblDica.ForeColor = vbBlack
 End If
 
 'UBound(P) = Número de objetos selecionados atualmente
 'como verificar se está vazio??? Estará vazio em algum momento???
On Error GoTo Erro_sem_objetos
 Select Case UBound(P)
 Case 1
 
  'habilita formatação e + algo?
 Case 2
 
 Case 3

 Case Else
  'Se não é possível construir um item,
  'desmarque todos os objetos selecionados
  'apenas inicie uma construção com a ferramenta atual
 End Select
On Error GoTo 0
Erro_sem_objetos:
End Sub
Private Sub Form_Paint()
 Dim N As Integer, N_Obj As Integer
 Dim D As Integer, Ini As Single, Fim As Single
 Dim Espec As Long
 'Tornar publico esse valor, atualizando quando adicionar ou remover objetos
 N_Obj = UBound(Obj)
 
 For N = 1 To N_Obj
  With Obj(N)
   If .Mostrar <> OCULTO Then
    Select Case .Tipo
    Case PONTO
     Me.DrawWidth = .Espessura
     If .P_rep(3) <> 0 Then
      Espec = Me.DrawWidth: Me.DrawWidth = 2
      If .Mostrar = SELECIONADO Then Me.Circle (Pixel_X(.P_rep(1) / .P_rep(3)), Pixel_Y(.P_rep(2) / .P_rep(3))), 3, vbRed
      Me.DrawWidth = Espec
      Me.PSet (Pixel_X(.P_rep(1) / .P_rep(3)), Pixel_Y(.P_rep(2) / .P_rep(3))), .Cor
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
     Me.DrawWidth = .Espessura
     
     Ini = Centro_X - (Visivel_X / 2)
     Fim = Centro_X + (Visivel_X / 2)
     Me.Line (Pixel_X(CSng(Ini)), Pixel_Y(0)) _
            -(Pixel_X(CSng(Fim)), Pixel_Y(0)), .Cor
     Me.DrawWidth = .Espessura * 3
      For D = Ini To Fim
       Me.PSet (Pixel_X(CSng(D)), Pixel_Y(0)), .Cor
      Next D
     Me.DrawWidth = .Espessura
     Ini = Centro_Y - (Visivel_Y / 2)
     Fim = Centro_Y + (Visivel_Y / 2)
     Me.Line (Pixel_X(0), Pixel_Y(CSng(Ini))) _
            -(Pixel_X(0), Pixel_Y(CSng(Fim))), .Cor
     Me.DrawWidth = .Espessura * 3
      For D = Ini To Fim
       Me.PSet (Pixel_X(0), Pixel_Y(CSng(D))), .Cor
      Next D
     
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
 
End Sub
Private Sub Form_Resize()
 Dim Visivel_antes_X As Single, Visivel_antes_Y As Single

 Visivel_antes_X = Visivel_X
 Visivel_antes_Y = Visivel_Y
 With Me
  'Mede a largura e a altura da área de desenho em "pixels"
  Visivel_X = .ScaleWidth - .VScroll1.Width
  Visivel_Y = .ScaleHeight - (tbrObjetos.Height + HScroll1.Height)
  'Converte a largura e a altura da área de desenho para "centímetros"
  Visivel_X = Visivel_X * TwipsPerPixelX_INICIAL / Twips_por_Cm
  Visivel_Y = Visivel_Y * TwipsPerPixelY_INICIAL / Twips_por_Cm
  'Atualiza as coordenadas que correspondem ao centro do form.
  Centro_X = Centro_X + (Visivel_X - Visivel_antes_X) / 2
  Centro_Y = Centro_Y - (Visivel_Y - Visivel_antes_Y) / 2
  
  
  On Error Resume Next
  'Como evitar que desapareça a área de desenho ao diminuir a largura e altura?
  HScroll1.Move .ScaleLeft, .ScaleTop + .ScaleHeight - TAM_BARRA - StatusBar1.Height, .ScaleWidth - TAM_BARRA, TAM_BARRA
  VScroll1.Move .ScaleLeft + .ScaleWidth - TAM_BARRA, .ScaleTop + tbrObjetos.Height, TAM_BARRA, .ScaleHeight - tbrObjetos.Height - TAM_BARRA - StatusBar1.Height
  picCanto.Move VScroll1.Left, HScroll1.Top
  On Error GoTo 0
  Timer1.Enabled = True
  
  .Cls
  .Refresh
  lblDica.Move 10, Me.ScaleTop + tbrObjetos.Height + 10
 End With
End Sub
Private Function Pixel_X(X_real As Single) As Single
'Incluir uma constante Pixel_por_Cm_X e outra Pixel_por_Cm_Y em algum lugar do programa...

 Pixel_X = (Me.ScaleWidth - VScroll1.Width) / 2 + _
       Zoom * Twips_por_Cm * (X_real - Centro_X) / TwipsPerPixelX_INICIAL
 
End Function
Private Function Pixel_Y(Y_real As Single) As Single
 Pixel_Y = tbrObjetos.Height + (Me.ScaleHeight - tbrObjetos.Height - HScroll1.Height) / 2 _
     - Zoom * Twips_por_Cm * (Y_real - Centro_Y) / TwipsPerPixelY_INICIAL
 
End Function
Private Function Cm_X(P_X As Single) As Single
  Cm_X = Centro_X + TwipsPerPixelX_INICIAL * (P_X - (Me.ScaleWidth - VScroll1.Width) / 2) / _
         (Zoom * Twips_por_Cm)
         
End Function
Private Function Cm_Y(P_Y As Single) As Single
  Cm_Y = Centro_Y - TwipsPerPixelY_INICIAL * (P_Y - tbrObjetos.Height - (Me.ScaleHeight - tbrObjetos.Height - HScroll1.Height) / 2) / _
         (Zoom * Twips_por_Cm)
         
End Function
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case vbKeyEscape
   tbrObjetos.Buttons(1).Value = tbrPressed
   tbrObjetos.Tag = tbrObjetos.Buttons(1).Key
   Me.lblDica.Visible = False
   Me.lstAjuda.Visible = False
  Case vbKeyDown
   If VScroll1.Value <> -MAX_Y Then VScroll1.Value = VScroll1.Value - VScroll1.SmallChange
  Case vbKeyUp
   If VScroll1.Value <> MAX_Y Then VScroll1.Value = VScroll1.Value + VScroll1.SmallChange
  Case vbKeyLeft
   If HScroll1.Value <> -MAX_X Then HScroll1.Value = HScroll1.Value - HScroll1.SmallChange
  Case vbKeyRight
   If HScroll1.Value <> MAX_X Then HScroll1.Value = HScroll1.Value + HScroll1.SmallChange
  Case vbKeyPageDown
   If VScroll1.Value - VScroll1.LargeChange > -MAX_Y Then
    VScroll1.Value = VScroll1.Value - VScroll1.LargeChange
   Else
    VScroll1.Value = -MAX_Y
   End If
  Case vbKeyPageUp
   If VScroll1.Value + VScroll1.LargeChange < MAX_Y Then
    VScroll1.Value = VScroll1.Value + VScroll1.LargeChange
   Else
    VScroll1.Value = MAX_Y
   End If
  Case Else
   'MsgBox KeyCode
 End Select
End Sub

Private Sub Timer1_Timer()
'Esse timer só existe para contornar um defeito na rotina Resize...
'O programa nao conhece a altura real da barra no instante do redimensionamento, só depois
With Me
 On Error Resume Next
 VScroll1.Move .ScaleLeft + .ScaleWidth - TAM_BARRA, .ScaleTop + tbrObjetos.Height, TAM_BARRA, .ScaleHeight - tbrObjetos.Height - TAM_BARRA
 On Error GoTo 0
 Timer1.Enabled = False
End With
End Sub

Private Sub HScroll1_Change()
'Tratar o bug que ocorre ao redimensionar o form:
'a rotina resize muda para decimais o valor dos centros X e Y,
'mas as scroll's tm value inteiro
 Centro_X = HScroll1.Value
 Me.Refresh
End Sub
Private Sub HScroll1_Scroll()
'Incluir mais valores entre os extremos de MAX_X e -MAX_X permitirá um scroll mais suave
'Conferir compatibilidade com o Zoom
 Centro_X = HScroll1.Value
 Me.Refresh
End Sub

Private Sub VScroll1_Change()
 Centro_Y = VScroll1.Value
 Me.Refresh
End Sub
Private Sub VScroll1_Scroll()
'Incluir mais valores entre os extremos de MAX_Y e -MAX_Y permitirá um scroll mais suave
'Conferir compatibilidade com o Zoom
 Centro_Y = VScroll1.Value
 Me.Refresh
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
 End Select
End Sub
