VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMDIGDD 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "GDD"
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9885
   Icon            =   "frmMDI_GDD.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar staInfo 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   9195
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuArquivoSalvar 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuArquivoNovo 
         Caption         =   "&Novo documento"
      End
      Begin VB.Menu mnuArquivoSair 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu mnuEditar 
      Caption         =   "&Editar"
      Begin VB.Menu mnuEditarDefPontos 
         Caption         =   "Definir &pontos"
         Begin VB.Menu mnuEditarDefPlano 
            Caption         =   "Usando planos &horizontais"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuEditarDefPlano 
            Caption         =   "Usando planos &frontais"
            Index           =   2
         End
         Begin VB.Menu mnuEditarDefPlano 
            Caption         =   "Usando planos de &perfil"
            Index           =   3
         End
      End
      Begin VB.Menu mnuEditarMagnetismo 
         Caption         =   "&Magnetismo"
      End
      Begin VB.Menu mnuEditarSelecionarTudo 
         Caption         =   "Selecionar &Tudo"
      End
      Begin VB.Menu mnuEditarInverter 
         Caption         =   "&Inverter Seleção"
      End
   End
   Begin VB.Menu mnuExibir 
      Caption         =   "E&xibir"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "A&juda"
      Begin VB.Menu mnuAjudaSobre 
         Caption         =   "&Sobre..."
      End
   End
End
Attribute VB_Name = "frmMDIGDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
 ChDir App.Path 'Pode ser útil usar a pasta atual
 GeraBarraStatus
 
 ReDim Doc(1 To 1)
 Sobre_Plano = PL_HORIZONTAL
 
 Doc(1).frm.Tag = 1
 Doc(1).frm.Caption = "Novo_" & 1
 
 Inicializa_Barra_Ferramentas 1
 Inicializa_Objetos 1
 Inicializa_OpenGL 1  'Ajusta formato dos pixels, iluminação, matrizes de projeção...
 
 Doc(1).frm.Show
End Sub
Private Sub GeraBarraStatus()
   Const TAM_BARRA = 400
   Dim i As Integer

   With staInfo
      For i = 1 To 6
         .Panels.Add  'Index, Key, Text, Style
      Next i
      With .Panels
         .Item(1).Style = sbrText
         .Item(1).AutoSize = sbrSpring
         .Item(1).MinWidth = 140
         
         .Item(2).Style = sbrText
         .Item(2).AutoSize = sbrSpring
         .Item(2).MinWidth = 140
         
         .Item(3).Style = sbrNum
         .Item(4).Style = sbrIns
         .Item(5).Style = sbrScrl
         .Item(6).Style = sbrTime
         .Item(7).Style = sbrDate

      End With
      .Height = TAM_BARRA
      .Align = vbAlignBottom
   End With
   
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
 'Se não foi cancelado o Unload em nenhum form, não restaram documentos abertos.
 If Not ExisteDocAberto() Then
     End
 End If
End Sub

Private Sub mnuAjudaSobre_Click()
frmSplash1.Show vbModal, Me
End Sub

Private Sub mnuArquivoNovo_Click()
 Dim Id As Integer
 Dim Existe As Integer
 'Busca o próximo índice disponível e exibe o form
 Id = GeraIdLivre
 Doc(Id).frm.Tag = Id
 Doc(Id).frm.Caption = "Novo_" & Id
 
 Existe = ExisteDocAberto
 If Existe Then Doc(Id).frm.WindowState = Doc(Existe).frm.WindowState
 
 Inicializa_Barra_Ferramentas Id
 Inicializa_Objetos Id
 Inicializa_OpenGL Id  'Ajusta formato dos pixels, iluminação, matrizes de projeção...
 
 Doc(Id).frm.Show
End Sub

Private Sub mnuArquivoSair_Click()
 'Descarregando o form MDI, ocorrerá o evento QueryUnload
 'para cada form filho, seguido do próprio MDI
 Unload frmMDIGDD
End Sub

Private Sub mnuEditarDefPLano_Click(Index As Integer)
   Dim i As Tipo_De_Plano
   For i = PL_HORIZONTAL To PL_PERFIL
      mnuEditarDefPlano(i).Checked = IIf(Index = i, True, False)
   Next i
   Sobre_Plano = Index
End Sub

Private Sub mnuEditarInverter_Click()
   Inverter_Todos ActiveForm.Tag
End Sub

Private Sub mnuEditarMagnetismo_Click()
   With mnuEditarMagnetismo
      .Checked = Not .Checked
   End With
End Sub

Private Sub mnuEditarSelecionarTudo_Click()
   Marcar_Todos ActiveForm.Tag, True
End Sub
