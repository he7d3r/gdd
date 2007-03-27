VERSION 5.00
Begin VB.MDIForm frmMDIGeo3d 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Geo3d"
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9885
   Icon            =   "frmMDIGeo3d.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
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
         Checked         =   -1  'True
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
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "frmMDIGeo3d"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
 ChDir App.Path 'Pode ser útil usar a pasta atual
 
 ReDim Doc(1 To 1)
 ReDim EstadoForm(1)
 Sobre_Plano = PL_HORIZONTAL
 
 Doc(1).frm.Tag = 1
 Doc(1).frm.Caption = "Novo_" & 1
 
 Inicializa_Barra_Ferramentas 1
 Inicializa_Objetos 1
 Inicializa_OpenGL 1  'Ajusta formato dos pixels, iluminação, matrizes de projeção...
 
 Doc(1).frm.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
 'Se não foi cancelado o Unload em nenhum form, não restaram documentos abertos.
 If Not ExisteDocAberto() Then
     End
 End If
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
 Unload frmMDIGeo3d
End Sub

Private Sub mnuEditarDefPLano_Click(Index As Integer)
   Dim i As Tipo_De_Plano
   For i = PL_HORIZONTAL To PL_PERFIL
      mnuEditarDefPLano(i).Checked = IIf(Index = i, True, False)
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
