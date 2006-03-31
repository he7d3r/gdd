VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTela_Desenho 
   Caption         =   "GDin"
   ClientHeight    =   7185
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11415
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   761
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.Toolbar tbrObjetos 
      Align           =   1  'Align Top
      Height          =   675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   1191
      ButtonWidth     =   1058
      ButtonHeight    =   1032
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ponteiro"
            Object.ToolTipText     =   "Seleção de objetos"
            ImageIndex      =   1
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ponto_livre"
            Object.ToolTipText     =   "Ponto livre"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ponto_sobre"
            Object.ToolTipText     =   "Ponto sobre objeto"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ponto_inter"
            Object.ToolTipText     =   "Pontos de Interseção"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Segmento"
            Object.ToolTipText     =   "Segmento de reta"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Vetor"
            Object.ToolTipText     =   "Vetor"
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Semi_reta"
            Object.ToolTipText     =   "Segmento de reta"
            ImageIndex      =   7
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reta"
            Object.ToolTipText     =   "Vetor"
            ImageIndex      =   8
            Style           =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Trângulo"
            Object.ToolTipText     =   "Triângulo"
            ImageIndex      =   9
            Style           =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Polígono"
            Object.ToolTipText     =   "Polígono"
            ImageIndex      =   10
            Style           =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Polígono_regular"
            Object.ToolTipText     =   "Polígono Regular"
            ImageIndex      =   11
            Style           =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Circunferência"
            Object.ToolTipText     =   "Circunferência"
            ImageIndex      =   12
            Style           =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Arco"
            Object.ToolTipText     =   "Arco de circunferência"
            ImageIndex      =   13
            Style           =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cônica"
            Object.ToolTipText     =   "Cônica por 5 pontos"
            ImageIndex      =   14
            Style           =   2
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reta_paralela"
            Object.ToolTipText     =   "Reta paralela "
            ImageIndex      =   15
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reta_perpendicular"
            Object.ToolTipText     =   "Reta perpendicular"
            ImageIndex      =   16
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Medriatriz"
            Object.ToolTipText     =   "Medriatriz"
            ImageIndex      =   17
            Style           =   2
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ponto_médio"
            Object.ToolTipText     =   "Ponto médio"
            ImageIndex      =   18
            Style           =   2
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bissetriz"
            Object.ToolTipText     =   "Bissetriz"
            ImageIndex      =   19
            Style           =   2
         EndProperty
      EndProperty
      MousePointer    =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   600
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
            Picture         =   "frmTela_Desenho.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":0D36
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":1A6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":34D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":420E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":4F44
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":5C7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":69B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":76E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":841C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":9152
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":9E88
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":ABBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":B8F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":C62A
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":D360
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":E096
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTela_Desenho.frx":EDCC
            Key             =   ""
         EndProperty
      EndProperty
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
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case 27
   tbrObjetos.Buttons(1).Value = tbrPressed
  Case Else
   MsgBox KeyCode
 End Select
End Sub
Private Sub Form_Load()

 Call Inicializar_Parametros
 Me.BackColor = vbWhite
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'Incluir rotina para:
 '-Iniciar seleção múltipla.
 '-Permitir seleção personalizada, com caixas de verificação para tipos de objetos...
 '-Seleção avançada é parte de um outro botão
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim N, N_Obj As Integer
 Dim Cor_Ponto_XY As Long
 
 Cor_Ponto_XY = Me.Point(X, Y)
 
 If Cor_Ponto_XY = Me.BackColor Then Exit Sub
 
 Me.Caption = ""
 N_Obj = UBound(Obj)
 For N = 1 To N_Obj
  With Obj(N)
   If .Mostrar Then
    If Cor_Ponto_XY = .Cor Then
     If .Tipo = PONTO Or .Tipo = PONTO_DE_INTERSECÇÃO Or .Tipo = PONTO_MEDIO Or .Tipo = PONTO_SOBRE Then
      'Me.MousePointer = vbSizeAll
      'If (X - .P_int(1)) ^ 2 + (Y - .P_int(2)) ^ 2 < 0.1 Then Me.Caption = "Ponto" & N: Exit For
     End If
    End If
   End If
  End With
 Next N
' Me.MousePointer = vbCrosshair
 Me.Caption = Me.Caption & " " & X & " " & Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Incluir rotinas para:
'-adicionar objetos
'-Selecionar objetos
'-Finalizar uma seleção múltipla

End Sub

Private Sub Form_Paint()
 Dim N, N_Obj As Integer
 
 N_Obj = UBound(Obj)
 
 For N = 1 To N_Obj
  With Obj(N)
   Select Case .Tipo
   Case PONTO
    Me.DrawWidth = .Espessura
    Me.PSet (Cv_X(.P_int(1)), Cv_Y(.P_int(2))) ', .Cor
    
   Case PONTO_SOBRE
   
   Case PONTO_DE_INTERSECÇÃO
   
   Case SEGMENTO
   
   Case VETOR
   
   Case RETA
   
   Case SEMI_RETA
   
   Case TRIÂNGULO
   
   Case POLÍGONO
   
   Case POLÍGONO_REGULAR
   
   Case EIXOS
   
   Case CIRCUNFERÊNCIA
   
   Case ARCO
   
   Case CÔNICA
   
   Case PERPENDICULAR
   
   Case PARALELA
   
   Case PONTO_MEDIO
   
   Case BISSETRIZ_PONTOS
   
   Case BISSETRIZ_RETAS
   
   Case COMPASSO
   
   Case REFLEXÃO
   
   Case SIMETRIA
   
   Case TRANSLAÇÃO
   
   Case INVERSO_CIRCUNFERÊNCIA
   
   Case TEXTO
   
   Case ÂNGULO

   Case Else
    
   End Select
  End With
 Next N
 
 For N = 0 To Tamanho_X \ 2
  Me.PSet (Cv_X(CSng(N)), Cv_Y(0))
  Me.PSet (Cv_X(0), Cv_Y(CSng(N)))
 Next N

End Sub
Private Function Cv_X(X_real As Single) As Long
'como forçar que a unidade padrão pareça medir 1 centímetro sobre a tela???
'Antigo: Cv_X = CLng(X_real * Zoom * (Me.ScaleWidth / Tamanho_X))
 Dim Pixel_por_Cm_X As Single

 Pixel_por_Cm_X = Twips_por_Cm / Screen.TwipsPerPixelX
 Cv_X = CLng(X_real * Zoom * Pixel_por_Cm_X)
 
End Function
Private Function Cv_Y(Y_real As Single) As Long
 Dim Pixel_por_Cm_X As Single

 Pixel_por_Cm_Y = Twips_por_Cm / Screen.TwipsPerPixelY
 aspec = -Me.ScaleWidth / (tbrObjetos.Height + Me.ScaleHeight)
 Cv_Y = CLng(-Y_real * aspec * Zoom * (tbrObjetos.Height + Me.ScaleHeight) / Tamanho_Y)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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

Private Sub Form_Resize()
'Parametriza a tela, obtendo com coordenadas inteiras (pixels):
'(-+)|(++)
'---------
'(--)|(+-)

'Dados:
'* (C_X, C_Y):  O Pixel central corresponde a esse ponto em medidas reais.
'* Pixels_X por Pixels_Y: Dimensões da tela de desenho

Me.Caption = Me.ScaleLeft & ", " & Me.ScaleTop & " : " & Me.ScaleWidth & " x " & Me.ScaleHeight
Me.Caption = Me.Caption & "   Twips: " & Me.Width & " x " & Me.Height

Exit Sub
 With Me
  .Cls
  .ScaleHeight = -Abs(.ScaleHeight)
  .ScaleTop = Abs(tbrObjetos.Height - .ScaleHeight) \ 2
  .ScaleLeft = -.ScaleWidth \ 2
  .Refresh
  '.Caption = "TAM: " & .ScaleWidth & " x " & .ScaleHeight _
  ' & "   CENTRO: " & Cv_X(0) & ", " & Cv_Y(0)
 End With
 
End Sub
