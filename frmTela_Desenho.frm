VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTela_Desenho 
   Caption         =   "GDin"
   ClientHeight    =   7185
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   479
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   760
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
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
      Top             =   2760
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
      Top             =   2160
      Width           =   300
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
      Top             =   1800
      Width           =   2265
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
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Ponteiro"
            Object.ToolTipText     =   "Sele��o de objetos"
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
            Object.ToolTipText     =   "Pontos de Interse��o"
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
            Key             =   "Tr�ngulo"
            Object.ToolTipText     =   "Tri�ngulo"
            ImageIndex      =   9
            Style           =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pol�gono"
            Object.ToolTipText     =   "Pol�gono"
            ImageIndex      =   10
            Style           =   2
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pol�gono_regular"
            Object.ToolTipText     =   "Pol�gono Regular"
            ImageIndex      =   11
            Style           =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Circunfer�ncia"
            Object.ToolTipText     =   "Circunfer�ncia"
            ImageIndex      =   12
            Style           =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Arco"
            Object.ToolTipText     =   "Arco de circunfer�ncia"
            ImageIndex      =   13
            Style           =   2
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "C�nica"
            Object.ToolTipText     =   "C�nica por 5 pontos"
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
            Key             =   "Ponto_m�dio"
            Object.ToolTipText     =   "Ponto m�dio"
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
      Left            =   375
      Top             =   2160
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
Private Const TAM_BARRA = 20
Private Sub Form_Initialize()
 Call Inicializar_Parametros
  With Me
  .BackColor = vbWhite
  HScroll1.Move .ScaleLeft, .ScaleTop + .ScaleHeight - TAM_BARRA, .ScaleWidth - TAM_BARRA, TAM_BARRA
  VScroll1.Move .ScaleLeft + .ScaleWidth - TAM_BARRA, .ScaleTop + tbrObjetos.Height, TAM_BARRA, .ScaleHeight - tbrObjetos.Height - TAM_BARRA
  HScroll1.Min = -MAX_X: HScroll1.Max = MAX_X
  VScroll1.Min = MAX_Y: VScroll1.Max = -MAX_Y
 End With

End Sub

Private Sub Form_LostFocus()
MsgBox "form.lostfocus"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'Incluir rotina para:
 '-Iniciar sele��o m�ltipla.
 '-Permitir sele��o personalizada, com caixas de verifica��o para tipos de objetos...
 '-Sele��o avan�ada � parte de um outro bot�o
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Incluir rotinas para:
'-adicionar objetos
'-Selecionar objetos
'-Finalizar uma sele��o m�ltipla
'
'Se um ponto clicado na tela ainda n�o existe em Obj(), criar um e guarda seu Id em P()
  
  
  Exit Sub
  
  
  
  Ferramenta = 0 'Item de menu atualmente selecionado.
  Select Case Ferramenta
    Case PONTO
    'Param=1
      ReDim Preserve Obj(1 To 1 + UBound(Obj))
      With Obj(UBound(Obj))
        '.Tipo = PONTO
        '.Nome = ""
        '.N_Param = 0
        ReDim .P_int(1 To 2)
        .P_int(1) = P(1)
        .P_int(2) = P(2)
        .Espessura = 4#
        .Mostrar = True
        
        '.Tra�o(0) = 1: .Tra�o(2) = 1'Irrelevante para pontos. Usar como X,O, . ou + ...
      End With
    Case PONTO_SOBRE
     'Param=2
    Case PONTO_DE_INTERSEC��O
    'Param=2
    Case SEGMENTO, VETOR
    'Param=2
      With Obj(UBound(Obj))
        .Tipo = Ferramenta 'SEGMENTO ou VETOR
        .Nome = ""
        .N_Param = 2
        ReDim .P_ext(1 To .N_Param)
        .P_ext(1) = P(1): .P_ext(2) = P(2) '(x,y) do ponto  inicial
        .P_ext(3) = P(3): .P_ext(4) = P(4) '(x,y) do ponto  final
        
        .Tra�o(1) = 1: .Tra�o(2) = 1
        .Cor = 0
        .Espessura = 4#
        .Mostrar = True
      End With
    Case RETA
    'Param=2
    Case SEMI_RETA
    'Param=2
    Case TRI�NGULO
    'Param=3
    Case POL�GONO
    'Param=N
    Case POL�GONO_REGULAR
    'Param=3
    Case EIXOS
    'Param=3
    Case CIRCUNFER�NCIA
    'Param=2
    Case ARCO
    'Param=3
    Case C�NICA
    'Param=5
    Case PERPENDICULAR
    'Param=2
    Case PARALELA
    'Param=2
    Case PONTO_MEDIO
    'Param=1 ou 2
    Case BISSETRIZ_PONTOS
    'Param=3
    Case BISSETRIZ_RETAS
    'Param=2
    Case COMPASSO
    'Param=2 ou 3
    Case REFLEX�O
    'Param=2
    Case SIMETRIA
    'Param=2
    Case TRANSLA��O
    'Param=2
    Case INVERSO_CIRCUNFER�NCIA
    'Param=2
    Case TEXTO
    'Param=1 + texto
    Case �NGULO
    'Param=2 ou 3
    Case Else
  End Select
End Sub

Private Sub Form_Paint()
 Dim N, N_Obj As Integer
 
 N_Obj = UBound(Obj)
 
 For N = 1 To N_Obj
  With Obj(N)
   Select Case .Tipo
   Case PONTO
    Me.DrawWidth = .Espessura
    Me.PSet (Pixel_X(.P_int(1)), Pixel_Y(.P_int(2))) ', .Cor
    
   Case PONTO_SOBRE
   
   Case PONTO_DE_INTERSEC��O
   
   Case SEGMENTO
   
   Case VETOR
   
   Case RETA
   
   Case SEMI_RETA
   
   Case TRI�NGULO
   
   Case POL�GONO
   
   Case POL�GONO_REGULAR
   
   Case EIXOS
   
   Case CIRCUNFER�NCIA
   
   Case ARCO
   
   Case C�NICA
   
   Case PERPENDICULAR
   
   Case PARALELA
   
   Case PONTO_MEDIO
   
   Case BISSETRIZ_PONTOS
   
   Case BISSETRIZ_RETAS
   
   Case COMPASSO
   
   Case REFLEX�O
   
   Case SIMETRIA
   
   Case TRANSLA��O
   
   Case INVERSO_CIRCUNFER�NCIA
   
   Case TEXTO
   
   Case �NGULO

   Case Else
    
   End Select
  End With
 Next N
 
 
 'Me.PSet (Pixel_X(Visivel_X / 2), Pixel_Y(Visivel_Y / 2)), vbGreen
 
 
 For N = 0 To 10
  Me.PSet (Pixel_X(CSng(N)), Pixel_Y(0))
 Next N
 For N = 1 To 10
  Me.PSet (Pixel_X(0), Pixel_Y(CSng(N))), vbRed
 Next N


End Sub

Private Sub Aponta_Objeto(ByVal X As Single, ByVal Y As Single)
 Const DIST_MIN = 0.4
 Dim N, N_Obj As Integer
 Dim Cor_Ponto_XY As Long
 
 Cor_Ponto_XY = Me.Point(X, Y)
 If Cor_Ponto_XY = Me.BackColor Then Exit Sub
 
 X = Cm_X(X): Y = Cm_Y(Y)
 N_Obj = UBound(Obj)
 For N = 1 To N_Obj
  With Obj(N)
   If .Mostrar Then
    If Cor_Ponto_XY = .Cor Then
     If .Tipo = PONTO Or .Tipo = PONTO_DE_INTERSEC��O Or .Tipo = PONTO_MEDIO Or .Tipo = PONTO_SOBRE Then
      'Me.MousePointer = vbSizeAll
      If Abs(X - .P_int(1)) + Abs(Y - .P_int(2)) < DIST_MIN Then Objeto_Prox = N: Exit Sub
     End If
    End If
   End If
  End With
 Next N
 Objeto_Prox = 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Objeto_Prox > 0 Then Exit Sub
  'Exibe R�tulo
 ' Me.MousePointer = vbCrosshair
 Me.Caption = Format(Cm_X(X), "0.0") & " " & Format(Cm_Y(Y), "0.0")

If Screen.TwipsPerPixelX <> TwipsPerPixelX_INICIAL Then
 MsgBox "TwipsPerPixelX mudou de " & TwipsPerPixelX_INICIAL & "para " & Screen.TwipsPerPixelX
End If
If Screen.TwipsPerPixelY <> TwipsPerPixelY_INICIAL Then
 MsgBox "TwipsPerPixelY mudou de " & TwipsPerPixelY_INICIAL & "para " & Screen.TwipsPerPixelY
End If
End Sub
Private Sub Form_Resize()
 Dim Visivel_antes_X, Visivel_antes_Y As Single

 Visivel_antes_X = Visivel_X
 Visivel_antes_Y = Visivel_Y
 With Me
  'Mede a largura e a altura da �rea de desenho em "pixels"
  Visivel_X = .ScaleWidth - .VScroll1.Width
  Visivel_Y = .ScaleHeight - (tbrObjetos.Height + HScroll1.Height)
  'Converte a largura e a altura da �rea de desenho para "cent�metros"
  Visivel_X = Visivel_X * TwipsPerPixelX_INICIAL / Twips_por_Cm
  Visivel_Y = Visivel_Y * TwipsPerPixelY_INICIAL / Twips_por_Cm
  'Atualiza as coordenadas que correspondem ao centro do form.
  Centro_X = Centro_X + (Visivel_X - Visivel_antes_X) / 2
  Centro_Y = Centro_Y - (Visivel_Y - Visivel_antes_Y) / 2
  
  .Cls
  .Refresh
  
  On Error Resume Next
  HScroll1.Move .ScaleLeft, .ScaleTop + .ScaleHeight - TAM_BARRA, .ScaleWidth - TAM_BARRA, TAM_BARRA
  VScroll1.Move .ScaleLeft + .ScaleWidth - TAM_BARRA, .ScaleTop + tbrObjetos.Height, TAM_BARRA, .ScaleHeight - tbrObjetos.Height - TAM_BARRA
  picCanto.Move .ScaleLeft + .ScaleWidth - TAM_BARRA, .ScaleTop + .ScaleHeight - TAM_BARRA
  On Error GoTo 0
  Timer1.Enabled = True
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
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Exit Sub
 Select Case MsgBox("Deseja salvar as altera��es?", vbQuestion + vbYesNoCancel, "Finalizando o aplicativo...")
 Case vbCancel
  Cancel = True
 Case vbNo
  Cancel = False
 Case vbYes
  Cancel = True
  'SalvarArquivo
 End Select
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case KeyCode
  Case vbKeyEscape
   tbrObjetos.Buttons(1).Value = tbrPressed
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


Private Sub tbrObjetos_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case UBound(P)
 Case 1
 'Button.Index
 Case Else
 
 End Select
End Sub

Private Sub Timer1_Timer()
'Esse timer s'o existe para contornar um defeito na rotina Resize...
'O programa nao conhece a altura real da barra no instante do redimensionamento, s� depois
With Me
 On Error Resume Next
 VScroll1.Move .ScaleLeft + .ScaleWidth - TAM_BARRA, .ScaleTop + tbrObjetos.Height, TAM_BARRA, .ScaleHeight - tbrObjetos.Height - TAM_BARRA
 On Error GoTo 0
 Timer1.Enabled = False
End With
End Sub

Private Sub HScroll1_Change()
 Centro_X = HScroll1.Value
 Me.Refresh
End Sub
Private Sub HScroll1_Scroll()
'Incluir mais valores entre os extremos de MAX_X e -MAX_X permitir� um scroll mais suave
'Conferir compatibilidade com o Zoom
 Centro_X = HScroll1.Value
 Me.Refresh
End Sub

Private Sub VScroll1_Change()
 Centro_Y = VScroll1.Value
 Me.Refresh
End Sub
Private Sub VScroll1_Scroll()
'Incluir mais valores entre os extremos de MAX_Y e -MAX_Y permitir� um scroll mais suave
'Conferir compatibilidade com o Zoom
 Centro_Y = VScroll1.Value
 Me.Refresh
End Sub

