Attribute VB_Name = "basGDin"
Public Enum Tipo_De_Objeto
 PONTO 'PONTO = 0. � o valor padr�o de novos objetos.
 PONTO_SOBRE
 PONTO_DE_INTERSEC��O
 SEGMENTO
 VETOR
 SEMI_RETA
 RETA
 TRI�NGULO
 POL�GONO
 POL�GONO_REGULAR
 CIRCUNFER�NCIA
 ARCO
 C�NICA
 PARALELA
 PERPENDICULAR
 MEDIATRIZ
 PONTO_MEDIO
 BISSETRIZ_PONTOS
 BISSETRIZ_RETAS
 
 COMPASSO
 REFLEX�O
 SIMETRIA
 TRANSLA��O
 INVERSO_CIRCUNFER�NCIA
 TEXTO
 �NGULO
 EIXOS
End Enum

Private Type Objeto
 Id As Integer 'Identifica exclusivamente cada objeto. Igual ao indice da matriz?
 Tipo As Tipo_De_Objeto 'Que item ser� guardado
 N_Param As Byte 'N�mero de objetos dos quais este � dependente
 Cor As Long 'Cor utilizada para desenhar na tela
 Espessura As Byte 'Raio dos pontos ou a largura de curvas e contornos
 Tra�o(1 To 2) As Byte 'Tipo de pontilhado
 Mostrar As Boolean 'Indica se o objeto ser� exibido
 Nome As String 'Um r�tulo para exibi��o em tela
 P_ext() As Integer 'Indices dos parametros (objetos)dos quais depende
 P_int() As Single 'Coordenadas e angulos livres
End Type

Public Const MAX_OBJETOS = 100
Public Const PI = 3.14159265358979
Public Const DEG = 1.74532925199433E-02 'PI / 180
Public Const Twips_por_Cm = 576
Public Const Twips_por_Polegada = 1440
Public Const Twips_por_Ponto = 20
Public Const MAX_X = 10
Public Const MAX_Y = 10

Public TwipsPerPixelX_INICIAL As Single, TwipsPerPixelY_INICIAL As Single
Public Centro_X As Single, Centro_Y As Single
Public Visivel_X As Single, Visivel_Y As Single 'Dimensoes que a tela parece ter
Public Zoom As Single
Public inc_Mov As Single, inc_Trans As Single

Public Obj() As Objeto
Public P() As Long
Public Objeto_Prox As Long

Public Sub Inicializar_Parametros()
'Atribui o valor inicial dos principais parametros.
'Futuramente, chamar� a fun��o que l� um arquivo de dados salvo.

 inc_Mov = 0.05
 inc_Trans = 1
 Objeto_Prox = 0
 
 Centro_X = 0#
 Centro_Y = 0#
 
 TwipsPerPixelX_INICIAL = Screen.TwipsPerPixelX
 TwipsPerPixelY_INICIAL = Screen.TwipsPerPixelY
 'O form mede N twips, cada cm cont�m M twips, logo o form mede N/M cm's
 
 'Mede a largura e a altura da �rea de desenho em "pixels"
 Visivel_X = frmTela_Desenho.ScaleWidth - frmTela_Desenho.VScroll1.Width
 Visivel_Y = frmTela_Desenho.ScaleHeight - _
      (frmTela_Desenho.tbrObjetos.Height + frmTela_Desenho.HScroll1.Height)
 'Converte a largura e a altura da �rea de desenho para "cent�metros"
 Visivel_X = Visivel_X * TwipsPerPixelX_INICIAL / Twips_por_Cm
 Visivel_Y = Visivel_Y * TwipsPerPixelY_INICIAL / Twips_por_Cm
 
 Zoom = 1#
 'Cria os objetos essenciais
 ReDim Obj(1 To 2)
 ReDim P(1 To 1) As Long
 
 With Obj(1)
  '.Tipo = PONTO
  ReDim .P_int(1 To 2)
  .P_int(1) = 0#
  .P_int(2) = 0#
  .Espessura = 4#
  .Mostrar = True
  .Nome = "Origem"
  '.Tra�o(0) = 1: .Tra�o(2) = 1'Irrelevante para pontos. Usar como X,O, . ou + ...
 End With
 
 With Obj(2)
  .Tipo = EIXOS
  ReDim .P_int(1 To 4)
  .P_int(1) = 0#: .P_int(2) = 1#
  .P_int(3) = 1#: .P_int(4) = 0#
  .Espessura = 1
  .Mostrar = True
  .Nome = "Eixo padr�o"
  .Tra�o(1) = 1: .Tra�o(2) = 1
  .N_Param = 1
  ReDim .P_ext(1 To .N_Param)
  .P_ext(1) = 1
 End With
 'frmTela_Desenho.Refresh
End Sub
