Attribute VB_Name = "basGDin"
Public Enum Tipo_De_Objeto
 PONTO 'PONTO = 0. � o valor padr�o de novos objetos.
 PONTO_SOBRE
 PONTO_DE_INTERSEC��O
 SEGMENTO
 VETOR
 RETA
 SEMI_RETA
 TRI�NGULO
 POL�GONO
 POL�GONO_REGULAR
 EIXOS
 CIRCUNFER�NCIA
 ARCO
 C�NICA
 PERPENDICULAR
 PARALELA
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
Public Const Twips_por_Pontos = 20

Public Centro_X, Centro_Y As Single
Public Tamanho_X, Tamanho_Y As Single
Public Zoom As Single
Public inc_Mov, inc_Trans As Single

Public Obj() As Objeto

Public Sub Inicializar_Parametros()
'Atribui o valor inicial dos principais parametros.
'Futuramente, chamar� a fun��o que l� um arquivo de dados salvo.
'
 inc_Mov = 0.05
 inc_Trans = 1
 
 Centro_X = 0#
 Centro_Y = 0#
 Tamanho_X = 10#
 Tamanho_Y = 10#
 Zoom = 1#
 'Cria os objetos essenciais
 ReDim Obj(1 To 2)
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
