Attribute VB_Name = "basObjGeometria"
Public Enum Tipo_De_Objeto
 PONTO 'PONTO = 0. É o valor padrão de novos objetos.
 PONTO_SOBRE
 PONTO_DE_INTERSECÇÃO
 SEGMENTO
 VETOR
 SEMI_RETA
 RETA
 TRIÂNGULO
 POLÍGONO
 POLÍGONO_REGULAR
 CIRCUNFERÊNCIA
 ARCO
 CÔNICA
 PARALELA
 PERPENDICULAR
 MEDIATRIZ
 PONTO_MÉDIO
 BISSETRIZ_PONTOS
 BISSETRIZ_RETAS
 
 COMPASSO
 REFLEXÃO
 SIMETRIA
 TRANSLAÇÃO
 INVERSO_CIRCUNFERÊNCIA
 TEXTO
 ÂNGULO
 EIXOS
End Enum
Public Enum Aparência
 OCULTO
 PADRAO
 SELECIONADO
End Enum

Private Type Objeto
 Id As Integer 'Identifica exclusivamente o objeto (por tipo??). Como o indice de Obj?
 Tipo As Tipo_De_Objeto 'Que item será guardado
 N_Param As Byte 'Número de objetos dos quais este é dependente
 Cor As Long 'Cor utilizada para desenhar na tela
 Espessura As Byte 'Raio dos pontos ou a largura de curvas e contornos
 Traço(1 To 2) As Byte 'Tipo de pontilhado
 Mostrar As Aparência 'Indica como o objeto será exibido
 Nome As String 'Um rótulo para exibição em tela
 P_dep() As Integer 'Indices dos parametros (objetos)dos quais depende
 P_rep() As Single 'Coordenadas e angulos livres
End Type

Public Const MAX_OBJETOS = 100
Public Const PI = 3.14159265358979
Public Const DEG = 1.74532925199433E-02 'PI / 180
Public Const Twips_por_Cm = 576
Public Const Twips_por_Polegada = 1440
Public Const Twips_por_Ponto = 20
Public Const MAX_X = 10
Public Const MAX_Y = 10
Public Const NENHUM = 0


Public TwipsPerPixelX_INICIAL As Single, TwipsPerPixelY_INICIAL As Single
Public Cm_por_Pixel_X, Cm_por_Pixel_Y As Single
Public Centro_X As Single, Centro_Y As Single
Public Visivel_X As Single, Visivel_Y As Single 'Dimensoes que a tela parece ter
Public Visivel_X_pix As GLsizei, Visivel_Y_pix As GLsizei
Public Zoom As Single
Public inc_Mov As Single, inc_Trans As Single

Public Obj() As Objeto
Public Nome(PONTO To EIXOS) As String
Public P() As Long
Public Objeto_Localizado() As Long

Public Sub Inicializar_Objetos()
'Atribui o valor inicial dos principais parametros.
'Futuramente, chamará a função que lê um arquivo de dados salvo.

Nome(PONTO) = "Ponto"
Nome(PONTO_SOBRE) = "Ponto"
Nome(PONTO_DE_INTERSECÇÃO) = "Ponto"
Nome(SEGMENTO) = "Segmento"
Nome(VETOR) = "Vetor"
Nome(SEMI_RETA) = "Semi-reta"
Nome(RETA) = "Reta"
Nome(TRIÂNGULO) = "Triângulo"
Nome(POLÍGONO) = "Polígono"
Nome(POLÍGONO_REGULAR) = "Polígono"
Nome(CIRCUNFERÊNCIA) = "Circuferência"
Nome(ARCO) = "Arco"
Nome(CÔNICA) = "Cônica"
Nome(PARALELA) = "Reta paralela"
Nome(PERPENDICULAR) = "Reta perpendicular"
Nome(MEDIATRIZ) = "Mediatriz"
Nome(PONTO_MÉDIO) = "Ponto médio"
Nome(BISSETRIZ_PONTOS) = "Bissetriz"
Nome(BISSETRIZ_RETAS) = "Bissetriz"
Nome(COMPASSO) = "Circunferência"
Nome(REFLEXÃO) = ""
Nome(SIMETRIA) = ""
Nome(TRANSLAÇÃO) = ""
Nome(INVERSO_CIRCUNFERÊNCIA) = ""
Nome(TEXTO) = "Texto"
Nome(ÂNGULO) = "Ângulo"
Nome(EIXOS) = "Eixo"

 'Cria os objetos essenciais
 ReDim Obj(1 To 2)
  
 ReDim P(1 To 1) As Long
 
 With Obj(1)
  '.Tipo = PONTO
  ReDim .P_rep(1 To 3)
  .P_rep(1) = 0#
  .P_rep(2) = 0#
  .P_rep(3) = 1#
  .Espessura = 4#
  .Mostrar = PADRAO
  .Nome = "Origem"
  '.Traço(0) = 1: .Traço(2) = 1'Irrelevante para pontos. Usar como X,O, . ou + ...
 End With
 
 With Obj(2)
  .Tipo = EIXOS
  ReDim .P_rep(1 To 4)
  .P_rep(1) = 0#: .P_rep(2) = 1#
  .P_rep(3) = 1#: .P_rep(4) = 0#
  ReDim .P_dep(1 To 1)
  .P_dep(1) = 1
  .Espessura = 1
  .Mostrar = PADRAO
  .Nome = "Eixo padrão"
  .Cor = RGB(200, 200, 200)
  .Traço(1) = 1: .Traço(2) = 1
  .N_Param = 1
  ReDim .P_dep(1 To .N_Param)
  .P_dep(1) = 1
 End With
 'frmTela_Desenho.Refresh?
End Sub
