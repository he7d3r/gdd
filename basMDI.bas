Attribute VB_Name = "basMDI"
'***  Módulo Global para o aplicativo MDI.  ***
'**********************************************
Option Explicit

'Identificadores das vistas individuais
Public Enum Vista
   PERSPECTIVA = 1
   FRONTAL
   LATERAL
   SUPERIOR
   EPURA
End Enum
'Constantes matemáticas utilizadas com frequência
Public Const PI = 3.14159265
Public Const DEG = PI / 180
Public Const ZERO = 0.0000001 'Usado em rotinas que tem problemas com o número zero=0.0
Public Const DIST_MAX_CENA = 100
Public Const DIST_MIN_CENA = 1

'Constantes limitadoras do tamanho de cada construção
Public Const MAX_OBJETOS = 10000       'usado como extremo superior da matriz Obj()
Public Const TAM_BUFER = MAX_OBJETOS   'usado no processo de 'apontar um objeto'

'O parâmetro 'SobrePlano' da rotina Des_Plano deve ser um destes
Public Enum Tipo_De_Plano
   PL_HORIZONTAL = 1
   PL_FRONTAL
   PL_PERFIL
End Enum

'O Atributo 'Tipo' de cada objeto deve ser um destes itens
Public Enum Tipo_De_Objeto
   PONTO 'PONTO = 0. É o valor padrão de novos objetos.
   'PONTO_SOBRE 'PONTO_DE_INTERSECÇÃO
   SEGMENTO
   'VETOR 'SEMI_RETA 'RETA 'TRIÂNGULO 'POLÍGONO 'POLÍGONO_REGULAR 'CIRCUNFERÊNCIA
   'ARCO 'CÔNICA 'PARALELA 'PERPENDICULAR 'MEDIATRIZ 'PONTO_MÉDIO 'BISSETRIZ_PONTOS
   'BISSETRIZ_RETAS  'COMPASSO 'REFLEXÃO 'SIMETRIA 'TRANSLAÇÃO 'INVERSO_CIRCUNFERÊNCIA
   'TEXTO 'ÂNGULO 'EIXOS
End Enum

'A estrutura de cada objeto na construção tem estes atributos
Public Type Objeto
   'Id As Integer 'Identifica exclusivamente o objeto (por tipo??). Como o indice de Obj?
   'Tipo As Tipo_De_Objeto 'Que item será guardado
   'N_Param As Byte 'Número de objetos dos quais este é dependente
   'Cor As Long 'Cor utilizada para desenhar na tela
   'Espessura As Byte 'Raio dos pontos ou a largura de curvas e contornos
   'Traço(1 To 2) As Byte 'Tipo de pontilhado
   Selec As Integer 'Indica que o objeto foi o "Selec-ésimo" a ser selecionado
   Mostrar As Boolean ' Aparência 'Indica como o objeto será exibido
   'Nome As String 'Um rótulo para exibição em tela
   'Id_Dep() As Long 'Indices dos parametros (objetos)dos quais este depende
   Coord(0 To 2) As GLdouble 'Coordenadas e angulos livres
End Type

'Informações disponíveis sobre cada documento aberto
Type udtDocumento
   frm As New frmMain         'Doc(nnn).frm.Tag será sempre igual ao índice 'nnn'
   Obj() As Objeto            'A contrução geometrica feita neste documento
   Obj_Sel() As Integer       'Matriz com os índices de cada objeto selecionado no doc
   Deletado As Boolean        'Indica se esse documento ainda esta em uso
   Alterado As Boolean        'Indica se esse documento precisará ser salvo
End Type


Public Doc() As udtDocumento        'Matriz contendo cada documento.
                                    '(Cada atributo 'frm' é um 'child' do frmMDIGeo3d)
Public P_Aux(0 To 2) As GLdouble    'Coordenadas de um ponto auxiliar para a definir objetos
Public Sobre_Plano As Tipo_De_Plano 'Indica plano usado ao definir pontos do espaço
Public Erro As glErrorConstants

Function ExisteDocAberto() As Integer
   Dim i As Integer
   'Retorna verdadeiro se houver ao menos um doc aberto.
   For i = 1 To UBound(Doc)
      If Not Doc(i).Deletado Then
         ExisteDocAberto = i
         Exit Function
      End If
   Next
End Function

Function GeraIdLivre() As Integer
   Dim i As Integer
   Dim Qtd As Integer
   
   Qtd = UBound(Doc)
   
   'Se algum documento foi deletado, reaproveite seu índice
   For i = 1 To Qtd
      If Doc(i).Deletado Then
         GeraIdLivre = i
         Doc(i).Deletado = False
         Exit Function
      End If
   Next
    
   'Se não havia doc deletado, crie e use um novo índice
   ReDim Preserve Doc(1 To Qtd + 1)
   GeraIdLivre = UBound(Doc)
End Function
