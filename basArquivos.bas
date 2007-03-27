Attribute VB_Name = "basArquivos"
Option Explicit

Private Type Ferramenta
 IdImg As Integer
 Key As String
 TipText As String
End Type

Sub Inicializa_Barra_Ferramentas(IdDoc As Integer)
 Const Arq_INI = "Tabela.ini"
 Dim imgX As ListImage
 Dim btnButton As Button
 
 Dim Qtd As Integer
 Dim FileNumber As Integer
 Dim N As Integer
 Dim F() As Ferramenta ' IdImg, Key e TipText
 
 FileNumber = FreeFile
 On Error GoTo ERRO
  Open App.Path & "\" & Arq_INI For Input As #FileNumber
 On Error GoTo 0
  
 N = 0
 'ReDim F(1 To N)
 
 With Doc(IdDoc).frm.ilsFerramentas
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
 
 With Doc(IdDoc).frm.tbrFerramentas
   .Buttons.Clear
   .ImageList = Doc(IdDoc).frm.ilsFerramentas
   For N = 1 To Qtd
    Set btnButton = .Buttons.Add(N, F(N).Key, "", tbrDefault, N)
    btnButton.ToolTipText = F(N).TipText
    btnButton.Style = tbrButtonGroup
   If N > 2 Then btnButton.Enabled = False
   Next N
   .Buttons(1).Value = tbrPressed
   .Tag = .Buttons.Item(1).Key
 End With
 
 Exit Sub
ERRO:
 'If Err.Number = 53 Then
  'Err.Clear
  'Recup_Arquivo
  'Inicializa
 'Else
  Err.Raise Err.Number
 'End If
End Sub



