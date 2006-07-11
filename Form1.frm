VERSION 5.00
Begin VB.Form frmMatriz 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Matrizes OpenGl"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Line Line8 
      X1              =   4680
      X2              =   4680
      Y1              =   4200
      Y2              =   5520
   End
   Begin VB.Line Line7 
      X1              =   4200
      X2              =   4200
      Y1              =   4200
      Y2              =   5520
   End
   Begin VB.Line Line6 
      X1              =   3840
      X2              =   3840
      Y1              =   4200
      Y2              =   5520
   End
   Begin VB.Line Line5 
      X1              =   3360
      X2              =   3360
      Y1              =   4200
      Y2              =   5520
   End
   Begin VB.Line Line4 
      X1              =   4680
      X2              =   4680
      Y1              =   1680
      Y2              =   3000
   End
   Begin VB.Line Line3 
      X1              =   4200
      X2              =   4200
      Y1              =   1680
      Y2              =   3000
   End
   Begin VB.Line Line2 
      X1              =   3840
      X2              =   3840
      Y1              =   1680
      Y2              =   3000
   End
   Begin VB.Line Line1 
      X1              =   3360
      X2              =   3360
      Y1              =   1680
      Y2              =   3000
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      TabIndex        =   50
      Top             =   4200
      Width           =   195
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      TabIndex        =   49
      Top             =   1680
      Width           =   195
   End
   Begin VB.Label Label13 
      Caption         =   "X' Y' Z' 1'"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4320
      TabIndex        =   48
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label7 
      Caption         =   "X Y Z 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3480
      TabIndex        =   47
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "X' Y' Z' 1'"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4320
      TabIndex        =   46
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label12 
      Caption         =   "X Y Z 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3480
      TabIndex        =   45
      Top             =   1680
      Width           =   255
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Altura"
      Height          =   195
      Left            =   2670
      TabIndex        =   44
      Top             =   840
      Width           =   405
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Largura"
      Height          =   195
      Left            =   1875
      TabIndex        =   43
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Inferior"
      Height          =   195
      Left            =   1185
      TabIndex        =   42
      Top             =   840
      Width           =   480
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Esquerda"
      Height          =   195
      Left            =   375
      TabIndex        =   41
      Top             =   840
      Width           =   675
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "|_____ Translação _____|"
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "|_____ Translação _____|"
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label lblProjectionMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   2520
      TabIndex        =   38
      Top             =   2760
      Width           =   705
   End
   Begin VB.Label lblProjectionMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   2520
      TabIndex        =   37
      Top             =   2400
      Width           =   705
   End
   Begin VB.Label lblProjectionMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   2520
      TabIndex        =   36
      Top             =   2040
      Width           =   705
   End
   Begin VB.Label lblProjectionMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   2520
      TabIndex        =   35
      Top             =   1680
      Width           =   705
   End
   Begin VB.Label lblProjectionMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   1800
      TabIndex        =   34
      Top             =   2760
      Width           =   705
   End
   Begin VB.Label lblProjectionMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   1800
      TabIndex        =   33
      Top             =   2400
      Width           =   705
   End
   Begin VB.Label lblProjectionMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   1800
      TabIndex        =   32
      Top             =   2040
      Width           =   705
   End
   Begin VB.Label lblProjectionMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   1800
      TabIndex        =   31
      Top             =   1680
      Width           =   705
   End
   Begin VB.Label lblProjectionMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   1080
      TabIndex        =   30
      Top             =   2760
      Width           =   705
   End
   Begin VB.Label lblProjectionMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   1080
      TabIndex        =   29
      Top             =   2400
      Width           =   705
   End
   Begin VB.Label lblProjectionMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   1080
      TabIndex        =   28
      Top             =   2040
      Width           =   705
   End
   Begin VB.Label lblProjectionMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   27
      Top             =   1680
      Width           =   705
   End
   Begin VB.Label lblProjectionMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   26
      Top             =   2760
      Width           =   705
   End
   Begin VB.Label lblProjectionMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   25
      Top             =   2400
      Width           =   705
   End
   Begin VB.Label lblProjectionMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   24
      Top             =   2040
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "GL_PROJECTION_MATRIX (glgProjectionMatrix):"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   1320
      Width           =   3540
   End
   Begin VB.Label lblProjectionMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   22
      Top             =   1680
      Width           =   705
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "GL_MODELVIEW_MATRIX (glgModelViewMatrix):"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   3840
      Width           =   3585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "GL_VIEWPORT (glgViewport):"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblViewPort 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   19
      Top             =   480
      Width           =   705
   End
   Begin VB.Label lblViewPort 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   18
      Top             =   480
      Width           =   705
   End
   Begin VB.Label lblViewPort 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   17
      Top             =   480
      Width           =   705
   End
   Begin VB.Label lblViewPort 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Top             =   480
      Width           =   705
   End
   Begin VB.Label lblModelViewMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   2520
      TabIndex        =   15
      Top             =   5280
      Width           =   705
   End
   Begin VB.Label lblModelViewMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   2520
      TabIndex        =   14
      Top             =   4920
      Width           =   705
   End
   Begin VB.Label lblModelViewMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   2520
      TabIndex        =   13
      Top             =   4560
      Width           =   705
   End
   Begin VB.Label lblModelViewMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   2520
      TabIndex        =   12
      Top             =   4200
      Width           =   705
   End
   Begin VB.Label lblModelViewMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   1800
      TabIndex        =   11
      Top             =   5280
      Width           =   705
   End
   Begin VB.Label lblModelViewMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   1800
      TabIndex        =   10
      Top             =   4920
      Width           =   705
   End
   Begin VB.Label lblModelViewMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   1800
      TabIndex        =   9
      Top             =   4560
      Width           =   705
   End
   Begin VB.Label lblModelViewMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   1800
      TabIndex        =   8
      Top             =   4200
      Width           =   705
   End
   Begin VB.Label lblModelViewMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   1080
      TabIndex        =   7
      Top             =   5280
      Width           =   705
   End
   Begin VB.Label lblModelViewMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   1080
      TabIndex        =   6
      Top             =   4920
      Width           =   705
   End
   Begin VB.Label lblModelViewMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   1080
      TabIndex        =   5
      Top             =   4560
      Width           =   705
   End
   Begin VB.Label lblModelViewMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   4
      Top             =   4200
      Width           =   705
   End
   Begin VB.Label lblModelViewMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   3
      Top             =   5280
      Width           =   705
   End
   Begin VB.Label lblModelViewMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   2
      Top             =   4920
      Width           =   705
   End
   Begin VB.Label lblModelViewMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   4560
      Width           =   705
   End
   Begin VB.Label lblModelViewMatrix 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   4200
      Width           =   705
   End
End
Attribute VB_Name = "frmMatriz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_GotFocus()
 AtualizaMatrizes
End Sub

Public Sub AtualizaMatrizes()
  Dim i As Integer, j As Integer
  glGetIntegerv glgViewport, Viewport(0)
  glGetFloatv glgModelViewMatrix, ModelViewMatrix(0)
  glGetFloatv glgProjectionMatrix, ProjectionMatrix(0)
    
  For i = 0 To 3
   lblViewPort(i) = Viewport(i)
  Next i
  For i = 0 To 3 'Para cada linha 'i' de 0 a 3,
   For j = 0 To 3 'considere as colunas 'j' de 0 a 3 e faça...
    lblModelViewMatrix(j * 4 + i) = Format(ModelViewMatrix(j * 4 + i), "0.0000")
    lblProjectionMatrix(j * 4 + i) = Format(ProjectionMatrix(j * 4 + i), "0.0000")
   Next j
  Next i
End Sub

