VERSION 5.00
Begin VB.Form frmEtiquetaTesteOpcoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Etiqueta Teste"
   ClientHeight    =   2625
   ClientLeft      =   1950
   ClientTop       =   11025
   ClientWidth     =   6090
   Icon            =   "frmEtiquetaTesteOpcoes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   6090
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   855
      Left            =   3728
      Picture         =   "frmEtiquetaTesteOpcoes.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1665
      Width           =   975
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Nova Etiqueta"
      Height          =   855
      Left            =   2558
      Picture         =   "frmEtiquetaTesteOpcoes.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1665
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   855
      Left            =   1388
      Picture         =   "frmEtiquetaTesteOpcoes.frx":0EEE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1665
      Width           =   975
   End
   Begin VB.TextBox txtSetor 
      Height          =   285
      Left            =   225
      TabIndex        =   3
      Top             =   1215
      Width           =   3525
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   4185
      TabIndex        =   2
      Top             =   495
      Width           =   960
   End
   Begin VB.TextBox txtNomeFuncionario 
      Height          =   285
      Left            =   225
      TabIndex        =   0
      Top             =   495
      Width           =   3525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Left            =   4185
      TabIndex        =   5
      Top             =   225
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome do Funcionário"
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   225
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Setor"
      Height          =   195
      Left            =   225
      TabIndex        =   1
      Top             =   945
      Width           =   375
   End
End
Attribute VB_Name = "frmEtiquetaTesteOpcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdImprimir_Click()
frmEtiquetaTeste.PrintForm
End Sub

Private Sub cmdNovo_Click()
frmEtiquetaTesteOpcoes.txtNomeFuncionario = ""
frmEtiquetaTesteOpcoes.txtCodigo = ""
frmEtiquetaTesteOpcoes.txtSetor = ""
frmEtiquetaTesteOpcoes.txtNomeFuncionario.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me

End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Height = 3030
Me.Width = 6210
frmEtiquetaTeste.Show
End Sub

Private Sub txtCodigo_Change()
frmEtiquetaTeste.lblCodigo.Caption = txtCodigo.Text
frmEtiquetaTeste.lblCodigoBarra.Caption = txtCodigo.Text
End Sub

Private Sub txtNomeFuncionario_Change()

frmEtiquetaTeste.lblFuncionario.Caption = txtNomeFuncionario.Text

End Sub


Private Sub txtSetor_Change()
frmEtiquetaTeste.lblNomeSetor.Caption = txtSetor.Text

End Sub


