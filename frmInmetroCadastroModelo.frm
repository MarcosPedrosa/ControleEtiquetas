VERSION 5.00
Begin VB.Form frmInmetroCadastroModelo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Modelo da Inmetro"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7350
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   60
      TabIndex        =   14
      Top             =   60
      Width           =   7185
      Begin VB.Frame Frame3 
         Caption         =   "DADOS"
         Height          =   2955
         Left            =   60
         TabIndex        =   17
         Top             =   780
         Width           =   7005
         Begin VB.TextBox TXT_REGISTRO_INMETRO 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1020
            MaxLength       =   12
            TabIndex        =   5
            Top             =   600
            Width           =   1605
         End
         Begin VB.TextBox TXT_DESCRICAO 
            Height          =   285
            Left            =   1020
            MaxLength       =   30
            TabIndex        =   4
            Top             =   240
            Width           =   4455
         End
         Begin VB.Frame Frame2 
            Caption         =   "Descreva os Modelos Aplicaveis"
            Height          =   1725
            Left            =   90
            TabIndex        =   18
            Top             =   1080
            Width           =   6825
            Begin VB.TextBox TXT_DESC_MOD4 
               Height          =   285
               Left            =   900
               MaxLength       =   42
               TabIndex        =   9
               Top             =   1290
               Width           =   5655
            End
            Begin VB.TextBox TXT_DESC_MOD1 
               Height          =   285
               Left            =   900
               MaxLength       =   42
               TabIndex        =   6
               Top             =   270
               Width           =   5655
            End
            Begin VB.TextBox TXT_DESC_MOD2 
               Height          =   285
               Left            =   900
               MaxLength       =   42
               TabIndex        =   7
               Top             =   600
               Width           =   5655
            End
            Begin VB.TextBox TXT_DESC_MOD3 
               Height          =   285
               Left            =   900
               MaxLength       =   42
               TabIndex        =   8
               Top             =   950
               Width           =   5655
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Linha 4:"
               Height          =   195
               Left            =   150
               TabIndex        =   24
               Top             =   1320
               Width           =   570
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Linha 3:"
               Height          =   195
               Left            =   150
               TabIndex        =   21
               Top             =   975
               Width           =   570
            End
            Begin VB.Label lblSenha 
               AutoSize        =   -1  'True
               Caption         =   "Linha 2:"
               Height          =   195
               Left            =   150
               TabIndex        =   20
               Top             =   615
               Width           =   570
            End
            Begin VB.Label lblLogin 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Linha 1:"
               Height          =   195
               Left            =   150
               TabIndex        =   19
               Top             =   270
               Width           =   570
            End
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Registro:"
            Height          =   195
            Left            =   90
            TabIndex        =   23
            Top             =   630
            Width           =   855
         End
         Begin VB.Label lblNome 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   90
            TabIndex        =   22
            Top             =   240
            Width           =   765
         End
      End
      Begin VB.CommandButton cmdNovo 
         Height          =   615
         Left            =   5100
         Picture         =   "frmInmetroCadastroModelo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Incluir registro"
         Top             =   3780
         Width           =   615
      End
      Begin VB.CommandButton cmdSalvar 
         Height          =   615
         Left            =   5760
         Picture         =   "frmInmetroCadastroModelo.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salvar Registro"
         Top             =   3780
         Width           =   615
      End
      Begin VB.CommandButton cmdExcluir 
         Height          =   615
         Left            =   4440
         Picture         =   "frmInmetroCadastroModelo.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Excluir registro"
         Top             =   3780
         Width           =   615
      End
      Begin VB.CommandButton btoCancelar 
         Height          =   615
         Left            =   6450
         Picture         =   "frmInmetroCadastroModelo.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Fecha a solicitação"
         Top             =   3780
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Confirmar Modelo"
         Height          =   705
         Left            =   60
         TabIndex        =   15
         Top             =   60
         Width           =   7005
         Begin VB.CommandButton cmdOk 
            Height          =   405
            Left            =   6000
            Picture         =   "frmInmetroCadastroModelo.frx":0E10
            Style           =   1  'Graphical
            TabIndex        =   2
            TabStop         =   0   'False
            ToolTipText     =   "Confirmar existência código digitado"
            Top             =   210
            Width           =   435
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Height          =   405
            Left            =   6480
            Picture         =   "frmInmetroCadastroModelo.frx":111A
            Style           =   1  'Graphical
            TabIndex        =   3
            TabStop         =   0   'False
            ToolTipText     =   "Limpar campos"
            Top             =   210
            Width           =   435
         End
         Begin VB.CommandButton cmd_pesquisa 
            Caption         =   "..."
            Height          =   255
            Left            =   1530
            Style           =   1  'Graphical
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   300
            Width           =   255
         End
         Begin VB.TextBox TXT_CODIGO 
            Height          =   315
            Left            =   1020
            MaxLength       =   3
            TabIndex        =   0
            Text            =   "0"
            Top             =   270
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo:"
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   300
            Width           =   540
         End
      End
   End
End
Attribute VB_Name = "frmInmetroCadastroModelo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável para MDIapp
Public Flag_ativo As Boolean 'Conterá true se o form ja foi ativado
Public cRec As ADODB.Recordset 'conterá os dados do registro corrente
Public Confirma_Mudanca As Boolean 'Servirá para confirmar as mudanças de alteracoes dos campos na tela
Public cTipo_Movimentacao As Integer  'Se for 1=Inclusão;2=Alteração

Private Sub btoCancelar_Click()
Unload Me
End Sub

Private Sub cmd_pesquisa_Click()
Dim oTela As frmPesquisarInmCadModelo

Set oTela = New frmPesquisarInmCadModelo

    oTela.Show 1
    If oTela.ccodigo_pesquisa = "" Then
        Me.TXT_CODIGO.Text = ""
        Me.TXT_CODIGO.BackColor = &H80000005
        Me.TXT_CODIGO.Enabled = True
        Me.TXT_CODIGO.SetFocus
    Else
        Me.TXT_CODIGO.Text = oTela.ccodigo_pesquisa
        Call cmdOk_Click
    End If
    Unload oTela: Set oTela = Nothing

End Sub

Private Sub cmdCancelar_Click()
Call Limpar_campos
Call Desabilitar_Campos

Me.cmdNovo.Enabled = True
Me.cmdSalvar.Enabled = False
Me.cmdExcluir.Enabled = False

Me.TXT_CODIGO.Enabled = True
Me.TXT_CODIGO.BackColor = &H80000005
Me.TXT_CODIGO.SetFocus

End Sub
Function Limpar_campos()
Me.TXT_CODIGO.Text = ""
Me.TXT_DESCRICAO.Text = ""
Me.TXT_REGISTRO_INMETRO.Text = ""
Me.TXT_DESC_MOD1.Text = ""
Me.TXT_DESC_MOD2.Text = ""
Me.TXT_DESC_MOD3.Text = ""
Me.TXT_DESC_MOD4.Text = ""

End Function

Private Sub cmdExcluir_Click()

On Error GoTo Erro

'Se confirmou a exclusão:
If MsgBox("Deseja Excluir este Modelo", vbQuestion + vbYesNo, "ATENÇÃO !!!") = vbYes Then
    Me.MousePointer = vbHourglass
    Call CCTempneInmetroCadmodelo.INM_CAD_MODELO_Excluir(sBancoMusashi, Me.TXT_CODIGO.Text)
    Me.MousePointer = vbDefault
    MsgBox "Exclusão realizada com sucesso!"
    Me.Confirma_Mudanca = False
    Call Limpar_campos
    Call Desabilitar_Campos
    Me.cmdNovo.Enabled = True
    Me.cmdSalvar.Enabled = False
    Me.cmdExcluir.Enabled = False
    Me.cTipo_Movimentacao = 0
    Me.cmd_pesquisa.Enabled = True
    Me.cmdOk.Enabled = True
    Me.TXT_CODIGO.Enabled = True
    Me.TXT_CODIGO.BackColor = &H80000005
    Me.TXT_CODIGO.SetFocus
End If

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

Private Sub cmdNovo_Click()
Call Limpar_campos
Call Habilitar_Campos
Me.TXT_CODIGO.BackColor = &H8000000F
Me.cmdCancelar.Default = False
cTipo_Movimentacao = 1
Confirma_Mudanca = True
Me.TXT_DESCRICAO.SetFocus

End Sub

Private Sub cmdOk_Click()

On Error GoTo Erro
Me.cmdOk.Default = False
Me.MousePointer = vbHourglass

If Trim(Len(Me.TXT_CODIGO.Text)) = 0 Then
   MsgBox "Digite um Valor para Confirmar o Código do Cliente", , Me.Caption
   Me.MousePointer = vbDefault
   Me.TXT_CODIGO.SetFocus
   Exit Sub
End If

TXT_CODIGO.Text = Format(TXT_CODIGO, "0")

Set cRec = CCTempneInmetroCadmodelo.INM_CAD_MODELO_Consultar(sBancoMusashi, Me.TXT_CODIGO.Text)

Call Habilitar_Campos
Call Carregar_campos

Me.TXT_CODIGO.BackColor = &H8000000F
Me.TXT_CODIGO.Enabled = False

Me.TXT_DESCRICAO.SetFocus
Confirma_Mudanca = True
Me.cmdExcluir.Enabled = True
cTipo_Movimentacao = 2

Set cRec = Nothing
Me.MousePointer = vbDefault
Exit Sub

Erro:
cTipo_Movimentacao = 0
Set cRec = Nothing
MsgBox Err.Description
Me.MousePointer = vbDefault
Me.TXT_CODIGO.SetFocus
Me.cmdOk.Default = True
End Sub

Private Sub cmdSalvar_Click()

Dim otemp As ADODB.Recordset
Dim cCodigo As String
Dim nx As Integer

Me.MousePointer = vbHourglass
On Error GoTo Erro


If cTipo_Movimentacao = 1 Then 'Inclusão da USUARIO
'   Me.TXT_REGISTRO_INMETRO.Text = Cripta(Me.TXT_REGISTRO_INMETRO.Text)
   Set otemp = CCTempneInmetroCadmodelo.INM_CAD_MODELO_Incluir(sBancoMusashi, _
                                                                "", _
                                                                Me.TXT_DESCRICAO.Text, _
                                                                Me.TXT_REGISTRO_INMETRO.Text, _
                                                                Me.TXT_DESC_MOD1.Text, _
                                                                Me.TXT_DESC_MOD2.Text, _
                                                                Me.TXT_DESC_MOD3.Text, _
                                                                Me.TXT_DESC_MOD4.Text)
   Me.TXT_CODIGO.Text = Format(otemp.Fields.Item(0), "0")

Else ' Alteração da USUARIO
'   Me.TXT_REGISTRO_INMETRO.Text = Cripta(Me.TXT_REGISTRO_INMETRO.Text)
   Call CCTempneInmetroCadmodelo.INM_CAD_MODELO_Alterar(sBancoMusashi, _
                                                          Me.TXT_CODIGO.Text, _
                                                          Me.TXT_DESCRICAO.Text, _
                                                          Me.TXT_REGISTRO_INMETRO.Text, _
                                                          Me.TXT_DESC_MOD1.Text, _
                                                          Me.TXT_DESC_MOD2.Text, _
                                                          Me.TXT_DESC_MOD3.Text, _
                                                          Me.TXT_DESC_MOD4.Text)
End If

Me.MousePointer = vbDefault
cCodigo = Me.TXT_CODIGO.Text
Desabilitar_Campos
Confirma_Mudanca = False
cTipo_Movimentacao = 0
Me.cmdNovo.Enabled = True
Me.cmdExcluir.Enabled = False
Me.cmdSalvar.Enabled = False
Me.cmd_pesquisa.Enabled = True
Me.cmdOk.Enabled = True
Me.cmdCancelar.SetFocus
Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

If Err.Number = 50001 Then
   Me.TXT_DESCRICAO.SetFocus
End If

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Flag_ativo = False
End Sub
Private Sub Form_Activate()
If Flag_ativo = True Then
   Exit Sub
End If
Me.Top = 0
Me.Left = 0
Flag_ativo = True
Call Limpar_campos
Call Desabilitar_Campos

Me.cmdNovo.Enabled = True
Me.cmdSalvar.Enabled = False
Me.cmdExcluir.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If cmdSalvar.Enabled = True Then
    If 7 = MsgBox("Possiveis alterações foram realizadas sem salvar,Deseja abandonar?", 32 + 4) Then
        'respondeu não
        Cancel = True
    End If
End If

End Sub
Private Function Habilitar_Campos()
Me.TXT_DESCRICAO.Enabled = True
Me.TXT_REGISTRO_INMETRO.Enabled = True
Me.TXT_DESC_MOD1.Enabled = True
Me.TXT_DESC_MOD2.Enabled = True
Me.TXT_DESC_MOD3.Enabled = True
Me.TXT_DESC_MOD4.Enabled = True

Me.TXT_DESCRICAO.BackColor = &H80000005
Me.TXT_REGISTRO_INMETRO.BackColor = &H80000005
Me.TXT_DESC_MOD1.BackColor = &H80000005
Me.TXT_DESC_MOD2.BackColor = &H80000005
Me.TXT_DESC_MOD3.BackColor = &H80000005
Me.TXT_DESC_MOD4.BackColor = &H80000005
End Function
Private Function Desabilitar_Campos()
Me.TXT_DESCRICAO.Enabled = False
Me.TXT_REGISTRO_INMETRO.Enabled = False
Me.TXT_DESC_MOD1.Enabled = False
Me.TXT_DESC_MOD2.Enabled = False
Me.TXT_DESC_MOD3.Enabled = False
Me.TXT_DESC_MOD4.Enabled = False

Me.TXT_DESCRICAO.BackColor = &H8000000F
Me.TXT_REGISTRO_INMETRO.BackColor = &H8000000F
Me.TXT_DESC_MOD1.BackColor = &H8000000F
Me.TXT_DESC_MOD2.BackColor = &H8000000F
Me.TXT_DESC_MOD3.BackColor = &H8000000F
Me.TXT_DESC_MOD4.BackColor = &H8000000F

End Function
Function Carregar_campos()
Me.TXT_DESCRICAO.Text = cRec!descricao
Me.TXT_REGISTRO_INMETRO.Text = IIf(IsNull(cRec!REGISTRO_INMETRO), "", cRec!REGISTRO_INMETRO)
Me.TXT_DESC_MOD1.Text = IIf(IsNull(cRec!DESC_MOD1), "", cRec!DESC_MOD1)
Me.TXT_DESC_MOD2.Text = IIf(IsNull(cRec!DESC_MOD2), "", cRec!DESC_MOD2)
Me.TXT_DESC_MOD3.Text = IIf(IsNull(cRec!DESC_MOD3), "", cRec!DESC_MOD3)
Me.TXT_DESC_MOD4.Text = IIf(IsNull(cRec!DESC_MOD4), "", cRec!DESC_MOD4)
End Function

Private Sub TXT_DESC_MOD4_Change()
Confirmar_Mudanca
End Sub

Private Sub TXT_DESCRICAO_Change()
Confirmar_Mudanca
End Sub

Private Sub TXT_DESC_MOD1_Change()
Confirmar_Mudanca
End Sub

Private Sub TXT_DESC_MOD2_Change()
Confirmar_Mudanca
End Sub

Private Sub TXT_DESC_MOD3_Change()
Confirmar_Mudanca
End Sub

Private Sub TXT_REGISTRO_INMETRO_Change()
Confirmar_Mudanca
End Sub
Public Function Confirmar_Mudanca()
If Confirma_Mudanca And Me.cmdSalvar.Visible = True Then
   Me.cmdExcluir.Enabled = False
   Me.cmdNovo.Enabled = False
   Me.cmdSalvar.Enabled = True
End If

End Function



