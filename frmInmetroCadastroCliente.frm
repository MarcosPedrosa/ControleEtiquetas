VERSION 5.00
Begin VB.Form frmInmetroCadastroCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Clientes da Inmetro"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5610
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   90
      TabIndex        =   10
      Top             =   60
      Width           =   5445
      Begin VB.CommandButton cmdNovo 
         Height          =   615
         Left            =   3330
         Picture         =   "frmInmetroCadastroCliente.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Incluir registro"
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton cmdSalvar 
         Height          =   615
         Left            =   3990
         Picture         =   "frmInmetroCadastroCliente.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Salvar Registro"
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton cmdExcluir 
         Height          =   615
         Left            =   2670
         Picture         =   "frmInmetroCadastroCliente.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Excluir registro"
         Top             =   1860
         Width           =   615
      End
      Begin VB.CommandButton btoCancelar 
         Height          =   615
         Left            =   4680
         Picture         =   "frmInmetroCadastroCliente.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Fecha a solicitação"
         Top             =   1860
         Width           =   615
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados"
         Height          =   1005
         Left            =   60
         TabIndex        =   13
         Top             =   810
         Width           =   5265
         Begin VB.TextBox txtNome 
            Height          =   285
            Left            =   720
            MaxLength       =   50
            TabIndex        =   4
            Top             =   270
            Width           =   4395
         End
         Begin VB.TextBox txtSAC 
            Height          =   285
            Left            =   720
            MaxLength       =   50
            TabIndex        =   5
            Top             =   600
            Width           =   4395
         End
         Begin VB.Label lblNome 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            Height          =   225
            Left            =   180
            TabIndex        =   15
            Top             =   300
            Width           =   420
         End
         Begin VB.Label lblLogin 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SAC.:"
            Height          =   195
            Left            =   180
            TabIndex        =   14
            Top             =   630
            Width           =   405
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Confirmar Cliente"
         Height          =   705
         Left            =   60
         TabIndex        =   11
         Top             =   60
         Width           =   5265
         Begin VB.CommandButton cmdOk 
            Height          =   405
            Left            =   4140
            Picture         =   "frmInmetroCadastroCliente.frx":0E10
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
            Left            =   4620
            Picture         =   "frmInmetroCadastroCliente.frx":111A
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
            Left            =   1230
            Style           =   1  'Graphical
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   300
            Width           =   255
         End
         Begin VB.TextBox txt_codigo 
            Height          =   315
            Left            =   720
            MaxLength       =   3
            TabIndex        =   0
            Text            =   "0"
            Top             =   270
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cod.:"
            Height          =   195
            Left            =   180
            TabIndex        =   12
            Top             =   300
            Width           =   375
         End
      End
   End
End
Attribute VB_Name = "frmInmetroCadastroCliente"
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
Dim oTela As frmPesquisarInmCadCliente

Set oTela = New frmPesquisarInmCadCliente

    oTela.Show 1
    If oTela.ccodigo_pesquisa = "" Then
        Me.TXT_CODIGO.Text = ""
        Me.TXT_CODIGO.BackColor = &H80000005
        Me.TXT_CODIGO.Enabled = True
        Me.TXT_CODIGO.SetFocus
    Else
        Me.TXT_CODIGO.Text = oTela.ccodigo_pesquisa
        cmdOk_Click
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
Me.txtNome.Text = ""
Me.txtSAC.Text = ""
End Function

Private Sub cmdExcluir_Click()

On Error GoTo Erro

'Se confirmou a exclusão:
If MsgBox("Deseja Excluir este Cliente?", vbQuestion + vbYesNo, "ATENÇÃO !!!") = vbYes Then
    Me.MousePointer = vbHourglass
    Call CCTempneInmetroCadCliente.INM_CAD_CLIENTE_Excluir(sBancoMusashi, Me.TXT_CODIGO.Text)
    Me.MousePointer = vbDefault
    MsgBox "Exclusão realizada com sucesso!", Me.Caption
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
Me.txtNome.SetFocus

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

Set cRec = CCTempneInmetroCadCliente.INM_CAD_CLIENTE_Consultar(sBancoMusashi, Me.TXT_CODIGO.Text)

Call Habilitar_Campos
Call Carregar_campos

Me.TXT_CODIGO.BackColor = &H8000000F
Me.TXT_CODIGO.Enabled = False

Me.txtNome.SetFocus
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
'   Me.txtSAC.Text = Cripta(Me.txtSAC.Text)
   Set otemp = CCTempneInmetroCadCliente.INM_CAD_CLIENTE_Incluir(sBancoMusashi, _
                                                                 "", _
                                                                 Me.txtNome.Text, _
                                                                 Me.txtSAC.Text)
                     
   Me.TXT_CODIGO.Text = Format(otemp.Fields.Item(0), "0")

Else ' Alteração da USUARIO
'   Me.txtSAC.Text = Cripta(Me.txtSAC.Text)
   Call CCTempneInmetroCadCliente.INM_CAD_CLIENTE_Alterar(sBancoMusashi, _
                                                          Me.TXT_CODIGO.Text, _
                                                          Me.txtNome.Text, _
                                                          Me.txtSAC.Text)
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
   Me.txtNome.SetFocus
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
Call cmdCancelar_Click

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
Me.txtNome.Enabled = True
Me.txtSAC.Enabled = True

Me.txtNome.BackColor = &H80000005
Me.txtSAC.BackColor = &H80000005
End Function
Private Function Desabilitar_Campos()
Me.txtNome.Enabled = False
Me.txtSAC.Enabled = False

Me.txtNome.BackColor = &H8000000F
Me.txtSAC.BackColor = &H8000000F

End Function
Function Carregar_campos()
Me.txtNome.Text = cRec!nome
Me.txtSAC.Text = IIf(IsNull(cRec!SAC), "", cRec!SAC)

End Function

Private Sub txtNome_Change()
Confirmar_Mudanca
End Sub

Private Sub txtSAC_Change()
Confirmar_Mudanca
End Sub
Public Function Confirmar_Mudanca()
If Confirma_Mudanca And Me.cmdSalvar.Visible = True Then
   Me.cmdExcluir.Enabled = False
   Me.cmdNovo.Enabled = False
   Me.cmdSalvar.Enabled = True
End If

End Function

