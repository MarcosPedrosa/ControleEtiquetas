VERSION 5.00
Begin VB.Form frmInmetroCadastroPecas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro Peças Inmetro"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7350
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   6495
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   7155
      Begin VB.Frame Frame1 
         Caption         =   "Confirmar Peça Musashi"
         Height          =   705
         Left            =   60
         TabIndex        =   13
         Top             =   60
         Width           =   7005
         Begin VB.CommandButton cmd_pesquisa 
            Caption         =   "..."
            Height          =   255
            Left            =   2610
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   300
            Width           =   255
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Height          =   405
            Left            =   6210
            Picture         =   "frmInmetroCadastroPecas.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "Limpar Campos"
            Top             =   210
            Width           =   435
         End
         Begin VB.CommandButton cmdOk 
            Height          =   405
            Left            =   5730
            Picture         =   "frmInmetroCadastroPecas.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            ToolTipText     =   "Confirmar Existência do Código da Peça da Musashi"
            Top             =   210
            Width           =   435
         End
         Begin VB.TextBox TXT_COD_PECA_MUSASHI 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1020
            MaxLength       =   15
            TabIndex        =   17
            ToolTipText     =   "Digite o Código da Peça para uma Nova ou Alteração da Etiqueta."
            Top             =   270
            Width           =   1875
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo:"
            Height          =   195
            Left            =   180
            TabIndex        =   18
            Top             =   300
            Width           =   540
         End
      End
      Begin VB.CommandButton btoCancelar 
         Height          =   615
         Left            =   6450
         Picture         =   "frmInmetroCadastroPecas.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Fecha a Solicitação"
         Top             =   5820
         Width           =   615
      End
      Begin VB.CommandButton cmdExcluir 
         Height          =   615
         Left            =   4440
         Picture         =   "frmInmetroCadastroPecas.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Excluir registro"
         Top             =   5820
         Width           =   615
      End
      Begin VB.CommandButton cmdSalvar 
         Height          =   615
         Left            =   5760
         Picture         =   "frmInmetroCadastroPecas.frx":0BA0
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Salvar Registro"
         Top             =   5820
         Width           =   615
      End
      Begin VB.CommandButton cmdNovo 
         Enabled         =   0   'False
         Height          =   615
         Left            =   5100
         Picture         =   "frmInmetroCadastroPecas.frx":0FE2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Incluir Novo Registro"
         Top             =   5820
         Width           =   615
      End
      Begin VB.Frame Frame3 
         Caption         =   "Dados"
         Height          =   4935
         Left            =   60
         TabIndex        =   1
         Top             =   780
         Width           =   7005
         Begin VB.Frame Frame6 
            Caption         =   "Selecione o Cliente e Digite a Peça"
            Height          =   945
            Left            =   120
            TabIndex        =   25
            Top             =   270
            Width           =   6765
            Begin VB.TextBox TXT_COD_PECA_CLIENTE 
               Height          =   285
               IMEMode         =   3  'DISABLE
               Left            =   930
               MaxLength       =   20
               TabIndex        =   27
               Top             =   600
               Width           =   2115
            End
            Begin VB.ComboBox CBO_CLIENTE 
               Height          =   315
               Left            =   930
               Style           =   2  'Dropdown List
               TabIndex        =   26
               ToolTipText     =   "Selecione o Cliente desta Etiqueta"
               Top             =   240
               Width           =   5595
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Peça:"
               Height          =   195
               Left            =   150
               TabIndex        =   29
               Top             =   660
               Width           =   420
            End
            Begin VB.Label lblNome 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cliente:"
               Height          =   195
               Left            =   150
               TabIndex        =   28
               Top             =   270
               Width           =   525
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Selecione o Modelo da Peça"
            Height          =   2055
            Left            =   120
            TabIndex        =   19
            Top             =   1290
            Width           =   6765
            Begin VB.TextBox TXT_DESC_MOD4 
               BackColor       =   &H8000000F&
               Height          =   285
               Left            =   900
               Locked          =   -1  'True
               MaxLength       =   42
               TabIndex        =   30
               Top             =   1680
               Width           =   5655
            End
            Begin VB.TextBox TXT_DESC_MOD3 
               BackColor       =   &H8000000F&
               Height          =   285
               Left            =   900
               Locked          =   -1  'True
               MaxLength       =   42
               TabIndex        =   24
               Top             =   1335
               Width           =   5655
            End
            Begin VB.TextBox TXT_DESC_MOD2 
               BackColor       =   &H8000000F&
               Height          =   285
               Left            =   900
               Locked          =   -1  'True
               MaxLength       =   42
               TabIndex        =   23
               Top             =   1005
               Width           =   5655
            End
            Begin VB.TextBox TXT_DESC_MOD1 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   285
               Left            =   900
               MaxLength       =   42
               TabIndex        =   22
               Top             =   660
               Width           =   5655
            End
            Begin VB.ComboBox CBO_CODIGO_MODELO 
               Height          =   315
               Left            =   900
               Style           =   2  'Dropdown List
               TabIndex        =   21
               ToolTipText     =   "Selecione o Modelo para esta Peça"
               Top             =   270
               Width           =   5655
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Modelo:"
               Height          =   195
               Left            =   90
               TabIndex        =   20
               Top             =   270
               Width           =   570
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Descreva os dados da Peça"
            Height          =   1365
            Left            =   120
            TabIndex        =   2
            Top             =   3420
            Width           =   6765
            Begin VB.TextBox TXT_DESC_PECA3 
               Height          =   285
               Left            =   900
               MaxLength       =   42
               TabIndex        =   5
               ToolTipText     =   "Dados da Peça Linha 3."
               Top             =   950
               Width           =   5655
            End
            Begin VB.TextBox TXT_DESC_PECA2 
               Height          =   285
               Left            =   900
               MaxLength       =   42
               TabIndex        =   4
               ToolTipText     =   "Dados da Peça Linha 2."
               Top             =   630
               Width           =   5655
            End
            Begin VB.TextBox TXT_DESC_PECA1 
               Height          =   285
               Left            =   900
               MaxLength       =   42
               TabIndex        =   3
               ToolTipText     =   "Dados da Peça Linha 1."
               Top             =   270
               Width           =   5655
            End
            Begin VB.Label lblLogin 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Linha 1:"
               Height          =   195
               Left            =   150
               TabIndex        =   8
               Top             =   270
               Width           =   570
            End
            Begin VB.Label lblSenha 
               AutoSize        =   -1  'True
               Caption         =   "Linha 2:"
               Height          =   195
               Left            =   150
               TabIndex        =   7
               Top             =   615
               Width           =   570
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Linha 3:"
               Height          =   195
               Left            =   150
               TabIndex        =   6
               Top             =   975
               Width           =   570
            End
         End
      End
   End
End
Attribute VB_Name = "frmInmetroCadastroPecas"
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
Private cDesc_Mod() As String
Private Sub btoCancelar_Click()
Unload Me
End Sub

Private Sub CBO_CLIENTE_Click()
Confirmar_Mudanca
End Sub

Private Sub CBO_CODIGO_MODELO_Click()
Dim nx As Integer

If Me.CBO_CODIGO_MODELO.ListIndex = -1 Then Exit Sub

If Me.CBO_CODIGO_MODELO.ListIndex = 0 Then
   Me.TXT_DESC_MOD1.Text = cDesc_Mod(1)
   Me.TXT_DESC_MOD2.Text = cDesc_Mod(2)
   Me.TXT_DESC_MOD3.Text = cDesc_Mod(3)
   Me.TXT_DESC_MOD4.Text = cDesc_Mod(4)
Else
   nx = 1 + (Me.CBO_CODIGO_MODELO.ListIndex * 4)
   Me.TXT_DESC_MOD1.Text = cDesc_Mod(nx)
   Me.TXT_DESC_MOD2.Text = cDesc_Mod(nx + 1)
   Me.TXT_DESC_MOD3.Text = cDesc_Mod(nx + 2)
   Me.TXT_DESC_MOD4.Text = cDesc_Mod(nx + 3)
End If
Confirmar_Mudanca

End Sub

Private Sub cmd_pesquisa_Click()
Dim oTela As frmPesquisarInmCadPeca

Set oTela = New frmPesquisarInmCadPeca

    oTela.Show 1
    If oTela.ccodigo_pesquisa = "" Then
        Me.TXT_COD_PECA_MUSASHI.Text = ""
        Me.TXT_COD_PECA_MUSASHI.BackColor = &H80000005
        Me.TXT_COD_PECA_MUSASHI.Enabled = True
        Me.TXT_COD_PECA_MUSASHI.SetFocus
    Else
        Me.TXT_COD_PECA_MUSASHI.Text = oTela.ccodigo_pesquisa
        Call cmdOk_Click
    End If
    Unload oTela: Set oTela = Nothing

End Sub

Private Sub cmdCancelar_Click()
Call Limpar_campos
Call Desabilitar_Campos

Me.TXT_COD_PECA_MUSASHI.Text = ""
Me.TXT_COD_PECA_MUSASHI.Enabled = True
Me.TXT_COD_PECA_MUSASHI.BackColor = &H80000005
Me.TXT_COD_PECA_MUSASHI.SetFocus

Me.cmdNovo.Enabled = False
Me.cmdSalvar.Enabled = False
Me.cmdExcluir.Enabled = False

End Sub
Function Limpar_campos()
Me.TXT_COD_PECA_CLIENTE.Text = ""
Me.TXT_DESC_MOD1.Text = ""
Me.TXT_DESC_MOD2.Text = ""
Me.TXT_DESC_MOD3.Text = ""
Me.TXT_DESC_MOD4.Text = ""
Me.TXT_DESC_PECA1.Text = ""
Me.TXT_DESC_PECA2.Text = ""
Me.TXT_DESC_PECA3.Text = ""
Me.CBO_CLIENTE.ListIndex = -1
Me.CBO_CODIGO_MODELO.ListIndex = -1
End Function

Private Sub cmdExcluir_Click()

On Error GoTo Erro

'Se confirmou a exclusão:
If MsgBox("Deseja Excluir esta Peça?", vbQuestion + vbYesNo, "ATENÇÃO !!!") = vbYes Then
    Me.MousePointer = vbHourglass
    Call CCTempneInmetroCadPeca.INM_CAD_PECA_Excluir(sBancoMusashi, Me.TXT_COD_PECA_MUSASHI.Text)
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
    Me.TXT_COD_PECA_MUSASHI.Enabled = True
    Me.TXT_COD_PECA_MUSASHI.BackColor = &H80000005
    Me.TXT_COD_PECA_MUSASHI.SetFocus
End If

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

Private Sub cmdNovo_Click()
Dim bResp As Boolean

If Len(Trim(Me.TXT_COD_PECA_MUSASHI.Text)) = 0 Then
   MsgBox "Digite o Código da Peça da MUSASHI para poder confirmar uma nova Peça da Inmetro."
   Me.TXT_COD_PECA_MUSASHI.SetFocus
   Exit Sub
End If

bResp = CCTempneInmetroCadPeca.INM_CAD_PECA_JaCadastrada(sBancoMusashi, Me.TXT_COD_PECA_MUSASHI.Text)

If bResp Then
   MsgBox "Peça da MUSASHI já cadastrada Confirme com o Botão ao lado para verificar os dados da Etiqueta da Inmetro."
   Me.cmdOk.SetFocus
   Exit Sub
End If

bResp = CCTempneInmetroCadPeca.INM_CAD_PECA_ExisteEtiqueta(sBancoMusashi, Me.TXT_COD_PECA_MUSASHI.Text)

If Not bResp Then
   MsgBox "Código da Peça na MUSASHI Não Existe no cadastro das ETIQUETAS. Verifique o código se está digitado corretamente."
   Me.TXT_COD_PECA_MUSASHI.SetFocus
   Exit Sub
End If

Call Limpar_campos
Call Habilitar_Campos
Me.CBO_CLIENTE.ListIndex = 0
Me.CBO_CODIGO_MODELO.ListIndex = 0

Me.TXT_COD_PECA_MUSASHI.BackColor = &H8000000F
Me.cmdCancelar.Default = False
cTipo_Movimentacao = 1
Confirma_Mudanca = True
Me.TXT_COD_PECA_CLIENTE.SetFocus

End Sub

Private Sub cmdOk_Click()

On Error GoTo Erro
Me.cmdOk.Default = False
Me.MousePointer = vbHourglass

If Trim(Len(Me.TXT_COD_PECA_MUSASHI.Text)) = 0 Then
   MsgBox "Digite o Código da Peça da Musashi", , Me.Caption
   Me.MousePointer = vbDefault
   Me.TXT_COD_PECA_MUSASHI.SetFocus
   Exit Sub
End If

Set cRec = CCTempneInmetroCadPeca.INM_CAD_PECA_Consultar(sBancoMusashi, Trim(Me.TXT_COD_PECA_MUSASHI.Text))

Call Habilitar_Campos
Call Carregar_campos

Me.TXT_COD_PECA_MUSASHI.BackColor = &H8000000F
Me.TXT_COD_PECA_MUSASHI.Enabled = False

Me.TXT_COD_PECA_CLIENTE.SetFocus
Confirma_Mudanca = True
Me.cmdExcluir.Enabled = True
Me.cmdNovo.Enabled = False
cTipo_Movimentacao = 2

Set cRec = Nothing
Me.MousePointer = vbDefault
Exit Sub

Erro:
cTipo_Movimentacao = 0
Set cRec = Nothing
MsgBox Err.Description
Me.MousePointer = vbDefault
Me.TXT_COD_PECA_MUSASHI.Enabled = True
Me.TXT_COD_PECA_MUSASHI.SetFocus
Me.cmdOk.Default = True
End Sub

Private Sub cmdSalvar_Click()

Dim otemp As ADODB.Recordset
Dim cCodigo As String
Dim nx As Integer

Me.MousePointer = vbHourglass
On Error GoTo Erro

If cTipo_Movimentacao = 1 Then 'Inclusão da USUARIO
   Set otemp = CCTempneInmetroCadPeca.INM_CAD_PECA_Incluir(sBancoMusashi, _
                                                           "", _
                                                           Me.TXT_COD_PECA_MUSASHI.Text, _
                                                           Me.TXT_COD_PECA_CLIENTE.Text, _
                                                           Me.CBO_CODIGO_MODELO.ItemData(Me.CBO_CODIGO_MODELO.ListIndex), _
                                                           Me.CBO_CLIENTE.ItemData(Me.CBO_CLIENTE.ListIndex), _
                                                           Me.TXT_DESC_PECA1.Text, _
                                                           Me.TXT_DESC_PECA2.Text, _
                                                           Me.TXT_DESC_PECA3.Text)
'   Me.TXT_COD_PECA_MUSASHI.Text = Format(oTemp.Fields.Item(0), "0")

Else ' Alteração da USUARIO
   Call CCTempneInmetroCadPeca.INM_CAD_PECA_Alterar(sBancoMusashi, _
                                                    "", _
                                                    Me.TXT_COD_PECA_MUSASHI.Text, _
                                                    Me.TXT_COD_PECA_CLIENTE.Text, _
                                                    Me.CBO_CODIGO_MODELO.ItemData(Me.CBO_CODIGO_MODELO.ListIndex), _
                                                    Me.CBO_CLIENTE.ItemData(Me.CBO_CLIENTE.ListIndex), _
                                                    Me.TXT_DESC_PECA1.Text, _
                                                    Me.TXT_DESC_PECA2.Text, _
                                                    Me.TXT_DESC_PECA3.Text)
End If

Me.MousePointer = vbDefault
cCodigo = Me.TXT_COD_PECA_MUSASHI.Text
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
   Me.TXT_COD_PECA_CLIENTE.SetFocus
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
Call Leitura_Dos_Combos
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
Me.TXT_COD_PECA_CLIENTE.Enabled = True
Me.TXT_COD_PECA_MUSASHI.Enabled = True
Me.TXT_DESC_PECA1.Enabled = True
Me.TXT_DESC_PECA2.Enabled = True
Me.TXT_DESC_PECA3.Enabled = True
Me.CBO_CLIENTE.Enabled = True
Me.CBO_CODIGO_MODELO.Enabled = True

Me.TXT_COD_PECA_CLIENTE.BackColor = &H80000005
Me.TXT_COD_PECA_MUSASHI.BackColor = &H80000005
Me.TXT_DESC_PECA1.BackColor = &H80000005
Me.TXT_DESC_PECA2.BackColor = &H80000005
Me.TXT_DESC_PECA3.BackColor = &H80000005
Me.CBO_CLIENTE.BackColor = &H80000005
Me.CBO_CODIGO_MODELO.BackColor = &H80000005
End Function
Private Function Desabilitar_Campos()
Me.TXT_COD_PECA_CLIENTE.Enabled = False
Me.TXT_COD_PECA_MUSASHI.Enabled = False
Me.TXT_DESC_PECA1.Enabled = False
Me.TXT_DESC_PECA2.Enabled = False
Me.TXT_DESC_PECA3.Enabled = False
Me.CBO_CLIENTE.Enabled = False
Me.CBO_CODIGO_MODELO.Enabled = False

Me.TXT_COD_PECA_CLIENTE.BackColor = &H8000000F
Me.TXT_COD_PECA_MUSASHI.BackColor = &H8000000F
Me.TXT_DESC_PECA1.BackColor = &H8000000F
Me.TXT_DESC_PECA2.BackColor = &H8000000F
Me.TXT_DESC_PECA3.BackColor = &H8000000F
Me.CBO_CLIENTE.BackColor = &H8000000F
Me.CBO_CODIGO_MODELO.BackColor = &H8000000F

End Function
Function Carregar_campos()
Dim nx As Integer

Me.TXT_COD_PECA_CLIENTE.Text = cRec!COD_PECA_CLIENTE
Me.TXT_DESC_PECA1.Text = IIf(IsNull(cRec!DESC_PECA1), "", cRec!DESC_PECA1)
Me.TXT_DESC_PECA2.Text = IIf(IsNull(cRec!DESC_PECA2), "", cRec!DESC_PECA2)
Me.TXT_DESC_PECA3.Text = IIf(IsNull(cRec!DESC_PECA3), "", cRec!DESC_PECA3)

If Not IsNull(cRec!CODIGO_CLIENTE) Then
    For nx = 1 To Me.CBO_CLIENTE.ListCount
        If Me.CBO_CLIENTE.ItemData(nx - 1) = cRec!CODIGO_CLIENTE Then
           Me.CBO_CLIENTE.ListIndex = nx - 1
           Exit For
        End If
    Next
    If Me.CBO_CLIENTE.ListIndex = -1 Then Me.CBO_CLIENTE.ListIndex = 0
Else
    If Me.CBO_CLIENTE.ListCount > 0 Then Me.CBO_CLIENTE.ListIndex = 0
End If

If Not IsNull(cRec!CODIGO_MODELO) Then
    For nx = 1 To Me.CBO_CODIGO_MODELO.ListCount
        If Me.CBO_CODIGO_MODELO.ItemData(nx - 1) = cRec!CODIGO_MODELO Then
           Me.CBO_CODIGO_MODELO.ListIndex = nx - 1
           Exit For
        End If
    Next
    If Me.CBO_CODIGO_MODELO.ListIndex = -1 Then Me.CBO_CODIGO_MODELO.ListIndex = 0
Else
    If Me.CBO_CODIGO_MODELO.ListCount > 0 Then Me.CBO_CODIGO_MODELO.ListIndex = 0
End If

End Function
Private Function Leitura_Dos_Combos()
Dim rs As ADODB.Recordset
Dim nx, ny, nz As Integer

On Error GoTo Erro

Me.MousePointer = vbHourglass
Set rs = New ADODB.Recordset

Me.CBO_CLIENTE.Clear
Me.CBO_CODIGO_MODELO.Clear

ReDim cDesc_Mod(100)

Set rs = CCTempneInmetroCadPeca.INM_CAD_PECA_Cons_CliMod(sBancoMusashi)

If rs.RecordCount > 0 Then
   rs.MoveFirst
   nx = 0: ny = 0: nz = 0
   While Not rs.EOF
         If rs!Tipo = "CLI" Then
            Me.CBO_CLIENTE.AddItem Format(rs!codigo, "0") & " - " & rs!descricao
            Me.CBO_CLIENTE.ItemData(ny) = Format(rs!codigo)
            ny = ny + 1
         Else
            Me.CBO_CODIGO_MODELO.AddItem Format(rs!codigo, "0") & " - " & rs!descricao
            Me.CBO_CODIGO_MODELO.ItemData(nz) = Format(rs!codigo)
            nz = nz + 1
            nx = nx + 1
            cDesc_Mod(nx) = rs!DESC_MOD1
            nx = nx + 1
            cDesc_Mod(nx) = IIf(IsNull(rs!DESC_MOD2), "", rs!DESC_MOD2)
            nx = nx + 1
            cDesc_Mod(nx) = IIf(IsNull(rs!DESC_MOD3), "", rs!DESC_MOD3)
            nx = nx + 1
            cDesc_Mod(nx) = IIf(IsNull(rs!DESC_MOD4), "", rs!DESC_MOD4)
         End If
         rs.MoveNext
   Wend
End If

If Me.CBO_CLIENTE.ListCount = 0 Then
   MsgBox "Não Existem Clientes cadastrados. Favor cadastrar Clientes para Inclusão das Peças."
   Call cmdCancelar_Click
   Exit Function
End If

If Me.CBO_CODIGO_MODELO.ListCount = 0 Then
   MsgBox "Não Existem Modelos cadastrados. Favor cadastrar Modelos da etiqueta para Inclusão das Peças."
   Call cmdCancelar_Click
   Exit Function
End If

Me.MousePointer = vbDefault

Exit Function

Erro:

Set rs = Nothing
Me.MousePointer = vbDefault
MsgBox Err.Description

End Function
Private Sub TXT_COD_PECA_CLIENTE_Change()
Confirmar_Mudanca
End Sub

Private Sub TXT_COD_PECA_MUSASHI_Change()
Me.cmdExcluir.Enabled = False
Me.cmdNovo.Enabled = True
Me.cmdSalvar.Enabled = False
End Sub

Private Sub TXT_COD_PECA_MUSASHI_GotFocus()
Me.cmdOk.Default = True
End Sub

Private Sub TXT_COD_PECA_MUSASHI_LostFocus()
Me.cmdOk.Default = False
End Sub

Private Sub TXT_DESC_PECA1_Change()
Confirmar_Mudanca
End Sub

Private Sub TXT_DESC_PECA2_Change()
Confirmar_Mudanca
End Sub

Private Sub TXT_DESC_PECA3_Change()
Confirmar_Mudanca
End Sub

Public Function Confirmar_Mudanca()
If Confirma_Mudanca And Me.cmdSalvar.Visible = True Then
   Me.cmdExcluir.Enabled = False
   Me.cmdNovo.Enabled = False
   Me.cmdSalvar.Enabled = True
End If

End Function





