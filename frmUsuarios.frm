VERSION 5.00
Begin VB.Form frmUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de usuários"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmUsuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5535
   Begin VB.Frame Frame4 
      BorderStyle     =   0  'None
      Height          =   3045
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   5445
      Begin VB.Frame Frame1 
         Caption         =   "Confirmar código usuário"
         Height          =   705
         Left            =   60
         TabIndex        =   15
         Top             =   60
         Width           =   5265
         Begin VB.CommandButton cmd_pesquisa 
            Caption         =   "..."
            Height          =   255
            Left            =   1230
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   300
            Width           =   255
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Height          =   405
            Left            =   4620
            Picture         =   "frmUsuarios.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "Limpar campos"
            Top             =   210
            Width           =   435
         End
         Begin VB.CommandButton cmdOk 
            Height          =   405
            Left            =   4140
            Picture         =   "frmUsuarios.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Confirmar existência código digitado"
            Top             =   210
            Width           =   435
         End
         Begin VB.TextBox txt_codigo 
            Height          =   315
            Left            =   720
            MaxLength       =   3
            TabIndex        =   19
            Text            =   "000"
            Top             =   270
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cod.:"
            Height          =   195
            Left            =   180
            TabIndex        =   20
            Top             =   300
            Width           =   375
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dados"
         Height          =   1395
         Left            =   60
         TabIndex        =   7
         Top             =   810
         Width           =   5265
         Begin VB.ComboBox cbo_tipo 
            Height          =   315
            ItemData        =   "frmUsuarios.frx":091E
            Left            =   3390
            List            =   "frmUsuarios.frx":0920
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   960
            Width           =   1725
         End
         Begin VB.TextBox txt_Matricula 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3390
            MaxLength       =   6
            TabIndex        =   10
            Top             =   600
            Width           =   1005
         End
         Begin VB.TextBox txtSenha 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   690
            MaxLength       =   20
            PasswordChar    =   "*"
            TabIndex        =   11
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txtLogin 
            Height          =   285
            Left            =   690
            MaxLength       =   15
            TabIndex        =   9
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtNome 
            Height          =   285
            Left            =   690
            MaxLength       =   40
            TabIndex        =   8
            Top             =   270
            Width           =   4395
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
            Height          =   195
            Left            =   2850
            TabIndex        =   22
            Top             =   990
            Width           =   315
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Matr."
            Height          =   195
            Left            =   2880
            TabIndex        =   21
            Top             =   630
            Width           =   360
         End
         Begin VB.Label lblSenha 
            AutoSize        =   -1  'True
            Caption         =   "Senha"
            Height          =   195
            Left            =   150
            TabIndex        =   14
            Top             =   990
            Width           =   465
         End
         Begin VB.Label lblLogin 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Login"
            Height          =   255
            Left            =   180
            TabIndex        =   13
            Top             =   630
            Width           =   390
         End
         Begin VB.Label lblNome 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            Height          =   225
            Left            =   180
            TabIndex        =   12
            Top             =   300
            Width           =   420
         End
      End
      Begin VB.CommandButton btoCancelar 
         Height          =   615
         Left            =   4680
         Picture         =   "frmUsuarios.frx":0922
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Fecha a solicitação"
         Top             =   2310
         Width           =   615
      End
      Begin VB.CommandButton cmdExcluir 
         Height          =   615
         Left            =   2670
         Picture         =   "frmUsuarios.frx":0A6C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir registro"
         Top             =   2310
         Width           =   615
      End
      Begin VB.CommandButton cmdSalvar 
         Height          =   615
         Left            =   3990
         Picture         =   "frmUsuarios.frx":0EAE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salvar Registro"
         Top             =   2310
         Width           =   615
      End
      Begin VB.CommandButton cmdNovo 
         Height          =   615
         Left            =   3330
         Picture         =   "frmUsuarios.frx":12F0
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Incluir registro"
         Top             =   2310
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "SENHA PARA CADASTRAMENTO"
      Height          =   2595
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5325
      Begin VB.TextBox SENHA 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         IMEMode         =   3  'DISABLE
         Left            =   1320
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   1
         Text            =   "321211"
         Top             =   810
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmUsuarios"
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

Private Sub cbo_tipo_Click()
Confirmar_Mudanca
End Sub

Private Sub cmd_pesquisa_Click()
Dim oTela As frmPesquisarUsuario

Set oTela = New frmPesquisarUsuario

    oTela.Show 1
    If oTela.ccodigo_pesquisa = "" Then
        Me.txt_codigo.Text = ""
        Me.txt_codigo.BackColor = &H80000005
        Me.txt_codigo.Enabled = True
        Me.txt_codigo.SetFocus
    Else
        Me.txt_codigo.Text = oTela.ccodigo_pesquisa
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

Me.txt_codigo.Enabled = True
Me.txt_codigo.BackColor = &H80000005
Me.txt_codigo.SetFocus
End Sub


Private Sub cmdExcluir_Click()

On Error GoTo Erro

'Se confirmou a exclusão:
If MsgBox("Deseja Excluir este Usuário?", vbQuestion + vbYesNo, "ATENÇÃO !!!") = vbYes Then
    Me.MousePointer = vbHourglass
    Call CCTempneUsuario.USUARIO_Excluir(sBancoMusashi, Me.txt_codigo.Text)
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
    Me.txt_codigo.Enabled = True
    Me.txt_codigo.BackColor = &H80000005
    Me.txt_codigo.SetFocus
End If

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

Private Sub cmdNovo_Click()
Call Limpar_campos
Call Habilitar_Campos
Me.txt_codigo.BackColor = &H8000000F
Me.cmdCancelar.Default = False
cTipo_Movimentacao = 1
Confirma_Mudanca = True
Me.cbo_tipo.ListIndex = 0
Me.txtNome.SetFocus

End Sub

Private Sub cmdOk_Click()

On Error GoTo Erro
Me.cmdOk.Default = False
Me.MousePointer = vbHourglass

If Trim(Len(Me.txt_codigo.Text)) = 0 Then
   MsgBox "Digite um valor para confirmar o código de Usuario", , Me.Caption
   Me.MousePointer = vbDefault
   Me.txt_codigo.SetFocus
   Exit Sub
End If

txt_codigo.Text = Format(txt_codigo, "000")

Set cRec = CCTempneUsuario.USUARIO_Consultar(sBancoMusashi, Me.txt_codigo.Text)

Call Habilitar_Campos
Call Carregar_campos

Me.txt_codigo.BackColor = &H8000000F
Me.txt_codigo.Enabled = False

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
Me.txt_codigo.SetFocus
Me.cmdOk.Default = True
End Sub

Private Sub cmdSalvar_Click()

Dim otemp As ADODB.Recordset
Dim cCodigo As String
Dim NX As Integer

Me.MousePointer = vbHourglass
On Error GoTo Erro


If cTipo_Movimentacao = 1 Then 'Inclusão da USUARIO
'   Me.txtSenha.Text = Cripta(Me.txtSenha.Text)
   Set otemp = CCTempneUsuario.USUARIO_Incluir(sBancoMusashi, _
                                               "", _
                                               Me.txtNome.Text, _
                                               Me.txtLogin.Text, _
                                               Me.txtSenha.Text, _
                                               Me.txt_Matricula.Text, _
                                               Me.cbo_tipo.ListIndex)
   
   Me.txt_codigo.Text = Format(otemp.Fields.Item(0), "000")

Else ' Alteração da USUARIO
'   Me.txtSenha.Text = Cripta(Me.txtSenha.Text)
   Call CCTempneUsuario.USUARIO_Alterar(sBancoMusashi, _
                                        Me.txt_codigo.Text, _
                                        Trim(Me.txtNome.Text), _
                                        Trim(Me.txtLogin.Text), _
                                        Trim(Me.txtSenha.Text), _
                                        Trim(Me.txt_Matricula.Text), _
                                        Me.cbo_tipo.ListIndex)
End If

Me.MousePointer = vbDefault
cCodigo = Me.txt_codigo.Text
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
ElseIf Err.Number = 50002 Then
   Me.txtLogin.SetFocus
ElseIf Err.Number = 50003 Then
   Me.txtSenha.SetFocus
ElseIf Err.Number = 50004 Then
   Me.txt_Matricula.SetFocus
End If
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
Me.SENHA.Text = ""
Me.SENHA.SetFocus

Me.cmdNovo.Enabled = True
Me.cmdSalvar.Enabled = False
Me.cmdExcluir.Enabled = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' para funcionar , tem que mudar o keyPreviwe=true
If KeyCode = 13 Then
   If Me.ActiveControl.TabIndex > 3 Then
      SendKeys "{TAB}"
   End If
ElseIf KeyCode = 27 Then
   If Me.ActiveControl.TabIndex = 0 Then
      If Me.cmdSalvar.Enabled = True Then
        If 6 = MsgBox("Deseja realmente sair deste módulo?", 32 + 4) Then
           Unload Me
        End If
      Else
        Unload Me
      End If
   Else
       SendKeys "+{TAB}" ' retornar campo
   End If
End If

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Me.Frame4.Visible = False
Flag_ativo = False
Me.cbo_tipo.AddItem "Expedição"
Me.cbo_tipo.AddItem "Inspeção Final"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If cmdSalvar.Enabled = True Then
    If 7 = MsgBox("Possiveis alterações foram realizadas sem salvar,Deseja abandonar?", 32 + 4) Then
        'respondeu não
        Cancel = True
    End If
End If


End Sub

Private Sub SENHA_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Unload Me
End If
End Sub

Private Sub SENHA_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
ElseIf KeyAscii = 13 Then
   If Me.SENHA.Text = "321211" Then
      Me.Frame3.Visible = False
      Me.Frame4.Visible = True
      Me.txt_codigo.Text = ""
      Me.txt_codigo.Enabled = True
      Me.txt_codigo.BackColor = &H80000005
      Me.txt_codigo.SetFocus
   Else
      MsgBox "Senha não confere, Tente novamente ou tecle <ESC>, para sair."
      Me.SENHA.Text = ""
      Me.SENHA.SetFocus
   End If
End If
End Sub

Private Sub SENHA_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Unload Me
End If

End Sub

Private Sub txt_codigo_GotFocus()
Me.cmdOk.Default = True
End Sub

Private Sub txt_codigo_LostFocus()
Me.cmdOk.Default = False
End Sub

Private Sub txt_Matricula_Change()
Confirmar_Mudanca
End Sub

Private Sub txtLogin_Change()
Confirmar_Mudanca
End Sub

Private Sub txtLogin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNome_Change()
Confirmar_Mudanca
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Function Habilitar_Campos()
Me.txtNome.Enabled = True
Me.txtLogin.Enabled = True
Me.txtSenha.Enabled = True
Me.txt_Matricula.Enabled = True
Me.cbo_tipo.Enabled = True

Me.txtNome.BackColor = &H80000005
Me.txtLogin.BackColor = &H80000005
Me.txtSenha.BackColor = &H80000005
Me.txt_Matricula.BackColor = &H80000005
Me.cbo_tipo.BackColor = &H80000005
End Function
Private Function Desabilitar_Campos()
Me.txtNome.Enabled = False
Me.txtLogin.Enabled = False
Me.txtSenha.Enabled = False
Me.txt_Matricula.Enabled = False
Me.cbo_tipo.Enabled = False

Me.txtNome.BackColor = &H8000000F
Me.txtLogin.BackColor = &H8000000F
Me.txtSenha.BackColor = &H8000000F
Me.txt_Matricula.BackColor = &H8000000F
Me.cbo_tipo.BackColor = &H8000000F

End Function

Private Sub txtSenha_Change()
Confirmar_Mudanca
End Sub
Public Function Confirmar_Mudanca()
If Confirma_Mudanca And Me.cmdSalvar.Visible = True Then
   Me.cmdExcluir.Enabled = False
   Me.cmdNovo.Enabled = False
   Me.cmdSalvar.Enabled = True
End If

End Function
Function Limpar_campos()
Me.txt_codigo.Text = ""
Me.txtNome.Text = ""
Me.txtLogin.Text = ""
Me.txtSenha.Text = ""
Me.txt_Matricula.Text = ""
Me.cbo_tipo.ListIndex = -1
End Function

Function Carregar_campos()
Me.txtLogin.Text = cRec!login
Me.txtNome.Text = cRec!nome
Me.txtSenha.Text = IIf(IsNull(cRec!SENHA), "", cRec!SENHA)
Me.txt_Matricula.Text = cRec!matricula
Me.cbo_tipo.ListIndex = cRec!Tipo

End Function
