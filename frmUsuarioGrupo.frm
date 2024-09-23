VERSION 5.00
Begin VB.Form frmUsuarioGrupo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupo de Usuários"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   8550
   Begin VB.Frame Frame1 
      Caption         =   "Usuários"
      Height          =   5505
      Left            =   60
      TabIndex        =   11
      Top             =   840
      Width           =   8355
      Begin VB.CommandButton cmd_excluir_todos 
         Caption         =   "Excluir T&odos"
         Height          =   315
         Left            =   6780
         TabIndex        =   14
         Top             =   2580
         Width           =   1245
      End
      Begin VB.CommandButton cmd_somar_todos 
         Caption         =   "Somar &Todos"
         Height          =   315
         Left            =   150
         TabIndex        =   12
         Top             =   2580
         Width           =   1245
      End
      Begin VB.ListBox lst_usuario 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1980
         Left            =   150
         TabIndex        =   17
         Top             =   450
         Width           =   8085
      End
      Begin VB.ListBox lst_UsuarioBanco 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   2220
         Left            =   150
         TabIndex        =   16
         Top             =   3030
         Width           =   8085
      End
      Begin VB.CommandButton cmd_excluir 
         Caption         =   "&Excluir"
         Height          =   315
         Left            =   7410
         TabIndex        =   15
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton cmd_somar 
         Caption         =   "&Somar"
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuários pertencentes ao grupo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   2745
         TabIndex        =   19
         Top             =   2760
         Width           =   2745
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Todos os usuários"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3015
         TabIndex        =   18
         Top             =   150
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   60
      TabIndex        =   4
      Top             =   60
      Width           =   8385
      Begin VB.TextBox TXT_DESCRICAO 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   3060
         MaxLength       =   80
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   5175
      End
      Begin VB.CommandButton cmd_Confirmar_Escala 
         Height          =   255
         Left            =   2340
         Picture         =   "frmUsuarioGrupo.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   270
         Width           =   345
      End
      Begin VB.CommandButton cmd_Cancelar_Escala 
         Height          =   255
         Left            =   2670
         Picture         =   "frmUsuarioGrupo.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   270
         Width           =   345
      End
      Begin VB.CommandButton cmd_Pesquisa_escala 
         Caption         =   "..."
         Height          =   255
         Left            =   1950
         Picture         =   "frmUsuarioGrupo.frx":0294
         TabIndex        =   5
         Top             =   270
         Width           =   375
      End
      Begin VB.TextBox txt_codigo 
         Height          =   315
         Left            =   1290
         MaxLength       =   6
         TabIndex        =   9
         Top             =   240
         Width           =   1755
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Código grupo : "
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   270
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdnovo 
      Caption         =   "&Novo"
      Height          =   330
      Left            =   3180
      TabIndex        =   3
      Top             =   6420
      Width           =   1275
   End
   Begin VB.CommandButton cmdsalvar 
      Caption         =   "&Salvar"
      Height          =   330
      Left            =   5835
      TabIndex        =   2
      Top             =   6420
      Width           =   1275
   End
   Begin VB.CommandButton cmdfechar 
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   7170
      TabIndex        =   1
      Top             =   6420
      Width           =   1275
   End
   Begin VB.CommandButton cmd_excluir_Grupo 
      Caption         =   "&Excluir"
      Height          =   330
      Left            =   4515
      TabIndex        =   0
      Top             =   6420
      Width           =   1275
   End
End
Attribute VB_Name = "frmUsuarioGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável para MDIapp
Public Flag_ativo As Boolean 'Conterá true se o form ja foi ativado
Public cRec As ADODB.Recordset 'conterá os dados do registro corrente
Public cData_alteracao As String 'Data de alteracao vinda do registro a ser alterado ou excluido
Public Confirma_Mudanca As Boolean 'Servirá para confirmar as mudanças de alteracoes dos campos na tela
Public sManutencao As Boolean ' Servirá para confirmar alteração dos campos
Public cTipo_Movimentacao As Integer  'Se for 1=Inclusão;2=Alteração

Private Sub cmd_Cancelar_Escala_Click()
Me.txt_codigo.Text = ""
Me.TXT_DESCRICAO.Text = ""
Me.TXT_DESCRICAO.BackColor = &H8000000F
Me.TXT_DESCRICAO.Enabled = False
Me.txt_codigo.Enabled = True
Me.txt_codigo.BackColor = &H80000005
Me.lst_UsuarioBanco.Clear

Me.cmdnovo.Enabled = True
Me.cmdsalvar.Enabled = False
Me.cmd_excluir_Grupo.Enabled = False
Me.cmd_Confirmar_Escala.Enabled = True
Me.cmd_Pesquisa_escala.Enabled = True

Me.cmd_somar.Enabled = False
Me.cmd_somar_todos.Enabled = False
Me.cmd_excluir.Enabled = False
Me.cmd_excluir_todos.Enabled = False

Me.lst_usuario.Enabled = False
Me.lst_UsuarioBanco.Clear
Me.lst_UsuarioBanco.Enabled = False

cTipo_Movimentacao = 0

Me.txt_codigo.SetFocus

End Sub

Private Sub cmd_Confirmar_Escala_Click()

On Error GoTo Erro

Me.cmd_Confirmar_Escala.Default = False
Me.MousePointer = vbHourglass
Set cRec = New ADODB.Recordset

If Trim(Len(Me.txt_codigo.Text)) = 0 And Me.txt_codigo.Enabled = True Then
   MsgBox "Digite um valor para confirmar o código do Grupo de Usuarios", , Me.Caption
   Me.MousePointer = vbDefault
   Me.txt_codigo.SetFocus
   Exit Sub
End If

Me.txt_codigo.Text = Format(txt_codigo, "000")

Set cRec = New ADODB.Recordset
Set cRec = CCTempneGrupoUsuario.GrupoUsuario_ConsultarGrupoPosto(sBancoMusashi, Me.txt_codigo.Text)

If cRec Is Nothing Then
   If cTipo_Movimentacao = 0 Then
      MsgBox "Não Existe Grupo de Usuarios com este Código, Tente Outro!"
   End If
   Me.MousePointer = vbDefault
   Set cRec = Nothing
   Exit Sub
End If
If cRec.RecordCount > 0 Then
   TXT_DESCRICAO.Text = cRec!Descricao
   Call CARREGA_GRUPO_POSTO
   Me.cmdnovo.Enabled = False
   Me.lst_usuario.Enabled = True
   cTipo_Movimentacao = 0
Else
   Me.txt_codigo.Text = ""
   Me.TXT_DESCRICAO.Text = ""
   MsgBox "Não Existe Grupo de Usuarios com este Código, Tente Outro!"
   Me.MousePointer = vbDefault
   Set cRec = Nothing
   Exit Sub
End If

Me.cmdnovo.Enabled = False
Me.cmdsalvar.Enabled = True
Me.cmd_excluir_Grupo.Enabled = True
Me.cmd_Confirmar_Escala.Enabled = False

Me.cmd_somar.Enabled = True
Me.cmd_somar_todos.Enabled = True
Me.cmd_excluir.Enabled = True
Me.cmd_excluir_todos.Enabled = True

Me.lst_usuario.Enabled = True
Me.lst_UsuarioBanco.Enabled = True

Me.TXT_DESCRICAO.BackColor = &H80000005
Me.TXT_DESCRICAO.Enabled = True
Me.txt_codigo.Enabled = False
Me.txt_codigo.BackColor = &H8000000F

Set cRec = Nothing
Me.MousePointer = vbDefault
Exit Sub

Erro:
Set cRec = Nothing
MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub

Private Sub cmd_Excluir_Click()
If Me.lst_UsuarioBanco.ListCount = 0 Then
   Me.cmdsalvar.Enabled = False
   Exit Sub
End If
Me.lst_UsuarioBanco.RemoveItem (Me.lst_UsuarioBanco.ListIndex)
Me.lst_UsuarioBanco.Refresh
Me.cmdsalvar.Enabled = True
End Sub

Private Sub cmd_excluir_Grupo_Click()
Dim nRet As Integer
On Error GoTo Erro

nRet = MsgBox("Confirma exclusão?", vbQuestion & vbYesNo, Me.Caption)
'Se confirmou a exclusão:
If nRet = 6 Then
    Me.MousePointer = vbDefault
    Call CCTempneGrupoUsuario.GrupoUsuario_Excluir(sBancoMusashi, Me.txt_codigo.Text)
    Me.MousePointer = vbDefault
    Me.txt_codigo.BackColor = &H80000005
    Me.txt_codigo.Enabled = True
    
    Me.cTipo_Movimentacao = 0
    Call cmd_Cancelar_Escala_Click
End If

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

Private Sub cmd_excluir_todos_Click()
Me.lst_UsuarioBanco.Clear
Me.lst_UsuarioBanco.Refresh
Me.cmdsalvar.Enabled = False
End Sub

Private Sub cmd_Pesquisa_escala_Click()
Dim oTela As frmPesquisarGrupoUsuario
Set oTela = New frmPesquisarGrupoUsuario

    oTela.Show 1
    If oTela.ccodigo_pesquisa = "" Then
        Me.txt_codigo.Text = ""
        Me.TXT_DESCRICAO.Text = ""
        Me.txt_codigo.Enabled = True
        Me.txt_codigo.BackColor = &H80000005
        Me.txt_codigo.SetFocus
    Else
        Me.txt_codigo.Text = oTela.ccodigo_pesquisa
        Me.TXT_DESCRICAO.Text = oTela.cnome
        cmd_Confirmar_Escala_Click
    End If
    Unload oTela: Set oTela = Nothing

End Sub

Private Sub cmd_somar_Click()
Call Somar_Posto
End Sub

Private Sub cmd_somar_todos_Click()
Call Somar_Todos_Posto
End Sub

Private Sub cmdfechar_Click()
Unload Me
End Sub

Private Sub cmdnovo_Click()
Me.txt_codigo.Text = ""
Me.txt_codigo.Enabled = False
Me.txt_codigo.BackColor = &H8000000F
Me.lst_UsuarioBanco.Clear
Me.TXT_DESCRICAO.BackColor = &H80000005
Me.TXT_DESCRICAO.Enabled = True
Me.cmd_Confirmar_Escala.Enabled = False
Me.cmd_Pesquisa_escala.Enabled = False
Me.cmd_excluir_Grupo.Enabled = False
Me.cmdnovo.Enabled = False
Me.cmd_somar.Enabled = True
Me.cmd_somar_todos.Enabled = True
Me.cmd_excluir.Enabled = True
Me.cmd_excluir_todos.Enabled = True
cTipo_Movimentacao = 1
Me.lst_usuario.Enabled = True
Me.lst_UsuarioBanco.Enabled = True

Me.TXT_DESCRICAO.SetFocus
End Sub

Private Sub cmdsalvar_Click()
Dim otemp As ADODB.Recordset
Dim cCodigo As String
Dim nx As Integer
Dim cFields As Collection

Me.MousePointer = vbHourglass
On Error GoTo Erro

Set cFields = New Collection

For nx = 1 To Me.lst_UsuarioBanco.ListCount
    cFields.Add Trim(Mid$(Me.lst_UsuarioBanco.List(nx - 1), 1, 15))
Next


If cTipo_Movimentacao = 1 Then 'Inclusão do Grupo de Usuarios
   Set otemp = CCTempneGrupoUsuario.GrupoUsuario_Incluir(sBancoMusashi, "", _
                                                       Me.TXT_DESCRICAO.Text, _
                                                       cFields, sUsuario)
   Me.txt_codigo.Text = Format(otemp.Fields(0), "000")
   Set otemp = Nothing
Else ' Alteração do Grupo de Usuarios
   Set otemp = CCTempneGrupoUsuario.GrupoUsuario_Alterar(sBancoMusashi, _
                                                       Me.txt_codigo.Text, _
                                                       Me.TXT_DESCRICAO.Text, _
                                                       cFields, _
                                                       sUsuario)
   Me.txt_codigo.Text = Format(otemp.Fields(0), "000")
   Set otemp = Nothing

End If

Me.MousePointer = vbDefault
Confirma_Mudanca = False
cTipo_Movimentacao = 0

Me.cmdnovo.Enabled = True
Me.cmdsalvar.Enabled = False
Me.cmd_excluir_Grupo.Enabled = False
Me.cmd_Confirmar_Escala.Enabled = True

Me.cmd_somar.Enabled = False
Me.cmd_somar_todos.Enabled = False
Me.cmd_excluir.Enabled = False
Me.cmd_excluir_todos.Enabled = False

Me.lst_usuario.Enabled = False
Me.lst_UsuarioBanco.Enabled = False

Me.TXT_DESCRICAO.BackColor = &H8000000F
Me.TXT_DESCRICAO.Enabled = False

Set cFields = Nothing

Exit Sub

Erro:

MsgBox Err.Description, , Me.Caption
Set cFields = Nothing

Me.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()
If Flag_ativo = True Then
   Exit Sub
End If
Me.Top = 0
Me.Left = 0
Flag_ativo = True
Confirma_Mudanca = False
Call cmd_Cancelar_Escala_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' para funcionar , tem que mudar o keyPreviwe=true
If KeyCode = 13 Then
   If Me.ActiveControl.TabIndex > 3 Then
      SendKeys "{TAB}"
   End If
ElseIf KeyCode = 27 Then
   Unload Me
End If

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Carrega_Usuario
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If cmdsalvar.Enabled = True Then
    If 7 = MsgBox("Possiveis alterações foram realizadas sem salvar,Deseja abandonar?", 32 + 4) Then
        'respondeu não
        Cancel = True
    End If
End If

End Sub
Public Sub Carrega_Usuario()
Dim cRec_Usu As ADODB.Recordset
Dim Y As Integer
Dim sNomePosto As String * 15
Dim sCodigoPosto As String * 6
Dim sArea As String * 20

On Error GoTo Erro
Me.MousePointer = vbHourglass
Set cRec_Usu = New ADODB.Recordset

lst_usuario.Clear
Set cRec_Usu = CCTempneTabUsuario.USUARIO_Consultar(sBancoMusashi, "1")

If cRec_Usu.RecordCount > 0 Then
    For Y = 1 To cRec_Usu.RecordCount
        sNomePosto = cRec_Usu!Login
        sCodigoPosto = cRec_Usu!Coligada
        sArea = cRec_Usu!nome
        If Len(Trim(sNomePosto)) > 0 Then lst_usuario.AddItem sNomePosto & " / " & sArea & " - " & sCodigoPosto
        If Y = cRec_Usu.RecordCount Then
        Else
           cRec_Usu.MoveNext
        End If
    Next
Else
    MsgBox "Não existem usuarios cadastrados, por favor cadastre-os!!"
End If


Set cRec_Usu = Nothing
Me.MousePointer = vbDefault

Exit Sub

Erro:
   Set cRec_Usu = Nothing
   MsgBox Err.Description, , Me.Caption
   Me.MousePointer = vbDefault

End Sub

Private Sub lst_Usuario_Click()
Call Somar_Posto
End Sub
Private Function Somar_Posto()
Dim nx As Integer


'ver se ja existe o mesmo selecionado
For nx = 1 To Me.lst_UsuarioBanco.ListCount
    If Mid$(Me.lst_usuario.List(Me.lst_usuario.ListIndex), 1, 15) = Mid$(Me.lst_UsuarioBanco.List(nx - 1), 1, 15) Then
       MsgBox "Item ja selecionado. Marque outro posto!"
       nx = Me.lst_UsuarioBanco.ListCount
       Exit Function
    End If
Next

Me.lst_UsuarioBanco.AddItem Me.lst_usuario.List(Me.lst_usuario.ListIndex)
Me.cmdsalvar.Enabled = True
Me.cmdnovo.Enabled = False
End Function
Private Function Somar_Todos_Posto()
Dim nx As Integer

Me.lst_UsuarioBanco.Clear

'ver se ja existe o mesmo selecionado
For nx = 1 To Me.lst_usuario.ListCount
    Me.lst_UsuarioBanco.AddItem Me.lst_usuario.List(nx - 1)
Next


End Function

Private Sub lst_UsuarioBanco_Click()
Me.lst_UsuarioBanco.RemoveItem (Me.lst_UsuarioBanco.ListIndex)
Me.lst_UsuarioBanco.Refresh
If Me.lst_UsuarioBanco.ListCount = 0 Then
   Me.cmdsalvar.Enabled = False
End If

End Sub
Private Function CARREGA_GRUPO_POSTO()
Dim nx As Integer
Dim sNomePosto As String * 15
Dim sCodigoPosto As String * 6

'ver se ja existe o mesmo selecionado
Me.lst_UsuarioBanco.Clear
For nx = 1 To Me.cRec.RecordCount
    sNomePosto = Me.cRec!Login
    sCodigoPosto = Me.cRec!codigo
    Me.lst_UsuarioBanco.AddItem sNomePosto & " - " & sCodigoPosto
    Me.cRec.MoveNext
Next
Set cRec = Nothing
Me.cmdsalvar.Enabled = True
Me.cmdnovo.Enabled = False

End Function

Private Sub TXT_codigo_GotFocus()
Me.cmd_Confirmar_Escala.Default = True
End Sub

Private Sub TXT_codigo_LostFocus()
Me.cmd_Confirmar_Escala.Default = False
End Sub


