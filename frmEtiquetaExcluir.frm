VERSION 5.00
Begin VB.Form frmEtiquetaExcluir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exclusão da etiqueta"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6405
   Begin VB.Frame Frame2 
      Caption         =   "Liberar exclusão"
      Height          =   1305
      Left            =   90
      TabIndex        =   10
      Top             =   1230
      Width           =   6225
      Begin VB.CommandButton cmd_excluir 
         Enabled         =   0   'False
         Height          =   735
         Left            =   4950
         Picture         =   "frmEtiquetaExcluir.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Excluir a etiqueta"
         Top             =   330
         Width           =   975
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   1515
         PasswordChar    =   "*"
         TabIndex        =   4
         ToolTipText     =   "Digite a senha do usuário"
         Top             =   750
         Width           =   2325
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1515
         TabIndex        =   3
         ToolTipText     =   "Digite o nome do usuário"
         Top             =   240
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Senha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   825
         Width           =   735
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Usuário:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   375
         Width           =   885
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Etiqueta"
      Height          =   1065
      Left            =   90
      TabIndex        =   8
      Top             =   90
      Width           =   6225
      Begin VB.CommandButton btoConfirma 
         Height          =   615
         Left            =   4830
         Picture         =   "frmEtiquetaExcluir.frx":27A2
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Confirmar sequencia da etiqueta"
         Top             =   300
         Width           =   555
      End
      Begin VB.CommandButton cmd_limpar 
         Height          =   615
         Left            =   5400
         Picture         =   "frmEtiquetaExcluir.frx":2AAC
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "limpar filtro para novo sequencial"
         Top             =   300
         Width           =   555
      End
      Begin VB.TextBox txtsequencial 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   1530
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "0000000000"
         ToolTipText     =   "digite o sequencial da etiqueta"
         Top             =   270
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sequencial :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   510
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdfechar 
      Height          =   735
      Left            =   5340
      Picture         =   "frmEtiquetaExcluir.frx":2EEE
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Saida da tela"
      Top             =   2610
      Width           =   975
   End
   Begin VB.CommandButton cmd_visualizar 
      Enabled         =   0   'False
      Height          =   735
      Left            =   4050
      Picture         =   "frmEtiquetaExcluir.frx":3330
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Imprimir"
      Top             =   1620
      Width           =   975
   End
   Begin VB.CommandButton cmd_Impressao 
      Enabled         =   0   'False
      Height          =   735
      Left            =   5070
      Picture         =   "frmEtiquetaExcluir.frx":3772
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Imprimir"
      Top             =   1590
      Width           =   975
   End
End
Attribute VB_Name = "frmEtiquetaExcluir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável para MDIapp
Public cRec As ADODB.Recordset 'conterá os dados do registro corrente
Public bAtivo As Boolean
Public bTelaImp As Boolean
Public bJafoi As Boolean
Public nCodUsuario As Integer

Private Sub btoConfirma_Click()

Dim nx As Integer
Dim sData As String

On Error GoTo Erro

Me.MousePointer = vbHourglass


Me.txtsequencial.Text = Format(CDbl(Me.txtsequencial.Text), "0000000000")

Set cRec = New ADODB.Recordset

Set cRec = CCTempneMov_Etiq.Mov_Etiq_Consultar_Filtro(sBancoMusashi, _
                                                      Me.txtsequencial.Text, _
                                                      "", _
                                                      "", _
                                                      "", _
                                                      "")

Me.cmd_excluir.Enabled = False
cmd_visualizar_Click

If cRec.RecordCount > 0 Then
   Call Liberar_campos
   Me.txtUserName.SetFocus
End If

Me.MousePointer = vbDefault
Exit Sub

Erro:

MsgBox Err.Description
Me.MousePointer = vbDefault
If Err.Number = 50000 Then
   cmd_limpar_Click
End If
If Err.Number = 50001 Then
   Me.txtsequencial.SetFocus
End If
End Sub

Private Sub cmd_excluir_Click()
Dim nx As Integer
'Dim cFields As Collection
Dim sData As String

If MsgBox("Confirma Exclusão desta etiqueta?", vbQuestion + vbYesNo, "ATENÇÃO !!!") = vbNo Then Exit Sub

On Error GoTo Erro

Me.MousePointer = vbHourglass

Me.txtsequencial.Text = Format(Me.txtsequencial.Text, "0000000000")

Set cRec = New ADODB.Recordset

Set cRec = CCTempneMov_Etiq.Mov_Etiq_Alt_Campos(sBancoMusashi, _
                                                Me.txtsequencial.Text, _
                                                "3", _
                                                "", _
                                                nCodUsuario)

cmd_limpar_Click

Me.MousePointer = vbDefault
MsgBox "Exclusão concluída com sucesso! ", vbOKOnly + vbInformation, "Tarefa Concluída"

Exit Sub

Erro:

MsgBox Err.Description
Me.MousePointer = vbDefault
If Err.Number = 50000 Then
   cmd_limpar_Click
End If
If Err.Number = 50001 Then
   Me.txtsequencial.SetFocus
End If
End Sub


Private Sub cmd_Impressao_Click()

Dim v As Integer

For v = 0 To (Forms.Count - 1)

    If Forms(v).Name = "frmExibicao1" Or Forms(v).Name = "frmExibicao4" Then
          If objApplication.filial = adMusashiDaAmazonia Then
             Printer.Orientation = 1
          Else
             Printer.Orientation = 2
          End If
          frmAvulsoPadraoPonteiro.PrintForm
          Printer.Orientation = 2: Printer.EndDoc
        Exit For
    End If
    If Forms(v).Name = "frmExibicao2" Then
        Printer.Orientation = 1
        frmExibicao2.PrintForm
        Printer.Orientation = 2: Printer.EndDoc
        Exit For
    End If
    If Forms(v).Name = "frmExibicao3" Then
        Printer.Orientation = 2
        frmExibicao3.PrintForm
        Printer.Orientation = 2: Printer.EndDoc
        Exit For
    End If
    If Forms(v).Name = "frmExibicao5" Then
        Printer.Orientation = 2
        frmExibicao5.PrintForm
        Printer.Orientation = 2: Printer.EndDoc
        Exit For
    End If
    If Forms(v).Name = "frmExibicao6" Then
        Printer.Orientation = 2
        frmExibicao6.PrintForm
        Printer.Orientation = 2: Printer.EndDoc
        Exit For
    End If
    If Forms(v).Name = "frmExibicao7UmProduto" Or Forms(v).Name = "frmExibicao7VariosProdutos" Then
        Printer.Orientation = 2
        frmExibicao7Ref.PrintForm
        Printer.Orientation = 2: Printer.EndDoc
        Exit For
    End If
Rem incluido para impressao do novo formulario da mvm international motores
    If Forms(v).Name = "frmExibicao9" Then
        If frmExibicao9.PICT_CRISTAL.Visible = True Then
           Call Imprime_etiqueta_MWM_Cristal
        Else
           Printer.Orientation = 2
           frmExibicao9.PrintForm
           Printer.Orientation = 2: Printer.EndDoc
        End If
        Exit For
    End If
    If Forms(v).Name = "frmExibicao10" Then
        Printer.Orientation = 2
        frmExibicao10.PrintForm
        Printer.Orientation = 2: Printer.EndDoc
        Exit For
    End If
    If Forms(v).Name = "frmExibicao11" Then
        Printer.Orientation = 2
        frmExibicao11.PrintForm
        Printer.Orientation = 2: Printer.EndDoc
        Exit For
    End If
Next
        
MsgBox "Impressão concluída com sucesso! a tela da etiqueta será encerrada!", vbOKOnly + vbInformation, "Tarefa Concluída"

Call Fechar_Form_Etiqueta
cmd_limpar_Click

End Sub

Private Sub cmd_limpar_Click()
Call Fechar_Form_Etiqueta
Call Limpar_campos
Me.txtsequencial.Text = ""

Me.cmd_Impressao.Enabled = False
Me.cmd_Visualizar.Enabled = False
Me.cmd_excluir.Enabled = False
Me.txtsequencial.SetFocus
End Sub

Private Sub cmdfechar_Click()
Unload Me
End Sub

Private Sub Form_Activate()

'If bTelaImp Then
'   cmd_limpar_Click
'   bTelaImp = False
'End If

If bAtivo Then
   If Me.cmd_excluir.Enabled = True Then
      Me.cmd_excluir.SetFocus
   End If
End If

Exit Sub

bAtivo = True
bJafoi = False
Me.txtsequencial.Text = ""
Me.txtsequencial.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' para funcionar , tem que mudar o keyPreviwe=true
If KeyCode = 13 Then
   If Me.ActiveControl.TabIndex > 0 And Me.ActiveControl.TabIndex < 4 Then
      SendKeys "{TAB}"
   End If
ElseIf KeyCode = 27 Then
   If Me.ActiveControl.TabIndex < 1 Then
      If 6 = MsgBox("Deseja realmente sair deste módulo?", 32 + 4) Then
         Unload Me
      End If
   Else
       SendKeys "+{TAB}" ' retornar campo
   End If
End If

End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
bTelaImp = False
Me.txtsequencial.Text = ""
Me.txtUserName.Text = ""
Me.txtPassword.Text = ""
Me.txtUserName.BackColor = &H8000000F
Me.txtPassword.BackColor = &H8000000F
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
bAtivo = False
'Fecha o form q estiver aberto
Call Fechar_Form_Etiqueta
End Sub


Private Sub txtPassword_KeyPress(KeyAscii As Integer)
Dim rRs As ADODB.Recordset

On Error GoTo Erro

nCodUsuario = 0

If KeyAscii = 13 Then
   Me.MousePointer = vbHourglass
   Set rRs = New ADODB.Recordset
   
   Set rRs = CCTempneUsuario.USUARIO_Confirmar_Login(sBancoMusashi, _
                                                     Me.txtUserName.Text, _
                                                     Me.txtPassword.Text)
   Me.MousePointer = vbDefault
   Me.cmd_excluir.Enabled = True
   Me.cmd_excluir.SetFocus
   nCodUsuario = rRs!codigo
   Set rRs = Nothing
End If

Exit Sub

Erro:

MsgBox Err.Description
Me.MousePointer = vbDefault
If Err.Number = 50002 Then
   Me.txtPassword.SetFocus
End If

End Sub

Private Sub txtsequencial_Change()
If Len(Trim(txtsequencial.Text)) = 10 Then
   btoConfirma_Click
End If

End Sub

Private Sub txtsequencial_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Me.btoConfirma.SetFocus
End If
If KeyAscii = 27 Then
   If MsgBox("Deseja sair deste módulo?", vbQuestion + vbYesNo, "ATENÇÃO !!!") = vbNo Then
      Me.txtsequencial.Text = ""
      Me.txtsequencial.SetFocus
   Else
      Unload Me
   End If
End If

End Sub

Private Sub cmd_visualizar_Click()
Rem  *********************************************************************************************************
Rem  *********************************************************************************************************
Rem  *********************************************************************************************************
Rem  *********************************************************************************************************
Rem  *********************************************************************************************************
    
Dim qtdeConteiners As Integer
Dim executaUnloadForm As Boolean
Dim sSeqId As String * 18 ' usado para formatar o campo ID, da MVM
Dim sSequencial As String
Dim nQuantidade As Integer
Dim sdata_aux As String
Dim oTela As Form
    
Rem leitura no banco para emissão da etiqueta
sSequencial = Format(Me.txtsequencial.Text, "0000000000")

Me.MousePointer = vbHourglass

Set cRec = CCTempneMov_Etiq.Mov_Etiq_Consultar(sBancoMusashi, _
                                               sSequencial)

If cRec.RecordCount = 0 Then
   MsgBox "Registro não encontrado! A etiqueta não será impressa!", vbInformation + vbOKOnly, "Tarefa com problemas"
   Exit Sub
End If

Me.MousePointer = vbDefault
    
executaUnloadForm = False

'Fecha o form q estiver aberto
Call Fechar_Form_Etiqueta


'qtdeConteiners = cRec.Fields("Tipo").Value
nQuantidade = cRec.Fields("Qtd_Etiq")
    
    '---------------------------------------------------------------------------------------------------
    'Atualiza o form frmExibicao de acordo com o tipo
    'Se tipo 1 opcao padrão etiqueta pequena
    If cRec.Fields("Tipo") = "1" Then
        frmAvulsoPadraoPonteiro.Show
        frmAvulsoPadraoPonteiro.Left = Me.Width
        If cRec.Fields("Cliente") <> "" Then
            frmAvulsoPadraoPonteiro.lblCliente.Caption = cRec.Fields("Cliente")
        End If
        
        If cRec.Fields("Cod_Cliente") <> "" Then
            frmAvulsoPadraoPonteiro.lblCodCliente2.Caption = cRec.Fields("Cod_Cliente")
        End If
        If cRec.Fields("Descr_Peca") <> "" Then
            frmAvulsoPadraoPonteiro.lblDescricao.Caption = cRec.Fields("Descr_Peca")
        End If
        If cRec.Fields("Lote") <> "" Then
            frmAvulsoPadraoPonteiro.lblLote2.Caption = cRec.Fields("Lote")
        End If
        If cRec.Fields("Peso") <> "" Then
            frmAvulsoPadraoPonteiro.lblPeso2.Caption = Format(cRec.Fields("Peso"), "0.00")
        End If
        If cRec.Fields("Qtd_Caixa") <> "" Then
            frmAvulsoPadraoPonteiro.lblQtd2.Caption = cRec.Fields("Qtd_Caixa")
        End If
        If cRec.Fields("Cod_Peca") <> "" Then
            frmAvulsoPadraoPonteiro.lblPeca.Caption = cRec.Fields("Cod_Peca")
        End If
        If Not IsNull(cRec.Fields("EMBALAGEM")) Then
            frmAvulsoPadraoPonteiro.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
        End If
        frmAvulsoPadraoPonteiro.lbl_data.Caption = Format(cRec.Fields("data_etiq"), "dd/mm/yyyy")
        frmAvulsoPadraoPonteiro.lbl_MES.Caption = Pega_Mes(Mid$(Format(cRec.Fields("data_etiq"), "dd/mm/yyyy"), 4, 2))

        If (cRec.Fields("Tipo") = "4") Then
           frmAvulsoPadraoPonteiro.lblMsgProduto.Visible = False
           If (Val((cRec.Fields("id_cliente")) = 1) Or (Val(cRec.Fields("id_cliente")) = 2) Or (Val(cRec.Fields("id_cliente")) = 3)) Then
              frmAvulsoPadraoPonteiro.lblMsgProduto.Visible = True
           End If
        End If

        If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
            frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption = Trim(cRec.Fields("ID_ETIQUETA"))
            frmAvulsoPadraoPonteiro.lblCodigoBarrasA.Caption = "*" & Trim(frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption) & "*"
            frmAvulsoPadraoPonteiro.lblCodigoBarrasB.Caption = "*" & Trim(frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption) & "*"
         Else
            frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption = ""
            frmAvulsoPadraoPonteiro.lblCodigoBarrasA.Caption = ""
            frmAvulsoPadraoPonteiro.lblCodigoBarrasB.Caption = ""
         End If
    End If
    
    '---------------------------------------------------------------------------------------------------
    'Se tipo 2 opcao FIAT
    If cRec.Fields("Tipo") = "2" Or _
       cRec.Fields("Tipo") = "F" Or _
       cRec.Fields("Tipo") = "Z" Then
        frmExibicao2.nTamannhowidth = Me.Width
        frmExibicao2.Show
        'Mostra FIAT
        frmExibicao2.lblCod_Peca.Caption = cRec.Fields("Cod_Peca")
        If cRec.Fields("Data_Expedicao") <> "" Then
            frmExibicao2.lblDataExpedicao2.Caption = cRec.Fields("Data_Expedicao")
        End If
        If cRec.Fields("Cod_Fornecedor") <> "" Then
            frmExibicao2.lblCodFornec2.Caption = Format(cRec.Fields("Cod_Fornecedor"), "000000")
        End If
        If cRec.Fields("Descr_Peca") <> "" Then
            frmExibicao2.lblDenominacao2.Caption = cRec.Fields("Descr_Peca")
        End If
        If cRec.Fields("Num_Doc_Fiscal") <> "" Then
            frmExibicao2.lblBam2.Caption = cRec.Fields("Num_Doc_Fiscal")
        End If
        If cRec.Fields("Cod_Cliente") <> "" Then
            frmExibicao2.lblDesenho2.Caption = Format(cRec.Fields("Cod_Cliente"), "00000000000")
        End If
        
        frmExibicao2.lblCodBarra.Caption = Format(cRec.Fields("Cod_Cliente"), "00000000000") & Format(cRec.Fields("Qtd_Caixa"), "00000") _
                                      & cRec.Fields("Cod_Embalagem_pw") & Format(cRec.Fields("Cod_Fornecedor"), "000000")
        frmExibicao2.lblCodBarra2.Caption = frmExibicao2.lblCodBarra.Caption
        'Adicionar o * de inicio e fim
        frmExibicao2.lblCodBarra.Caption = "*" & frmExibicao2.lblCodBarra.Caption & "*"
        frmExibicao2.lblCodBarraCp1.Caption = frmExibicao2.lblCodBarra.Caption
        frmExibicao2.lblCodBarraCp2.Caption = frmExibicao2.lblCodBarra.Caption
        
        If cRec.Fields("Data_Lote") <> "" Then
            frmExibicao2.lblDataProducao2.Caption = cRec.Fields("Data_Lote")
        End If
        If cRec.Fields("Cod_Embalagem") <> "" Then
            frmExibicao2.lblCodEmbalagem2.Caption = cRec.Fields("Cod_Embalagem")
        End If
        If cRec.Fields("Lote") <> "" Then
            frmExibicao2.lblNumLote2.Caption = cRec.Fields("Lote")
        End If
        If cRec.Fields("Qtd_Lote") <> "" Then
            frmExibicao2.lblQtdLote2.Caption = cRec.Fields("Qtd_Lote")
        End If
        If cRec.Fields("Qtd_Caixa") <> "" Then
            frmExibicao2.lblQtdEmbalagem2.Caption = Format(cRec.Fields("Qtd_Caixa"), "00000")
        End If
        If cRec.Fields("Classe_Funcional") <> "" Then
            frmExibicao2.lblClasseFuncional2.Caption = cRec.Fields("Classe_Funcional")
        End If
        If cRec.Fields("Vinculo") <> "" Then
            frmExibicao2.lblVinculo2.Caption = cRec.Fields("Vinculo")
        End If
        If cRec.Fields("Ind_Suplementar") <> "" Then
            frmExibicao2.lblIndicacaoSuplementar2.Caption = cRec.Fields("Ind_Suplementar")
        End If
        If cRec.Fields("Embarque_Controlado") <> "" Then
            frmExibicao2.lblEmbarqueControlado2.Caption = cRec.Fields("Embarque_Controlado")
        End If
        If cRec.Fields("Desvio") <> "" Then
            frmExibicao2.lblLoteSobDesvio2.Caption = cRec.Fields("Desvio")
        End If
        If cRec.Fields("DUM") <> "" Then
            frmExibicao2.lblDum2.Caption = cRec.Fields("DUM")
        End If
        If Not IsNull(cRec.Fields("EMBALAGEM")) Then
            frmExibicao2.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
        Else
            frmExibicao2.lblEmbalagem2.Caption = ""
        End If
    Rem acrescentado o pto de entrega - 09-09-2016
        If cRec.Fields("Pto_Entrega") <> "" Then
            frmExibicao2.lblPontoEntrega2.Caption = Trim(cRec.Fields("Pto_Entrega"))
        End If
        
        If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
            frmExibicao2.lblCodigoBarras.Caption = ""
            frmExibicao2.lblCodigoBarras.Caption = Trim(cRec.Fields("ID_ETIQUETA"))
            frmExibicao2.lblCodigoBarrasA.Caption = "*" & Trim(frmExibicao2.lblCodigoBarras.Caption) & "*"
            frmExibicao2.lblCodigoBarrasB.Caption = "*" & Trim(frmExibicao2.lblCodigoBarras.Caption) & "*"
    '            frmExibicao2.lblCodigoBarrasC.Caption = "*" & frmExibicao2.lblCodigoBarras.Caption & "*"
    '            frmExibicao2.lblCodigoBarrasD.Caption = "*" & frmExibicao2.lblCodigoBarras.Caption & "*"
         Else
            frmExibicao2.lblCodigoBarras.Caption = ""
            frmExibicao2.lblCodigoBarrasA.Caption = ""
            frmExibicao2.lblCodigoBarrasB.Caption = ""
         End If
         
         If cRec.Fields("Tipo") = "2" Then ' aqui será preenchido os dados para a outra etiqueta tipo expotacao da fiat italiana
    
            Set oTela = New frmExibicao12
            oTela.slbl_01 = IIf(IsNull(cRec.Fields("Cliente")), " ", Trim(cRec.Fields("Cliente")))
            oTela.slbl_02 = IIf(IsNull(cRec.Fields("Pto_Entrega")), " ", Trim(cRec.Fields("Pto_Entrega")))
            oTela.slbl_03 = " "
            oTela.slbl_04 = "MUSASHI DO BRASIL LTDA"
            oTela.slbl_07 = " "
            oTela.slbl_08 = " " ' FALTA VER COM MAURO
            oTela.slbl_09 = IIf(IsNull(cRec.Fields("Qtd_Caixa")), " ", Trim(cRec.Fields("Qtd_Caixa")))
            oTela.slbl_10 = IIf(IsNull(cRec.Fields("Descr_Peca")), " ", Trim(cRec.Fields("Descr_Peca")))
            oTela.slbl_11 = IIf(IsNull(cRec.Fields("Cod_Fornecedor")), " ", Trim(cRec.Fields("Cod_Fornecedor")))
            oTela.slbl_12 = IIf(IsNull(cRec.Fields("Cod_Embalagem")), " ", Trim(cRec.Fields("Cod_Embalagem")))
            oTela.slbl_13 = Format(Now(), "DD/MM/YYYY")
            oTela.slbl_14 = " "
            oTela.slbl_15 = " "
            oTela.slbl_16 = IIf(IsNull(cRec.Fields("Lote")), " ", Trim(cRec.Fields("Lote")))
            oTela.ldl_usuario.Caption = IIf(IsNull(cRec.Fields("EMBALAGEM")), " ", cRec.Fields("EMBALAGEM"))
            oTela.lbl_Sequencial.Caption = Trim(cRec.Fields("ID_ETIQUETA"))
            oTela.lbl_barras1.Caption = "*" & oTela.lbl_Sequencial.Caption & "*"
            oTela.lbl_barras2.Caption = "*" & oTela.lbl_Sequencial.Caption & "*"
            oTela.Visible = False
            
         End If
    
         If cRec.Fields("Tipo") = "Z" Then ' aqui será preenchido os dados para a outra etiqueta tipo expotacao da fiat italiana
    
            Set oTela = New frmExibicao12
            oTela.slbl_01 = IIf(IsNull(cRec.Fields("Cliente")), " ", Trim(cRec.Fields("Cliente")))
            oTela.slbl_02 = IIf(IsNull(cRec.Fields("Pto_Entrega")), " ", Trim(cRec.Fields("Pto_Entrega")))
            oTela.slbl_03 = " "
            oTela.slbl_04 = "MUSASHI DO BRASIL LTDA"
            oTela.slbl_07 = " "
            oTela.slbl_08 = " " ' FALTA VER COM MAURO
            oTela.slbl_09 = IIf(IsNull(cRec.Fields("Qtd_Caixa")), " ", Trim(cRec.Fields("Qtd_Caixa")))
            oTela.slbl_10 = IIf(IsNull(cRec.Fields("Descr_Peca")), " ", Trim(cRec.Fields("Descr_Peca")))
            oTela.slbl_11 = IIf(IsNull(cRec.Fields("Cod_Fornecedor")), " ", Trim(cRec.Fields("Cod_Fornecedor")))
            oTela.slbl_12 = IIf(IsNull(cRec.Fields("Cod_Embalagem")), " ", Trim(cRec.Fields("Cod_Embalagem")))
            oTela.slbl_13 = Format(Now(), "DD/MM/YYYY")
            oTela.slbl_14 = " "
            oTela.slbl_15 = " "
            oTela.slbl_16 = IIf(IsNull(cRec.Fields("Lote")), " ", Trim(cRec.Fields("Lote")))
            oTela.ldl_usuario.Caption = IIf(IsNull(cRec.Fields("EMBALAGEM")), " ", cRec.Fields("EMBALAGEM"))
            oTela.lbl_Sequencial.Caption = Trim(cRec.Fields("ID_ETIQUETA"))
            oTela.lbl_barras1.Caption = "*" & oTela.lbl_Sequencial.Caption & "*"
            oTela.lbl_barras2.Caption = "*" & oTela.lbl_Sequencial.Caption & "*"
            oTela.Visible = False
         End If
    ''''    If cRec.Fields("Tipo") = "2" Then
    ''''        frmExibicao2.Show
    ''''        'Mostra FIAT
    ''''        frmExibicao2.lblCod_Peca.Caption = cRec.Fields("Cod_Peca")
    ''''        If cRec.Fields("Data_Expedicao") <> "" Then
    ''''            frmExibicao2.lblDataExpedicao2.Caption = cRec.Fields("Data_Expedicao")
    ''''        End If
    ''''        If cRec.Fields("Cod_Fornecedor") <> "" Then
    ''''            frmExibicao2.lblCodFornec2.Caption = Format(cRec.Fields("Cod_Fornecedor"), "000000")
    ''''        End If
    ''''        If cRec.Fields("Descr_Peca") <> "" Then
    ''''            frmExibicao2.lblDenominacao2.Caption = cRec.Fields("Descr_Peca")
    ''''        End If
    ''''        If cRec.Fields("Num_Doc_Fiscal") <> "" Then
    ''''            frmExibicao2.lblBam2.Caption = cRec.Fields("Num_Doc_Fiscal")
    ''''        End If
    ''''        If cRec.Fields("Cod_Cliente") <> "" Then
    ''''            frmExibicao2.lblDesenho2.Caption = Format(cRec.Fields("Cod_Cliente"), "00000000000")
    ''''        End If
    ''''
    ''''        frmExibicao2.lblCodBarra.Caption = Format(cRec.Fields("Cod_Cliente"), "00000000000") & Format(cRec.Fields("Qtd_Caixa"), "00000") _
    ''''                                      & cRec.Fields("Cod_Embalagem_pw") & Format(cRec.Fields("Cod_Fornecedor"), "000000")
    ''''        frmExibicao2.lblCodBarra2.Caption = frmExibicao2.lblCodBarra.Caption
    ''''        'Adicionar o * de inicio e fim
    ''''        frmExibicao2.lblCodBarra.Caption = "*" & frmExibicao2.lblCodBarra.Caption & "*"
    ''''        frmExibicao2.lblCodBarraCp1.Caption = frmExibicao2.lblCodBarra.Caption
    ''''        frmExibicao2.lblCodBarraCp2.Caption = frmExibicao2.lblCodBarra.Caption
    ''''
    ''''        If cRec.Fields("Data_Lote") <> "" Then
    ''''            frmExibicao2.lblDataProducao2.Caption = cRec.Fields("Data_Lote")
    ''''        End If
    ''''        If cRec.Fields("Cod_Embalagem") <> "" Then
    ''''            frmExibicao2.lblCodEmbalagem2.Caption = cRec.Fields("Cod_Embalagem")
    ''''        End If
    ''''        If cRec.Fields("Lote") <> "" Then
    ''''            frmExibicao2.lblNumLote2.Caption = cRec.Fields("Lote")
    ''''        End If
    ''''        If cRec.Fields("Qtd_Lote") <> "" Then
    ''''            frmExibicao2.lblQtdLote2.Caption = cRec.Fields("Qtd_Lote")
    ''''        End If
    ''''        If cRec.Fields("Qtd_Caixa") <> "" Then
    ''''            frmExibicao2.lblQtdEmbalagem2.Caption = Format(cRec.Fields("Qtd_Caixa"), "00000")
    ''''        End If
    ''''        If cRec.Fields("Classe_Funcional") <> "" Then
    ''''            frmExibicao2.lblClasseFuncional2.Caption = cRec.Fields("Classe_Funcional")
    ''''        End If
    ''''        If cRec.Fields("Vinculo") <> "" Then
    ''''            frmExibicao2.lblVinculo2.Caption = cRec.Fields("Vinculo")
    ''''        End If
    ''''        If cRec.Fields("Ind_Suplementar") <> "" Then
    ''''            frmExibicao2.lblIndicacaoSuplementar2.Caption = cRec.Fields("Ind_Suplementar")
    ''''        End If
    ''''        If cRec.Fields("Embarque_Controlado") <> "" Then
    ''''            frmExibicao2.lblEmbarqueControlado2.Caption = cRec.Fields("Embarque_Controlado")
    ''''        End If
    ''''        If cRec.Fields("Desvio") <> "" Then
    ''''            frmExibicao2.lblLoteSobDesvio2.Caption = cRec.Fields("Desvio")
    ''''        End If
    ''''        If cRec.Fields("DUM") <> "" Then
    ''''            frmExibicao2.lblDum2.Caption = cRec.Fields("DUM")
    ''''        End If
    ''''        If Not IsNull(cRec.Fields("EMBALAGEM")) Then
    ''''            frmExibicao2.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
    ''''        Else
    ''''            frmExibicao2.lblEmbalagem2.Caption = ""
    ''''        End If
    ''''Rem acrescentado o pto de entrega - 09-09-2016
    ''''        If cRec.Fields("Pto_Entrega") <> "" Then
    ''''            frmExibicao2.lblPontoEntrega2.Caption = Trim(cRec.Fields("Pto_Entrega"))
    ''''        End If
    ''''
    ''''        If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
    ''''            frmExibicao2.lblCodigoBarras.Caption = ""
    ''''            frmExibicao2.lblCodigoBarras.Caption = Trim(cRec.Fields("ID_ETIQUETA"))
    ''''            frmExibicao2.lblCodigoBarrasA.Caption = "*" & Trim(frmExibicao2.lblCodigoBarras.Caption) & "*"
    ''''            frmExibicao2.lblCodigoBarrasB.Caption = "*" & Trim(frmExibicao2.lblCodigoBarras.Caption) & "*"
    '''''            frmExibicao2.lblCodigoBarrasC.Caption = "*" & frmExibicao2.lblCodigoBarras.Caption & "*"
    '''''            frmExibicao2.lblCodigoBarrasD.Caption = "*" & frmExibicao2.lblCodigoBarras.Caption & "*"
    ''''         Else
    ''''            frmExibicao2.lblCodigoBarras.Caption = ""
    ''''            frmExibicao2.lblCodigoBarrasA.Caption = ""
    ''''            frmExibicao2.lblCodigoBarrasB.Caption = ""
    ''''         End If
    ''''
    ''''
    End If
        
    '---------------------------------------------------------------------------------------------------
    'Se tipo 3 opcao FORD
    If cRec.Fields("Tipo") = "3" Then
        sdata_aux = Mid$(Trim(cRec.Fields("data_etiq")), 1, 2) & _
                    Pega_Mes(Val(Mid$(Trim(cRec.Fields("data_etiq")), 4, 2))) & _
                    Mid$(Trim(cRec.Fields("data_etiq")), 7, 4)
        frmExibicao3.lbl_Cliente.Caption = "MUSASHI DO BRASIL LTDA" 'Trim(crec.Fields("Cliente"))
        frmExibicao3.lbl_Cod_Fornecedor.Caption = "CFEOA"  'Left(crec.Fields("Cod_Fornecedor"), 16)
        frmExibicao3.lbl_Cod_Fornecedor_Barras.Caption = "CFEOA" 'Trim(crec.Fields("Cod_Fornecedor"))
        frmExibicao3.lbl_qtd.Caption = Left(Trim(cRec.Fields("Qtd_Caixa")), 16)
        frmExibicao3.lbl_qtd_barras.Caption = Left(Trim(cRec.Fields("Qtd_Caixa")), 16)
        frmExibicao3.lbl_peso.Caption = Left(Trim(cRec.Fields("Peso")), 16)
        frmExibicao3.lbl_container.Caption = "KLT 4314 CFEOA" 'Left(Trim(crec.Fields("Lote"), 16))
        frmExibicao3.lbl_lote.Caption = Trim(cRec.Fields("Lote"))
        frmExibicao3.lbl_data.Caption = sdata_aux
        frmExibicao3.lbl_Cod_cliente.Caption = Left(Trim(cRec.Fields("Cod_Cliente")), 16)
        frmExibicao3.lbl_cod_cliente_1.Caption = Left(Trim(cRec.Fields("Cod_Cliente")), 16)
        frmExibicao3.lbl_cod_cliente_Barras.Caption = Left(Trim(cRec.Fields("Cod_Cliente")), 16)
        frmExibicao3.lbl_Cod_Peca.Caption = Trim(cRec.Fields("Cod_Peca"))
        frmExibicao3.lbl_line_feed_loc2.Caption = Trim(cRec.Fields("Cod_Peca")) & " " & Trim(cRec.Fields("Lote"))

        frmExibicao3.lbl_descr_peca.Caption = Trim(cRec.Fields("Descr_Peca"))
        frmExibicao3.lbl_id_etiqueta.Caption = Left(Trim(cRec.Fields("id_etiqueta")), 16)
        frmExibicao3.lbl_id_etiqueta_barra.Caption = Left(Trim(cRec.Fields("id_etiqueta")), 16)
        sdata_aux = Mid$(Trim(cRec.Fields("data_etiq")), 9, 2) & _
                    Mid$(Trim(cRec.Fields("data_etiq")), 4, 2) & _
                    Mid$(Trim(cRec.Fields("data_etiq")), 1, 2)
        frmExibicao3.DataToEncodeText.Text = Format(cRec.Fields("id_etiqueta"), "00000") & " (P)" & _
                                             Format(Trim(cRec.Fields("Qtd_Caixa")), "0") & " (Q)" & _
                                             Trim(cRec.Fields("Cod_Fornecedor")) & " (V)" & _
                                             sdata_aux & " (D)" & _
                                             Format(Trim(cRec.Fields("Serial")), "000") & " (S)"
        
        frmExibicao3.lbl_to.Caption = "FORD TAUBATE"
        frmExibicao3.lbl_cust.Caption = "FI05D"
        frmExibicao3.lbl_doc_code.Caption = "R3"
    
        frmExibicao3.PDF1.DataToEncode = frmExibicao3.DataToEncodeText.Text
        
''        frmExibicao3ant.Show
''        frmExibicao3ant.lblNumPeca.Caption = Left(cRec.Fields("Cod_Cliente"), 16)
''        frmExibicao3ant.lblNumPecaA.Caption = "*P" & Trim(Left(cRec.Fields("Cod_Cliente"), 16)) & "*"
''        frmExibicao3ant.lblNumPecaB.Caption = frmExibicao3ant.lblNumPecaA.Caption
''        frmExibicao3ant.lblNumPeca.Caption = Trim(frmExibicao3ant.lblNumPeca.Caption)
''        frmExibicao3ant.lblCod_Peca = cRec.Fields("Cod_Peca")
''        frmExibicao3ant.lblLote = cRec.Fields("Lote")
''        frmExibicao3ant.lblQtd.Caption = cRec.Fields("Qtd_Caixa")
''        frmExibicao3ant.lblQtdA.Caption = "*Q" & cRec.Fields("Qtd_Caixa") & "*"
''        frmExibicao3ant.lblQtdB.Caption = frmExibicao3ant.lblQtdA.Caption
''        If Not IsNull(cRec.Fields("Cod_Fornecedor")) Then
''            frmExibicao3ant.lblNumFornec.Caption = cRec.Fields("Cod_Fornecedor")
''            frmExibicao3ant.lblNumFornecA.Caption = "*V" & cRec.Fields("Cod_Fornecedor") & "*"
''            frmExibicao3ant.lblNumFornecB.Caption = frmExibicao3ant.lblNumFornecA.Caption
''        End If
''        If Not IsNull(cRec.Fields("Serial")) Then
''            frmExibicao3ant.lblNumSerial.Caption = cRec.Fields("Serial")
''            frmExibicao3ant.lblNumSerialA.Caption = "*S" & cRec.Fields("Serial") & "*"
''            frmExibicao3ant.lblNumSerialB.Caption = frmExibicao3ant.lblNumSerialA.Caption
''        End If
''        If Not IsNull(cRec.Fields("Cod_Util")) Then
''            frmExibicao3ant.lblCodUtil.Caption = cRec.Fields("Cod_Util")
''        End If
''        If Not IsNull(cRec.Fields("Linha_Util")) Then
''            frmExibicao3ant.lblLinhaUtil.Caption = cRec.Fields("Linha_Util")
''        End If
''        frmExibicao3ant.lblSufixo.Caption = Trim(Right((cRec.Fields("Cod_Cliente")), 5))
''        frmExibicao3ant.lblSufixoA.Caption = "*C" & frmExibicao3ant.lblSufixo.Caption & "*"
''        frmExibicao3ant.lblSufixoB.Caption = frmExibicao3ant.lblSufixoA.Caption
''        If (cRec.Fields("Desvio")) <> "" Then
''            frmExibicao3ant.lblDestino.Caption = Trim(cRec.Fields("Desvio"))
''        End If
''        frmExibicao3ant.lblDestinoA.Caption = "*D" & frmExibicao3ant.lblDestino.Caption & "*"
''        frmExibicao3ant.lblDestinoB.Caption = frmExibicao3ant.lblDestinoA.Caption
''
''        If Not IsNull(cRec.Fields("EMBALAGEM")) Then
''            frmExibicao3ant.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
''        End If
''
''        If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
''            frmExibicao3ant.lblCodigoBarras.Caption = Format(cRec.Fields("ID_ETIQUETA"), "0000000000")
''            frmExibicao3ant.lblCodigoBarrasA.Caption = "*" & frmExibicao3ant.lblCodigoBarras.Caption & "*"
''            frmExibicao3ant.lblCodigoBarrasB.Caption = "*" & frmExibicao3ant.lblCodigoBarras.Caption & "*"
''         Else
''            frmExibicao3ant.lblCodigoBarras.Caption = ""
''            frmExibicao3ant.lblCodigoBarrasA.Caption = ""
''            frmExibicao3ant.lblCodigoBarrasB.Caption = ""
''         End If
        
        
    End If
    
    '---------------------------------------------------------------------------------------------------
    'Se tipo 4 opcao padrão etiqueta grande
    If cRec.Fields("Tipo") = "4" Or cRec.Fields("Tipo") = "8" Then
        'frmExibicao4.Show
        frmAvulsoPadraoPonteiro.Show
        frmAvulsoPadraoPonteiro.Left = Me.Width
        If cRec.Fields("Cliente") <> "" Then
            'frmExibicao4.lblCliente.Caption = cRec.Fields("Cliente")
            frmAvulsoPadraoPonteiro.lblCliente.Caption = cRec.Fields("Cliente")
        End If
        If cRec.Fields("Cod_Cliente") <> "" Then
            'frmExibicao4.lblCodCliente2.Caption = cRec.Fields("Cod_Cliente")
            frmAvulsoPadraoPonteiro.lblCodCliente2.Caption = cRec.Fields("Cod_Cliente")
        End If
        If cRec.Fields("Descr_Peca") <> "" Then
            'frmExibicao4.lblDescricao.Caption = cRec.Fields("Descr_Peca")
            frmAvulsoPadraoPonteiro.lblDescricao.Caption = cRec.Fields("Descr_Peca")
        End If
        If cRec.Fields("Lote") <> "" Then
            'frmExibicao4.lblLote2.Caption = cRec.Fields("Lote")
            frmAvulsoPadraoPonteiro.lblLote2.Caption = cRec.Fields("Lote")
        End If
        If cRec.Fields("Peso") <> "" Then
            'frmExibicao4.lblPeso2.Caption = cRec.Fields("Peso")
            frmAvulsoPadraoPonteiro.lblPeso2.Caption = Format(cRec.Fields("Peso"), "0.00")
        End If
        If cRec.Fields("Qtd_Caixa") <> "" Then
            'frmExibicao4.lblQtd2.Caption = cRec.Fields("Qtd_Caixa")
            frmAvulsoPadraoPonteiro.lblQtd2.Caption = cRec.Fields("Qtd_Caixa")
        End If
        If cRec.Fields("Cod_Peca") <> "" Then
            'frmExibicao4.lblPeca.Caption = cRec.Fields("Cod_Peca")
            frmAvulsoPadraoPonteiro.lblPeca.Caption = cRec.Fields("Cod_Peca")
        End If
        If Not IsNull(cRec.Fields("EMBALAGEM")) Then
            frmAvulsoPadraoPonteiro.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
        End If
        frmAvulsoPadraoPonteiro.lbl_data.Caption = Format(cRec.Fields("data_etiq"), "dd/mm/yyyy")
        frmAvulsoPadraoPonteiro.lbl_MES.Caption = Pega_Mes(Mid$(Format(cRec.Fields("data_etiq"), "dd/mm/yyyy"), 4, 2))
        
        If (cRec.Fields("Tipo") = "4") Then
           frmAvulsoPadraoPonteiro.lblMsgProduto.Visible = False
           If (Val((cRec.Fields("id_cliente")) = 1) Or (Val(cRec.Fields("id_cliente")) = 2) Or (Val(cRec.Fields("id_cliente")) = 3)) Then
              frmAvulsoPadraoPonteiro.lblMsgProduto.Visible = True
           End If
        End If
        
        If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
           frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption = Format(cRec.Fields("ID_ETIQUETA"), "0000000000")
           frmAvulsoPadraoPonteiro.lblCodigoBarrasA.Caption = "*" & frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption & "*"
           frmAvulsoPadraoPonteiro.lblCodigoBarrasB.Caption = "*" & frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption & "*"
        Else
           frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption = ""
           frmAvulsoPadraoPonteiro.lblCodigoBarrasA.Caption = ""
           frmAvulsoPadraoPonteiro.lblCodigoBarrasB.Caption = ""
        End If
        
        If (cRec.Fields("Tipo") = "8") Then
           frmAvulsoPadraoPonteiro.lblCodCliente1.Caption = "CODE:"
           frmAvulsoPadraoPonteiro.lblQtd1.Caption = "QTY.:"
           frmAvulsoPadraoPonteiro.lblPeso1.Caption = "WEIGHT:"
           frmAvulsoPadraoPonteiro.lblLote1.Caption = "LOT.:"
        End If
         
    
    End If
    
    '---------------------------------------------------------------------------------------------------
    'Se tipo 5 opcao GM etiqueta grande
    'MODIFICADO EM 23-08-2007 (MARCOS PEDROSA) ACRESCENTAR CAMPOS LAY-OUT NOVO
    If cRec.Fields("Tipo") = "5" Then
        frmExibicao5.tWithOpcoes = Me.Width
        frmExibicao5.Show
        If cRec.Fields("Cliente") <> "" Then
            frmExibicao5.lblTo.Caption = cRec.Fields("Cliente")
        End If
        Rem incluido (marcos pedrosa) em 23-08-2007
        If cRec.Fields("MOTIVO_ALTERACAO_OUTROS") <> "" Then
            frmExibicao5.lblMotivo_alteracao_outros.Caption = Trim(cRec("MOTIVO_ALTERACAO_OUTROS").Value)
        End If
        If cRec.Fields("Ind_Suplementar") <> "" Then
            frmExibicao5.lblPlant.Caption = cRec.Fields("Ind_Suplementar")
        End If
        Rem DEFINICAO ATE 23-08-2007 (MARCOS PEDROSA)
'''        If cRec.Fields("Cod_Util") <> "" Then
'''            frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & cRec.Fields("Cod_Util")
'''        End If
'''        If cRec.Fields("Desvio") <> "" Then
'''            frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & cRec.Fields("Desvio")
'''        End If
        Rem INCLUIDO MARCOS PEDROSA EM 23-08-2007
        If cRec.Fields("Cod_Embalagem_Pw") <> "" Then
            frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & cRec.Fields("Cod_Embalagem_Pw")
        End If
        Rem INCLUIDO MARCOS PEDROSA EM 23-08-2007
        If cRec.Fields("Desvio") <> "" Then
            frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & cRec.Fields("Desvio")
        End If

        If cRec.Fields("Cod_Cliente") <> "" Then
            frmExibicao5.lblPartNumber.Caption = cRec.Fields("Cod_Cliente")
        End If
        If cRec.Fields("Descr_Peca") <> "" Then
            frmExibicao5.lblPeca.Caption = cRec.Fields("Descr_Peca")
        End If
        If cRec.Fields("Qtd_Caixa") <> 0 Then
            frmExibicao5.lblQtd.Caption = cRec.Fields("Qtd_Caixa")
        End If
        If cRec.Fields("Modelo") <> "" Then
            frmExibicao5.lblMaterial.Caption = cRec.Fields("Modelo")
        End If
        If cRec.Fields("Cod_Embalagem_Pw") <> "" Then
            frmExibicao5.lblReference.Caption = cRec.Fields("Cod_Embalagem_Pw")
        End If
'        If cRec.Fields("Pto_Entrega") <> "" Then
'            frmExibicao5.lblLicense.Caption = cRec.Fields("Pto_Entrega")
'            frmExibicao5.lblLicenseA.Caption = "*" & cRec.Fields("Pto_Entrega") & "*"
'            frmExibicao5.lblLicenseB.Caption = "*" & cRec.Fields("Pto_Entrega") & "*"
'        End If
        Rem comentado em 23-08-2007 (marcos pedrosa)
'''        If cRec.Fields("Cod_Embalagem") <> "" Then
'''            frmExibicao5.lblContainerType.Caption = cRec.Fields("Cod_Embalagem")
'''        End If
        Rem INCLUIDO EM 23-08-2007
        If cRec.Fields("compl_peca1") <> "" Then
            frmExibicao5.lblContainerType.Caption = cRec.Fields("compl_peca1")
        End If
        If cRec.Fields("Peso") <> 0 Then
            frmExibicao5.lblgrossWeight.Caption = Format(cRec.Fields("Peso"), "0.00")
        End If
        Rem comentado em 23-08-2007 (marcos pedrosa)
'''        If cRec.Fields("Embarque_Controlado") <> "" Then
'''            frmExibicao5.lblRoute.Caption = cRec.Fields("Embarque_Controlado")
'''        End If
        Rem INCLUIDO EM 23-08-2007
        If cRec.Fields("compl_peca2") <> "" Then
            frmExibicao5.lblRoute.Caption = cRec.Fields("compl_peca2")
        End If
        If cRec.Fields("Lote") <> "" Then
            frmExibicao5.lblLot.Caption = cRec.Fields("Lote")
        End If
        Rem comentado em 23-08-2007 (marcos pedrosa)
'''        If cRec.Fields("Dum") <> "" Then
'''            frmExibicao5.lblEng.Caption = cRec.Fields("Dum")
'''        End If
        Rem INCLUIDO EM 23-08-2007
        If cRec.Fields("data_lote") <> "" Then
            frmExibicao5.lblEng.Caption = Mid$(cRec.Fields("data_lote"), 1, 2) & _
                                          Pega_Mes(Val(Mid$(cRec.Fields("data_lote"), 4, 2))) & _
                                          Mid$(cRec.Fields("data_lote"), 7, 4)
        End If
        Rem INCLUIDO EM 23-08-2007
        If cRec.Fields("envio_lote") = "1" Then
            frmExibicao5.lblvalidade.Caption = "N"
        Else
            frmExibicao5.lblvalidade.Caption = ""
        End If
        
        If cRec.Fields("Data_expedicao") <> "" Then
            frmExibicao5.lblMfgDate.Caption = Mid$(cRec.Fields("Data_expedicao"), 1, 2) & _
                                              Pega_Mes(Val(Mid$(cRec.Fields("Data_expedicao"), 4, 2))) & _
                                              Mid$(cRec.Fields("Data_expedicao"), 7, 4)
        End If
        If cRec.Fields("Cod_Peca") <> "" Then
            frmExibicao5.lblCodMSB.Caption = cRec.Fields("Cod_Peca")
        End If
        If Not IsNull(cRec.Fields("EMBALAGEM")) Then
            frmExibicao5.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
        End If
        If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
            frmExibicao5.lblCodigoBarras.Caption = Format(cRec.Fields("ID_ETIQUETA"), "0000000000")
            frmExibicao5.lblCodigoBarrasA.Caption = "*" & frmExibicao5.lblCodigoBarras.Caption & "*"
            frmExibicao5.lblCodigoBarrasB.Caption = "*" & frmExibicao5.lblCodigoBarras.Caption & "*"
         Else
            frmExibicao5.lblCodigoBarras.Caption = ""
            frmExibicao5.lblCodigoBarrasA.Caption = ""
            frmExibicao5.lblCodigoBarrasB.Caption = ""
         End If
    End If
    
'-----------------------------------------------------------------------------------------------------
    'Se tipo 6 opcao Cliente MVM INTERNACIONAL MOTORES
    If cRec.Fields("Tipo") = "6" Or cRec.Fields("Tipo") = "9" Then
        frmExibicao9.Show
        If cRec.Fields("Tipo") = "9" Then
           frmExibicao9.PICT_CRISTAL.Visible = True
        Else
           frmExibicao9.PICT_CRISTAL.Visible = False
        End If
        frmExibicao9.Left = Me.Width
        If cRec.Fields("Cliente") <> "" Then
            frmExibicao9.lblCliente.Caption = cRec.Fields("Cliente")
        End If
        If cRec.Fields("Cod_Cliente") <> "" Then
            frmExibicao9.lblCodBar_Cod_cliente.Caption = "*" & Trim(cRec.Fields("Cod_Cliente")) & "*"
            frmExibicao9.lblCodBar_Cod_cliente1.Caption = "*" & Trim(cRec.Fields("Cod_Cliente")) & "*"
            frmExibicao9.lbl_Cod_cliente.Caption = Trim(cRec.Fields("Cod_Cliente"))
        End If
        If cRec.Fields("Descr_Peca") <> "" Then
            frmExibicao9.lbl_desc_peca.Caption = cRec.Fields("Descr_Peca")
        End If
        If cRec.Fields("cod_fornecedor") <> "" Then
            frmExibicao9.lblCod_Fornecedor.Caption = cRec.Fields("cod_fornecedor")
        End If

        frmExibicao9.lbl_Fornecedor.Caption = "MUSASHI"
        If cRec.Fields("data_expedicao") <> "" Then
            frmExibicao9.lbl_data_expedicao.Caption = cRec.Fields("data_expedicao")
        End If
        If cRec.Fields("qtd_caixa") <> "" Then
            frmExibicao9.lblCodBarQtd_caixa.Caption = "*" & cRec.Fields("qtd_caixa") & "*"
            frmExibicao9.lblqtd_caixa.Caption = cRec.Fields("qtd_caixa")
        End If
        If cRec.Fields("Lote") <> "" Then
            frmExibicao9.lblCodBar_lote.Caption = "*" & cRec.Fields("Lote") & "*"
            frmExibicao9.lblLote.Caption = cRec.Fields("Lote")
        End If
        If cRec.Fields("Cod_Peca") <> "" Then
            frmExibicao9.lbl_id_etiqueta.Caption = Format(cRec.Fields("ID_ETIQUETA"), "0000000000")
            frmExibicao9.lblCodBar_Cod_Peca.Caption = "*" & cRec.Fields("id_etiqueta") & "*"
            frmExibicao9.lbl_Cod_Peca.Caption = cRec.Fields("cod_peca")
        End If
        
        If Not IsNull(cRec.Fields("EMBALAGEM")) Then
            frmExibicao9.lblCodFunc.Caption = cRec.Fields("EMBALAGEM")
        Else
            frmExibicao9.lblCodFunc.Caption = ""
        End If
        
        Rem preparar o "id" da etiqueta com sua formatacao conforme documento fornecedor(6)+dtEtiqueta(6-aammdd)+seq. do dia
        sSeqId = " "
        
        If Not IsNull(cRec.Fields("desvio_aviso_mod")) Then
'            frmExibicao9.lblDesvio_Aviso_Mod.Caption = cRec.Fields("desvio_aviso_mod")
            sSeqId = Mid$(cRec.Fields("desvio_aviso_mod"), 1, 6) ' fornecedor
            sSeqId = Trim(sSeqId) & Mid$(cRec.Fields("desvio_aviso_mod"), 13, 2) ' (aa)mmdd
            sSeqId = Trim(sSeqId) & Mid$(cRec.Fields("desvio_aviso_mod"), 9, 2)  ' aa(mm)dd
            sSeqId = Trim(sSeqId) & Mid$(cRec.Fields("desvio_aviso_mod"), 7, 2) ' aamm(dd)
            sSeqId = Trim(sSeqId) & Format(Mid$(cRec.Fields("desvio_aviso_mod"), 15, 11), "000000") ' sequencial
            
            frmExibicao9.lblCodBar_Desvio_Aviso_Mod.Caption = "*" & sSeqId & "*"
            frmExibicao9.lblCodBar_Desvio_Aviso_Mod1.Caption = "*" & sSeqId & "*"
            frmExibicao9.lblDesvio_Aviso_Mod.Caption = sSeqId
            
        End If
         
    End If

'-----------------------------------------------------------------------------------------------------
    'Se tipo 7 opcao Palete GM
    If cRec.Fields("Tipo") = "7" Then
        If (IsNull(cRec.Fields("CODIGO_PRODUTO2").Value)) Then
            Set frmExibicao7Ref = frmExibicao7UmProduto
        ElseIf (Trim(cRec.Fields("CODIGO_PRODUTO2").Value) = "") Then
            Set frmExibicao7Ref = frmExibicao7UmProduto
        Else
            Set frmExibicao7Ref = frmExibicao7VariosProdutos
        End If
        
        
        frmExibicao7Ref.Show
        
        'Cliente
        If Not IsNull(cRec.Fields("Cliente").Value) Then
            frmExibicao7Ref.lblTo.Caption = cRec.Fields("Cliente").Value
        Else
            frmExibicao7Ref.lblTo.Caption = ""
        End If
        If Not IsNull(cRec.Fields("IND_SUPLEMentar").Value) Then
            frmExibicao7Ref.lblPlant.Caption = cRec.Fields("IND_SUPLEMentar").Value
        Else
            frmExibicao7Ref.lblPlant.Caption = ""
        End If
        Rem INCLUIDO MARCOS PEDROSA EM 28-08-2007
        If cRec.Fields("Cod_Embalagem_Pw") <> "" Then
            frmExibicao7Ref.lblPlant.Caption = frmExibicao7Ref.lblPlant.Caption & "-" & cRec.Fields("Cod_Embalagem_Pw")
        End If
        Rem INCLUIDO MARCOS PEDROSA EM 28-08-2007
        If cRec.Fields("Embarque_Controlado") <> "" Then
            frmExibicao7Ref.lblPlant.Caption = frmExibicao7Ref.lblPlant.Caption & "-" & cRec.Fields("Embarque_Controlado")
        End If
        Rem INCLUIDO MARCOS PEDROSA EM 28-08-2007
        If cRec.Fields("Modelo") <> "" Then
            frmExibicao7Ref.lblMaterial.Caption = cRec.Fields("Modelo")
        End If
        
        Rem INCLUIDO EM 28-08-2007
        If cRec.Fields("envio_lote") = "1" Then
            frmExibicao7Ref.lblvalidade.Caption = "N"
        Else
            frmExibicao7Ref.lblvalidade.Caption = ""
        End If
        
        'Pode ser gravado nulo no caso da etiqueta N: 7 (Palete GM)
        If Not IsNull(cRec.Fields("Pto_Entrega").Value) Then
            frmExibicao7Ref.lblLicense.Caption = cRec.Fields("Pto_Entrega").Value
        Else
            frmExibicao7Ref.lblLicense.Caption = ""
        End If
        frmExibicao7Ref.lblLicenseA.Caption = "*" & cRec.Fields("Pto_Entrega").Value & "*"
        frmExibicao7Ref.lblLicenseB.Caption = "*" & cRec.Fields("Pto_Entrega").Value & "*"
        frmExibicao7Ref.lblShipmentDate.Caption = UCase(Format(Date, "DD/MMM/YYYY"))
        frmExibicao7Ref.lblPeso.Caption = Format(cRec.Fields("Peso"), "0.00")
        
        frmExibicao7Ref.lblCodigoProduto1.Caption = cRec.Fields("CODIGO_PRODUTO1").Value
        If (IsNumeric(cRec.Fields("QTDE_CAIXA1").Value) And IsNumeric(cRec.Fields("PECAS_CAIXA1").Value)) Then
            frmExibicao7Ref.lblQtde1.Caption = cRec.Fields("QTDE_CAIXA1").Value & " X " & CStr(CInt(cRec.Fields("PECAS_CAIXA1").Value)) & " PC"
        Else
            frmExibicao7Ref.lblQtde1.Caption = ""
        End If
        If Not IsNull(cRec.Fields("COMPL_PECA1").Value) Then
            frmExibicao7Ref.lblComplPeca1.Caption = cRec.Fields("COMPL_PECA1").Value
        Else
            frmExibicao7Ref.lblComplPeca1.Caption = ""
        End If
        
        
        If frmExibicao7Ref.Name = "frmExibicao7UmProduto" Then
            If (IsNumeric(cRec.Fields("QTDE_CAIXA1").Value) And IsNumeric(cRec.Fields("PECAS_CAIXA1").Value)) Then
                If CInt(cRec.Fields("QTDE_CAIXA1").Value) <> 0 And CInt(cRec.Fields("PECAS_CAIXA1").Value) <> 0 Then
                    frmExibicao7UmProduto.lblQtdeTot1.Caption = cRec.Fields("QTDE_CAIXA1").Value * cRec.Fields("PECAS_CAIXA1").Value
                End If
            End If
        ElseIf frmExibicao7Ref.Name = "frmExibicao7VariosProdutos" Then
            frmExibicao7Ref.lblCodigoProduto2.Caption = cRec.Fields("CODIGO_PRODUTO2").Value
            If (IsNumeric(cRec.Fields("QTDE_CAIXA2").Value) And IsNumeric(cRec.Fields("PECAS_CAIXA2").Value)) Then
                If CInt(cRec.Fields("QTDE_CAIXA2").Value) <> 0 And CInt(cRec.Fields("PECAS_CAIXA2").Value) <> 0 Then
                    frmExibicao7Ref.lblQtde2.Caption = cRec.Fields("QTDE_CAIXA2").Value & " X " & CStr(CInt(cRec.Fields("PECAS_CAIXA2").Value)) & " PC"
                End If
            Else
                frmExibicao7Ref.lblQtde2.Caption = ""
            End If
            If Not IsNull(cRec.Fields("COMPL_PECA2").Value) Then
                frmExibicao7Ref.lblComplPeca2.Caption = cRec.Fields("COMPL_PECA2").Value
            Else
                frmExibicao7Ref.lblComplPeca2.Caption = ""
            End If
            
            If Not IsNull(cRec.Fields("CODIGO_PRODUTO3").Value) Then
                frmExibicao7Ref.lblCodigoProduto3.Caption = cRec.Fields("CODIGO_PRODUTO3").Value
            Else
                frmExibicao7Ref.lblCodigoProduto3.Caption = ""
            End If
            If (IsNumeric(cRec.Fields("QTDE_CAIXA3").Value) And IsNumeric(cRec.Fields("PECAS_CAIXA3").Value)) Then
                If CInt(cRec.Fields("QTDE_CAIXA3").Value) <> 0 And CInt(cRec.Fields("PECAS_CAIXA3").Value) <> 0 Then
                    frmExibicao7Ref.lblQtde3.Caption = cRec.Fields("QTDE_CAIXA3").Value & " X " & CStr(CInt(cRec.Fields("PECAS_CAIXA3").Value)) & " PC"
                End If
            Else
                frmExibicao7Ref.lblQtde3.Caption = ""
            End If
            If Not IsNull(cRec.Fields("COMPL_PECA3").Value) Then
                frmExibicao7Ref.lblComplPeca3.Caption = cRec.Fields("COMPL_PECA3").Value
            Else
                frmExibicao7Ref.lblComplPeca3.Caption = ""
            End If
            
            If Not IsNull(cRec.Fields("CODIGO_PRODUTO4").Value) Then
                frmExibicao7Ref.lblCodigoProduto4.Caption = cRec.Fields("CODIGO_PRODUTO4").Value
            Else
                frmExibicao7Ref.lblCodigoProduto4.Caption = ""
            End If
            If (IsNumeric(cRec.Fields("QTDE_CAIXA4").Value) And IsNumeric(cRec.Fields("PECAS_CAIXA4").Value)) Then
                If CInt(cRec.Fields("QTDE_CAIXA4").Value) <> 0 And CInt(cRec.Fields("PECAS_CAIXA4").Value) <> 0 Then
                    frmExibicao7Ref.lblQtde4.Caption = cRec.Fields("QTDE_CAIXA4").Value & " X " & CStr(CInt(cRec.Fields("PECAS_CAIXA4").Value)) & " PC"
                End If
            Else
                frmExibicao7Ref.lblQtde4.Caption = ""
            End If
            If Not IsNull(cRec.Fields("COMPL_PECA4").Value) Then
                frmExibicao7Ref.lblComplPeca4.Caption = cRec.Fields("COMPL_PECA4").Value
            Else
                frmExibicao7Ref.lblComplPeca4.Caption = ""
            End If
        End If
        
        qtdeConteiners = 0
        If (IsNumeric(cRec.Fields("QTDE_CAIXA1").Value)) Then
            qtdeConteiners = qtdeConteiners + CInt(cRec.Fields("QTDE_CAIXA1").Value)
        End If
        If (IsNumeric(cRec.Fields("QTDE_CAIXA2").Value)) Then
            qtdeConteiners = qtdeConteiners + CInt(cRec.Fields("QTDE_CAIXA2").Value)
        End If
        If (IsNumeric(cRec.Fields("QTDE_CAIXA3").Value)) Then
            qtdeConteiners = qtdeConteiners + CInt(cRec.Fields("QTDE_CAIXA3").Value)
        End If
        If (IsNumeric(cRec.Fields("QTDE_CAIXA4").Value)) Then
            qtdeConteiners = qtdeConteiners + CInt(cRec.Fields("QTDE_CAIXA4").Value)
        End If
        
        
        frmExibicao7Ref.lblQtdeContainers.Caption = CStr(qtdeConteiners)
        
    End If
    
    
    '-------------------------------------------------------------------
    'Se tipo A - Identificação de alterações do produto e processo
    If cRec.Fields("Tipo") = "A" Then
        
        frmExibicao6.Show
        
        With frmExibicao6
            .lblDesenho.Caption = Trim(cRec("COD_CLIENTE").Value & "")
            .lblDesvio.Caption = Trim(cRec("DESVIO_AVISO_MOD").Value & "")
            .lblData1.Caption = Format(Date, "dd/mm/yyyy")
            .lblData2.Caption = .lblData1.Caption
            .lblNotaFiscal1.Caption = Trim(cRec("NUM_DOC_FISCAL").Value & "")
            .lblNotaFiscal2.Caption = .lblNotaFiscal1.Caption
            
            .lblOptDefinitiva.Visible = False
            .lblOptProvisoria.Visible = False
            .lblOptLoteUnico.Visible = False
            
            Select Case Trim(cRec("TIPO_ALTERACAO").Value)
            Case "1"
                .lblOptDefinitiva.Visible = True
            Case "2"
                .lblOptProvisoria.Visible = True
            Case "3"
                .lblOptLoteUnico.Visible = True
            End Select
            
            .lblOptDesvio.Visible = False
            .lblOptMaterialSelecionado.Visible = False
            .lblOptOutros.Visible = False
            .lblOptReparoRetrabaho.Visible = False
            .lblOptProdutoNovo.Visible = False
            Select Case Trim(cRec("MOTIVO_ALTERACAO").Value)
            Case "1"
                frmExibicao6.lblOptDesvio.Visible = True
            Case "2"
                frmExibicao6.lblOptMaterialSelecionado.Visible = True
            Case "3"
                frmExibicao6.lblOptOutros.Visible = True
            Case "4"
                frmExibicao6.lblOptReparoRetrabaho.Visible = True
            Case "5"
                frmExibicao6.lblOptProdutoNovo.Visible = True
            End Select
            
            .lblMotivoAlteracaoOutros.Caption = Trim(cRec("MOTIVO_ALTERACAO_OUTROS").Value & "")
            
            .lblOptPrimeiroEnvio.Visible = False
            .lblOptLoteIntermediario.Visible = False
            .lblOptUltimoLote.Visible = False
            Select Case Trim(cRec("ENVIO_LOTE").Value)
            Case "1"
                .lblOptPrimeiroEnvio.Visible = True
            Case "2"
                .lblOptLoteIntermediario.Visible = True
            Case "3"
                .lblOptUltimoLote.Visible = True
            End Select
            
            .lblNumAm.Caption = Trim(cRec("NUM_AM").Value & "")
            
        End With
        
        nQuantidade = cRec("QTD_ETIQ").Value
        
    End If
    
    bTelaImp = False
    Me.cmd_Impressao.Enabled = True
    Me.SetFocus
    
    '-------------------------------------------------------------------
End Sub
Private Function Fechar_Form_Etiqueta()

Dim v As Integer

For v = 0 To (Forms.Count - 1)
    If Forms(v).Name = "frmExibicao1" Or Forms(v).Name = "frmExibicao4" Then
        Unload frmAvulsoPadraoPonteiro
        Exit For
    End If
    If Forms(v).Name = "frmExibicao2" Then
        Unload frmExibicao2
        Exit For
    End If
    If Forms(v).Name = "frmExibicao3" Then
        Unload frmExibicao3
        Exit For
    End If
    If Forms(v).Name = "frmExibicao5" Then
        Unload frmExibicao5
        Exit For
    End If
    If Forms(v).Name = "frmExibicao7UmProduto" Then
        Unload frmExibicao7UmProduto
        Exit For
    End If
    If Forms(v).Name = "frmExibicao7VariosProdutos" Then
        Unload frmExibicao7VariosProdutos
        Exit For
    End If
    If Forms(v).Name = "frmExibicao9" Then
        Unload frmExibicao9
        Exit For
    End If
    If Forms(v).Name = "frmExibicao10" Then
        Unload frmExibicao10
        Exit For
    End If
    If Forms(v).Name = "frmExibicao11" Then
        Unload frmExibicao11
        Exit For
    End If
Next

End Function

Private Function Limpar_campos()

Me.txtUserName.Text = ""
Me.txtPassword.Text = ""
Me.txtUserName.BackColor = &H8000000F
Me.txtPassword.BackColor = &H8000000F

End Function
Private Function Liberar_campos()

Me.txtUserName.Text = ""
Me.txtPassword.Text = ""
Me.txtUserName.BackColor = &H80000005
Me.txtPassword.BackColor = &H80000005

End Function

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
