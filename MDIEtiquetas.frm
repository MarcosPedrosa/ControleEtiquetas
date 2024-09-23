VERSION 5.00
Begin VB.MDIForm MDIEtiquetas 
   BackColor       =   &H8000000C&
   Caption         =   "Emissão de Etiquetas - Versão 19/09/2024"
   ClientHeight    =   6000
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11055
   Icon            =   "MDIEtiquetas.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuetiqueta 
      Caption         =   "&Etiqueta"
      Visible         =   0   'False
      Begin VB.Menu mnuPrevia 
         Caption         =   "Prévia de impressão"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "I&mprimir"
      End
   End
   Begin VB.Menu mnuExclusao 
      Caption         =   "&Excluir registros"
   End
   Begin VB.Menu mnuAvulsa 
      Caption         =   "&Impressão Avulsa"
      Begin VB.Menu mnuPadrãoSemConexao 
         Caption         =   "Padrão Honda - Sem conexão com R/3"
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPadrao 
         Caption         =   "Padrão - Honda"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFiat 
         Caption         =   "Fiat"
      End
      Begin VB.Menu mnuFord 
         Caption         =   "Ford"
      End
      Begin VB.Menu mnuGm 
         Caption         =   "GM"
      End
      Begin VB.Menu MnuIdentAltFiat 
         Caption         =   "Identificação de alteração Fiat"
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPaleteGmUmProduto 
         Caption         =   "&Palete GM (um produto)"
      End
      Begin VB.Menu mnuPaleteGmVariosProdutos 
         Caption         =   "&Palete GM (vários produtos)"
      End
      Begin VB.Menu mnuTelaTeste 
         Caption         =   "Imprime Tela Teste"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuFerramentas 
      Caption         =   "&Ferramentas"
      Begin VB.Menu mnuImportacaoProduto 
         Caption         =   "&Importação dos produtos"
      End
      Begin VB.Menu mnuReimpressaoEtiquetas 
         Caption         =   "&Re-impressão das etiquetas"
      End
      Begin VB.Menu MnuCadastroUsuario 
         Caption         =   "&Cadastro de usuários"
      End
      Begin VB.Menu MnuExluirEtiqueta 
         Caption         =   "Exclusão de Etiquetas"
      End
      Begin VB.Menu MnuDesmembra 
         Caption         =   "Desmembrar uma etiqueta"
      End
      Begin VB.Menu MnuEtiqProduto 
         Caption         =   "Emissão etiquetas produto"
      End
      Begin VB.Menu MnuEtiqAjuste 
         Caption         =   "Emitir etiquetas de Ajuste"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuEtiqAjusteEmb 
         Caption         =   "Re-emitir etiqueta de embarque"
         Visible         =   0   'False
      End
      Begin VB.Menu MnuEtiqAjusteEmbAju 
         Caption         =   "Re-emitir etiqueta de embarque ajuste"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuetiq 
      Caption         =   "Eti&queta"
      Index           =   0
      Begin VB.Menu mnuEtiquetas 
         Caption         =   "&Transpotadora..."
         Index           =   1
      End
      Begin VB.Menu mnuEtiquetas 
         Caption         =   "1-International Motores..."
         Index           =   2
      End
      Begin VB.Menu mnuEtiquetas 
         Caption         =   "2-Etiqueta Honda Manaus..."
         Index           =   3
      End
      Begin VB.Menu mnuEtiquetas 
         Caption         =   "3-Etiquetas FIAT..."
         Index           =   4
      End
      Begin VB.Menu mnuEtiquetas 
         Caption         =   "4-Etiquetas KRUPP..."
         Index           =   5
      End
      Begin VB.Menu mnuEtiquetas 
         Caption         =   "5-Etiquetas Fiat Latam..."
         Index           =   6
      End
      Begin VB.Menu mnuEtiquetas 
         Caption         =   "Etiqueta Ford..."
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEtiquetas 
         Caption         =   "6-Etiqueta Yamaha..."
         Index           =   8
      End
      Begin VB.Menu mnuEtiquetas 
         Caption         =   "7-Etiqueta Fiat Itália..."
         Index           =   9
      End
      Begin VB.Menu mnuEtiquetas 
         Caption         =   "8-Etiquetas Honda (NOVA)"
         Index           =   10
         Begin VB.Menu mnuetiquetaHondaPallet 
            Caption         =   "1-Etiquetas Honda por N.Fiscal + Pallet..."
            Index           =   0
         End
         Begin VB.Menu mnuetiquetaHondaPallet 
            Caption         =   "2-Etiquetas Honda Por NFE..."
            Index           =   1
         End
      End
      Begin VB.Menu mnuEtiquetas 
         Caption         =   "9-Etiquetas GM (NOVA)..."
         Index           =   11
      End
      Begin VB.Menu mnuEtiquetas 
         Caption         =   "10-Etiquetas Inmetro"
         Index           =   12
         Begin VB.Menu mnuEtiquetasInmetro 
            Caption         =   "Emissão de etiquetas..."
            Index           =   0
         End
         Begin VB.Menu mnuEtiquetasInmetro 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuEtiquetasInmetro 
            Caption         =   "Cadastro de Peças..."
            Index           =   2
         End
         Begin VB.Menu mnuEtiquetasInmetro 
            Caption         =   "Cadastro de Modelo..."
            Index           =   3
         End
         Begin VB.Menu mnuEtiquetasInmetro 
            Caption         =   "Cadastro de Clientes"
            Index           =   4
         End
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "MDIEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim gEtiqueta As Arquivo
Dim gNumeroArquivo As Integer
Dim gTamanhoRegistro As Long
Dim gRegistroAtual As Long
Dim gUltimoRegistro As Long
Dim proximaEtqLivreAR As Double

Public forcaSaida As Boolean

Dim objEtiqueta As Etiqueta

Private Sub MDIForm_Activate()

If App.PrevInstance Then
   MsgBox "Este Programa JÁ esta sendo processado neste computador", 16, "<ENTER>=Para Finalizar"
   Close: End
End If

End Sub

Private Sub MDIForm_Load()
Dim sArq As String
Dim sArquivo As String


    On Error GoTo Erro
     
    Set objApplication = New Application
    Rem receber os parametros do arquivo \param.txt
    objApplication.setTipoBanco = bdAccess
    
    'Calcula o tamanho do registro do showa
    gTamanhoRegistro41 = Len(Arq_Mov_Showa)
 
    
    'Obtém o próximo número de arquivo disponível
    gNumeroArquivo41 = Len(Arq_Mov_Showa)
    
    Rem verificar a existência do arquivo na mda caso exista.
    sArquivo = objApplication.caminhoImportacao & "\etiqshowa.txt"
    sArq = VBA.Dir(sArquivo)
    If Len(Trim(sArq)) > 0 Then 'caso exista a nota na origem continua
       'Abre o arquivo para acesso aleatório. Se não
       Open objApplication.caminhoImportacao & "\etiqshowa.txt" For Random As gNumeroArquivo41 Len = gTamanhoRegistro41
       
       'Descobre qual é o último número de registro do arquivo
       gUltimoRegistro41 = FileLen(objApplication.caminhoImportacao & "\etiqshowa.txt") / gTamanhoRegistro41
       
       If gUltimoRegistro41 > 0 Then ' existe o arquivo texto contendo dados para emissao da etiqueta
          frmExibicao41.Show
          Exit Sub
       End If
    
    End If
    
    objApplication.AbrirConexao
    Set objEtiquetaControlador = New EtiquetaControlador
    Set objPecaAvulsoControlador = New PecaAvulsoControlador
    Set objTransporteControlador = New Mov_Transporte
    
    Rem aqui marcos 19/10/2012, sempre pegar o banco do sql server das duas empresas
    If objApplication.filial = adMusashiDaAmazonia Then
        objApplication.setTipoBancoInterface = bdAccess
        objEtiquetaControlador.setTipoBanco = bdAccess
    Else
' retirar aqui marcos
        objApplication.setTipoBancoInterface = bdSqlServer
        objEtiquetaControlador.setTipoBanco = bdSqlServer
    End If

    objPecaAvulsoControlador.setTipoBanco = bdAccess
' retirar aqui marcos
    objApplication.AbrirConexaoInterface
    Set objEtiquetaControlador.setConnection = objApplication.CnnInterface
    Set objPecaAvulsoControlador.setConnection = objApplication.cnn
    
    Rem VER AQUI MARCOS PEDROSA 13/03/2007
    If objApplication.filial = adMusashiDaAmazonia Then
        Set frmAvulsoPadraoPonteiro = frmExibicao1
    Else
        Rem AQUI MARCOS, MUDAR O NOME DO FORMULARIO PARA O Set frmAvulsoPadraoPonteiro = frmExibicao4AJU
        Set frmAvulsoPadraoPonteiro = frmExibicao4
    End If
    
    dteEtiquetas.Connections("ConEtiquetas").connectionString = objApplication.connectionString
    
  '  MsgBox "vai abrir mdb do arquivo - " & dteEtiquetas.Connections("ConEtiquetas").connectionString
    dteEtiquetas.rsEtiquetas.Open
 '   MsgBox "Abriu mdb do arquivo vb "
    
    Rem Foi incluida a rotina de delecao dos registros que nao foram impressos no access
    Rem esta rotina assegura que não havera registros no access
    Rem incluida em 01/02/2007 (Marcos Pedrosa)
    If dteEtiquetas.rsEtiquetas.RecordCount > 0 Then
       While Not dteEtiquetas.rsEtiquetas.EOF
           dteEtiquetas.rsEtiquetas.Delete
           dteEtiquetas.rsEtiquetas.MoveNext
       Wend
    End If
    
    'Calcula o tamanho do registro
     gTamanhoRegistro = Len(gEtiqueta) + 2
    
    'Obtém o próximo número de arquivo disponível
    gNumeroArquivo = FreeFile
    
    Open objApplication.caminhoImportacao & "\etiq.txt" For Random As gNumeroArquivo Len = gTamanhoRegistro
    gUltimoRegistro = FileLen(objApplication.caminhoImportacao & "\etiq.txt") / gTamanhoRegistro

Rem parar aqui marcos, continua em AdicionaRegistro
    If gUltimoRegistro = 0 Then ' existe o arquivo texto contendo dados para emissao da etiqueta
'''''        Close gNumeroArquivo
'''''        Kill objApplication.caminhoImportacao & "\etiq.txt"
    Else
        Rem aqui leitura do arquivo etiq.txt
        AdicionaRegistro
    End If
    
    If dteEtiquetas.rsEtiquetas.RecordCount = 0 Then
        If adMusashiDoBrasil Then
           If MsgBox("Nenhum arquivo encontrado para importação! Deseja imprimir avulso ?", vbQuestion + vbYesNo, "ATENÇÃO !!!") = vbNo Then
               dteEtiquetas.rsEtiquetas.Close
               Set objApplication = Nothing
               End
           End If
        End If
    End If
    
    'LimpaEtiquetas
    
    Close gNumeroArquivo
    If dteEtiquetas.rsEtiquetas.RecordCount <> 0 Then
        Rem aqui prepara para saber o tipo do form a ser apresentado na tela
        frmOpcoes.Show
    End If
    
    Exit Sub
Erro:
    MsgErro "Ocorreu um erro ao carregar o programa."
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    If forcaSaida = False Then
        If MsgBox("Deseja encerrar o programa ?", vbQuestion + vbYesNo, "ATENÇÃO!!!") = vbYes Then
            Cancel = False
        Else
            Cancel = True
        End If
    End If
    
End Sub

Private Sub MnuCadastroUsuario_Click()
frmUsuarios.Show
End Sub

Private Sub MnuDesmembra_Click()
frmEtiquetaDesmembra.Show
End Sub

Private Sub MnuEtiqAjuste_Click()
    frmEtiquetaAjuste.Show
End Sub

Private Sub MnuEtiqAjusteEmb_Click()
    frmEtiquetaAjusteEmb.Show
End Sub

Private Sub MnuEtiqAjusteEmbAju_Click()
    frmEtiquetaAjusteEmbAju.Show
End Sub

Private Sub MnuEtiqProduto_Click()
'frmEtiquetaEmissaoProduto.Show
End Sub

Private Sub mnuetiquetaHondaPallet_Click(Index As Integer)

If Index = 0 Then
  frmEtiquetaHondaQrcodePallet.Show
ElseIf Index = 1 Then
  frmEtiquetaHondaQrcode.Show
End If

End Sub

Private Sub MnuEtiquetas_Click(Index As Integer)
If Index = 1 Then
   frmExibicao8Transporte.Show
ElseIf Index = 2 Then
  frmExibicaoAvulsaInternational.Show
ElseIf Index = 3 Then
  frmEtiquetaHondaManaus.Show
ElseIf Index = 4 Then
  frmEtiquetaFiat.Show
ElseIf Index = 5 Then
  frmEtiquetaKRUPP.Show
ElseIf Index = 6 Then
  frmEtiquetaFiatLatamFCA.Show
ElseIf Index = 7 Then
  frmEtiquetaFord.Show
ElseIf Index = 8 Then
  frmExibicao10.Show
ElseIf Index = 9 Then
  frmExibicao2.Show
'ElseIf Index = 10 Then
'  frmEtiquetaHondaQrcodePallet.Show
'  frmEtiquetaHondaQrcode.Show
'    frmExibicaoHONDA.Show
ElseIf Index = 11 Then
  frmEtiquetaGMMaster.Show
End If

End Sub

Private Sub mnuEtiquetasInmetro_Click(Index As Integer)
If Index = 0 Then
   frmInmetroEmissaoEtiquetas.Show
ElseIf Index = 2 Then
  frmInmetroCadastroPecas.Show
ElseIf Index = 3 Then
  frmInmetroCadastroModelo.Show
ElseIf Index = 4 Then
  frmInmetroCadastroCliente.Show
End If

End Sub

Private Sub mnuExclusao_Click()
    
    If MsgBox("Tem certeza que deseja excluir todos os registros que ainda não foram impressos ?", vbYesNo + vbQuestion, "ATENÇÃO !!!") = vbNo Then
        Exit Sub
    Else
        If dteEtiquetas.rsEtiquetas.RecordCount > 0 Then
            dteEtiquetas.rsEtiquetas.MoveFirst
            
            While Not dteEtiquetas.rsEtiquetas.EOF
                dteEtiquetas.rsEtiquetas.Delete
                dteEtiquetas.rsEtiquetas.MoveNext
            Wend
        End If
    End If
    
    MsgBox "Não há mais registros a imprimir. O aplicativo será encerrado!", vbOKOnly + vbInformation, "INFORMAÇÃO !!!"
    forcaSaida = True
    Unload Me
    
End Sub

Private Sub MnuExluirEtiqueta_Click()
frmEtiquetaExcluir.Show
End Sub

Private Sub mnuFiat_Click()
    frmFiat.Show
End Sub

Private Sub mnuFord_Click()
    frmFord.Show
End Sub

Private Sub mnuGm_Click()
    frmGm.Show
End Sub

Private Sub MnuIdentAltFiat_Click()
    frmIdentAlteracao.Show
End Sub

Private Sub mnuImportacaoProduto_Click()
    If MsgBox("Deseja executar a importação dos produtos?" & vbNewLine & "Esta operação pode levar vários minutos.", vbInformation + vbYesNo, "Importar os produtos?") = vbYes Then
        MousePointer = vbHourglass
        If objPecaAvulsoControlador.importaProdutosXml() = True Then
            MsgBox "Importação concluída.", vbInformation, objApplication.tituloPrograma
        Else
            MsgBox "A importação não foi realizada.", vbExclamation, objApplication.tituloPrograma
        End If
        MousePointer = vbNormal
    End If
End Sub

Private Sub mnuPadrao_Click()
    frmPadrao.Show
End Sub

Private Sub mnuPadrãoSemConexao_Click()
    frmEtiquetaAvulsoPadrao.Show
End Sub

Private Sub mnuPaleteGmUmProduto_Click()
    Set frmExibicao7Ref = frmExibicao7UmProduto
    frmPaleteGm.Show
    
    frmPaleteGm.txtCodigoProduto2.Enabled = False
    frmPaleteGm.txtCodigoProduto2.BackColor = vbButtonFace
    frmPaleteGm.txtQtdCaixas2.Enabled = False
    frmPaleteGm.txtQtdCaixas2.BackColor = vbButtonFace
    frmPaleteGm.txtPecasPorCaixa2.Enabled = False
    frmPaleteGm.txtPecasPorCaixa2.BackColor = vbButtonFace
    frmPaleteGm.txtComplPeca2.Enabled = False
    frmPaleteGm.txtComplPeca2.BackColor = vbButtonFace
    
    frmPaleteGm.txtCodigoProduto3.Enabled = False
    frmPaleteGm.txtCodigoProduto3.BackColor = vbButtonFace
    frmPaleteGm.txtQtdCaixas3.Enabled = False
    frmPaleteGm.txtQtdCaixas3.BackColor = vbButtonFace
    frmPaleteGm.txtPecasPorCaixa3.Enabled = False
    frmPaleteGm.txtPecasPorCaixa3.BackColor = vbButtonFace
    frmPaleteGm.txtComplPeca3.Enabled = False
    frmPaleteGm.txtComplPeca3.BackColor = vbButtonFace
    
    frmPaleteGm.txtCodigoProduto4.Enabled = False
    frmPaleteGm.txtCodigoProduto4.BackColor = vbButtonFace
    frmPaleteGm.txtQtdCaixas4.Enabled = False
    frmPaleteGm.txtQtdCaixas4.BackColor = vbButtonFace
    frmPaleteGm.txtPecasPorCaixa4.Enabled = False
    frmPaleteGm.txtPecasPorCaixa4.BackColor = vbButtonFace
    frmPaleteGm.txtComplPeca4.Enabled = False
    frmPaleteGm.txtComplPeca4.BackColor = vbButtonFace
    
End Sub

Private Sub mnuPaleteGmVariosProdutos_Click()
    Set frmExibicao7Ref = frmExibicao7VariosProdutos
    frmPaleteGm.Show
End Sub

Private Sub mnuReimpressaoEtiquetas_Click()
    frmEtiquetaReimprime.Show
End Sub

Private Sub mnuSair_Click()
    Unload Me
End Sub

Public Sub AdicionaRegistro()
    Dim oTela As frmTelaFormacaoPDF417
    Dim Data As Date
    Dim contEtiqueta As Integer
    
    On Error GoTo Erro
    Rem aqui marcos recebe os dados do aquivo texto e copia para o banco access
    
    Rem MARCOS PEDROSA CRIADO PARA CONTROLAR A SEQUENCIA DAS ETIQUETAS
    proximaEtqLivreAR = 0: nSequencia_Yamaha = 0
    
    For gRegistroAtual = 1 To gUltimoRegistro
        Rem AQUI VER CAMPOS A ACRESCENTAR. 23-08-2007 GM
        Get #gNumeroArquivo, gRegistroAtual, gEtiqueta
        '''For contEtiqueta = 1 To CInt(Trim(gEtiqueta.Qtd_Etiq)) Step 1
            
            dteEtiquetas.rsEtiquetas.AddNew
            
            If Trim(gEtiqueta.Cod_Peca) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Cod_Peca") = Trim(gEtiqueta.Cod_Peca)
            End If
            If Trim(gEtiqueta.Lote) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Lote") = Trim(gEtiqueta.Lote)
            End If
            If Trim(gEtiqueta.Qtd_Caixa) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa") = Trim(gEtiqueta.Qtd_Caixa)
            End If
            If (gEtiqueta.Dia <> "  ") And (gEtiqueta.Mes <> "  ") And (gEtiqueta.Ano <> "    ") Then
                Data = gEtiqueta.Dia & "/" & gEtiqueta.Mes & "/" & gEtiqueta.Ano
                dteEtiquetas.rsEtiquetas.Fields("Data_Etiq") = Data
            End If
            If Trim(gEtiqueta.Cliente) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Cliente") = Trim(gEtiqueta.Cliente)
            Else
                MsgBox "Codigo do Cliente sem preenchimento, não sera Importada, Favor avisar ao depto de TI Urgente!"
                GoTo Erro
            End If
            
            If Trim(gEtiqueta.Cod_Cliente) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente") = (gEtiqueta.Cod_Cliente)
            End If
            
            If Trim(gEtiqueta.Peso) <> Empty Then
                If IsNumeric(Trim(gEtiqueta.Peso)) Then
                    dteEtiquetas.rsEtiquetas.Fields("Peso") = Round(CDbl(Replace(Trim(gEtiqueta.Peso), ".", ",")), 2)
                End If
            End If
            
Rem marcos ****** incluido 14/04/2006
            If Trim(gEtiqueta.Tipo_caixa) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("tipo_caixa") = (gEtiqueta.Tipo_caixa)
            Else
                dteEtiquetas.rsEtiquetas.Fields("tipo_caixa") = "1"
            End If
            Rem aqui ,tipo_caixa
Rem ****** incluido

            If Trim(gEtiqueta.Descr_Peca) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Descr_Peca") = Trim(gEtiqueta.Descr_Peca)
            End If
            
            'Como na versão antiga (ERRADA)
            If Trim(gEtiqueta.Qtd_Etiq) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Qtd_Etiq") = Trim(gEtiqueta.Qtd_Etiq)
                '''dteEtiquetas.rsEtiquetas.Fields("Qtd_Etiq") = "1" 'A qtde está expressa no for
            End If
            If Trim(gEtiqueta.Tipo) <> Empty Then
                Rem aqui marcos retirar
                dteEtiquetas.rsEtiquetas.Fields("Tipo") = Trim(gEtiqueta.Tipo)
'                dteEtiquetas.rsEtiquetas.Fields("Tipo") = "S"
            End If
            If (gEtiqueta.Dia2 <> "  ") And (gEtiqueta.Mes2 <> "  ") And (gEtiqueta.Ano2 <> "    ") Then
                Data = gEtiqueta.Dia2 & "/" & gEtiqueta.Mes2 & "/" & gEtiqueta.Ano2
                dteEtiquetas.rsEtiquetas.Fields("Data_Expedicao") = Data
            End If
            If Trim(gEtiqueta.Cod_Embalagem_Pw) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Cod_Embalagem_Pw") = Trim(gEtiqueta.Cod_Embalagem_Pw)
            End If
            If (gEtiqueta.Dia3 <> "  ") And (gEtiqueta.Mes3 <> "  ") And (gEtiqueta.Ano3 <> "    ") Then
                If gEtiqueta.Ano3 <> "0000" Then
                   Data = gEtiqueta.Dia3 & "/" & gEtiqueta.Mes3 & "/" & gEtiqueta.Ano3
                   dteEtiquetas.rsEtiquetas.Fields("Data_Lote") = Data
                End If
                
            End If
            If Trim(gEtiqueta.Classe_Funcional) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Classe_Funcional") = Trim(gEtiqueta.Classe_Funcional)
            End If
            If Trim(gEtiqueta.Embarque_Controlado) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Embarque_Controlado") = Trim(gEtiqueta.Embarque_Controlado)
            End If
            If Trim(gEtiqueta.Cod_Fornec) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor") = Trim(gEtiqueta.Cod_Fornec)
            End If
            If Trim(gEtiqueta.Num_Doc_Fiscal) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Num_Doc_Fiscal") = Trim(gEtiqueta.Num_Doc_Fiscal)
            End If
            If Trim(gEtiqueta.Cod_Embalagem) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Cod_Embalagem") = Trim(gEtiqueta.Cod_Embalagem)
            End If
            If Trim(gEtiqueta.Ind_Suplementar) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Ind_Suplementar") = Trim(gEtiqueta.Ind_Suplementar)
            End If
            If Trim(gEtiqueta.Pto_Entrega) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega") = Trim(gEtiqueta.Pto_Entrega)
            End If
            If (gEtiqueta.Dia4 <> "  " And gEtiqueta.Dia4 <> "00") And (gEtiqueta.Mes4 <> "  " And gEtiqueta.Mes4 <> "00") And (gEtiqueta.Ano4 <> "    " And gEtiqueta.Ano4 <> "0000") Then
                Data = Mid$(gEtiqueta.Ano4, 3, 2) & "/" & Mid$(gEtiqueta.Ano4, 1, 2) & "/" & gEtiqueta.Dia4 & gEtiqueta.Mes4
                dteEtiquetas.rsEtiquetas.Fields("DUM") = Data
            End If
            If Trim(gEtiqueta.Qtd_Lote) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Qtd_Lote") = Trim(gEtiqueta.Qtd_Lote)
            End If
            If Trim(gEtiqueta.Vinculo) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Vinculo") = Trim(gEtiqueta.Vinculo)
            End If
            If Trim(gEtiqueta.desvio) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Desvio") = Trim(gEtiqueta.desvio)
            End If
            If Trim(gEtiqueta.Serial) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Serial") = Trim(gEtiqueta.Serial)
            End If
            If Trim(gEtiqueta.Cod_Util) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Cod_Util") = Trim(gEtiqueta.Cod_Util)
            End If
            If Trim(gEtiqueta.Linha_Util) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Linha_Util") = Trim(gEtiqueta.Linha_Util)
            End If
            If Trim(gEtiqueta.Modelo) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("Modelo") = Trim(gEtiqueta.Modelo)
            End If
            If Trim(gEtiqueta.desvio_aviso_mod) <> "" Then
                dteEtiquetas.rsEtiquetas.Fields("DESVIO_AVISO_MOD") = Trim(gEtiqueta.desvio_aviso_mod)
            End If
            If Trim(gEtiqueta.tipo_alteracao) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("TIPO_ALTERACAO") = Trim(gEtiqueta.tipo_alteracao)
            End If
            If Trim(gEtiqueta.motivo_alteracao) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("MOTIVO_ALTERACAO") = Trim(gEtiqueta.motivo_alteracao)
            End If
            If Trim(gEtiqueta.motivo_alteracao_outros) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("MOTIVO_ALTERACAO_OUTROS") = Trim(gEtiqueta.motivo_alteracao_outros)
            End If
            If Trim(gEtiqueta.envio_lote) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("ENVIO_LOTE") = Trim(gEtiqueta.envio_lote)
            End If
            If Trim(gEtiqueta.num_am) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("NUM_AM") = Trim(gEtiqueta.num_am)
            End If
            
            If Trim(gEtiqueta.Embalagem) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM") = Trim(gEtiqueta.Embalagem)
            End If
            
            If Trim(Trim(gEtiqueta.id_cliente)) <> Empty Then
                dteEtiquetas.rsEtiquetas.Fields("ID_CLIENTE").Value = Trim(gEtiqueta.id_cliente)
            End If
            
            dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value = Trim(gEtiqueta.codigo_produto1)
            dteEtiquetas.rsEtiquetas.Fields("QTDE_CAIXA1").Value = Trim(gEtiqueta.qtde_caixa1)
            dteEtiquetas.rsEtiquetas.Fields("PECAS_CAIXA1").Value = Trim(gEtiqueta.pecas_caixa1)
            dteEtiquetas.rsEtiquetas.Fields("COMPL_PECA1").Value = Trim(gEtiqueta.compl_peca1)
            
            dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO2").Value = Trim(gEtiqueta.codigo_produto2)
            dteEtiquetas.rsEtiquetas.Fields("QTDE_CAIXA2").Value = Trim(gEtiqueta.qtde_caixa2)
            dteEtiquetas.rsEtiquetas.Fields("PECAS_CAIXA2").Value = Trim(gEtiqueta.pecas_caixa2)
            dteEtiquetas.rsEtiquetas.Fields("COMPL_PECA2").Value = Trim(gEtiqueta.compl_peca2)
            
            dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO3").Value = Trim(gEtiqueta.codigo_produto3)
            dteEtiquetas.rsEtiquetas.Fields("QTDE_CAIXA3").Value = Trim(gEtiqueta.qtde_caixa3)
            dteEtiquetas.rsEtiquetas.Fields("PECAS_CAIXA3").Value = Trim(gEtiqueta.pecas_caixa3)
            dteEtiquetas.rsEtiquetas.Fields("COMPL_PECA3").Value = Trim(gEtiqueta.compl_peca3)
            
            dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO4").Value = Trim(gEtiqueta.codigo_produto4)
            dteEtiquetas.rsEtiquetas.Fields("QTDE_CAIXA4").Value = Trim(gEtiqueta.qtde_caixa4)
            dteEtiquetas.rsEtiquetas.Fields("PECAS_CAIXA4").Value = Trim(gEtiqueta.pecas_caixa4)
            dteEtiquetas.rsEtiquetas.Fields("COMPL_PECA4").Value = Trim(gEtiqueta.compl_peca4)
            
Rem marcos ****** incluido 15/05/2006 novos campos - AJUSTE EM 13/06/2016

            dteEtiquetas.rsEtiquetas.Fields("remessa").Value = " "
            dteEtiquetas.rsEtiquetas.Fields("fatur").Value = " "
            
            dteEtiquetas.rsEtiquetas.Fields("xblnr").Value = " "
            dteEtiquetas.rsEtiquetas.Fields("posnr").Value = 0
            
            dteEtiquetas.rsEtiquetas.Fields("pallet").Value = " "
            dteEtiquetas.rsEtiquetas.Fields("conf_pallet").Value = " "
            dteEtiquetas.rsEtiquetas.Fields("usuario").Value = " "
            dteEtiquetas.rsEtiquetas.Fields("tp_transporte").Value = 0
            dteEtiquetas.rsEtiquetas.Fields("tipo_padrao_cx").Value = 0
            dteEtiquetas.rsEtiquetas.Fields("placa").Value = " "
            dteEtiquetas.rsEtiquetas.Fields("sequencia_placa").Value = 0
            dteEtiquetas.rsEtiquetas.Fields("id_coletor").Value = 0
            dteEtiquetas.rsEtiquetas.Fields("tipo_frete").Value = " "
            dteEtiquetas.rsEtiquetas.Fields("indforjimport").Value = gEtiqueta.indforjimport
            If Trim(gEtiqueta.Tipo) = "F" Then
'               If Trim(gEtiqueta.fatur) <> Empty Then dteEtiquetas.rsEtiquetas.Fields("fatur").Value = Trim(gEtiqueta.fatur)
'               If Trim(gEtiqueta.xblnr) <> Empty Then dteEtiquetas.rsEtiquetas.Fields("xblnr").Value = Trim(gEtiqueta.xblnr)
'               If Trim(gEtiqueta.posnr) <> Empty Then dteEtiquetas.rsEtiquetas.Fields("posnr").Value = Trim(gEtiqueta.posnr)
'               If Trim(gEtiqueta.pallet) <> Empty Then dteEtiquetas.rsEtiquetas.Fields("pallet").Value = Trim(gEtiqueta.pallet)
'               If Trim(gEtiqueta.conf_pallet) <> Empty Then dteEtiquetas.rsEtiquetas.Fields("conf_pallet").Value = Trim(gEtiqueta.conf_pallet)
'               If Trim(gEtiqueta.usuario) <> Empty Then dteEtiquetas.rsEtiquetas.Fields("usuario").Value = Trim(gEtiqueta.usuario)
               If Trim(gEtiqueta.tp_transporte) <> Empty Then dteEtiquetas.rsEtiquetas.Fields("tp_transporte").Value = Trim(gEtiqueta.tp_transporte)
'               If Trim(gEtiqueta.tipo_padrao_cx) <> Empty Then dteEtiquetas.rsEtiquetas.Fields("tipo_padrao_cx").Value = Trim(gEtiqueta.tipo_padrao_cx)
'               If Trim(gEtiqueta.Placa) <> Empty Then dteEtiquetas.rsEtiquetas.Fields("placa").Value = Trim(gEtiqueta.Placa)
'               If Trim(gEtiqueta.sequencia_placa) <> Empty Then dteEtiquetas.rsEtiquetas.Fields("sequencia_placa").Value = Trim(gEtiqueta.sequencia_placa)
'               If Trim(gEtiqueta.id_coletor) <> Empty Then dteEtiquetas.rsEtiquetas.Fields("id_coletor").Value = Trim(gEtiqueta.id_coletor)
'               If Trim(gEtiqueta.tipo_frete) <> Empty Then dteEtiquetas.rsEtiquetas.Fields("tipo_frete").Value = Trim(gEtiqueta.tipo_frete)
               If Trim(gEtiqueta.indforjimport) <> Empty Then dteEtiquetas.rsEtiquetas.Fields("indforjimport").Value = Trim(gEtiqueta.indforjimport)
            End If
            Rem caso seja etiqueta da yamaha, será verificado o quantitativo de emitidaspara iniciar a sequencia da etiqueta do dia
            If Trim(gEtiqueta.Tipo) = "Y" Then
               If nSequencia_Yamaha = 0 Then
                  nSequencia_Yamaha = CCTempneSequenciaDia.SEQUENCIA_DIA_Consultar(sBancoMusashi, "Y")
                  dteEtiquetas.rsEtiquetas.Fields("Sequencia_Dia") = nSequencia_Yamaha + 1
               Else
                  nSequencia_Yamaha = nSequencia_Yamaha + 1
                  dteEtiquetas.rsEtiquetas.Fields("Sequencia_Dia") = nSequencia_Yamaha
               End If
             End If
Rem ****** incluido
            
            Rem acrescentado em 15/05/2006
            
            'Ainda não usado
            'gEtiqueta.tipo_transporte '(A/R)
            
            
            If Trim(gEtiqueta.Tipo) = "1" Or _
               Trim(gEtiqueta.Tipo) = "2" Or _
               Trim(gEtiqueta.Tipo) = "3" Or _
               Trim(gEtiqueta.Tipo) = "4" Or _
               Trim(gEtiqueta.Tipo) = "5" Or _
               Trim(gEtiqueta.Tipo) = "6" Or _
               Trim(gEtiqueta.Tipo) = "7" Or _
               Trim(gEtiqueta.Tipo) = "9" Or _
               Trim(gEtiqueta.Tipo) = "Y" Or _
               Trim(gEtiqueta.Tipo) = "F" Or _
               Trim(gEtiqueta.Tipo) = "S" Or _
               Trim(gEtiqueta.Tipo) = "M" Or _
               Trim(gEtiqueta.Tipo) = "Z" Then
                
                Set objEtiqueta = New Etiqueta
                objEtiqueta.setTipoBanco = objApplication.getTipoBancoInterface

INC_NOVA_ETIQUTA:

Rem parar aqui marcos, continua em proximaEtqLivreAR = proximaEtqLivreAR + 1
'''' retirar aqui marcos VER AQUI MAURO
                If proximaEtqLivreAR = 0 Then
                   If Trim(gEtiqueta.Tipo) <> "S" Then
                      proximaEtqLivreAR = Val(Left(objEtiquetaControlador.getProximaEtiquetaCaixaLivre(), 10))
                      objEtiqueta.itens.Item("ID_ETIQUETA") = Format(proximaEtqLivreAR, "0000000000")
                   Else
                      proximaEtqLivreAR = proximaEtqLivreAR + 1
                      objEtiqueta.itens.Item("ID_ETIQUETA") = Format(proximaEtqLivreAR, "0000000000")
                   End If
                Else
                   proximaEtqLivreAR = proximaEtqLivreAR + 1
                   objEtiqueta.itens.Item("ID_ETIQUETA") = Format(proximaEtqLivreAR, "0000000000")
                End If

Rem AQUI MARCOS COMENTADO                   objEtiqueta.itens.Item("ID_ETIQUETA") = Left(objEtiquetaControlador.getProximaEtiquetaCaixaLivre(), 10)
                objEtiqueta.itens.Item("ID_CLIENTE") = Left(Trim(gEtiqueta.id_cliente), 15)
                objEtiqueta.itens.Item("ID_PECA") = Left(Trim(gEtiqueta.Cod_Peca), 15)
                'objEtiqueta.itens.Item("ID_BORDERO") = Left("NOVO_BORDERO!", 15) Será preenchido na associação com o borderô
                objEtiqueta.itens.Item("LOTE") = Left(Trim(gEtiqueta.Lote), 15)
                objEtiqueta.itens.Item("QTDE") = CLng(Trim(gEtiqueta.Qtd_Caixa))
                objEtiqueta.itens.Item("TIPO_EMBALAGEM") = Left("", 2)
                objEtiqueta.itens.Item("STATUS") = "AB"
Rem AQUI MARCOS INCLUIDO
                objEtiqueta.itens.Item("COD_NO_CLIENTE") = Left(Trim(gEtiqueta.Cod_Cliente), 20)
                objEtiqueta.itens.Item("PESO") = Left(Trim(gEtiqueta.Peso), 16)
                objEtiqueta.itens.Item("TIPO_CAIXA") = Left(Trim(gEtiqueta.Tipo_caixa), 1)
Rem AQUI MARCOS INCLUIDO 15/05/2005
                objEtiqueta.itens.Item("TIPO_PADRAO_CX") = gEtiqueta.tp_transporte
Rem AQUI MARCOS INCLUIDO 15/08/2011
                objEtiqueta.itens.Item("indforjimport") = gEtiqueta.indforjimport
                nVarInsEtiqDup = 0
                
Rem INCLUIR FUNÇÃO PARA GERAR A IMAGEM DA ETIQUETA PDF417 EM UMA TELA DE AUXILIO E GRAVAR NO BANCO DE DADOS
'                If Trim(gEtiqueta.Tipo) = "3" Then
'
'                   Set oTela = New frmTelaFormacaoPDF417
'                   oTela.sNomeArquivo = objEtiqueta.itens.Item("ID_ETIQUETA")
'                   oTela.DataToEncodeText.Text = Format(proximaEtqLivreAR, "0") & " (P)" & Format(Trim(gEtiqueta.Qtd_Caixa), "0") & " (Q)" & Trim(gEtiqueta.Cod_Fornec) & " (V)" & Mid$(gEtiqueta.Ano, 3, 2) & Trim(gEtiqueta.Mes) & Trim(gEtiqueta.Dia) & " (D)" & Trim(gEtiqueta.Serial) & " (S)"
'                   Unload oTela: Set oTela = Nothing
'                End If
                
Rem parar aqui marcos, continua em If nVarInsEtiqDup = 1 Then
'''' retirar aqui marcos
                If Trim(gEtiqueta.Tipo) <> "S" Then
                   objEtiqueta.save (adInserir)
                End If
                
                If nVarInsEtiqDup = 1 Then
                   proximaEtqLivreAR = 0
                   GoTo INC_NOVA_ETIQUTA
                End If
                
                dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA") = objEtiqueta.itens.Item("ID_ETIQUETA")
                
                dteEtiquetas.rsEtiquetas.Update
                
                Rem aqui gravara o mov_etiq, mas que não seja de Manaus
                Rem 19/10/2012, sera gravada o mov etiq em manaus
'                If objApplication.filial = adMusashiDoBrasil Then

Rem parar aqui marcos, continua em dteEtiquetas.rsEtiquetas.Update
'''' retirar aqui marcos
                If Trim(gEtiqueta.Tipo) <> "S" Then
                   Call Gravar_Arquivo_MovEtiq
                End If
                
            Else 'Trim(gEtiqueta.Tipo) = "A" Then
Rem AQUI MARCOS INCLUIDO
                If proximaEtqLivreAR = 0 Then
                   proximaEtqLivreAR = Val(Left(objEtiquetaControlador.getProximaEtiquetaCaixaLivre(), 10))
                   dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA") = Left(objEtiquetaControlador.getProximaEtiquetaCaixaLivre(), 10)
                Else
                   proximaEtqLivreAR = proximaEtqLivreAR + 1
                   dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA") = Format(proximaEtqLivreAR, "0000000000")
                End If
                
                dteEtiquetas.rsEtiquetas.Update


Rem AQUI MARCOS INCLUIDO
Rem RETIRADO MARCOS                    dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA") = Left(objEtiquetaControlador.getProximaEtiquetaCaixaLivre(), 10)
                '''dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA") = ""
            End If
            
'''            dteEtiquetas.rsEtiquetas.Update
        '''Next
    Next
Rem parar aqui marcos, continua em Exit Sub

'''' retirar aqui marcos
    If UCase(Dir(objApplication.caminhoImportacao & "\etiq.txt")) = "ETIQ.TXT" Then
        Close gNumeroArquivo
        Kill objApplication.caminhoImportacao & "\etiq.txt"
    End If
    
    Exit Sub
Erro:
    If Err.Number = -2147217873 Then proximaEtqLivreAR = 0: GoTo INC_NOVA_ETIQUTA
    MsgErro "Ocorreu um erro ao importar as etiquetas do arquivo texto."
End Sub

Public Sub LimpaEtiquetas()
    'FIAT
    frmExibicao2.lblCod_Peca = ""
    frmExibicao2.lblDataExpedicao2 = ""
    frmExibicao2.lblCodFornec2 = ""
    frmExibicao2.lblDenominacao2 = ""
    frmExibicao2.lblBam2 = ""
    frmExibicao2.lblDesenho2 = ""
    frmExibicao2.lblCodBarra = ""
    frmExibicao2.lblCodBarraCp1 = ""
    frmExibicao2.lblCodBarraCp2 = ""
    frmExibicao2.lblDataProducao2 = ""
    frmExibicao2.lblCodEmbalagem2 = ""
    frmExibicao2.lblNumLote2 = ""
    frmExibicao2.lblQtdLote2 = ""
    frmExibicao2.lblQtdEmbalagem2 = ""
    frmExibicao2.lblClasseFuncional2 = ""
    frmExibicao2.lblVinculo2 = ""
    frmExibicao2.lblIndicacaoSuplementar2 = ""
    frmExibicao2.lblPontoEntrega2 = ""
    frmExibicao2.lblEmbarqueControlado2 = ""
    frmExibicao2.lblLoteSobDesvio2 = ""
    frmExibicao2.lblDum2 = ""
    
    'FORD
    frmExibicao3ant.lblNumPeca = ""
    frmExibicao3ant.lblNumPecaA = ""
    frmExibicao3ant.lblNumPecaB = ""
    frmExibicao3ant.lblLote = ""
    frmExibicao3ant.lblQtd = ""
    frmExibicao3ant.lblQtdA = ""
    frmExibicao3ant.lblQtdB = ""
    frmExibicao3ant.lblSufixo = ""
    frmExibicao3ant.lblSufixoA = ""
    frmExibicao3ant.lblSufixoB = ""
    frmExibicao3ant.lblNumFornec = ""
    frmExibicao3ant.lblNumFornecA = ""
    frmExibicao3ant.lblNumFornecB = ""
    frmExibicao3ant.lblCodUtil = ""
    frmExibicao3ant.lblLinhaUtil = ""
    frmExibicao3ant.lblNumSerial = ""
    frmExibicao3ant.lblNumSerialA = ""
    frmExibicao3ant.lblNumSerialB = ""
    frmExibicao3ant.lblDestino = ""
    frmExibicao3ant.lblDestinoA = ""
    frmExibicao3ant.lblDestinoB = ""
End Sub

Public Sub Gravar_Arquivo_MovEtiq()

Dim cFields As Collection
Dim cRec_Aux As ADODB.Recordset

On Error GoTo Erro
'Rem aqui marcos recebe os dados do aquivo texto e grava no banco SQL
'


Set cFields = New Collection

cFields.Add IIf(IsNull(gEtiqueta.Cod_Peca), "NULL", "'" & Trim(gEtiqueta.Cod_Peca) & "'"), "Cod_Peca"                                   '01    As String * 7
cFields.Add Trim(gEtiqueta.Lote), "Lote"                                             '02    As String * 15
cFields.Add gEtiqueta.Qtd_Caixa, "Qtd_Caixa"                                         '03    As numeric * 10
cFields.Add gEtiqueta.Mes & "-" & gEtiqueta.Dia & "-" & gEtiqueta.Ano, "Data_Etiq"   '04    As ddmmyyyy * 8
cFields.Add Trim(gEtiqueta.Cod_Tabela), "Cod_Tabela"                                 '05    As String * 4
cFields.Add Trim(gEtiqueta.Cliente), "Cliente"                                       '06    As String * 30
cFields.Add Trim(gEtiqueta.Cod_Cliente), "Cod_Cliente"                               '07    As String * 20
cFields.Add Format(Trim(Replace(gEtiqueta.Peso, ".", ",")), "0.00"), "Peso"                                          '08    As String * 16
cFields.Add Trim(gEtiqueta.Descr_Peca), "Descr_Peca"                                 '09    As String * 40
cFields.Add gEtiqueta.Qtd_Etiq, "Qtd_Etiq"                                           '10    As numeric * 4
cFields.Add gEtiqueta.Tipo, "Tipo"                                                   '11    As String * 1

If gEtiqueta.Mes2 = "00" And gEtiqueta.Dia2 = "00" And gEtiqueta.Ano2 = "0000" Then
    cFields.Add "", "Data_Expedicao"                                                 '12    As ddmmyyyy * 8
Else
    If Len(Trim(gEtiqueta.Mes2)) > 0 And Len(Trim(gEtiqueta.Dia2)) > 0 And Len(Trim(gEtiqueta.Ano2)) > 0 Then
       cFields.Add gEtiqueta.Mes2 & "-" & gEtiqueta.Dia2 & "-" & gEtiqueta.Ano2, "Data_Expedicao" '12    As ddmmyyyy * 8
    Else
       cFields.Add "", "Data_Expedicao"                                              '12    As ddmmyyyy * 8
    End If
End If

cFields.Add Trim(gEtiqueta.Cod_Embalagem_Pw), "Cod_Embalagem_Pw"                     '13    As String * 3

If gEtiqueta.Mes3 = "00" And gEtiqueta.Dia3 = "00" And gEtiqueta.Ano3 = "0000" Then
    cFields.Add "", "Data_Lote"                                                      '14    As ddmmyyyy * 8
Else
    If Len(Trim(gEtiqueta.Mes3)) > 0 And Len(Trim(gEtiqueta.Dia3)) > 0 And Len(Trim(gEtiqueta.Ano3)) > 0 Then
       cFields.Add Trim(gEtiqueta.Mes3 & "-" & gEtiqueta.Dia3 & "-" & gEtiqueta.Ano3), "Data_Lote" '14    As ddmmyyyy * 8
    Else
       cFields.Add "", "Data_Lote"                                                    '14    As ddmmyyyy * 8
    End If
End If

cFields.Add Trim(gEtiqueta.Classe_Funcional), "Classe_Funcional"                     '15    As String * 2
cFields.Add Trim(gEtiqueta.Embarque_Controlado), "Embarque_Controlado"               '16    As String * 3
cFields.Add Trim(gEtiqueta.Cod_Fornec), "Cod_Fornec"                                 '17    As String * 10
cFields.Add Trim(gEtiqueta.Num_Doc_Fiscal), "Num_Doc_Fiscal"                         '18    As String * 6
cFields.Add Trim(gEtiqueta.Cod_Embalagem), "Cod_Embalagem"                           '19    As String * 3
cFields.Add Trim(gEtiqueta.Ind_Suplementar), "Ind_Suplementar"                       '20    As String * 2 'Apos este campos pos 200
cFields.Add Trim(gEtiqueta.Pto_Entrega), "Pto_Entrega"                               '21    As String * 15

If gEtiqueta.Dia4 = "00" And gEtiqueta.Mes4 = "00" And gEtiqueta.Ano4 = "0000" Then
    cFields.Add "", "Dum"                                                            '22    As ddmmyyyy * 8
Else
    If Len(Trim(gEtiqueta.Dia4)) > 0 And Len(Trim(gEtiqueta.Mes4)) > 0 And Len(Trim(gEtiqueta.Ano4)) > 0 Then
       If gEtiqueta.Tipo = "F" Or gEtiqueta.Tipo = "Z" Then
          cFields.Add Mid$(Trim(gEtiqueta.Ano4), 1, 2) & "-" & Mid$(Trim(gEtiqueta.Ano4), 3, 2) & "-" & gEtiqueta.Dia4 & gEtiqueta.Mes4, "Dum" '22    As ddmmyyyy * 8
       Else
          cFields.Add Trim(gEtiqueta.Mes4 & "-" & gEtiqueta.Dia4 & "-" & gEtiqueta.Ano4), "Dum" '22    As ddmmyyyy * 8
       End If
    Else
       cFields.Add "", "Dum"                                                         '22    As ddmmyyyy * 8
    End If
'    cFields.Add Trim(gEtiqueta.Mes4 & "-" & gEtiqueta.Dia4 & "-" & gEtiqueta.Ano4), "Dum" '22    As ddmmyyyy * 8
End If

cFields.Add gEtiqueta.Qtd_Lote, "Qtd_Lote"                                           '23    As numeric * 6
cFields.Add Trim(gEtiqueta.Vinculo), "Vinculo"                                       '24    As String * 2
cFields.Add Trim(gEtiqueta.desvio), "desvio"                                         '25    As String * 5
cFields.Add gEtiqueta.Serial, "Serial"                                               '26    As String * 10
cFields.Add Trim(gEtiqueta.Cod_Util), "Cod_Util"                                     '27    As String * 4
cFields.Add gEtiqueta.Linha_Util, "Linha_Util"                                       '28    As String * 4
cFields.Add Trim(gEtiqueta.Modelo), "Modelo"                                         '29    As String * 10
cFields.Add Trim(gEtiqueta.Tipo_caixa), "Tipo_caixa"                                 '30    As String * 1
cFields.Add Trim(gEtiqueta.desvio_aviso_mod), "desvio_aviso_mod"                     '31    As String * 25
cFields.Add Trim(gEtiqueta.tipo_alteracao), "tipo_alteracao"                         '32    As String * 1
cFields.Add Trim(gEtiqueta.motivo_alteracao), "motivo_alteracao"                     '33    As String * 1
cFields.Add Trim(gEtiqueta.motivo_alteracao_outros), "motivo_alteracao_outros"       '24    As String * 25
cFields.Add Trim(gEtiqueta.envio_lote), "envio_lote"                                 '35    As String * 1
cFields.Add Trim(gEtiqueta.num_am), "num_am"                                         '36    As String * 8
cFields.Add Trim(gEtiqueta.Embalagem), "Embalagem"                                   '37    As String * 5
cFields.Add Trim(gEtiqueta.id_cliente), "id_cliente"                                 '38    As String * 10
cFields.Add Trim(gEtiqueta.tp_transporte), "tp_transporte"                           '39    As String * 1    '(A/R)
cFields.Add Trim(gEtiqueta.codigo_produto1), "codigo_produto1"                       '40    As String * 10
cFields.Add Trim(gEtiqueta.qtde_caixa1), "qtde_caixa1"                               '41    As String * 3
cFields.Add Trim(gEtiqueta.pecas_caixa1), "pecas_caixa1"                             '42    As String * 4 'Apos este campo pos 358
cFields.Add Trim(gEtiqueta.compl_peca1), "compl_peca1"                               '43    As String * 10
cFields.Add Trim(gEtiqueta.codigo_produto2), "codigo_produto2"                       '44    As String * 10
cFields.Add Trim(gEtiqueta.qtde_caixa2), "qtde_caixa2"                               '45    As String * 3
cFields.Add Trim(gEtiqueta.pecas_caixa2), "pecas_caixa2"                             '46    As String * 4 'Apos este campo pos 385
cFields.Add Trim(gEtiqueta.compl_peca2), "compl_peca2"                               '47    As String * 10
cFields.Add Trim(gEtiqueta.codigo_produto3), "codigo_produto3"                       '48    As String * 10
cFields.Add Trim(gEtiqueta.qtde_caixa3), "qtde_caixa3"                               '49    As String * 3
cFields.Add Trim(gEtiqueta.pecas_caixa3), "pecas_caixa3"                             '50    As String * 4 'Apos este campo 412
cFields.Add Trim(gEtiqueta.compl_peca3), "compl_peca3"                               '51    As String * 10
cFields.Add Trim(gEtiqueta.codigo_produto4), "codigo_produto4"                       '52    As String * 10
cFields.Add Trim(gEtiqueta.qtde_caixa4), "qtde_caixa4"                               '53    As String * 3
cFields.Add Trim(gEtiqueta.pecas_caixa4), "pecas_caixa4"                             '54    As String * 4
cFields.Add Trim(gEtiqueta.compl_peca4), "compl_peca4"                               '55    As String * 10
cFields.Add Format(proximaEtqLivreAR, "0000000000"), "sequencia"                     '56    As String * 10
cFields.Add Format(proximaEtqLivreAR, "0000000000"), "id_etiqueta"                   '57    As String * 10
cFields.Add "", "sequencia_org"                                                      '58    As String * 10
cFields.Add "0", "statusmov"                                                         '59    As String * 4
cFields.Add "0", "usuario_org"                                                        '60    As String * 4
cFields.Add "PP04", "lgort"                                                          '61    As String * 4
cFields.Add Format(Now(), "mm-dd-yyyy"), "dt_alteracao"                              '62    As String * 8
cFields.Add gEtiqueta.indforjimport, "indForjImport"
cFields.Add nSequencia_Yamaha
Set cRec_Aux = New ADODB.Recordset

Rem INCLUIR AS ETIQUETAS NA TABELAMOV_ETIQ
Set cRec_Aux = CCTempneMov_Etiq.Mov_Etiq_Incluir(sBancoMusashi, _
                                                 cFields)
Exit Sub

Erro:
'If Err.Number = -2147217873 Then proximaEtqLivreAR = 0: GoTo INC_NOVA_ETIQUETA
MsgErro "Ocorreu um erro ao importar as etiquetas do arquivo texto."


End Sub

Private Sub mnuTelaTeste_Click()
    frmExibicao41.Show
End Sub
