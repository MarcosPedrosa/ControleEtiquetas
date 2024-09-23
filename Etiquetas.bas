Attribute VB_Name = "ArqEtiqueta"
Global sBancoMusashi As String
Global sBancoAccess As String
Global proximaEtqLivreAR As Double

Rem aqui cod_peca, tamanho = 7 para 8 (10/09/2010)
Rem aqui cod_peca, tamanho = 8 para 9 (21/09/2012)

Type Arquivo
    Cod_Peca    As String * 9 ' 1-9
    Lote        As String * 15 ' 10-24
    Qtd_Caixa   As String * 10 ' 25-34
    'Data_Etiq
    Dia         As String * 2 ' 35-36
    Mes         As String * 2 ' 37-38
    Ano         As String * 4 ' 39-42
    Cod_Tabela  As String * 4 ' 43-46
    Cliente     As String * 30 ' 47-76
    Cod_Cliente As String * 20 ' 77-96
    Peso        As String * 16 ' 97-112
    Descr_Peca  As String * 40 ' 113-152
    Qtd_Etiq    As String * 4 '153-156
    Tipo        As String * 1 ' 157-157
    'Data_Expedicao
    Dia2        As String * 2 ' 158-159
    Mes2        As String * 2 ' 160-161
    Ano2        As String * 4 ' 162-165
    Cod_Embalagem_Pw    As String * 3 ' 166-168
    'Data_Lote
    Dia3        As String * 2 ' 169-170
    Mes3        As String * 2 ' 171-172
    Ano3        As String * 4 ' 173-176
    Classe_Funcional As String * 2 '177-178
    Embarque_Controlado As String * 3 ' 179-181
    Cod_Fornec  As String * 10 ' 182-191
    Num_Doc_Fiscal  As String * 6 '192-197
    Cod_Embalagem   As String * 3 ' 198-200
    Ind_Suplementar   As String * 2 '201-202
    Pto_Entrega As String * 15 ' 203-217
    'Dum
    Dia4        As String * 2 '218-219
    Mes4        As String * 2 ' 220-221
    Ano4        As String * 4 ' 222-225
    Qtd_Lote    As String * 6 ' 226-231
    Vinculo     As String * 2 ' 232-233
    desvio      As String * 5 ' 234-238
    Serial      As String * 10 ' 239-248
    Cod_Util    As String * 4 ' 249-252
    Linha_Util  As String * 4 ' 253-257
    Modelo      As String * 10 ' 258-267
    'Novos campos:
    desvio_aviso_mod        As String * 25 ' 268-292
    tipo_alteracao          As String * 1 ' 293-293
    motivo_alteracao        As String * 1 ' 294-294
    motivo_alteracao_outros As String * 25 ' 295-319
    envio_lote              As String * 1 ' 320-320
    num_am                  As String * 8 ' 321-328
    Embalagem               As String * 5 ' 329-333
    id_cliente              As String * 10 ' 334-343
    tp_transporte         As String * 1    '(A/R) ' 344-344
    codigo_produto1         As String * 10 ' 345-354
    qtde_caixa1             As String * 3 ' 355-357
    pecas_caixa1            As String * 4 '358-361
    compl_peca1             As String * 10 ' 362-371
    codigo_produto2         As String * 10 ' 372-381
    qtde_caixa2             As String * 3 ' 382-384
    pecas_caixa2            As String * 4 ' 385-388
    compl_peca2             As String * 10 ' 389-398
    codigo_produto3         As String * 10 ' 399-408
    qtde_caixa3             As String * 3 ' 409-411
    pecas_caixa3            As String * 4 '412-415
    compl_peca3             As String * 10 ' 416-425
    codigo_produto4         As String * 10 ' 426-435
    qtde_caixa4             As String * 3 ' 436-438
    pecas_caixa4            As String * 4 ' 439-342
    compl_peca4             As String * 10 ' 343-352
    Tipo_caixa              As String * 1 ' 353-353
    indforjimport           As String * 1  ' 354-354  se "x", entao imprime etiqueta com a tarja da opcao 4 para o form etiqueta4
    Sequencia_Dia           As String * 5 ' 355-359
End Type
'****************************************************************************************************************

Type Arquivo_Transporte
    Placa       As String * 11
    Sequencial  As String * 6
    Tipo_transp As String * 1
    Tipo_caixa  As String * 1
    Cod_transp  As String * 10
    Nome_transp As String * 35
    Motorista   As String * 35
    final_arq   As String * 2
End Type

Global Arq_Transporte As Arquivo_Transporte
'****************************************************************************************************************

Type Arquivo_Honda_Manaus
    pedido         As String * 26
    linha1         As String * 1
    tipo_pedido    As String * 26
    linha2         As String * 1
    linha_pedido   As String * 26
    linha3         As String * 1
    nota_fiscal    As String * 26
    linha4         As String * 1
    serie          As String * 26
    linha5         As String * 1
    codigo_item    As String * 26
    linha6         As String * 1
    descricao      As String * 30
    linha7         As String * 1
    quantidade     As String * 26
    linha8         As String * 1
    unidade        As String * 26
    linha9         As String * 1
    empresa        As String * 26
    linha10        As String * 1
    data_entrega   As String * 26
    linha11        As String * 1
    hora_entrega   As String * 26
    linha12        As String * 1
    setor          As String * 26
    linha13        As String * 1
    volume         As String * 26
    linha14        As String * 1
    total_volume   As String * 27
    linha15        As String * 1
    outros         As String * 50
    linha16        As String * 1
    cod_bar_pedido As String * 27
    linha17        As String * 1
    cod_bar_item   As String * 40
    linha18        As String * 1
    local_entrega  As String * 40
    linha19        As String * 1
    psv            As String * 15
    linha20        As String * 1
    final_arq      As String * 2
End Type

Global Arquivo_AM As Arquivo_Honda_Manaus
'****************************************************************************************************************

Public Function CCTempneMov_Etiq() As neMov_Etiq
     Set CCTempneMov_Etiq = New neMov_Etiq
End Function

Public Function CCTempneUsuario() As neUsuario
     Set CCTempneUsuario = New neUsuario
End Function
Public Function CCTempneEtiqueta() As neEtiqueta
     Set CCTempneEtiqueta = New neEtiqueta
End Function
Public Function CCTempPecaAvulso() As PecaAvulso
     Set CCTempPecaAvulso = New PecaAvulso
End Function
Public Function CCTempneSequenciaDia() As neSequenciaDia
     Set CCTempneSequenciaDia = New neSequenciaDia
End Function
Public Function CCTempneInmetroCadCliente() As neInmetroCadCliente
     Set CCTempneInmetroCadCliente = New neInmetroCadCliente
End Function
Public Function CCTempneInmetroCadmodelo() As neInmetroCadModelo
     Set CCTempneInmetroCadmodelo = New neInmetroCadModelo
End Function
Public Function CCTempneInmetroCadPeca() As neInmetroCadPeca
     Set CCTempneInmetroCadPeca = New neInmetroCadPeca
End Function

