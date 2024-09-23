VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "neMov_Etiq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'**************** Ações Atualizacao Tab. Mov. etiq
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE NEGOCIOS DO Ações na tabela movEtiq
' ***********************************************************
' ***********************************************************
' ***********************************************************

Public Function Mov_Etiq_Consultar_Filtro(ByVal sBanco As String, _
                                          ByVal c_sequencial As String, _
                                          ByVal c_DataSelecao As String, _
                                          ByVal c_Lote As String, _
                                          ByVal c_peca As String, _
                                          ByVal c_Qtd_Caixa As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim bDigita As Boolean ' campo auxiliar para criticar os campos obrigatorios que devem ser preenchidos

If Len(Trim(c_sequencial)) = 0 And _
   Len(Trim(c_Qtd_Caixa)) = 0 And _
   Len(Trim(c_Lote)) = 0 And _
   Len(Trim(c_peca)) = 0 Then
   Err.Raise 50000, Err.Source, "Não há nada digitado no filtro para seleção de registros!"
End If

bDigita = False

If Len(Trim(c_Qtd_Caixa)) > 0 Or _
   Len(Trim(c_Lote)) > 0 Or _
   Len(Trim(c_peca)) > 0 Then
   bDigita = True ' caso algum campo acima tenha algo,
End If

If bDigita = True Then ' se houver campo preenchido, todos os tres dvem estar
   If Len(Trim(c_Qtd_Caixa)) = 0 Or _
      Len(Trim(c_Lote)) = 0 Or _
      Len(Trim(c_peca)) = 0 Then
      Err.Raise 50001, Err.Source, "Não digitando o sequencial, os outros tres campos são obrigatorios para seleção de registros!"
   End If
End If

If Len(Trim(c_sequencial)) > 0 Then
   If Not IsNumeric(c_sequencial) Then
      Err.Raise 50001, Err.Source, "Sequencial nâo numérico, redigite-o para seleção de registro!"
   End If
End If

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "", c_sequencial), "sequencial"
cFields.Add IIf(Len(Trim(c_DataSelecao)) = 0, "", "'" & Mid$(c_DataSelecao, 7, 4) & Mid$(c_DataSelecao, 4, 2) & Mid$(c_DataSelecao, 1, 2) & "'"), "DataSelecao"
cFields.Add IIf(Len(Trim(c_Lote)) = 0, "", "'" & c_Lote & "'"), "Lote"
cFields.Add IIf(Len(Trim(UCase(c_peca))) = 0, "", "'" & c_peca & "'"), "peca"
cFields.Add IIf(Len(Trim(c_Qtd_Caixa)) = 0, "", "'" & c_Qtd_Caixa & "'"), "Qtd_Caixa"

Set Mov_Etiq_Consultar_Filtro = Mov_Etiq_Critica(ADOConnection, cFields, 1)

If Mov_Etiq_Consultar_Filtro.RecordCount = 0 Then
   Err.Raise 50001, Err.Source, "Não há registro(s) encontrado(s) com esta Seleção!"
End If

If Mov_Etiq_Consultar_Filtro!statusMov = 3 Then
   Err.Raise 50000, Err.Source, "Sequência desta Etiqueta já excluida!"
End If

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:

If ADOConnection Is Nothing Then
Else
   If ADOConnection <> "" Then
      ADOConnection.RollbackTrans
   End If
End If

Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Incluir(ByVal sBanco As String, _
                                 ByVal cFields As Collection) As ADODB.Recordset

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cConect As daAbertura
Dim nSequencia As Double


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Set Mov_Etiq_Incluir = Mov_Etiq_Critica(ADOConnection, cFields, 2)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function Mov_Etiq_Consultar(ByVal sBanco As String, _
                                   ByVal c_sequencial As String) As ADODB.Recordset
                                   

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "", c_sequencial), "sequencial"

Set Mov_Etiq_Consultar = Mov_Etiq_Critica(ADOConnection, cFields, 12)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Alt_Campos(ByVal sBanco As String, _
                                    ByVal c_sequencial As String, _
                                    ByVal c_statusmov As String, _
                                    ByVal c_sequencia_org As String, _
                                    ByVal c_usuario_org As String) As ADODB.Recordset
                                   

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

If Len(Trim(c_sequencial)) = 0 Then
   Err.Raise 50000, Err.Source, "O sequencial não foi digitado, não deverá atualizar o registro!"
End If

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "", "'" & c_sequencial & "'"), "sequencial"
cFields.Add IIf(Len(Trim(c_statusmov)) = 0, "", c_statusmov), "c_statusmov"
cFields.Add IIf(Len(Trim(c_sequencia_org)) = 0, "", "'" & c_sequencia_org & "'"), "c_sequencia_org"
cFields.Add IIf(Len(Trim(c_usuario_org)) = 0, "0", c_usuario_org), "c_usuario_org"

Set Mov_Etiq_Alt_Campos = Mov_Etiq_Critica(ADOConnection, cFields, 13)

Rem caso seja uma exclusão será excluida da tabela etiquetas status = 3
If c_statusmov = "3" Then
   Set cFields = New Collection
   cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "", c_sequencial), "sequencial"
   Call CCTempneEtiqueta.Etiqueta_Critica(ADOConnection, cFields, 14)
End If

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Desmembrar(ByVal sBanco As String, _
                                    ByVal c_sequencial As String, _
                                    ByVal c_usuario As String, _
                                    ByVal c_fields As Collection, _
                                    ByVal c_fieldsDes As Collection) As ADODB.Recordset
                                  
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cQtde As Collection
Dim cConect As daAbertura
Dim nx As Integer
Dim nQtdeAux As Integer
Dim nProxSeq As Double
Dim cRec As ADODB.Recordset
Dim nPeso As Double
Dim nStatus As String

On Error GoTo Erro

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Rem *******************************************************************************
Rem verificar as quantidade e criticar se bate com o total da etiqueta original
Rem *******************************************************************************
nQtdeAux = 0
For nx = 3 To c_fields.Count
    nQtdeAux = nQtdeAux + CInt(c_fields(nx))
Next
If CInt(c_fields(1)) <> nQtdeAux Then
   Err.Raise 50001, Err.Source, "Soma das quantidades não confere com o total. Diferença é de " & str(CInt(c_fields(1)) - nQtdeAux)
End If
Rem *******************************************************************************

Rem *******************************************************************************
Rem pegar a ultima sequencia no banco para adicionar a proxima sequencia disponivel na tabela mov_etiq
Rem *******************************************************************************
nProxSeq = 0
cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "", c_sequencial), "sequencial"
Set Mov_Etiq_Desmembrar = Mov_Etiq_Critica(ADOConnection, cFields, 11)
If Mov_Etiq_Desmembrar.RecordCount = 0 Then
   nProxSeq = Mov_Etiq_Desmembrar!ULTIMA_ETIQUETA
Else
   nProxSeq = Mov_Etiq_Desmembrar!ULTIMA_ETIQUETA + 1
End If
Rem *******************************************************************************

Rem *******************************************************************************
Rem pegar os dados(recordset) da mov_etiq que sera desmembrada para gerar as novas
Rem *******************************************************************************
Set cRec = New ADODB.Recordset
Set cRec = Mov_Etiq_Critica(ADOConnection, cFields, 12)

If cRec.RecordCount = 0 Then
   Err.Raise 50002, Err.Source, "Etiqueta que será desmenbrada não encontrada, consulte-a novamente,caso não tenha sido excluida por outro usuário!"
End If
Rem *******************************************************************************

Rem *******************************************************************************
Rem Calcular o peso para o desmenbramento da etiqueta E CALCULAR PARA CADA NOVA ETIQUETA
Rem *******************************************************************************
nPeso = VBA.CDbl(cRec!Peso) / nQtdeAux
Set cQtde = New Collection
cQtde.Add "" 'dados em branco para acompanhar a mesma sequencia da mariz das quantidades(c_fields)
For nx = 2 To c_fields.Count
    cQtde.Add Format(VBA.CDbl(Trim(c_fields(nx))) * nPeso, "0.00")
Next

Rem *******************************************************************************
Rem *******************************************************************************
Rem Loop para determinar quantas inclusoes serao feitas pela quantidade de divisoes
Rem de caixas feita pelo usuario
Rem *******************************************************************************

For nx = 3 To c_fields.Count
     Set cFields = New Collection

GRAVAR_DESMEMBRAMENTO:
     If Len(c_fieldsDes(nx - 2)) = 0 Then
        nStatus = "3"
     Else
        nStatus = "1"
     End If
     Set cFields = Carrega_Campos_MovEtiq(Format(nProxSeq, "0000000000"), _
                                          c_fields(nx), _
                                          cQtde(nx), _
                                          nStatus, _
                                          c_usuario, _
                                          c_fields(2), _
                                          cRec)   ' carregar campos do mov_etiq, PARA Incluir a nova etiqueta
                                          
     Call Mov_Etiq_Critica(ADOConnection, cFields, 2) ' incluir dados do mov_etiq
      
     Set cFields = Carrega_Campos_Etiqueta(Format(nProxSeq, "0000000000"), _
                                                  c_fields(nx), _
                                                  cQtde(nx), _
                                                  c_fields(2), _
                                                  cRec) ' carregar os dados do tabela etiquetas para incluir dados da etiqueta
                                                  
     Call CCTempneEtiqueta.Etiqueta_Critica(ADOConnection, cFields, 2) ' incluir dados da etiqueta
     
     nProxSeq = nProxSeq + 1
Next
Rem *******************************************************************************

Rem *******************************************************************************
Rem Atualizar a etiqueta original para o status de excluida e excluir a etiqueta desmembrada
Rem *******************************************************************************
Set cFields = New Collection
cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "", "'" & c_sequencial & "'"), "sequencial"
cFields.Add "3", "c_statusmov"
cFields.Add "", "c_sequencia_org"
cFields.Add IIf(Len(Trim(c_usuario)) = 0, "0", c_usuario), "c_usuario_org"

Call Mov_Etiq_Critica(ADOConnection, cFields, 13) ' rotina para mudar o status da etiqueta original na tabela mov_etiq

Call CCTempneEtiqueta.Etiqueta_Critica(ADOConnection, cFields, 14) ' rotina para exclusao da etiqueta na tabela etiqueta
Rem *******************************************************************************

Rem *******************************************************************************
Rem Selecionar as etiquetas para impressão destas, filtrando pela sequencia original
Rem *******************************************************************************

Set Mov_Etiq_Desmembrar = Mov_Etiq_Critica(ADOConnection, cFields, 15) ' rotina para ler as novas etiquetas geradas filtrando pela sequencia original e atributo sequencia_org. na tabela mov_etiq

Rem *******************************************************************************

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:

If Err.Number = 1 Then ' ERRO DE DUPLICACAO DE CHAVE, PEGAR A PROXIMA E VOLTAR A GRAVAR
   nProxSeq = nProxSeq + 1
   GoTo GRAVAR_DESMEMBRAMENTO
End If

If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Pallet_Consulta(ByVal sBanco As String, _
                                    ByVal cFields As Collection) As ADODB.Recordset

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cConect As daAbertura
Dim nSequencia As Double


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Set Mov_Pallet_Consulta = Mov_Etiq_Critica(ADOConnection, cFields, 16)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Consultar_Ajuste(ByVal sBanco As String, _
                                          ByVal c_sequencial As String, _
                                          ByVal c_peca As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim bDigita As Boolean ' campo auxiliar para criticar os campos obrigatorios que devem ser preenchidos

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "", c_sequencial), "sequencial"
cFields.Add IIf(Len(Trim(UCase(c_peca))) = 0, "", "'" & c_peca & "'"), "peca"

Set Mov_Etiq_Consultar_Ajuste = Mov_Etiq_Critica(ADOConnection, cFields, 17)

If Mov_Etiq_Consultar_Ajuste.RecordCount = 0 Then
   Err.Raise 50001, Err.Source, "Não há registro(s) encontrado(s) com esta Seleção!"
End If

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:

If ADOConnection Is Nothing Then
Else
   If ADOConnection <> "" Then
      ADOConnection.RollbackTrans
   End If
End If

Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Aju_Campos(ByVal sBanco As String, _
                                    ByVal c_sequencial As String, _
                                    ByVal c_statusmov As String, _
                                    ByVal c_sequencia_org As String, _
                                    ByVal c_usuario_org As String) As ADODB.Recordset
                                   

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

If Len(Trim(c_sequencial)) = 0 Then
   Err.Raise 50000, Err.Source, "O sequencial não foi digitado, não deverá atualizar o registro!"
End If

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "", "'" & c_sequencial & "'"), "sequencial"
cFields.Add IIf(Len(Trim(c_statusmov)) = 0, "", c_statusmov), "c_statusmov"
cFields.Add IIf(Len(Trim(c_sequencia_org)) = 0, "", "'" & c_sequencia_org & "'"), "c_sequencia_org"
cFields.Add IIf(Len(Trim(c_usuario_org)) = 0, "0", c_usuario_org), "c_usuario_org"

Set Mov_Etiq_Aju_Campos = Mov_Etiq_Critica(ADOConnection, cFields, 18)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Consultar_Ajus(ByVal sBanco As String, _
                                        ByVal c_sequencial As String) As ADODB.Recordset
                                   

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "", c_sequencial), "sequencial"

Set Mov_Etiq_Consultar_Ajus = Mov_Etiq_Critica(ADOConnection, cFields, 19)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Consultar_Ajuste_Emb(ByVal sBanco As String, _
                                              ByVal c_sequencial As String, _
                                              ByVal c_peca As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim bDigita As Boolean ' campo auxiliar para criticar os campos obrigatorios que devem ser preenchidos

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "", c_sequencial), "sequencial"
cFields.Add IIf(Len(Trim(UCase(c_peca))) = 0, "", "'" & c_peca & "'"), "peca"

Set Mov_Etiq_Consultar_Ajuste_Emb = Mov_Etiq_Critica(ADOConnection, cFields, 20)

If Mov_Etiq_Consultar_Ajuste_Emb.RecordCount = 0 Then
   Err.Raise 50001, Err.Source, "Não há registro(s) encontrado(s) com esta Seleção!"
End If

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:

If ADOConnection Is Nothing Then
Else
   If ADOConnection <> "" Then
      ADOConnection.RollbackTrans
   End If
End If

Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Consultar_Ajus_Emb(ByVal sBanco As String, _
                                            ByVal c_sequencial As String) As ADODB.Recordset
                                   

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "", c_sequencial), "sequencial"

Set Mov_Etiq_Consultar_Ajus_Emb = Mov_Etiq_Critica(ADOConnection, cFields, 21)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Consultar_Ajuste1(ByVal sBanco As String, _
                                           ByVal c_sequencial As String, _
                                           ByVal c_peca As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim bDigita As Boolean ' campo auxiliar para criticar os campos obrigatorios que devem ser preenchidos

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "", c_sequencial), "sequencial"
cFields.Add IIf(Len(Trim(UCase(c_peca))) = 0, "", "'" & c_peca & "'"), "peca"

Set Mov_Etiq_Consultar_Ajuste1 = Mov_Etiq_Critica(ADOConnection, cFields, 22)

If Mov_Etiq_Consultar_Ajuste1.RecordCount = 0 Then
   Err.Raise 50001, Err.Source, "Não há registro(s) encontrado(s) com esta Seleção!"
End If

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:

If ADOConnection Is Nothing Then
Else
   If ADOConnection <> "" Then
      ADOConnection.RollbackTrans
   End If
End If

Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Consultar_Ajus1_Emb(ByVal sBanco As String, _
                                             ByVal c_sequencial As String) As ADODB.Recordset
                                   

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "", c_sequencial), "sequencial"

Set Mov_Etiq_Consultar_Ajus1_Emb = Mov_Etiq_Critica(ADOConnection, cFields, 23)

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Consultar_Filtro_FIAT(ByVal sBanco As String, _
                                               ByVal c_sequencial As String, _
                                               ByVal c_Selecao As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim bDigita As Boolean ' campo auxiliar para criticar os campos obrigatorios que devem ser preenchidos

If Len(Trim(c_sequencial)) = 0 Then
   Err.Raise 50000, Err.Source, "Não há nada digitado no filtro para seleção de registros!"
End If

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "111", c_sequencial), "sequencial"
cFields.Add c_Selecao

Set Mov_Etiq_Consultar_Filtro_FIAT = Mov_Etiq_Critica(ADOConnection, cFields, 24)

If Mov_Etiq_Consultar_Filtro_FIAT.RecordCount = 0 Then
   Err.Raise 50001, Err.Source, "Não há registro(s) encontrado(s) com esta Seleção!"
End If

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:

If ADOConnection Is Nothing Then
Else
   If ADOConnection <> "" Then
      ADOConnection.RollbackTrans
   End If
End If

Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Consultar_FIAT_FTP(ByVal sBanco As String, _
                                            ByVal c_sequencial As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim bDigita As Boolean ' campo auxiliar para criticar os campos obrigatorios que devem ser preenchidos

If Len(Trim(c_sequencial)) = 0 Then
   Err.Raise 50000, Err.Source, "Não há Numero da etiqueta informado para seleção de registros!"
End If
Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "", c_sequencial), "sequencial"

Set Mov_Etiq_Consultar_FIAT_FTP = Mov_Etiq_Critica(ADOConnection, cFields, 25)

If Mov_Etiq_Consultar_FIAT_FTP.RecordCount = 0 Then
   Err.Raise 50001, Err.Source, "Não há registro(s) encontrado(s) com esta Seleção!"
End If

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:

If ADOConnection Is Nothing Then
Else
   If ADOConnection <> "" Then
      ADOConnection.RollbackTrans
   End If
End If

Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Consultar_Filtro_KRUPP(ByVal sBanco As String, _
                                                ByVal c_sequencial As String, _
                                                ByVal c_Selecao As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura

If Len(Trim(c_sequencial)) = 0 Then
   Err.Raise 50000, Err.Source, "Não há nada digitado no filtro para seleção de registros!"
End If

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "111", c_sequencial), "sequencial"
cFields.Add c_Selecao

Set Mov_Etiq_Consultar_Filtro_KRUPP = Mov_Etiq_Critica(ADOConnection, cFields, 26)

If Mov_Etiq_Consultar_Filtro_KRUPP.RecordCount = 0 Then
   Err.Raise 50001, Err.Source, "Não há registro(s) encontrado(s) com esta Seleção!"
End If

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:

If ADOConnection Is Nothing Then
Else
   If ADOConnection <> "" Then
      ADOConnection.RollbackTrans
   End If
End If

Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Consultar_Filtro_FIAT_LATAM(ByVal sBanco As String, _
                                                     ByVal c_sequencial As String, _
                                                     ByVal c_Selecao As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim bDigita As Boolean ' campo auxiliar para criticar os campos obrigatorios que devem ser preenchidos

'If Len(Trim(c_sequencial)) = 0 Then
'   Err.Raise 50000, Err.Source, "Não há nada digitado no filtro para seleção de registros!"
'End If

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "111", c_sequencial), "sequencial"
cFields.Add c_Selecao

Set Mov_Etiq_Consultar_Filtro_FIAT_LATAM = Mov_Etiq_Critica(ADOConnection, cFields, 27)

If Mov_Etiq_Consultar_Filtro_FIAT_LATAM.RecordCount = 0 Then
   Err.Raise 50001, Err.Source, "Não há registro(s) encontrado(s) com esta Seleção!"
End If

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:

If ADOConnection Is Nothing Then
Else
   If ADOConnection <> "" Then
      ADOConnection.RollbackTrans
   End If
End If

Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Consultar_FIAT_FTP_LATAM(ByVal sBanco As String, _
                                                  ByVal c_sequencial As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim bDigita As Boolean ' campo auxiliar para criticar os campos obrigatorios que devem ser preenchidos

If Len(Trim(c_sequencial)) = 0 Then
   Err.Raise 50000, Err.Source, "Não há Numero da etiqueta informado para seleção de registros!"
End If
Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "", c_sequencial), "sequencial"

Set Mov_Etiq_Consultar_FIAT_FTP_LATAM = Mov_Etiq_Critica(ADOConnection, cFields, 28)

If Mov_Etiq_Consultar_FIAT_FTP_LATAM.RecordCount = 0 Then
   Err.Raise 50001, Err.Source, "Não há registro(s) encontrado(s) com esta Seleção!"
End If

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:

If ADOConnection Is Nothing Then
Else
   If ADOConnection <> "" Then
      ADOConnection.RollbackTrans
   End If
End If

Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Consultar_Filtro_HONDA_QRCODE(ByVal sBanco As String, _
                                                       ByVal c_sequencial As String, _
                                                       ByVal c_Selecao As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim bDigita As Boolean ' campo auxiliar para criticar os campos obrigatorios que devem ser preenchidos

'If Len(Trim(c_sequencial)) = 0 Then
'   Err.Raise 50000, Err.Source, "Não há nada digitado no filtro para seleção de registros!"
'End If

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "111", c_sequencial), "sequencial"
cFields.Add c_Selecao

Set Mov_Etiq_Consultar_Filtro_HONDA_QRCODE = Mov_Etiq_Critica(ADOConnection, cFields, 29)

If Mov_Etiq_Consultar_Filtro_HONDA_QRCODE.RecordCount = 0 Then
   Err.Raise 50001, Err.Source, "Não há registro(s) encontrado(s) com esta Seleção!"
End If

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:

If ADOConnection Is Nothing Then
Else
   If ADOConnection <> "" Then
      ADOConnection.RollbackTrans
   End If
End If

Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Consultar_Filtro_HONDA_QRCODE_PALLET(ByVal sBanco As String, _
                                                              ByVal c_sequencial As String, _
                                                              ByVal c_txtSeqPallet As String, _
                                                              ByVal c_Selecao As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim bDigita As Boolean ' campo auxiliar para criticar os campos obrigatorios que devem ser preenchidos

'If Len(Trim(c_sequencial)) = 0 Then
'   Err.Raise 50000, Err.Source, "Não há nada digitado no filtro para seleção de registros!"
'End If

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_Selecao
cFields.Add c_sequencial
cFields.Add c_txtSeqPallet

Set Mov_Etiq_Consultar_Filtro_HONDA_QRCODE_PALLET = Mov_Etiq_Critica(ADOConnection, cFields, 291)

If Mov_Etiq_Consultar_Filtro_HONDA_QRCODE_PALLET.RecordCount = 0 Then
   Err.Raise 50001, Err.Source, "Não há registro(s) encontrado(s) com esta Seleção!"
End If

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:

If ADOConnection Is Nothing Then
Else
   If ADOConnection <> "" Then
      ADOConnection.RollbackTrans
   End If
End If

Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Consultar_GM_P(ByVal sBanco As String, _
                                        ByVal c_sequencial As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim bDigita As Boolean ' campo auxiliar para criticar os campos obrigatorios que devem ser preenchidos

If Len(Trim(c_sequencial)) = 0 Then
   Err.Raise 50000, Err.Source, "Não há Numero da etiqueta informado para seleção de registros!"
End If
Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "", c_sequencial), "sequencial"

Set Mov_Etiq_Consultar_GM_P = Mov_Etiq_Critica(ADOConnection, cFields, 25)

If Mov_Etiq_Consultar_GM_P.RecordCount = 0 Then
   Err.Raise 50001, Err.Source, "Não há registro(s) encontrado(s) com esta Seleção!"
End If

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:

If ADOConnection Is Nothing Then
Else
   If ADOConnection <> "" Then
      ADOConnection.RollbackTrans
   End If
End If

Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function Mov_Etiq_Consultar_Filtro_GM_MASTER(ByVal sBanco As String, _
                                                    ByVal c_sequencial As String, _
                                                    ByVal c_Selecao As String) As ADODB.Recordset
On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim bDigita As Boolean ' campo auxiliar para criticar os campos obrigatorios que devem ser preenchidos

'If Len(Trim(c_sequencial)) = 0 Then
'   Err.Raise 50000, Err.Source, "Não há nada digitado no filtro para seleção de registros!"
'End If

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add IIf(Len(Trim(c_sequencial)) = 0, "111", c_sequencial), "sequencial"
cFields.Add c_Selecao

Set Mov_Etiq_Consultar_Filtro_GM_MASTER = Mov_Etiq_Critica(ADOConnection, cFields, 31)

If Mov_Etiq_Consultar_Filtro_GM_MASTER.RecordCount = 0 Then
   Err.Raise 50001, Err.Source, "Não há registro(s) encontrado(s) com esta Seleção!"
End If

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:

If ADOConnection Is Nothing Then
Else
   If ADOConnection <> "" Then
      ADOConnection.RollbackTrans
   End If
End If

Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function Mov_Etiq_Critica(ByVal cCONECCAO As Connection, _
                                 ByVal cFields As Collection, _
                                 ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim sSQL As String
Dim sName As String

If nAction > 1 And nAction < 11 Then
   'Critica dos Campos dos Mov_Etiq
    If nAction = 1 And Len(Trim(cFields(1))) = 0 Then
       Err.Raise 50000, Err.Source, "Código do Mov_Etiq Inválido"
    End If
    If Len(Trim(cFields(2))) = 0 Then
       Err.Raise 50000, Err.Source, "Nome do Mov_Etiq Inválido"
    End If
    For nx = 2 To cFields.Count
        If IsNull(cFields(nx)) Then
            'sName = cFields(nx).Name
            cFields.Remove nx
            cFields.Add "NULL", , , nx - 1
        Else
            Select Case nx
                'Verifica campos alfanuméricos NULL
                Case 2 To 60
                 If Len(Trim(cFields(nx))) = 0 Then
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add "", , , nx - 1
                 End If
                
                'Verifica campos numéricos NULL
                'Case 32, 35
                '    If Val(cFields(nx)) = 0 Then
                '        'sName = cFields(nx).Name
                '        cFields.Remove nx
                '        cFields.Add "NULL", , , nx - 1
                '    End If
            End Select
        End If
        
        If cFields(nx) <> "NULL" Then
            Select Case nx
                'Altera campos alfanuméricos
                Case 2, 5 To 9, 11, 13, 15 To 21, 24 To 58, 61, 63
                    sSQL = "'" & Trim(cFields(nx)) & "'"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                
                ' Altera campos tipo data
                Case 0
                    sSQL = "#" & Format(cFields(nx), "mm/dd/yyyy") & "#"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                 'Altera campos tipo data da alteração
                Case 4, 12, 14, 22, 62
                    sSQL = "CONVERT(DATETIME, '" & cFields(nx) & "  00:00:00', 102)"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                'Altera campos Boolean
                Case 0
                If Len(Trim(cFields(nx))) = 0 Then
                   If cFields(nx) Then
                       sSQL = "1"
                   Else
                       sSQL = "0"
                   End If
                   'sName = cFields(nx).Name
                   cFields.Remove nx
                   cFields.Add sSQL, , , nx - 1
                End If
                'Altera campos numéricos fracionários(strings)
                Case 0
'                    If InStr(cFields(nx), ",") > 0 Then
                        sSQL = Replace(cFields(nx), ".", "")
                        sSQL = Replace(sSQL, ",", ".")
                        cFields.Remove nx
                        cFields.Add sSQL, , , nx - 1
'                    End If
                'Altera campos numéricos fracionários
                Case 10
'                    If InStr(cFields(nx), ",") > 0 Then
                        sSQL = Replace(cFields(nx), ".", "")
                        sSQL = Replace(sSQL, ",", ".")
                        cFields.Remove nx
                        cFields.Add sSQL, , , nx - 1
'                    End If
            End Select
        End If
    
    Next nx
End If

Set Mov_Etiq_Critica = Mov_Etiq_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE DADOS DO Mov_Etiq
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Mov_Etiq_dados(ByVal cCONECCAO As Connection, _
                               ByVal cFields As Collection, _
                               ByVal nAction As Integer) As ADODB.Recordset
On Error GoTo Erro
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim sDate As String
Dim nseq As String
Dim bWere As Boolean
Dim bTemCampo As Boolean

Rem ************************************************************************************************************
Rem Ações da camada de negocios
Rem ************************************************************************************************************
Select Case nAction

Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 1 'Consultar Mov_Etiq(filtrando por campos)
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    bWere = False
    
    'id_bordero = "100 cod" e conf_palet <> 'E'
    
    sSQL = "SELECT Mov_Etiq.sequencia," & _
                  "Mov_Etiq.cod_peca," & _
                  "Mov_Etiq.Descr_Peca," & _
                  "Mov_Etiq.Cod_Cliente," & _
                  "Mov_Etiq.Lote," & _
                  "Mov_Etiq.Qtd_Caixa," & _
                  "Mov_Etiq.Data_Etiq," & _
                  "Mov_Etiq.tipo," & _
                  "Mov_Etiq.Cliente," & _
                  "Mov_Etiq.statusMov, " & _
                  "Mov_Etiq.indforjimport "

    sSQL = sSQL & "FROM Mov_Etiq "
    
    sSQL = sSQL & "INNER JOIN ETIQUETA ON Mov_Etiq.sequencia = ETIQUETA.ID_ETIQUETA "
    
    sSQL = sSQL & " WHERE (Mov_Etiq.conf_pallet = '' OR Mov_Etiq.conf_pallet IS NULL) ": bWere = True
    
    
    If Len(Trim(cFields(1))) > 0 Then ' SEQUENCIA
       sSQL = sSQL & " and Mov_Etiq.sequencia  = " & cFields(1) & "": bWere = True
    End If
    
    If Len(Trim(cFields(2))) > 0 Then ' data da digitacao
       If bWere Then
          sSQL = sSQL & " and Mov_Etiq.Data_Etiq = " & cFields(2)
       Else
          sSQL = sSQL & " where Mov_Etiq.Data_Etiq = " & cFields(2): bWere = True
       End If
    End If
   
    If Len(Trim(cFields(3))) > 0 Then ' Lote
       If bWere Then
          sSQL = sSQL & " and Mov_Etiq.Lote = " & cFields(3)
       Else
          sSQL = sSQL & " where Mov_Etiq.Lote = " & cFields(3): bWere = True
       End If
    End If
   
    If Len(Trim(cFields(4))) > 0 Then ' peca
       If bWere Then
          sSQL = sSQL & " and Mov_Etiq.cod_peca = " & cFields(4)
       Else
          sSQL = sSQL & " where Mov_Etiq.cod_peca = " & cFields(4): bWere = True
       End If
    End If
   
    If Len(Trim(cFields(5))) > 0 Then ' peca
       If bWere Then
          sSQL = sSQL & " and Mov_Etiq.Qtd_Caixa = " & cFields(5)
       Else
          sSQL = sSQL & " where Mov_Etiq.Qtd_Caixa = " & cFields(5)
       End If
    End If
   
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 2 'Incluir Mov_Etiq
Rem ************************************************************************************************************
Rem ************************************************************************************************************

    sSQL = "INSERT INTO Mov_Etiq (" & _
           "Cod_Peca, " & _
           "Lote, " & _
           "Qtd_Caixa, " & _
           "Data_Etiq, " & _
           "Cod_Tabela, " & _
           "Cliente, " & _
           "Cod_Cliente, " & _
           "Peso, " & _
           "Descr_Peca, " & _
           "Qtd_Etiq, " & _
           "Tipo, " & _
           "Data_Expedicao, " & _
           "Cod_Embalagem_Pw, " & _
           "Data_Lote, " & _
           "Classe_Funcional, " & _
           "Embarque_Controlado, " & _
           "Cod_Fornecedor, " & _
           "Num_Doc_Fiscal, " & _
           "Cod_Embalagem, " & _
           "Ind_Suplementar, " & _
           "Pto_Entrega, " & _
           "Dum, " & _
           "Qtd_Lote, "
    
    sSQL = sSQL & _
           "Vinculo, " & _
           "desvio, " & _
           "Serial, " & _
           "Cod_Util, " & _
           "Linha_Util, " & _
           "Modelo, " & _
           "Tipo_caixa, " & _
           "desvio_aviso_mod, " & _
           "tipo_alteracao, " & _
           "motivo_alteracao, " & _
           "motivo_alteracao_outros, " & _
           "envio_lote, " & _
           "num_am, " & _
           "Embalagem, " & _
           "id_cliente, " & _
           "tp_transporte, "
    
    sSQL = sSQL & _
           "codigo_produto1, " & _
           "qtde_caixa1, " & _
           "pecas_caixa1, " & _
           "compl_peca1, " & _
           "codigo_produto2, " & _
           "qtde_caixa2, " & _
           "pecas_caixa2, " & _
           "compl_peca2, " & _
           "codigo_produto3, " & _
           "qtde_caixa3, " & _
           "pecas_caixa3, " & _
           "compl_peca3, " & _
           "codigo_produto4, " & _
           "qtde_caixa4, " & _
           "pecas_caixa4, " & _
           "compl_peca4, " & _
           "sequencia, " & _
           "id_etiqueta, " & _
           "sequencia_org, " & _
           "statusMov, " & _
           "usuario_org, " & _
           "lgort, " & _
           "dt_alteracao, " & _
           "indforjimport, Sequencia_Dia "

    sSQL = sSQL & ") VALUES (" & _
                    cFields(1) & ", " & cFields(2) & ", " & _
                    cFields(3) & ", " & cFields(4) & ", " & _
                    cFields(5) & ", " & cFields(6) & ", " & _
                    cFields(7) & ", " & cFields(8) & ", " & _
                    cFields(9) & ", " & cFields(10) & ", " & _
                    cFields(11) & ", " & cFields(12) & ", " & _
                    cFields(13) & ", " & cFields(14) & ", " & _
                    cFields(15) & ", " & cFields(16) & ", " & _
                    cFields(17) & ", " & cFields(18) & ", " & _
                    cFields(19) & ", " & cFields(20) & ", " & _
                    cFields(21) & ", " & cFields(22) & ", "

     sSQL = sSQL & cFields(23) & ", " & cFields(24) & ", " & _
                    cFields(25) & ", " & cFields(26) & ", " & _
                    cFields(27) & ", " & cFields(28) & ", " & _
                    cFields(29) & ", " & cFields(30) & ", " & _
                    cFields(31) & ", " & cFields(32) & ", " & _
                    cFields(33) & ", " & cFields(34) & ", " & _
                    cFields(35) & ", " & cFields(36) & ", " & _
                    cFields(37) & ", " & cFields(38) & ", " & _
                    cFields(39) & ", " & cFields(40) & ", "

     sSQL = sSQL & cFields(41) & ", " & cFields(42) & ", " & _
                    cFields(43) & ", " & cFields(44) & ", " & _
                    cFields(45) & ", " & cFields(46) & ", " & _
                    cFields(47) & ", " & cFields(48) & ", " & _
                    cFields(49) & ", " & cFields(50) & ", " & _
                    cFields(51) & ", " & cFields(52) & ", " & _
                    cFields(53) & ", " & cFields(54) & ", " & _
                    cFields(55) & ", " & cFields(56) & ", " & _
                    cFields(57) & ", " & cFields(58) & ", " & _
                    cFields(59) & ", " & cFields(60) & ", " & _
                    cFields(61) & ", " & cFields(62) & ", " & _
                    cFields(63) & ", " & cFields(64) & ")"

    
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 11 'Consultar proxima sequencia livre
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    sSQL = "SELECT " & _
           "MAX(CONVERT(BIGINT, ID_ETIQUETA)) AS ULTIMA_ETIQUETA FROM ETIQUETA"

Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 12 'Consultar Mov_Etiq(filtrando pela sequencia)
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    sSQL = "SELECT Mov_Etiq.sequencia," & _
                  "Mov_Etiq.Cod_Peca, " & _
                  "Mov_Etiq.Lote, " & _
                  "Mov_Etiq.Qtd_Caixa, " & _
                  "Mov_Etiq.Data_Etiq, " & _
                  "Mov_Etiq.Cod_Tabela, " & _
                  "Mov_Etiq.Cliente, " & _
                  "Mov_Etiq.Cod_Cliente, " & _
                  "Mov_Etiq.Peso, " & _
                  "Mov_Etiq.Descr_Peca, " & _
                  "Mov_Etiq.Qtd_Etiq, " & _
                  "Mov_Etiq.Tipo, " & _
                  "Mov_Etiq.Data_Expedicao, " & _
                  "Mov_Etiq.Cod_Embalagem_Pw, " & _
                  "Mov_Etiq.Data_Lote, " & _
                  "Mov_Etiq.Classe_Funcional, " & _
                  "Mov_Etiq.Embarque_Controlado, " & _
                  "Mov_Etiq.Cod_Fornecedor, " & _
                  "Mov_Etiq.Num_Doc_Fiscal, " & _
                  "Mov_Etiq.Cod_Embalagem, " & _
                  "Mov_Etiq.Ind_Suplementar, " & _
                  "Mov_Etiq.Pto_Entrega, " & _
                  "Mov_Etiq.Dum, " & _
                  "Mov_Etiq.Qtd_Lote, "
    sSQL = sSQL & _
                  "Mov_Etiq.Vinculo, " & _
                  "Mov_Etiq.desvio, " & _
                  "Mov_Etiq.Serial, " & _
                  "Mov_Etiq.Cod_Util, " & _
                  "Mov_Etiq.Linha_Util, " & _
                  "Mov_Etiq.Modelo, " & _
                  "Mov_Etiq.Tipo_caixa, " & _
                  "Mov_Etiq.desvio_aviso_mod, " & _
                  "Mov_Etiq.tipo_alteracao, " & _
                  "Mov_Etiq.motivo_alteracao, " & _
                  "Mov_Etiq.motivo_alteracao_outros, " & _
                  "Mov_Etiq.envio_lote, " & _
                  "Mov_Etiq.num_am, " & _
                  "Mov_Etiq.Embalagem, " & _
                  "Mov_Etiq.id_cliente, " & _
                  "Mov_Etiq.tp_transporte, "
    
    sSQL = sSQL & _
                  "Mov_Etiq.codigo_produto1, " & _
                  "Mov_Etiq.qtde_caixa1, " & _
                  "Mov_Etiq.pecas_caixa1, " & _
                  "Mov_Etiq.compl_peca1, " & _
                  "Mov_Etiq.codigo_produto2, " & _
                  "Mov_Etiq.qtde_caixa2, " & _
                  "Mov_Etiq.pecas_caixa2, " & _
                  "Mov_Etiq.compl_peca2, " & _
                  "Mov_Etiq.codigo_produto3, " & _
                  "Mov_Etiq.qtde_caixa3, " & _
                  "Mov_Etiq.pecas_caixa3, " & _
                  "Mov_Etiq.compl_peca3, " & _
                  "Mov_Etiq.codigo_produto4, " & _
                  "Mov_Etiq.qtde_caixa4, " & _
                  "Mov_Etiq.pecas_caixa4, " & _
                  "Mov_Etiq.compl_peca4, " & _
                  "Mov_Etiq.sequencia, " & _
                  "Mov_Etiq.id_etiqueta, " & _
                  "Mov_Etiq.sequencia_org, " & _
                  "Mov_Etiq.statusMov, " & _
                  "Mov_Etiq.usuario_org, " & _
                  "Mov_Etiq.lgort, " & _
                  "Mov_Etiq.dt_alteracao, " & _
                  "Mov_Etiq.indforjimport,Mov_etiq.Sequencia_Dia "
                  
    sSQL = sSQL & "FROM Mov_Etiq "
    
    sSQL = sSQL & "INNER JOIN ETIQUETA ON Mov_Etiq.sequencia = ETIQUETA.ID_ETIQUETA "
''''    sSQL = sSQL & "INNER JOIN cont_pallet ON etiqueta.id_etiqueta = cont_pallet.ID_ETIQUETA AND etiqueta.pallet = cont_pallet.pallet "
    
    Rem retirar depois de salvar aqui marcos marcos  31/01/2024
    sSQL = sSQL & " WHERE (ETIQUETA.xblnr = '' or ETIQUETA.xblnr IS NULL) "
    
    
    If Len(Trim(cFields(1))) > 0 Then ' SEQUENCIA
       sSQL = sSQL & " and Mov_Etiq.sequencia  = " & cFields(1)
    End If

''''    sSQL = sSQL & " ORDER BY CONT_PALLET.ID ASC "

Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 13 'Alterar campos status, da tabela movetiq,
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    bTemCampo = False
    
    sSQL = "UPDATE MOV_ETIQ SET "
    
    If Len(Trim(cFields(2))) > 0 Then
       sSQL = sSQL & "statusMov = " & cFields(2): bTemCampo = True      'STATUS DO REGISTRO
    End If
    
    If Len(Trim(cFields(3))) > 0 Then ' sequencia_org 'sequencia original, caso seja desmembrada
       If bTemCampo Then
          sSQL = sSQL & ",sequencia_org = " & cFields(3)
       Else
          sSQL = sSQL & "sequencia_org = " & cFields(3): bTemCampo = True
       End If
    End If
    
    If Len(Trim(cFields(4))) > 0 Then ' usuario_org 'usuario, que fez o desmembramento
       If bTemCampo Then
          sSQL = sSQL & ",usuario_org = " & cFields(4)
       Else
          sSQL = sSQL & "usuario_org = " & cFields(4): bTemCampo = True
       End If
    End If
    
    sSQL = sSQL & _
           " WHERE sequencia = " & cFields(1)
           
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 15 'FILTRAR AS ETIQUETAS PELA SEQUENCIA QUE FOI GERADAS POR UMA OUTRA, A ORIGINAL, SEM AS EXCLUIDAS
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    sSQL = "SELECT * FROM MOV_ETIQ " & _
           " WHERE SEQUENCIA_ORG  = " & cFields(1) & _
           " AND   statusMov <> 3"

Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 16 'FILTRAR as qtde de caixas,qtde e notas fiscais de um determinado pallet.
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    sSQL = "SELECT " & _
           "XBLNR, " & _
           "COUNT(*) AS QTDE, " & _
           "SUM(PESO) AS PESO "
           
    sSQL = sSQL & _
           "FROM ETIQUETA "
           
    sSQL = sSQL & _
           "WHERE PALLET = '" & cFields(1) & "'" & _
           "  AND ID_CLIENTE = '0000000007' GROUP BY XBLNR"


Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 17 'Consultar Mov_Etiq(filtrando-as para fazer ajustes) - 18-08-2015
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    bTemCampo = False
    
    sSQL = "SELECT Mov_Etiq.sequencia," & _
                  "Mov_Etiq.cod_peca," & _
                  "Mov_Etiq.Descr_Peca," & _
                  "Mov_Etiq.Cod_Cliente," & _
                  "Mov_Etiq.Lote," & _
                  "Mov_Etiq.Qtd_Caixa," & _
                  "Mov_Etiq.Data_Etiq," & _
                  "Mov_Etiq.tipo," & _
                  "Mov_Etiq.Cliente," & _
                  "Mov_Etiq.statusMov, " & _
                  "Mov_Etiq.indforjimport "

    sSQL = sSQL & "FROM Mov_Etiq "
    
    sSQL = sSQL & "INNER JOIN ETIQUETA ON Mov_Etiq.sequencia = ETIQUETA.ID_ETIQUETA "
    
    sSQL = sSQL & " WHERE ETIQUETA.conf_pallet <> 'E' " & _
                  " AND   ETIQUETA.id_bordero = 'SEM COD' " & _
                  " AND   ETIQUETA.SEQUENCIA_PLACA = 0 " & _
                  " AND   ETIQUETA.TIPO_EMBALAGEM <> 'OK' "
                  
    If Len(Trim(cFields(1))) > 0 Then ' SEQUENCIA
       sSQL = sSQL & " AND Mov_Etiq.sequencia  = " & cFields(1)
    End If
    
    If Len(Trim(cFields(2))) > 0 Then ' peca
       sSQL = sSQL & " AND Mov_Etiq.cod_peca = " & cFields(2)
    End If
    
    sSQL = sSQL & " ORDER BY ETIQUETA.ID_PECA"
    
    'id_bordero = "100 cod" e conf_palet <> 'E'
    ' DEPOIS ATUALIZAR O CAMPO TIPO_EMBALAGEM
    
Rem ************************************************************************************************************
Rem ************************************************************************************************************
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 18 'Alterar TIPO_EMBALAGEM status, da tabela etiqueta referente aos ajustes efetuados,
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    
    sSQL = "UPDATE ETIQUETA SET "
    sSQL = sSQL & " ETIQUETA.TIPO_EMBALAGEM = 'OK'"
    sSQL = sSQL & " WHERE ETIQUETA.ID_ETIQUETA = " & cFields(1)

Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 19 'Consultar Mov_Etiq(filtrando pela sequencia)
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    sSQL = "SELECT Mov_Etiq.sequencia," & _
                  "Mov_Etiq.Cod_Peca, " & _
                  "Mov_Etiq.Lote, " & _
                  "Mov_Etiq.Qtd_Caixa, " & _
                  "Mov_Etiq.Data_Etiq, " & _
                  "Mov_Etiq.Cod_Tabela, " & _
                  "Mov_Etiq.Cliente, " & _
                  "Mov_Etiq.Cod_Cliente, " & _
                  "Mov_Etiq.Peso, " & _
                  "Mov_Etiq.Descr_Peca, " & _
                  "Mov_Etiq.Qtd_Etiq, " & _
                  "Mov_Etiq.Tipo, " & _
                  "Mov_Etiq.Data_Expedicao, " & _
                  "Mov_Etiq.Cod_Embalagem_Pw, " & _
                  "Mov_Etiq.Data_Lote, " & _
                  "Mov_Etiq.Classe_Funcional, " & _
                  "Mov_Etiq.Embarque_Controlado, " & _
                  "Mov_Etiq.Cod_Fornecedor, " & _
                  "Mov_Etiq.Num_Doc_Fiscal, " & _
                  "Mov_Etiq.Cod_Embalagem, " & _
                  "Mov_Etiq.Ind_Suplementar, " & _
                  "Mov_Etiq.Pto_Entrega, " & _
                  "Mov_Etiq.Dum, " & _
                  "Mov_Etiq.Qtd_Lote, "
    sSQL = sSQL & _
                  "Mov_Etiq.Vinculo, " & _
                  "Mov_Etiq.desvio, " & _
                  "Mov_Etiq.Serial, " & _
                  "Mov_Etiq.Cod_Util, " & _
                  "Mov_Etiq.Linha_Util, " & _
                  "Mov_Etiq.Modelo, " & _
                  "Mov_Etiq.Tipo_caixa, " & _
                  "Mov_Etiq.desvio_aviso_mod, " & _
                  "Mov_Etiq.tipo_alteracao, " & _
                  "Mov_Etiq.motivo_alteracao, " & _
                  "Mov_Etiq.motivo_alteracao_outros, " & _
                  "Mov_Etiq.envio_lote, " & _
                  "Mov_Etiq.num_am, " & _
                  "Mov_Etiq.Embalagem, " & _
                  "Mov_Etiq.id_cliente, " & _
                  "Mov_Etiq.tp_transporte, "
    
    sSQL = sSQL & _
                  "Mov_Etiq.codigo_produto1, " & _
                  "Mov_Etiq.qtde_caixa1, " & _
                  "Mov_Etiq.pecas_caixa1, " & _
                  "Mov_Etiq.compl_peca1, " & _
                  "Mov_Etiq.codigo_produto2, " & _
                  "Mov_Etiq.qtde_caixa2, " & _
                  "Mov_Etiq.pecas_caixa2, " & _
                  "Mov_Etiq.compl_peca2, " & _
                  "Mov_Etiq.codigo_produto3, " & _
                  "Mov_Etiq.qtde_caixa3, " & _
                  "Mov_Etiq.pecas_caixa3, " & _
                  "Mov_Etiq.compl_peca3, " & _
                  "Mov_Etiq.codigo_produto4, " & _
                  "Mov_Etiq.qtde_caixa4, " & _
                  "Mov_Etiq.pecas_caixa4, " & _
                  "Mov_Etiq.compl_peca4, " & _
                  "Mov_Etiq.sequencia, " & _
                  "Mov_Etiq.id_etiqueta, " & _
                  "Mov_Etiq.sequencia_org, " & _
                  "Mov_Etiq.statusMov, " & _
                  "Mov_Etiq.usuario_org, " & _
                  "Mov_Etiq.lgort, " & _
                  "Mov_Etiq.dt_alteracao, " & _
                  "Mov_Etiq.indforjimport, Mov_Etiq.Sequencia_Dia "
                  
'    sSQL = sSQL & "ETIQUETA.DATA_FATURAMENTO "
    
    sSQL = sSQL & "FROM Mov_Etiq "
    
    sSQL = sSQL & "INNER JOIN ETIQUETA ON Mov_Etiq.sequencia = ETIQUETA.ID_ETIQUETA "
    
    sSQL = sSQL & " WHERE ETIQUETA.conf_pallet <> 'E' " & _
                  " AND   ETIQUETA.id_bordero = 'SEM COD' " & _
                  " AND   ETIQUETA.SEQUENCIA_PLACA = 0 " & _
                  " AND   ETIQUETA.TIPO_EMBALAGEM <> 'OK' "
'    sSQL = sSQL & " WHERE ETIQUETA.xblnr = '' "
    
    
    If Len(Trim(cFields(1))) > 0 Then ' SEQUENCIA
       sSQL = sSQL & " and Mov_Etiq.sequencia  = " & cFields(1)
    End If
       
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 20 'Consultar Mov_Etiq(filtrando-as para fazer ajustes nas etiquetas de de embarque) - 18-08-2015
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    bTemCampo = False
    
    sSQL = "SELECT Mov_Etiq.sequencia," & _
                  "Mov_Etiq.cod_peca," & _
                  "Mov_Etiq.Descr_Peca," & _
                  "Mov_Etiq.Cod_Cliente," & _
                  "Mov_Etiq.Lote," & _
                  "Mov_Etiq.Qtd_Caixa," & _
                  "Mov_Etiq.Data_Etiq," & _
                  "Mov_Etiq.tipo," & _
                  "Mov_Etiq.Cliente," & _
                  "Mov_Etiq.statusMov, " & _
                  "Mov_Etiq.indforjimport "

    sSQL = sSQL & "FROM Mov_Etiq "
    
    sSQL = sSQL & "INNER JOIN ETIQUETA ON Mov_Etiq.sequencia = ETIQUETA.ID_ETIQUETA "
    
    sSQL = sSQL & " WHERE ETIQUETA.conf_pallet = 'E' " & _
                  " AND   ETIQUETA.id_bordero = 'SEM COD' " & _
                  " AND   ETIQUETA.SEQUENCIA_PLACA = 0 " & _
                  " AND   ETIQUETA.TIPO_EMBALAGEM <> 'OK' "
                  
    If Len(Trim(cFields(1))) > 0 Then ' SEQUENCIA
       sSQL = sSQL & " AND Mov_Etiq.sequencia  = " & cFields(1)
    End If
    
    If Len(Trim(cFields(2))) > 0 Then ' peca
       sSQL = sSQL & " AND Mov_Etiq.cod_peca = " & cFields(2)
    End If
    
    sSQL = sSQL & " ORDER BY ETIQUETA.ID_PECA"

Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 22 'Consultar Mov_Etiq(filtrando-as para fazer ajustes) - 26-08-2015
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    bTemCampo = False
    
    sSQL = "SELECT Mov_Etiq.sequencia," & _
                  "Mov_Etiq.cod_peca," & _
                  "Mov_Etiq.Descr_Peca," & _
                  "Mov_Etiq.Cod_Cliente," & _
                  "Mov_Etiq.Lote," & _
                  "Mov_Etiq.Qtd_Caixa," & _
                  "Mov_Etiq.Data_Etiq," & _
                  "Mov_Etiq.tipo," & _
                  "Mov_Etiq.Cliente," & _
                  "Mov_Etiq.statusMov, " & _
                  "Mov_Etiq.indforjimport, " & _
                  "ETIQUETA.xblnr, " & _
                  "ETIQUETA.pallet, " & _
                  "ETIQUETA.placa, " & _
                  "Mov_Etiq.Qtd_Etiq "

    sSQL = sSQL & "FROM Mov_Etiq "
    
    sSQL = sSQL & "INNER JOIN ETIQUETA ON Mov_Etiq.sequencia = ETIQUETA.ID_ETIQUETA "
    
    sSQL = sSQL & " WHERE ETIQUETA.DATA_CRIACAO >= '20150801' " & _
         " and ETIQUETA.ID_BORDERO = 'Sem Cod'" & _
         " and ETIQUETA.REMESSA <> 0" & _
         " and ETIQUETA.sequencia_placa not in (0,'',150487,150492,150503,150507,150488) " & _
         " and ETIQUETA.conf_pallet <> 'F' " & _
         " and ETIQUETA.ID_CLIENTE in ('0000000001','0000000002') "
                  
    If Len(Trim(cFields(1))) > 0 Then ' SEQUENCIA
       sSQL = sSQL & " AND Mov_Etiq.sequencia  = " & cFields(1)
    End If
    
    If Len(Trim(cFields(2))) > 0 Then ' peca
       sSQL = sSQL & " AND Mov_Etiq.cod_peca = " & cFields(2)
    End If
    
    sSQL = sSQL & " ORDER BY ETIQUETA.sequencia_placa,ETIQUETA.pallet "

Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 23 'Consultar Mov_Etiq(filtrando-as para fazer ajustes nas etiquetas de de embarque) - 18-08-2015
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    bTemCampo = False
    
    sSQL = "SELECT Mov_Etiq.sequencia," & _
                  "Mov_Etiq.cod_peca," & _
                  "Mov_Etiq.Descr_Peca," & _
                  "Mov_Etiq.Cod_Cliente," & _
                  "Mov_Etiq.Lote," & _
                  "Mov_Etiq.Qtd_Caixa," & _
                  "Mov_Etiq.Data_Etiq," & _
                  "Mov_Etiq.tipo," & _
                  "Mov_Etiq.Cliente," & _
                  "Mov_Etiq.statusMov, " & _
                  "Mov_Etiq.indforjimport, " & _
                  "ETIQUETA.xblnr, " & _
                  "ETIQUETA.pallet, " & _
                  "ETIQUETA.placa, " & _
                  "Mov_Etiq.Qtd_Etiq, " & _
                  "Mov_Etiq.Peso, " & _
                  "ETIQUETA.ID_ETIQUETA "

    sSQL = sSQL & "FROM Mov_Etiq "
    
    sSQL = sSQL & "INNER JOIN ETIQUETA ON Mov_Etiq.sequencia = ETIQUETA.ID_ETIQUETA "
    
    sSQL = sSQL & " WHERE ETIQUETA.conf_pallet = 'E' "
    
    If Len(Trim(cFields(1))) > 0 Then ' SEQUENCIA
       sSQL = sSQL & " AND Mov_Etiq.sequencia  = " & cFields(1)
    End If
    
'    If Len(Trim(cFields(2))) > 0 Then ' peca
'       sSQL = sSQL & " AND Mov_Etiq.cod_peca = " & cFields(2)
'    End If
    
    sSQL = sSQL & " ORDER BY ETIQUETA.ID_PECA"

Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 24 'Consultar Mov_Etiq(filtrando as etiquetas para a FIAT) 13-05-2016
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    bTemCampo = False

    sSQL = "SELECT Mov_etiq.lote AS Num_Lote, " & _
           "ETIQUETA.QTDE AS Qtde_Emb, " & _
           "classe_funcional AS Classe_Func, " & _
           "ind_suplementar AS Indicacao_Supl, " & _
           "data_lote AS Data_Fab_Lote,  " & _
           "Mov_Etiq.Cod_embalagem_Pw AS Cod_Emb, " & _
           "Vinculo, " & _
           "desvio as Lote_Sob_Desv, " & _
           "Qtd_lote as Qtde_Lote, " & _
           "Mov_etiq.Motivo_alteracao_outros as Aplicacao, " & _
           "DUM, " & _
           "embarque_controlado as Embarque_Ctrl, " & _
           "'13093' as Cod_Fornecedor, " & _
           "ETIQUETA.xblnr as Num_Doc_Fis_BAM, " & _
           " CONVERT(VARCHAR(11),GETDATE(),103) as Data, " & _
           "Mov_etiq.Pto_entrega as Ponto_Entrega, " & _
           "Mov_etiq.descr_Peca as Denominacao, " & _
           "ETIQUETA.cod_no_cliente as Num_Desenho, " & _
           "ETIQUETA.id_peca as Ctrl_Interno, " & _
           "ETIQUETA.id_etiqueta as Ctrl_Oper_Log, " & _
           "' ' as Codigo_Numero, " & _
           "' ' as codigo_barras, " & _
           "ETIQUETA.pallet, " & _
           "ETIQUETA.xblnr "

           
    sSQL = sSQL & "FROM Mov_Etiq "
    
    sSQL = sSQL & "INNER JOIN ETIQUETA ON Mov_Etiq.sequencia = ETIQUETA.ID_ETIQUETA "
    
    If Trim(cFields(2)) = "1" Then ' pallet
       sSQL = sSQL & " WHERE ETIQUETA.pallet  = '" & cFields(1) & "'"
    Else 'etiqueta
       sSQL = sSQL & " WHERE ETIQUETA.id_etiqueta  = '" & cFields(1) & "'"
    End If

    sSQL = sSQL & " ORDER BY ETIQUETA.ID_PECA"
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 25 'Consultar Mov_Etiq(filtrando a etiqueta especifica para a FIAT) 13-05-2016
Rem ************************************************************************************************************
Rem ************************************************************************************************************

    sSQL = "SELECT Mov_etiq.lote AS Num_Lote, " & _
           "ETIQUETA.QTDE AS Qtde_Emb, " & _
           "classe_funcional AS Classe_Func, " & _
           "ind_suplementar AS Indicacao_Supl, " & _
           "data_lote AS Data_Fab_Lote,  " & _
           "Mov_Etiq.Cod_embalagem_Pw AS Cod_Emb, " & _
           "Vinculo, " & _
           "desvio as Lote_Sob_Desv, " & _
           "Qtd_lote as Qtde_Lote, " & _
           "Mov_etiq.Motivo_alteracao_outros as Aplicacao, " & _
           "DUM, " & _
           "embarque_controlado as Embarque_Ctrl, " & _
           "'13093' as Cod_Fornecedor, " & _
           "ETIQUETA.xblnr as Num_Doc_Fis_BAM, " & _
           " CONVERT(VARCHAR(11),GETDATE(),103) as Data, " & _
           "Mov_etiq.Pto_entrega as Ponto_Entrega, " & _
           "Mov_etiq.descr_Peca as Denominacao, " & _
           "ETIQUETA.cod_no_cliente as Num_Desenho, " & _
           "ETIQUETA.id_peca as Ctrl_Interno, " & _
           "ETIQUETA.id_etiqueta as Ctrl_Oper_Log, " & _
           "' ' as Codigo_Numero, " & _
           "' ' as codigo_barras, " & _
           "ETIQUETA.pallet, Mov_etiq.embalagem "
           
    sSQL = sSQL & "FROM Mov_Etiq "
    
    sSQL = sSQL & "INNER JOIN ETIQUETA ON Mov_Etiq.sequencia = ETIQUETA.ID_ETIQUETA "
    
    sSQL = sSQL & " WHERE MOV_ETIQ.ID_ETIQUETA = '" & Trim(cFields(1)) & "'"
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 26 'Consultar Mov_Etiq(filtrando as etiquetas para a KRUUP) 12-08-2016
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    bTemCampo = False

    sSQL = "SELECT ETIQUETA.lote AS Num_Lote, " & _
           "ETIQUETA.QTDE AS Qtde_Emb, " & _
           "classe_funcional AS Classe_Func, " & _
           "ind_suplementar AS Indicacao_Supl, " & _
           "data_lote AS Data_Fab_Lote,  " & _
           "Mov_Etiq.embalagem AS Cod_Emb, " & _
           "Vinculo, " & _
           "desvio as Lote_Sob_Desv, " & _
           "Qtd_lote as Qtde_Lote, " & _
           "Mov_etiq.Motivo_alteracao_outros as Aplicacao, " & _
           "DUM, " & _
           "embarque_controlado as Embarque_Ctrl, " & _
           "'13093' as Cod_Fornecedor, " & _
           "ETIQUETA.xblnr as Num_Doc_Fis_BAM, " & _
           "CONVERT(VARCHAR(11),GETDATE(),103) as Data, " & _
           "Mov_etiq.Pto_entrega as Ponto_Entrega, " & _
           "Mov_etiq.descr_Peca as Denominacao, " & _
           "ETIQUETA.cod_no_cliente as Num_Desenho, " & _
           "ETIQUETA.id_peca as Ctrl_Interno, " & _
           "ETIQUETA.id_etiqueta as Ctrl_Oper_Log, " & _
           "' ' as Codigo_Numero, " & _
           "' ' as codigo_barras, " & _
           "ETIQUETA.pallet, " & _
           "ETIQUETA.xblnr, " & _
           "Mov_Etiq.cod_peca, "
    sSQL = sSQL & _
           "mov_etiq.Num_Doc_Fiscal "

    sSQL = sSQL & "FROM Mov_Etiq "
    
    sSQL = sSQL & "INNER JOIN ETIQUETA ON Mov_Etiq.sequencia = ETIQUETA.ID_ETIQUETA "
    
    sSQL = sSQL & " AND ETIQUETA.ID_CLIENTE  = '0000000009' "
    
    If Trim(cFields(2)) = "1" Then ' pallet
       sSQL = sSQL & " AND ETIQUETA.pallet  = '" & cFields(1) & "'"
    Else 'etiqueta
       sSQL = sSQL & " AND ETIQUETA.id_etiqueta  = '" & cFields(1) & "'"
    End If

    sSQL = sSQL & " ORDER BY ETIQUETA.ID_PECA"
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 27 'Consultar Mov_Etiq(filtrando as etiquetas para a FIAT LATAM) 07-11-2016
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    bTemCampo = False

    sSQL = "SELECT " & _
           "ETIQUETA.peso as Peso_Bruto, " & _
           "Mov_etiq.compl_peca4 AS ODM, " & _
           "Mov_etiq.data_lote as Data_Producao, " & _
           "' ' AS Data_validade,  " & _
           "CONVERT(VARCHAR(11),GETDATE(),103) AS Data_Expedicao,  " & _
           "Mov_Etiq.Cod_embalagem_Pw AS Cod_Emb, " & _
           "ETIQUETA.QTDE AS Quantidade, " & _
           "'09' AS DOCA, " & _
           "SUBSTRING(Mov_etiq.Pto_entrega,1,3) as Ponto_Entrega, " & _
           "' ' AS Control_Log_Qual,  " & _
           "'13093' as Cod_Fornecedor, " & _
           "ETIQUETA.xblnr AS Num_Doc_Fiscal, " & _
           "Mov_Etiq.desvio as Lote_Sob_Desv, " & _
           "1 as Qtde_emb, " & _
           "'200055' AS Num_Sheda_Serial,  " & _
           "ETIQUETA.id_peca as Id_Inter_Fornecedor, " & _
           "Mov_Etiq.embarque_controlado as Embarque_Ctrl, " & _
           "Mov_Etiq.ind_suplementar AS Indicacao_Supl, " & _
           "Mov_Etiq.classe_funcional AS Classe_Func, "
      
      sSQL = sSQL & _
           "Mov_Etiq.compl_peca3 AS Dados_Transporte, " & _
           "Mov_Etiq.Qtd_lote as Qtde_Lote, " & _
           "ETIQUETA.lote AS Num_Lote, " & _
           "'MUSASHI DO BRASIL LTDA' AS Razao_Social, " & _
           "' ' AS Cod_Barras, " & _
           "' ' AS Desenho_Chrysler, "
      
      sSQL = sSQL & _
           "Mov_etiq.descr_Peca as Descricao_Produto, " & _
           "ETIQUETA.cod_no_cliente as Num_Desenho, " & _
           "'1046' AS Destino, " & _
           "' ' AS Cod_Destino, " & _
           "Mov_Etiq.Vinculo, " & _
           "Mov_Etiq.compl_peca2 AS Restricoes, " & _
           "' ' AS QR_Codes, " & _
           "' ' AS LogoMarca, " & _
           "Mov_Etiq.DUM, " & _
           "' ' AS PDF_417, " & _
           "ETIQUETA.Tipo_Frete AS Incoterms, " & _
           "ETIQUETA.ID_ETIQUETA, " & _
           "ETIQUETA.ID_CLIENTE "
           
    sSQL = sSQL & "FROM Mov_Etiq "
    
    sSQL = sSQL & "INNER JOIN ETIQUETA ON Mov_Etiq.sequencia = ETIQUETA.ID_ETIQUETA "
'    sSQL = sSQL & " WHERE ETIQUETA.pallet  = '04112016*5'"
'    sSQL = sSQL & " WHERE ETIQUETA.xblnr  = '369312'"
'"
    If Trim(cFields(2)) = "1" Then ' pallet
       sSQL = sSQL & " WHERE ETIQUETA.pallet  = '" & cFields(1) & "'"
    Else 'etiqueta
       sSQL = sSQL & " WHERE ETIQUETA.id_etiqueta  = " & cFields(1) & ""
    End If

    sSQL = sSQL & " ORDER BY ETIQUETA.ID_PECA"
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 28 'Consultar Mov_Etiq(filtrando a etiqueta especifica para a FIAT LATAM) 07-11-2016
Rem ************************************************************************************************************
Rem ************************************************************************************************************

    sSQL = "SELECT Mov_etiq.lote AS Num_Lote, " & _
           "ETIQUETA.QTDE AS Qtde_Emb, " & _
           "classe_funcional AS Classe_Func, " & _
           "ind_suplementar AS Indicacao_Supl, " & _
           "data_lote AS Data_Fab_Lote,  " & _
           "Mov_Etiq.Cod_embalagem_Pw AS Cod_Emb, " & _
           "Vinculo, " & _
           "desvio as Lote_Sob_Desv, " & _
           "Qtd_lote as Qtde_Lote, " & _
           "Mov_etiq.Motivo_alteracao_outros as Aplicacao, " & _
           "DUM, " & _
           "embarque_controlado as Embarque_Ctrl, " & _
           "'13093' as Cod_Fornecedor, " & _
           "ETIQUETA.xblnr as Num_Doc_Fis_BAM, " & _
           " CONVERT(VARCHAR(11),GETDATE(),103) as Data, " & _
           "Mov_etiq.Pto_entrega as Ponto_Entrega, " & _
           "Mov_etiq.descr_Peca as Denominacao, " & _
           "ETIQUETA.cod_no_cliente as Num_Desenho, " & _
           "ETIQUETA.id_peca as Ctrl_Interno, " & _
           "ETIQUETA.id_etiqueta as Ctrl_Oper_Log, " & _
           "' ' as Codigo_Numero, " & _
           "' ' as codigo_barras, " & _
           "ETIQUETA.pallet "
           
    sSQL = sSQL & "FROM Mov_Etiq "
    
    sSQL = sSQL & "INNER JOIN ETIQUETA ON Mov_Etiq.sequencia = ETIQUETA.ID_ETIQUETA "
    
    sSQL = sSQL & " WHERE MOV_ETIQ.ID_ETIQUETA = '" & Trim(cFields(1)) & "'"

Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 29 'Consultar Mov_Etiq(filtrando as etiquetas para a HONDA QRCODE) 14/05/2019
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Rem AQUI MARCOS 14/05/2019, AJUSTAR OS DADOS PARA A HONDA
    bTemCampo = False

    sSQL = "SELECT " & _
           "ROW_NUMBER() OVER(ORDER BY  ETIQUETA.id_etiqueta ASC) AS seq, " & _
           "ETIQUETA.id_peca, " & _
           "mov_etiq.Descr_Peca, " & _
           "etiqueta.qtde, " & _
           "Etiqueta.Lote as num_lote, " & _
           "ETIQUETA.ID_ETIQUETA AS id, " & _
           "'NULO' as tipo_embalagem, " & _
           "etiqueta.xblnr as num_doc_fiscal, " & _
           "etiqueta.id_cliente, " & _
           "'MUSASHI DO BRASIL LTDA' as fornecedor, " & _
           "ETIQUETA.cod_no_cliente, " & _
           "mov_etiq.cod_util, " & _
           "CONVERT(VARCHAR(11),GETDATE(),103) as data_sistema, " & _
           "(SELECT COUNT(*) FROM ETIQUETA WHERE ETIQUETA.xblnr  = '" & cFields(1) & "') AS total_volume, " & _
           "EMBALAGEM , " & _
           "DATA_ETIQ, " & _
           "ETIQUETA.Tipo_caixa, " & _
           "ETIQUETA.DATA_FATURAMENTO "
           
    sSQL = sSQL & " FROM Mov_Etiq "
    
    sSQL = sSQL & " INNER JOIN ETIQUETA ON Mov_Etiq.sequencia = ETIQUETA.ID_ETIQUETA "
    
    If Trim(cFields(2)) = "1" Then ' Nota fiscal
       sSQL = sSQL & " WHERE ETIQUETA.xblnr  = '" & cFields(1) & "'"
    Else 'etiqueta
       sSQL = sSQL & " WHERE ETIQUETA.id_etiqueta  = " & cFields(1) & ""
    End If

    sSQL = sSQL & " ORDER BY ETIQUETA.DATA_FATURAMENTO"

Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 291 'Consultar Mov_Etiq(filtrando as etiquetas para a HONDA QRCODE E PALLET) 07/08/2024
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Rem AQUI MARCOS 14/05/2019, AJUSTAR OS DADOS PARA A HONDA
    bTemCampo = False

    sSQL = "SELECT " & _
           "ROW_NUMBER() OVER(ORDER BY  ETIQUETA.id_etiqueta ASC) AS seq, " & _
           "ETIQUETA.id_peca, " & _
           "mov_etiq.Descr_Peca, " & _
           "etiqueta.qtde, " & _
           "Etiqueta.Lote as num_lote, " & _
           "ETIQUETA.ID_ETIQUETA AS id, " & _
           "'NULO' as tipo_embalagem, " & _
           "etiqueta.xblnr as num_doc_fiscal, " & _
           "etiqueta.id_cliente, " & _
           "'MUSASHI DO BRASIL LTDA' as fornecedor, " & _
           "ETIQUETA.cod_no_cliente, " & _
           "mov_etiq.cod_util, " & _
           "CONVERT(VARCHAR(11),GETDATE(),103) as data_sistema, " & _
           "(SELECT COUNT(*) FROM ETIQUETA WHERE ETIQUETA.xblnr  = '" & cFields(1) & "') AS total_volume, " & _
           "EMBALAGEM , " & _
           "DATA_ETIQ, " & _
           "ETIQUETA.Tipo_caixa, " & _
           "ETIQUETA.DATA_FATURAMENTO, " & _
           "ETIQUETA.pallet "
           
    sSQL = sSQL & " FROM Mov_Etiq "
    
    sSQL = sSQL & " INNER JOIN ETIQUETA ON Mov_Etiq.sequencia = ETIQUETA.ID_ETIQUETA "
    
'    If Trim(cFields(1)) = "1" Then ' pallet
'       If Len(Trim(cFields(3))) > 0 Then ' pallet INTERVALO
'          sSQL = sSQL & " WHERE ETIQUETA.pallet BETWEEN  '" & cFields(2) & "'"
'          sSQL = sSQL & " AND   '" & cFields(3) & "'"
'       Else
'          sSQL = sSQL & " WHERE ETIQUETA.pallet  = '" & cFields(3) & "'"
'       End If
'    End If

    sSQL = sSQL & " WHERE ETIQUETA.xblnr  = '" & cFields(2) & "'"
    sSQL = sSQL & " AND   ETIQUETA.pallet  = '" & cFields(3) & "'"
    sSQL = sSQL & " ORDER BY ETIQUETA.DATA_FATURAMENTO"

Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 30 'Consultar Mov_Etiq(filtrando a etiqueta especifica para a GM) 05/10/2020
Rem ************************************************************************************************************
Rem ************************************************************************************************************

    sSQL = "SELECT Mov_etiq.lote AS Num_Lote, " & _
           "ETIQUETA.QTDE AS Qtde_Emb, " & _
           "classe_funcional AS Classe_Func, " & _
           "ind_suplementar AS Indicacao_Supl, " & _
           "data_lote AS Data_Fab_Lote,  " & _
           "Mov_Etiq.Cod_embalagem_Pw AS Cod_Emb, " & _
           "Vinculo, " & _
           "desvio as Lote_Sob_Desv, " & _
           "Qtd_lote as Qtde_Lote, " & _
           "Mov_etiq.Motivo_alteracao_outros as Aplicacao, " & _
           "DUM, " & _
           "embarque_controlado as Embarque_Ctrl, " & _
           "'13093' as Cod_Fornecedor, " & _
           "ETIQUETA.xblnr as Num_Doc_Fis_BAM, " & _
           " CONVERT(VARCHAR(11),GETDATE(),103) as Data, " & _
           "Mov_etiq.Pto_entrega as Ponto_Entrega, " & _
           "Mov_etiq.descr_Peca as Denominacao, " & _
           "ETIQUETA.cod_no_cliente as Num_Desenho, " & _
           "ETIQUETA.id_peca as Ctrl_Interno, " & _
           "ETIQUETA.id_etiqueta as Ctrl_Oper_Log, " & _
           "' ' as Codigo_Numero, " & _
           "' ' as codigo_barras, " & _
           "ETIQUETA.pallet, Mov_etiq.embalagem "
           
    sSQL = sSQL & "FROM Mov_Etiq "
    
    sSQL = sSQL & "INNER JOIN ETIQUETA ON Mov_Etiq.sequencia = ETIQUETA.ID_ETIQUETA "
    
    sSQL = sSQL & " WHERE MOV_ETIQ.ID_ETIQUETA = '" & Trim(cFields(1)) & "'"
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 31 'Consultar Mov_Etiq(filtrando as etiquetas para a GM MASTER) 10-11-2020
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    bTemCampo = False

    sSQL = "SELECT " & _
           "'4J D30' as lblPlant, " & _
           "CONVERT(VARCHAR(11),GETDATE(),103) AS data_etiq,  " & _
           "Mov_etiq.codigo_produto1, " & _
           "ETIQUETA.peso, " & _
           "ETIQUETA.QTDE AS Quantidade, " & _
           "Mov_etiq.cod_cliente, " & _
           "Mov_Etiq.compl_peca2 AS Route, " & _
           "Mov_etiq.Pto_entrega as Ponto_Entrega, "
Rem aqui marcos gm

      sSQL = sSQL & _
           "Mov_etiq.data_lote as Data_Producao, " & _
           "' ' AS Data_validade,  " & _
           "CONVERT(VARCHAR(11),GETDATE(),103) AS Data_Expedicao,  " & _
           "Mov_Etiq.Cod_embalagem_Pw AS Cod_Emb, " & _
           "'09' AS DOCA, " & _
           "ETIQUETA.xblnr AS Num_Doc_Fiscal, " & _
           "Mov_Etiq.desvio as Lote_Sob_Desv, " & _
           "Mov_Etiq.embarque_controlado as Embarque_Ctrl, " & _
           "Mov_Etiq.ind_suplementar AS Indicacao_Supl, " & _
           "Mov_Etiq.classe_funcional AS Classe_Func, " & _
           "Mov_Etiq.compl_peca3 AS Dados_Transporte, " & _
           "Mov_Etiq.Qtd_lote as Qtde_Lote, " & _
           "ETIQUETA.lote AS Num_Lote, " & _
           "'MUSASHI DO BRASIL LTDA' AS Razao_Social, "
      
      sSQL = sSQL & _
           "Mov_etiq.descr_Peca as Descricao_Produto, " & _
           "ETIQUETA.cod_no_cliente as Num_Desenho, " & _
           "Mov_Etiq.Vinculo, " & _
           "Mov_Etiq.compl_peca2 AS Restricoes, " & _
           "Mov_Etiq.DUM, " & _
           "ETIQUETA.ID_PECA, " & _
           "ETIQUETA.ID_ETIQUETA, " & _
           "ETIQUETA.ID_CLIENTE, " & _
           "ETIQUETA.DATA_CRIACAO "
           
    sSQL = sSQL & "FROM Mov_Etiq "
    
    sSQL = sSQL & "INNER JOIN ETIQUETA ON Mov_Etiq.sequencia = ETIQUETA.ID_ETIQUETA "
    
    If Trim(cFields(2)) = "1" Then ' pallet
       sSQL = sSQL & " WHERE ETIQUETA.pallet  = '" & cFields(1) & "'"
    Else 'etiqueta
       sSQL = sSQL & " WHERE ETIQUETA.id_etiqueta  = " & cFields(1) & ""
    End If

    sSQL = sSQL & " ORDER BY ETIQUETA.ID_PECA"

End Select

Set rs = New ADODB.Recordset

cCONECCAO.CursorLocation = adUseClientBatch

rs.Open sSQL, cCONECCAO

Set Mov_Etiq_dados = rs
Set rs = Nothing
Set rsSeq = Nothing

Exit Function

Erro:
Set rs = Nothing
Set rsSeq = Nothing
If InStr(1, Err.Description, "duplicate key") > 0 Then
   Err.Number = 1
End If

Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function Carrega_Campos_MovEtiq(sSequencia As String, _
                                        sQtdEetiq As String, _
                                        sPeso As String, _
                                        sStatusMov As String, _
                                        sUsuario As String, _
                                        sMatricula As String, _
                                        cRec As Recordset) As Collection
Dim cFields As Collection

Set cFields = New Collection


cFields.Add IIf(IsNull(cRec!Cod_Peca), "", "'" & Trim(cRec!Cod_Peca) & "'"), "Cod_Peca"  '01    As String * 7
cFields.Add Trim(cRec!Lote), "Lote"                                                      '02    As String * 15
cFields.Add sQtdEetiq, "Qtd_Caixa"                                                       '03    As numeric * 10
cFields.Add Format(cRec!Data_Etiq, "mm/dd/yyyy")                                         '04    As ddmmyyyy * 8
cFields.Add IIf(IsNull(cRec!Cod_Tabela), "", Trim(cRec!Cod_Tabela)), "Cod_Tabela"        '05    As String * 4
cFields.Add IIf(IsNull(cRec!Cliente), "", Trim(cRec!Cliente)), "Cliente"                 '06    As String * 30
cFields.Add IIf(IsNull(cRec!Cod_Cliente), "", Trim(cRec!Cod_Cliente)), "Cod_Cliente"     '07    As String * 20
cFields.Add sPeso, "Peso"                                                                '08    As String * 16
cFields.Add Trim(cRec!Descr_Peca), "Descr_Peca"                                          '09    As String * 40
cFields.Add cRec!Qtd_Etiq, "Qtd_Etiq"                                                    '10    As numeric * 4
cFields.Add cRec!Tipo, "Tipo"                                                            '11    As String * 1
cFields.Add IIf(IsNull(cRec!Data_Expedicao), "", Format(cRec!Data_Expedicao, "mm/dd/yyyy")), "Data_Expedicao" '12    As ddmmyyyy * 8
cFields.Add IIf(IsNull(cRec!Cod_Embalagem_Pw), "", Trim(cRec!Cod_Embalagem_Pw)), "Cod_Embalagem_Pw"           '13    As String * 3
cFields.Add IIf(IsNull(cRec!data_lote), "", Format(cRec!data_lote, "mm/dd/yyyy")), "Data_Lote"                '14    As ddmmyyyy * 8
cFields.Add IIf(IsNull(cRec!Classe_Funcional), "", Trim(cRec!Classe_Funcional)), "Classe_Funcional"           '15    As String * 2
cFields.Add IIf(IsNull(cRec!Embarque_Controlado), "", Trim(cRec!Embarque_Controlado)), "Embarque_Controlado"  '16    As String * 3
cFields.Add IIf(IsNull(cRec!Cod_Fornecedor), "", Trim(cRec!Cod_Fornecedor)), "Cod_Fornecedor"                 '17    As String * 10
cFields.Add IIf(IsNull(cRec!Num_Doc_Fiscal), "", Trim(cRec!Num_Doc_Fiscal)), "Num_Doc_Fiscal"       '18    As String * 6
cFields.Add IIf(IsNull(cRec!Cod_Embalagem), "", Trim(cRec!Cod_Embalagem)), "Cod_Embalagem"          '19    As String * 3
cFields.Add IIf(IsNull(cRec!Ind_Suplementar), "", Trim(cRec!Ind_Suplementar)), "Ind_Suplementar"    '20    As String * 2 'Apos este campos pos 200
cFields.Add IIf(IsNull(cRec!Pto_Entrega), "", Trim(cRec!Pto_Entrega)), "Pto_Entrega"                '21    As String * 15
cFields.Add IIf(IsNull(cRec!Dum), "", Format(cRec!Dum, "mm/dd/yyyy")), "Dum"                        '22    As ddmmyyyy * 8
cFields.Add cRec!Qtd_Lote, "Qtd_Lote"                                                               '23    As numeric * 6
cFields.Add IIf(IsNull(cRec!Vinculo), "", Trim(cRec!Vinculo)), "Vinculo"                            '24    As String * 2
cFields.Add IIf(IsNull(cRec!desvio), "", Trim(cRec!desvio)), "desvio"                               '25    As String * 5
cFields.Add IIf(IsNull(cRec!Serial), "", cRec!Serial), "Serial"                                     '26    As String * 10
cFields.Add IIf(IsNull(cRec!Cod_Util), "", Trim(cRec!Cod_Util)), "Cod_Util"                         '27    As String * 4
cFields.Add IIf(IsNull(cRec!Linha_Util), "", cRec!Linha_Util), "Linha_Util"                         '28    As String * 4
cFields.Add IIf(IsNull(cRec!Modelo), "", Trim(cRec!Modelo)), "Modelo"                               '29    As String * 10
cFields.Add IIf(IsNull(cRec!Tipo_caixa), "1", Trim(cRec!Tipo_caixa)), "Tipo_caixa"                   '30    As String * 1
cFields.Add IIf(IsNull(cRec!desvio_aviso_mod), "", Trim(cRec!desvio_aviso_mod)), "desvio_aviso_mod" '31    As String * 25
cFields.Add IIf(IsNull(cRec!tipo_alteracao), "", Trim(cRec!tipo_alteracao)), "tipo_alteracao"       '32    As String * 1
cFields.Add IIf(IsNull(cRec!motivo_alteracao), "", Trim(cRec!motivo_alteracao)), "motivo_alteracao" '33    As String * 1
cFields.Add IIf(IsNull(cRec!motivo_alteracao_outros), "", Trim(cRec!motivo_alteracao_outros)), "motivo_alteracao_outros"  '24    As String * 25
cFields.Add IIf(IsNull(cRec!envio_lote), "", Trim(cRec!envio_lote)), "envio_lote"                   '35    As String * 1
cFields.Add IIf(IsNull(cRec!num_am), "", Trim(cRec!num_am)), "num_am"                               '36    As String * 8
cFields.Add IIf(IsNull(sMatricula), "", Trim(sMatricula)), "Embalagem"                              '37    As String * 5
cFields.Add IIf(IsNull(cRec!id_cliente), "", Trim(cRec!id_cliente)), "id_cliente"                   '38    As String * 10
cFields.Add IIf(IsNull(cRec!tp_transporte), "", Trim(cRec!tp_transporte)), "tp_transporte"          '39    As String * 1    '(A/R)
cFields.Add IIf(IsNull(cRec!codigo_produto1), "", Trim(cRec!codigo_produto1)), "codigo_produto1"    '40    As String * 10
cFields.Add IIf(IsNull(cRec!qtde_caixa1), "", Trim(cRec!qtde_caixa1)), "qtde_caixa1"                '41    As String * 3
cFields.Add IIf(IsNull(cRec!pecas_caixa1), "", Trim(cRec!pecas_caixa1)), "pecas_caixa1"             '42    As String * 4 'Apos este campo pos 358
cFields.Add IIf(IsNull(cRec!compl_peca1), "", Trim(cRec!compl_peca1)), "compl_peca1"                '43    As String * 10
cFields.Add IIf(IsNull(cRec!codigo_produto2), "", Trim(cRec!codigo_produto2)), "codigo_produto2"    '44    As String * 10
cFields.Add IIf(IsNull(cRec!qtde_caixa2), "", Trim(cRec!qtde_caixa2)), "qtde_caixa2"                '45    As String * 3
cFields.Add IIf(IsNull(cRec!pecas_caixa2), "", Trim(cRec!pecas_caixa2)), "pecas_caixa2"             '46    As String * 4 'Apos este campo pos 385
cFields.Add IIf(IsNull(cRec!compl_peca2), "", Trim(cRec!compl_peca2)), "compl_peca2"                '47    As String * 10
cFields.Add IIf(IsNull(cRec!codigo_produto3), "", Trim(cRec!codigo_produto3)), "codigo_produto3"    '48    As String * 10
cFields.Add IIf(IsNull(cRec!qtde_caixa3), "", Trim(cRec!qtde_caixa3)), "qtde_caixa3"                '49    As String * 3
cFields.Add IIf(IsNull(cRec!pecas_caixa3), "", Trim(cRec!pecas_caixa3)), "pecas_caixa3"             '50    As String * 4 'Apos este campo 412
cFields.Add IIf(IsNull(cRec!compl_peca3), "", Trim(cRec!compl_peca3)), "compl_peca3"                '51    As String * 10
cFields.Add IIf(IsNull(cRec!codigo_produto4), "", Trim(cRec!codigo_produto4)), "codigo_produto4"    '52    As String * 10
cFields.Add IIf(IsNull(cRec!qtde_caixa4), "", Trim(cRec!qtde_caixa4)), "qtde_caixa4"                '53    As String * 3
cFields.Add IIf(IsNull(cRec!pecas_caixa4), "", Trim(cRec!pecas_caixa4)), "pecas_caixa4"             '54    As String * 4
cFields.Add IIf(IsNull(cRec!compl_peca4), "", Trim(cRec!compl_peca4)), "compl_peca4"                '55    As String * 10
cFields.Add sSequencia, "sequencia"                                                                 '56    As String * 10
cFields.Add sSequencia, "id_etiqueta"                                                               '57    As String * 10
cFields.Add cRec!sequencia, "sequencia_org"                                                         '58    As String * 10
cFields.Add sStatusMov, "statusmov"                                                                 '59    As String * 4
cFields.Add sUsuario, "usuario_org"                                                                 '60    As String * 4
cFields.Add IIf(IsNull(cRec!Lgort), "PP04", Trim(cRec!Lgort)), "lgort"                                                                         '61    As String * 4
cFields.Add Format(Now(), "mm-dd-yyyy"), "dt_alteracao"                                                                        '61    As String * 4
cFields.Add IIf(IsNull(cRec!indforjimport), " ", Trim(cRec!indforjimport)), "indforjimport"                                                                        '61    As String * 4
cFields.Add IIf(IsNull(cRec!Sequencia_Dia), "0", Trim(cRec!Sequencia_Dia)), "sequencia_dia"                                                                        '61    As String * 4

Set Carrega_Campos_MovEtiq = cFields

End Function

Private Function Carrega_Campos_Etiqueta(sSequencia As String, _
                                         sQtdEetiq As String, _
                                         sPeso As String, _
                                         sMatricula As String, _
                                         cRec As Recordset) As Collection
Dim cFields As Collection

Set cFields = New Collection

cFields.Add "'" & sSequencia & "'", "ID_ETIQUETA"
cFields.Add IIf(IsNull(cRec!id_cliente), " ", Trim(cRec!id_cliente)), "id_cliente"
cFields.Add IIf(IsNull(cRec!Cod_Peca), " ", Trim(cRec!Cod_Peca)), "ID_PECA"
cFields.Add " ", "ID_BORDERO"
cFields.Add IIf(IsNull(cRec!Lote), " ", Trim(cRec!Lote)), "Lote"
cFields.Add sQtdEetiq, "QTDE"
cFields.Add " ", "TIPO_EMBALAGEM"
cFields.Add Format(Now(), "mm/dd/yyyy"), "DATA_CRIACAO"
cFields.Add Format(Now(), "hh:mm:ss"), "HORA_CRIACAO"
cFields.Add "AB", "Status"
cFields.Add IIf(IsNull(cRec!Cod_Cliente), " ", Trim(cRec!Cod_Cliente)), "cod_no_cliente"
cFields.Add IIf(IsNull(cRec!Tipo_caixa), "1", Trim(cRec!Tipo_caixa)), "Tipo_caixa"
cFields.Add sPeso, "Peso"
cFields.Add " ", "REMESSA"
cFields.Add " ", "fatur"
cFields.Add " ", "xblnr"
cFields.Add "0", "posnr"
cFields.Add " ", "pallet"
cFields.Add " ", "conf_pallet"
cFields.Add " ", "Usuario"
cFields.Add "0", "tp_transporte"
cFields.Add IIf(IsNull(cRec!tp_transporte), "0", Trim(cRec!tp_transporte)), "tipo_padrao_cx"
cFields.Add " ", "tipo_frete"
cFields.Add " ", "Placa"
cFields.Add "0", "sequencia_placa"
cFields.Add "0", "id_coletor"
cFields.Add IIf(IsNull(cRec!Lgort), "PP04", Trim(cRec!Lgort)), "lgort"
cFields.Add IIf(IsNull(sMatricula), "1111", Trim(sMatricula)), "Embalagem"

Set Carrega_Campos_Etiqueta = cFields

End Function



