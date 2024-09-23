VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "neInmetroCadPeca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Variável pública que conterá o MDIform
'Option Explicit
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE NEGOCIOS CADASTRO MODELO
' ***********************************************************
' ***********************************************************
' ***********************************************************

Public Function INM_CAD_PECA_Incluir(ByVal sBanco As String, _
                                     ByVal c_CODIGO As String, _
                                     ByVal c_COD_PECA_MUSASHI As String, _
                                     ByVal c_COD_PECA_CLIENTE As String, _
                                     ByVal c_CODIGO_MODELO As String, _
                                     ByVal c_CODIGO_CLIENTE As String, _
                                     ByVal c_DESC_PECA1 As String, _
                                     ByVal c_DESC_PECA2 As String, _
                                     ByVal c_DESC_PECA3 As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_CODIGO, "c_CODIGO"
cFields.Add c_COD_PECA_MUSASHI
cFields.Add c_COD_PECA_CLIENTE
cFields.Add c_CODIGO_MODELO
cFields.Add c_CODIGO_CLIENTE
cFields.Add c_DESC_PECA1
cFields.Add c_DESC_PECA2
cFields.Add c_DESC_PECA3

Set INM_CAD_PECA_Incluir = INM_CAD_PECA_Critica(ADOConnection, cFields, 2)

Rem consultar o ultimo registro incluso no banco
Set INM_CAD_PECA_Incluir = INM_CAD_PECA_Critica(ADOConnection, cFields, 6)

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

Public Function INM_CAD_PECA_Consultar(ByVal sBanco As String, _
                               Optional ByVal c_CODIGO As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

If IsNull(c_CODIGO) Then
Else
   cFields.Add c_CODIGO, "c_CODIGO"
End If

Set INM_CAD_PECA_Consultar = INM_CAD_PECA_Critica(ADOConnection, cFields, 1)

If INM_CAD_PECA_Consultar.RecordCount = 0 And cFields.Count = 0 Then
   Err.Raise 50001, Err.Source, "Não há Registros no Cadastro de Peças da Inmetro Encontrados, Sem Peças Cadastradas !"
End If

If INM_CAD_PECA_Consultar.RecordCount = 0 Then
   Err.Raise 50000, Err.Source, "Não há Registro de Peça da Inmetro com Este Código!"
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

Public Function INM_CAD_PECA_Alterar(ByVal sBanco As String, _
                                     ByVal c_CODIGO As String, _
                                     ByVal c_COD_PECA_MUSASHI As String, _
                                     ByVal c_COD_PECA_CLIENTE As String, _
                                     ByVal c_CODIGO_MODELO As String, _
                                     ByVal c_CODIGO_CLIENTE As String, _
                                     ByVal c_DESC_PECA1 As String, _
                                     ByVal c_DESC_PECA2 As String, _
                                     ByVal c_DESC_PECA3 As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_CODIGO, "c_CODIGO"
cFields.Add c_COD_PECA_MUSASHI
cFields.Add c_COD_PECA_CLIENTE
cFields.Add c_CODIGO_MODELO
cFields.Add c_CODIGO_CLIENTE
cFields.Add c_DESC_PECA1
cFields.Add c_DESC_PECA2
cFields.Add c_DESC_PECA3

Set INM_CAD_PECA_Alterar = INM_CAD_PECA_Critica(ADOConnection, cFields, 3)

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
Public Function INM_CAD_PECA_Excluir(ByVal sBanco As String, _
                                     ByVal c_CODIGO As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_CODIGO, "c_CODIGO"
cFields.Add Null '2
cFields.Add Null '3
cFields.Add Null '4
cFields.Add Null '5
cFields.Add Null '6
cFields.Add Null '7
cFields.Add Null '8

Rem verificar se existe movimento na tabela
Set INM_CAD_PECA_Excluir = INM_CAD_PECA_Critica(ADOConnection, cFields, 5)

If INM_CAD_PECA_Excluir.Fields(0) > 0 Then
   Set INM_CAD_PECA_Excluir = INM_CAD_PECA_Critica(ADOConnection, cFields, 4)
Else
   Err.Raise 50000, Err.Source, "Não Existe esta Peça da Etiqueta da Inmetro com este código, Operação Cancelada."
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
Public Function INM_CAD_PECA_Cons_CliMod(ByVal sBanco As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Set INM_CAD_PECA_Cons_CliMod = INM_CAD_PECA_Critica(ADOConnection, cFields, 7)

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
Public Function INM_CAD_PECA_JaCadastrada(ByVal sBanco As String, _
                                          ByVal c_CODIGO As String) As Boolean
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim cRec As ADODB.Recordset

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_CODIGO, "c_CODIGO"

Set cRec = INM_CAD_PECA_Critica(ADOConnection, cFields, 1)

If cRec.RecordCount = 0 Then
   INM_CAD_PECA_JaCadastrada = False
Else
   INM_CAD_PECA_JaCadastrada = True
End If

ADOConnection.CommitTrans
Set cRec = Nothing
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set cRec = Nothing
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function INM_CAD_PECA_ExisteEtiqueta(ByVal sBanco As String, _
                                            ByVal c_CODIGO As String) As Boolean
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim cRec As ADODB.Recordset

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_CODIGO, "c_CODIGO"

Set cRec = INM_CAD_PECA_Critica(ADOConnection, cFields, 8)

If cRec.RecordCount = 0 Then
   INM_CAD_PECA_ExisteEtiqueta = False
Else
   INM_CAD_PECA_ExisteEtiqueta = True
End If

ADOConnection.CommitTrans
Set cRec = Nothing
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing

Exit Function

Erro:
If ADOConnection <> "" Then
   ADOConnection.RollbackTrans
End If
Set cRec = Nothing
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function INM_CAD_PECA_Cons_Impressao(ByVal sBanco As String, _
                                            ByVal c_CODIGO As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_CODIGO, "c_CODIGO"

Set INM_CAD_PECA_Cons_Impressao = INM_CAD_PECA_Critica(ADOConnection, cFields, 9)

If INM_CAD_PECA_Cons_Impressao.RecordCount = 0 Then
   Err.Raise 50001, Err.Source, "Não há Registros no Cadastro de Peças da Inmetro com este Código. Verifique o código !"
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

Public Function INM_CAD_PECA_Critica(ByVal cCONECCAO As Connection, _
                                     ByVal cFields As Collection, _
                                     ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim sSQL As String
Dim sName As String

If nAction > 1 And nAction < 4 Then
   'Critica dos Campos dos INM_CAD_PECA
'    If nAction > 2 And Len(Trim(cFields(1))) = 0 Then
'       Err.Raise 50000, Err.Source, "Código do Modelos Inválido"
'    End If
'cFields.Add c_CODIGO, "c_CODIGO"
'cFields.Add c_COD_PECA_MUSASHI
'cFields.Add c_COD_PECA_CLIENTE
'cFields.Add c_CODIGO_MODELO
'cFields.Add c_CODIGO_CLIENTE
'cFields.Add c_DESC_PECA1
'cFields.Add c_DESC_PECA2
'cFields.Add c_DESC_PECA3
    
    If Len(Trim(cFields(2))) = 0 Then
       Err.Raise 50001, Err.Source, "Codigo da peça da MUSASHI sem preenchimento.Selecione o seu código."
    End If
    If Len(Trim(cFields(3))) = 0 Then
       Err.Raise 50002, Err.Source, "Código da peça do Cliente em Brancos. Preencha o codigo da peça do cliente."
    End If
    If Len(Trim(cFields(4))) = 0 Then
       Err.Raise 50003, Err.Source, "Modelo Não Selecionado. Selecione o Modelo a ser Usado."
    End If
    If Len(Trim(cFields(5))) = 0 Then
       Err.Raise 50004, Err.Source, "Cliente Não Selecionado. Selecione o Cliente."
    End If
    If Len(Trim(cFields(6))) = 0 Then
       Err.Raise 50005, Err.Source, "Descrição da Peça sem Preenchimento. Digite pelo menos uma linha, a Descriçao 1."
    End If
   
   'Por aspas,Nulls,virguas nos campos correspondentes
    For nx = 2 To cFields.Count
        If IsNull(cFields(nx)) Then
            'sName = cFields(nx).Name
            cFields.Remove nx
            cFields.Add "NULL", , , nx - 1
        Else
            Select Case nx
                'Verifica campos alfanuméricos NULL
                Case 2, 3, 6, 7, 8
                 If Len(Trim(cFields(nx))) = 0 Then
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add "NULL", , , nx - 1
                 End If
                
                'Verifica campos numéricos NULL
                Case 0
                    If Val(cFields(nx)) = 0 Then
                        'sName = cFields(nx).Name
                        cFields.Remove nx
                        cFields.Add "NULL", , , nx - 1
                    End If
            End Select
        End If
        
        If cFields(nx) <> "NULL" Then
            Select Case nx
                'Altera campos alfanuméricos
                Case 2, 3, 6, 7, 8
                    sSQL = "'" & Replace(cFields(nx), "'", "`") & "'"
'                    sSQL = "'" & cFields(nx) & "'"
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
                Case 0
                    sSQL = "#" & Format(cFields(nx), "mm/dd/yyyy HH:MM:SS") & "#"
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
                'Altera campos numéricos fracionários
                Case 0
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

Set INM_CAD_PECA_Critica = INM_CAD_PECA_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE DADOS DO INM_CAD_PECA
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function INM_CAD_PECA_dados(ByVal cCONECCAO As Connection, _
                                   ByVal cFields As Collection, _
                                   ByVal nAction As Integer) As ADODB.Recordset
On Error GoTo Erro
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim sDate As String
Dim nseq As Integer
Dim rsSeq As ADODB.Recordset


Select Case nAction
    Case 1 'Consultar INM_CAD_PECA
    
    sSQL = "SELECT  COD_PECA_MUSASHI, " & _
                    "(SELECT NOME FROM INM_CAD_CLIENTE WHERE INM_CAD_CLIENTE.CODIGO = INM_CAD_PECA.CODIGO_CLIENTE) AS NOME_CLIENTE, " & _
                    "(SELECT DESCRICAO FROM INM_CAD_MODELO WHERE INM_CAD_MODELO.CODIGO = INM_CAD_PECA.CODIGO_MODELO) AS DESC_MODELO, " & _
                    "DESC_PECA1, " & _
                    "CODIGO, " & _
                    "COD_PECA_CLIENTE, " & _
                    "CODIGO_MODELO, " & _
                    "CODIGO_CLIENTE, " & _
                    "DESC_PECA2, " & _
                    "DESC_PECA3 " & _
                    "FROM INM_CAD_PECA "
                    
    If cFields.Count > 0 Then
       If Len(Trim(cFields(1))) > 0 Then
          sSQL = sSQL & " WHERE COD_PECA_MUSASHI = '" & cFields(1) & "'"
       Else
          sSQL = sSQL & " ORDER BY COD_PECA_MUSASHI"
       End If
    Else
       sSQL = sSQL & " ORDER BY COD_PECA_MUSASHI"
    End If
    
    Case 2 'Incluir INM_CAD_PECA
    
    sDate = "'" & Format(Now, "mm/dd/yyyy HH:MM:SS") & "'"

    sSQL = "INSERT INTO INM_CAD_PECA (" & _
                    "COD_PECA_MUSASHI, " & _
                    "COD_PECA_CLIENTE, " & _
                    "CODIGO_MODELO, " & _
                    "CODIGO_CLIENTE, " & _
                    "DESC_PECA1, " & _
                    "DESC_PECA2, " & _
                    "DESC_PECA3 )  "
                    
    sSQL = sSQL & _
                  "VALUES (" & _
                  cFields(2) & ", " & _
                  cFields(3) & ", " & _
                  cFields(4) & ", " & _
                  cFields(5) & ", " & _
                  cFields(6) & ", " & _
                  cFields(7) & ", " & _
                  cFields(8) & ")"

    Case 3 'Alterar INM_CAD_PECA
    
    sDate = "#" & Format(Now, "mm/dd/yyyy HH:MM:SS") & "#"
'
    sSQL = "UPDATE INM_CAD_PECA SET" & _
                 "  INM_CAD_PECA.COD_PECA_MUSASHI = " & cFields(2) & _
                 ", INM_CAD_PECA.COD_PECA_CLIENTE = " & cFields(3) & _
                 ", INM_CAD_PECA.CODIGO_MODELO = " & cFields(4) & _
                 ", INM_CAD_PECA.CODIGO_CLIENTE = " & cFields(5) & _
                 ", INM_CAD_PECA.DESC_PECA1 = " & cFields(6) & _
                 ", INM_CAD_PECA.DESC_PECA2 = " & cFields(7) & _
                 ", INM_CAD_PECA.DESC_PECA3 = " & cFields(8) & _
                 "  WHERE INM_CAD_PECA.COD_PECA_MUSASHI = " & cFields(2)
    
    Case 4 'Excluir INM_CAD_PECA
    
    sSQL = "DELETE  FROM INM_CAD_PECA" & _
                 "  WHERE INM_CAD_PECA.COD_PECA_MUSASHI = '" & cFields(1) & "'"
    
    Case 5 'Consultar INM_CAD_PECA PARA VER SE HA REGISTROS COM ESTE Modelos
    sSQL = "SELECT   COUNT(*) " & _
                    "FROM INM_CAD_PECA "
    sSQL = sSQL & " WHERE INM_CAD_PECA.COD_PECA_MUSASHI = '" & cFields(1) & "'"

    Case 6 'Consultar o registro incluso para retornar o CODIGO para interface
    sSQL = "SELECT   MAX(CODIGO) AS CODIGO FROM INM_CAD_PECA "

    Case 7 'Consultar o registro de clientes e modelos para carga na tela de pecas para interface
    sSQL = "SELECT   'CLI' AS TIPO, CODIGO, NOME AS DESCRICAO, '' AS DESC_MOD1,'' AS DESC_MOD2,'' AS DESC_MOD3,'' AS DESC_MOD4 FROM INM_CAD_CLIENTE "
    sSQL = sSQL & " UNION " & _
           "SELECT   'MOD' AS TIPO, CODIGO, descricao, DESC_MOD1, DESC_MOD2, DESC_MOD3, DESC_MOD4 DESCRICAO FROM INM_CAD_MODELO "
    
    Case 8 'Consultar o registro de uma etiqueta para saber se existe pelo menos uma no cadsatro de etiquetas
    sSQL = "SELECT TOP 1 * FROM ETIQUETA WHERE ID_PECA = '" & cFields(1) & "'"
    
    Case 9 'Consultar INM_CAD_PECA para impressão da etiqueta
    
    sSQL = "SELECT  COD_PECA_MUSASHI, " & _
                    "(SELECT NOME FROM INM_CAD_CLIENTE WHERE INM_CAD_CLIENTE.CODIGO = INM_CAD_PECA.CODIGO_CLIENTE) AS NOME_CLIENTE, " & _
                    "COD_PECA_CLIENTE, " & _
                    "CODIGO_MODELO, " & _
                    "CODIGO_CLIENTE, " & _
                    "DESC_PECA1, " & _
                    "DESC_PECA2, " & _
                    "DESC_PECA3, " & _
                    "INM_CAD_MODELO.DESCRICAO AS DESC_MODELO, " & _
                    "REGISTRO_INMETRO, " & _
                    "DESC_MOD1, " & _
                    "DESC_MOD2, " & _
                    "DESC_MOD3, " & _
                    "DESC_MOD4 " & _
                    "FROM INM_CAD_PECA "
    sSQL = sSQL & _
           " INNER JOIN INM_CAD_MODELO ON INM_CAD_MODELO.CODIGO = INM_CAD_PECA.CODIGO_MODELO   "
    sSQL = sSQL & " WHERE COD_PECA_MUSASHI = '" & cFields(1) & "'"
    
End Select

Set rs = New ADODB.Recordset

cCONECCAO.CursorLocation = adUseClientBatch

rs.Open sSQL, cCONECCAO

Set INM_CAD_PECA_dados = rs

Set rs = Nothing
Set rsSeq = Nothing

Exit Function

Erro:
    Set rs = Nothing
    Set rsSeq = Nothing
    If Err.Number = -2147217873 Then Err.Description = "Usuário já cadastrado, Tente com outro COD_PECA_MUSASHI!"
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function







