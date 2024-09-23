VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "neInmetroCadCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Variável pública que conterá o MDIform
'Option Explicit
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE NEGOCIOS CADASTRO CLIENTE
' ***********************************************************
' ***********************************************************
' ***********************************************************

Public Function INM_CAD_CLIENTE_Incluir(ByVal sBanco As String, _
                                        ByVal c_CODIGO As String, _
                                        ByVal c_nome As String, _
                                        ByVal c_SAC As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_CODIGO, "c_CODIGO"
cFields.Add c_nome, "c_NOME"
cFields.Add c_SAC, "c_SAC"

Set INM_CAD_CLIENTE_Incluir = INM_CAD_CLIENTE_Critica(ADOConnection, cFields, 2)

Rem consultar o ultimo registro incluso no banco
Set INM_CAD_CLIENTE_Incluir = INM_CAD_CLIENTE_Critica(ADOConnection, cFields, 6)

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

Public Function INM_CAD_CLIENTE_Consultar(ByVal sBanco As String, _
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

Set INM_CAD_CLIENTE_Consultar = INM_CAD_CLIENTE_Critica(ADOConnection, cFields, 1)

If INM_CAD_CLIENTE_Consultar.RecordCount = 0 And cFields.Count = 0 Then
   Err.Raise 50001, Err.Source, "Não há registros encontrados, Sem Clientes cadastrados!"
End If

If INM_CAD_CLIENTE_Consultar.RecordCount = 0 Then
   Err.Raise 50000, Err.Source, "Não há Registro de Cliente com este Código!"
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

Public Function INM_CAD_CLIENTE_Alterar(ByVal sBanco As String, _
                                        ByVal c_CODIGO As String, _
                                        ByVal c_nome As String, _
                                        ByVal c_SAC As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_CODIGO, "c_CODIGO" '1
cFields.Add c_nome, "c_NOME" '2
cFields.Add c_SAC, "c_SAC" '3

Set INM_CAD_CLIENTE_Alterar = INM_CAD_CLIENTE_Critica(ADOConnection, cFields, 3)

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
Public Function INM_CAD_CLIENTE_Excluir(ByVal sBanco As String, _
                                        ByVal c_CODIGO As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_CODIGO, "c_CODIGO" '1
cFields.Add Null, "c_NOME" '2
cFields.Add Null, "c_SAC" '3
cFields.Add Null, "c_Senha" '4
cFields.Add Null, "c_Matricula" '5
cFields.Add Null, "c_Tipo" '6

Rem verificar se existe movimento na tabela mov_etiq
Set INM_CAD_CLIENTE_Excluir = INM_CAD_CLIENTE_Critica(ADOConnection, cFields, 5)

If INM_CAD_CLIENTE_Excluir.RecordCount = 0 Then
   Set INM_CAD_CLIENTE_Excluir = INM_CAD_CLIENTE_Critica(ADOConnection, cFields, 4)
Else
   Err.Raise 50000, Err.Source, "Existe Peça da Inmetro de Num. " & INM_CAD_CLIENTE_Excluir!COD_PECA_MUSASHI & " com uso deste Cliente. Operação será Cancelada."
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
Public Function INM_CAD_CLIENTE_Critica(ByVal cCONECCAO As Connection, _
                                        ByVal cFields As Collection, _
                                        ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim sSQL As String
Dim sName As String

If nAction > 1 And nAction < 4 Then
   'Critica dos Campos dos INM_CAD_CLIENTE
    If nAction > 2 And Len(Trim(cFields(1))) = 0 Then
       Err.Raise 50000, Err.Source, "Código do CLIENTE Inválido"
    End If
    
    If Len(Trim(cFields(2))) = 0 Then
       Err.Raise 50001, Err.Source, "NOME do CLIENTE Inválido"
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
                Case 2, 3
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
                Case 2, 3
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

Set INM_CAD_CLIENTE_Critica = INM_CAD_CLIENTE_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE DADOS DO INM_CAD_CLIENTE
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function INM_CAD_CLIENTE_dados(ByVal cCONECCAO As Connection, _
                                      ByVal cFields As Collection, _
                                      ByVal nAction As Integer) As ADODB.Recordset
On Error GoTo Erro
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim sDate As String
Dim nseq As Integer
Dim rsSeq As ADODB.Recordset


Select Case nAction
    Case 1 'Consultar INM_CAD_CLIENTE
    
    sSQL = "SELECT   CODIGO, " & _
                    "NOME, " & _
                    "SAC " & _
                    "FROM INM_CAD_CLIENTE "
                    
    If cFields.Count > 0 Then
       If Len(Trim(cFields(1))) > 0 Then
          sSQL = sSQL & " WHERE CODIGO = " & cFields(1)
       Else
          sSQL = sSQL & " ORDER BY NOME"
       End If
    Else
       sSQL = sSQL & " ORDER BY NOME"
    End If
    
    Case 2 'Incluir INM_CAD_CLIENTE
    
    sDate = "'" & Format(Now, "mm/dd/yyyy HH:MM:SS") & "'"

    sSQL = "INSERT INTO INM_CAD_CLIENTE (" & _
                    "NOME, " & _
                    "SAC) "
                    
    sSQL = sSQL & _
                  "VALUES (" & _
                  cFields(2) & ", " & _
                  cFields(3) & ")"

    Case 3 'Alterar INM_CAD_CLIENTE
    
    sDate = "#" & Format(Now, "mm/dd/yyyy HH:MM:SS") & "#"
'                 "  INM_CAD_CLIENTE.CODIGO = " & cFields(1)
    sSQL = "UPDATE INM_CAD_CLIENTE SET" & _
                 "  INM_CAD_CLIENTE.NOME = " & cFields(2) & _
                 ", INM_CAD_CLIENTE.SAC = " & cFields(3) & _
                 "  WHERE INM_CAD_CLIENTE.CODIGO = " & cFields(1)
    
    Case 4 'Excluir INM_CAD_CLIENTE
    
    sSQL = "DELETE  FROM INM_CAD_CLIENTE" & _
                 "  WHERE INM_CAD_CLIENTE.CODIGO = " & cFields(1)
    
    Case 5 'Consultar INM_CAD_PECA PARA VER SE HA REGISTROS COM ESTE CLIENTE
    sSQL = "SELECT   COD_PECA_MUSASHI " & _
                    "FROM INM_CAD_PECA "
    sSQL = sSQL & " WHERE CODIGO_CLIENTE = " & cFields(1)

    Case 6 'Consultar o registro incluso para retornar o CODIGO para interface
    sSQL = "SELECT   MAX(CODIGO) AS CODIGO FROM INM_CAD_CLIENTE "

End Select

Set rs = New ADODB.Recordset

cCONECCAO.CursorLocation = adUseClientBatch

rs.Open sSQL, cCONECCAO

Set INM_CAD_CLIENTE_dados = rs

Set rs = Nothing
Set rsSeq = Nothing

Exit Function

Erro:
    Set rs = Nothing
    Set rsSeq = Nothing
    If Err.Number = -2147217873 Then Err.Description = "Usuário já cadastrado, Tente com outro NOME!"
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function



