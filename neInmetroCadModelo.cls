VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "neInmetroCadModelo"
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

Public Function INM_CAD_MODELO_Incluir(ByVal sBanco As String, _
                                       ByVal c_CODIGO As String, _
                                       ByVal c_DESCRICAO As String, _
                                       ByVal c_REGISTRO_INMETRO As String, _
                                       ByVal c_DESC_MOD1 As String, _
                                       ByVal c_DESC_MOD2 As String, _
                                       ByVal c_DESC_MOD3 As String, _
                                       ByVal c_DESC_MOD4 As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_CODIGO, "c_CODIGO"
cFields.Add c_DESCRICAO, "c_DESCRICAO"
cFields.Add c_REGISTRO_INMETRO, "c_REGISTRO_INMETRO"
cFields.Add c_DESC_MOD1
cFields.Add c_DESC_MOD2
cFields.Add c_DESC_MOD3
cFields.Add c_DESC_MOD4

Set INM_CAD_MODELO_Incluir = INM_CAD_MODELO_Critica(ADOConnection, cFields, 2)

Rem consultar o ultimo registro incluso no banco
Set INM_CAD_MODELO_Incluir = INM_CAD_MODELO_Critica(ADOConnection, cFields, 6)

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

Public Function INM_CAD_MODELO_Consultar(ByVal sBanco As String, _
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

Set INM_CAD_MODELO_Consultar = INM_CAD_MODELO_Critica(ADOConnection, cFields, 1)

If INM_CAD_MODELO_Consultar.RecordCount = 0 And cFields.Count = 0 Then
   Err.Raise 50001, Err.Source, "Não há registros encontrados, Sem Modelos cadastrados!"
End If

If INM_CAD_MODELO_Consultar.RecordCount = 0 Then
   Err.Raise 50000, Err.Source, "Não há Registro de Modelos com este Código!"
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

Public Function INM_CAD_MODELO_Alterar(ByVal sBanco As String, _
                                       ByVal c_CODIGO As String, _
                                       ByVal c_DESCRICAO As String, _
                                       ByVal c_REGISTRO_INMETRO As String, _
                                       ByVal c_DESC_MOD1 As String, _
                                       ByVal c_DESC_MOD2 As String, _
                                       ByVal c_DESC_MOD3 As String, _
                                       ByVal c_DESC_MOD4 As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_CODIGO, "c_CODIGO" '1
cFields.Add c_DESCRICAO, "c_DESCRICAO" '2
cFields.Add c_REGISTRO_INMETRO, "c_REGISTRO_INMETRO" '3
cFields.Add c_DESC_MOD1
cFields.Add c_DESC_MOD2
cFields.Add c_DESC_MOD3
cFields.Add c_DESC_MOD4

Set INM_CAD_MODELO_Alterar = INM_CAD_MODELO_Critica(ADOConnection, cFields, 3)

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
Public Function INM_CAD_MODELO_Excluir(ByVal sBanco As String, _
                                       ByVal c_CODIGO As Integer) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_CODIGO, "c_CODIGO" '1
cFields.Add Null '2
cFields.Add Null '3
cFields.Add Null '4
cFields.Add Null '5
cFields.Add Null '6
cFields.Add Null '7

Rem verificar se existe movimento na tabela
Set INM_CAD_MODELO_Excluir = INM_CAD_MODELO_Critica(ADOConnection, cFields, 5)

If INM_CAD_MODELO_Excluir.RecordCount = 0 Then
   Set INM_CAD_MODELO_Excluir = INM_CAD_MODELO_Critica(ADOConnection, cFields, 4)
Else
   Err.Raise 50000, Err.Source, "Existe Peça da Inmetro, " & INM_CAD_MODELO_Excluir.Fields(0) & " com uso deste Modelo. Operação será Cancelada."
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
Public Function INM_CAD_MODELO_Critica(ByVal cCONECCAO As Connection, _
                                       ByVal cFields As Collection, _
                                       ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim sSQL As String
Dim sName As String

If nAction > 1 And nAction < 4 Then
   'Critica dos Campos dos INM_CAD_MODELO
    If nAction > 2 And Len(Trim(cFields(1))) = 0 Then
       Err.Raise 50000, Err.Source, "Código do Modelos Inválido"
    End If
    
    If Len(Trim(cFields(2))) = 0 Then
       Err.Raise 50001, Err.Source, "DESCRICAO do Modelo Inválido"
    End If
    If Len(Trim(cFields(3))) = 0 Then
       Err.Raise 50001, Err.Source, "Número do Registro deste Modelo Inválido"
    End If
    If Len(Trim(cFields(4))) = 0 Then
       Err.Raise 50001, Err.Source, "Descreva pelo Menos um Modelo Aplicável"
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
                Case 2, 3, 4, 5, 6, 7
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
                Case 2, 3, 4, 5, 6, 7
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

Set INM_CAD_MODELO_Critica = INM_CAD_MODELO_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE DADOS DO INM_CAD_MODELO
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function INM_CAD_MODELO_dados(ByVal cCONECCAO As Connection, _
                                     ByVal cFields As Collection, _
                                     ByVal nAction As Integer) As ADODB.Recordset
On Error GoTo Erro
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim sDate As String
Dim nseq As Integer
Dim rsSeq As ADODB.Recordset


Select Case nAction
    Case 1 'Consultar INM_CAD_MODELO
    
    sSQL = "SELECT   CODIGO, " & _
                    "DESCRICAO, " & _
                    "REGISTRO_INMETRO, " & _
                    "DESC_MOD1, " & _
                    "DESC_MOD2, " & _
                    "DESC_MOD3, " & _
                    "DESC_MOD4 " & _
                    "FROM INM_CAD_MODELO "
                    
    If cFields.Count > 0 Then
       If Len(Trim(cFields(1))) > 0 Then
          sSQL = sSQL & " WHERE CODIGO = " & cFields(1)
       Else
          sSQL = sSQL & " ORDER BY DESCRICAO"
       End If
    Else
       sSQL = sSQL & " ORDER BY DESCRICAO"
    End If
    
    Case 2 'Incluir INM_CAD_MODELO
    
    sDate = "'" & Format(Now, "mm/dd/yyyy HH:MM:SS") & "'"

    sSQL = "INSERT INTO INM_CAD_MODELO (" & _
                    "DESCRICAO, " & _
                    "REGISTRO_INMETRO, " & _
                    "DESC_MOD1, " & _
                    "DESC_MOD2, " & _
                    "DESC_MOD3, " & _
                    "DESC_MOD4 )  "
                    
    sSQL = sSQL & _
                  "VALUES (" & _
                  cFields(2) & ", " & _
                  cFields(3) & ", " & _
                  cFields(4) & ", " & _
                  cFields(5) & ", " & _
                  cFields(6) & ", " & _
                  cFields(7) & ")"

    Case 3 'Alterar INM_CAD_MODELO
    
    sDate = "#" & Format(Now, "mm/dd/yyyy HH:MM:SS") & "#"
'                 "  INM_CAD_MODELO.CODIGO = " & cFields(1)
    sSQL = "UPDATE INM_CAD_MODELO SET" & _
                 "  INM_CAD_MODELO.DESCRICAO = " & cFields(2) & _
                 ", INM_CAD_MODELO.REGISTRO_INMETRO = " & cFields(3) & _
                 ", INM_CAD_MODELO.DESC_MOD1 = " & cFields(4) & _
                 ", INM_CAD_MODELO.DESC_MOD2 = " & cFields(5) & _
                 ", INM_CAD_MODELO.DESC_MOD3 = " & cFields(6) & _
                 ", INM_CAD_MODELO.DESC_MOD4 = " & cFields(7) & _
                 "  WHERE INM_CAD_MODELO.CODIGO = " & cFields(1)
    
    Case 4 'Excluir INM_CAD_MODELO
    
    sSQL = "DELETE  FROM INM_CAD_MODELO" & _
                 "  WHERE INM_CAD_MODELO.CODIGO = " & cFields(1)
    
    Case 5 'Consultar INM_CAD_PECA PARA VER SE HA REGISTROS COM ESTE Modelos
    sSQL = "SELECT   COD_PECA_MUSASHI " & _
                    "FROM INM_CAD_PECA "
    sSQL = sSQL & " WHERE INM_CAD_PECA.CODIGO_MODELO = " & cFields(1)

    Case 6 'Consultar o registro incluso para retornar o CODIGO para interface
    sSQL = "SELECT   MAX(CODIGO) AS CODIGO FROM INM_CAD_MODELO "

End Select

Set rs = New ADODB.Recordset

cCONECCAO.CursorLocation = adUseClientBatch

rs.Open sSQL, cCONECCAO

Set INM_CAD_MODELO_dados = rs

Set rs = Nothing
Set rsSeq = Nothing

Exit Function

Erro:
    Set rs = Nothing
    Set rsSeq = Nothing
    If Err.Number = -2147217873 Then Err.Description = "Modelo já cadastrado, Tente novemente ou chame o Responsável!"
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function





