VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "neUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Variável pública que conterá o MDIform
'Option Explicit
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE NEGOCIOS DO TURMA
' ***********************************************************
' ***********************************************************
' ***********************************************************

Public Function USUARIO_Incluir(ByVal sBanco As String, _
                                ByVal c_codigo As String, _
                                ByVal c_nome As String, _
                                ByVal c_Login As String, _
                                ByVal c_senha As String, _
                                ByVal c_Matricula As String, _
                                ByVal c_Tipo As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_codigo, "c_codigo"
cFields.Add c_nome, "c_nome"
cFields.Add c_Login, "c_Login"
cFields.Add c_senha, "c_Senha"
cFields.Add c_Matricula, "c_Matricula"
cFields.Add c_Tipo, "c_Tipo"

Set USUARIO_Incluir = USUARIO_Critica(ADOConnection, cFields, 2)

ADOConnection.CommitTrans

Rem consultar o ultimo registro incluso no banco
Set USUARIO_Incluir = USUARIO_Critica(ADOConnection, cFields, 12)

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

Public Function USUARIO_Consultar(ByVal sBanco As String, _
                         Optional ByVal c_codigo As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

If IsNull(c_codigo) Then
Else
   cFields.Add c_codigo, "c_codigo"
End If

Set USUARIO_Consultar = USUARIO_Critica(ADOConnection, cFields, 1)

If USUARIO_Consultar.RecordCount = 0 And cFields.Count = 0 Then
   Err.Raise 50001, Err.Source, "Não há registros encontrados, Sem usuários cadastrados!"
End If

If USUARIO_Consultar.RecordCount = 0 Then
   Err.Raise 50000, Err.Source, "Não há Usuário encontrado, com este código!"
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
Public Function USUARIO_Confirmar_Login(ByVal sBanco As String, _
                                        ByVal c_codigo As String, _
                                        ByVal c_senha As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add "'" & c_codigo & "'", "c_codigo"

Set USUARIO_Confirmar_Login = USUARIO_Critica(ADOConnection, cFields, 13)

If USUARIO_Confirmar_Login.RecordCount = 0 And cFields.Count = 0 Then
   Err.Raise 50001, Err.Source, "Não há registros encontrados, Sem usuários cadastrados!"
End If

If USUARIO_Confirmar_Login.RecordCount = 0 Then
   Err.Raise 50000, Err.Source, "Não há Usuário encontrado, com este código!"
End If

If c_senha <> USUARIO_Confirmar_Login!SENHA Then
   Err.Raise 50002, Err.Source, "Senha não inválida, redigite!"
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

Public Function USUARIO_Alterar(ByVal sBanco As String, _
                                ByVal c_codigo As String, _
                                ByVal c_nome As String, _
                                ByVal c_Login As String, _
                                ByVal c_senha As String, _
                                ByVal c_Matricula As String, _
                                ByVal c_Tipo As String) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_codigo, "c_codigo" '1
cFields.Add c_nome, "c_nome" '2
cFields.Add c_Login, "c_Login" '3
cFields.Add c_senha, "c_Senha" '4
cFields.Add c_Matricula, "c_Matricula" '5
cFields.Add c_Tipo, "c_Tipo" '6

Set USUARIO_Alterar = USUARIO_Critica(ADOConnection, cFields, 3)

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
Public Function USUARIO_Excluir(ByVal sBanco As String, _
                                ByVal c_codigo As Integer) As ADODB.Recordset
On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_codigo, "c_codigo" '1
cFields.Add Null, "c_nome" '2
cFields.Add Null, "c_Login" '3
cFields.Add Null, "c_Senha" '4
cFields.Add Null, "c_Matricula" '5
cFields.Add Null, "c_Tipo" '6

Rem verificar se existe movimento na tabela mov_etiq
Set USUARIO_Excluir = USUARIO_Critica(ADOConnection, cFields, 11)

If USUARIO_Excluir.Fields(0) = 0 Then
   Set USUARIO_Excluir = USUARIO_Critica(ADOConnection, cFields, 4)
Else
   Err.Raise 50000, Err.Source, "Existe registro de movimentação de etiquetas com este usuário, Operação cancelada."
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
Public Function USUARIO_Critica(ByVal cCONECCAO As Connection, _
                                ByVal cFields As Collection, _
                                ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim sSQL As String
Dim sName As String

If nAction > 1 And nAction < 11 Then
   'Critica dos Campos dos USUARIO
    If nAction > 2 And Len(Trim(cFields(1))) = 0 Then
       Err.Raise 50000, Err.Source, "Código do USUARIO Inválido"
    End If
    
    If Len(Trim(cFields(2))) = 0 Then
       Err.Raise 50001, Err.Source, "Nome do USUARIO Inválido"
    End If
   
    If Len(Trim(cFields(3))) = 0 Then
       Err.Raise 50002, Err.Source, "Login USUÁRIO Inválido"
    End If
    
    If Len(Trim(cFields(4))) = 0 Then
       Err.Raise 50003, Err.Source, "Senha USUÁRIO Inválido"
    End If
    
    If Len(Trim(cFields(5))) = 0 Then
       Err.Raise 50004, Err.Source, "Matricula USUÁRIO Inválido"
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
                Case 2, 3, 4
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
                Case 2, 3, 4
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

Set USUARIO_Critica = USUARIO_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE DADOS DO USUARIO
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function USUARIO_dados(ByVal cCONECCAO As Connection, _
                              ByVal cFields As Collection, _
                              ByVal nAction As Integer) As ADODB.Recordset
On Error GoTo Erro
Dim rs As ADODB.Recordset
Dim sSQL As String
Dim sDate As String
Dim nseq As Integer
Dim rsSeq As ADODB.Recordset


Select Case nAction
    Case 1 'Consultar USUARIO
    
    sSQL = "SELECT   codigo, " & _
                    "nome, " & _
                    "Login, " & _
                    "senha, " & _
                    "Matricula, " & _
                    "Tipo " & _
                    "FROM USUARIO "
                    
    If cFields.Count > 0 Then
       If Len(Trim(cFields(1))) > 0 Then
          sSQL = sSQL & " WHERE codigo = " & cFields(1)
       Else
          sSQL = sSQL & " ORDER BY nome"
       End If
    Else
       sSQL = sSQL & " ORDER BY nome"
    End If
    
    Case 2 'Incluir USUARIO
    
    sDate = "'" & Format(Now, "mm/dd/yyyy HH:MM:SS") & "'"

    sSQL = "INSERT INTO USUARIO (" & _
                    "nome, " & _
                    "Login, " & _
                    "senha, " & _
                    "Matricula, " & _
                    "Tipo) "
                    
    sSQL = sSQL & _
                  "VALUES (" & _
                  cFields(2) & ", " & _
                  cFields(3) & ", " & _
                  cFields(4) & ", " & _
                  cFields(5) & ", " & _
                  cFields(6) & ")"

    Case 3 'Alterar USUARIO
    
    sDate = "#" & Format(Now, "mm/dd/yyyy HH:MM:SS") & "#"
'                 "  USUARIO.codigo = " & cFields(1)
    sSQL = "UPDATE USUARIO SET" & _
                 "  USUARIO.nome = " & cFields(2) & _
                 ", USUARIO.Login = " & cFields(3) & _
                 ", USUARIO.senha = " & cFields(4) & _
                 ", USUARIO.Matricula = " & cFields(5) & _
                 ", USUARIO.Tipo = " & cFields(6) & _
                 "  WHERE USUARIO.codigo = " & cFields(1)
    
    Case 4 'Excluir USUARIO
    
    sSQL = "DELETE  FROM USUARIO" & _
                 "  WHERE USUARIO.codigo = " & cFields(1)
    
    Case 11 'Consultar MOV_ETIQ PARA VER SE HA REGISTROS COM ESTE USUARIOS
    sSQL = "SELECT   COUNT(*) " & _
                    "FROM MOV_ETIQ "
    sSQL = sSQL & " WHERE usuario_org = " & cFields(1)

    Case 12 'Consultar o registro incluso para retornar o codigo para interface
    sSQL = "SELECT   max(codigo) as codigo FROM USUARIO "

    Case 13 'Consultar USUARIO PELO NOME DO LOGIN
    
    sSQL = "SELECT   codigo, " & _
                    "nome, " & _
                    "Login, " & _
                    "senha, " & _
                    "Matricula, " & _
                    "Tipo " & _
                    "FROM USUARIO "
    
    sSQL = sSQL & " WHERE Login = " & cFields(1)

End Select

Set rs = New ADODB.Recordset

cCONECCAO.CursorLocation = adUseClientBatch

rs.Open sSQL, cCONECCAO

Set USUARIO_dados = rs

Set rs = Nothing
Set rsSeq = Nothing

Exit Function

Erro:
    Set rs = Nothing
    Set rsSeq = Nothing
    If Err.Number = -2147217873 Then Err.Description = "Usuário já cadastrado, Tente com outro nome!"
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

