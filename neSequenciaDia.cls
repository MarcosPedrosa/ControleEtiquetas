VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "neSequenciaDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'**************** Ações SEQUENCIA_DIA
'Consultar Ações SEQUENCIA_DIA = 1
'Incluir Ações SEQUENCIA_DIA = 2
'Alterar Ações SEQUENCIA_DIA = 3
'Excluir Ações SEQUENCIA_DIA = 4
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE NEGOCIOS DO Ações na tabela movEtiq
' ***********************************************************
' ***********************************************************
' ***********************************************************

Public Function SEQUENCIA_DIA_Consultar(ByVal sBanco As String, _
                                        ByVal c_Tipo As String) As Integer
                                   

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura
Dim rRec As ADODB.Recordset

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add c_Tipo

Set rRec = SEQUENCIA_DIA_Critica(ADOConnection, cFields, 1)

If rRec.RecordCount = 0 Then
   SEQUENCIA_DIA_Consultar = 0
Else
   SEQUENCIA_DIA_Consultar = rRec!SEQUENCIA_DIA
End If

ADOConnection.CommitTrans
Set ADOConnection = Nothing
Set cConect = Nothing
Set cFields = Nothing
Set rRec = Nothing

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
Set rRec = Nothing

Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Function SEQUENCIA_DIA_Incluir(ByVal sBanco As String, _
                                 ByVal cFields As Collection) As ADODB.Recordset

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cConect As daAbertura
Dim nSequencia As Double


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Set SEQUENCIA_DIA_Incluir = SEQUENCIA_DIA_Critica(ADOConnection, cFields, 2)

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
Public Function SEQUENCIA_DIA_Excluir(ByVal sBanco As String, _
                                      ByVal cFields As Collection) As ADODB.Recordset

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cConect As daAbertura
Dim nSequencia As Double


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Set SEQUENCIA_DIA_Excluir = SEQUENCIA_DIA_Critica(ADOConnection, cFields, 3)

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
Public Function SEQUENCIA_DIA_Alterar(ByVal sBanco As String, _
                                      ByVal cFields As Collection) As ADODB.Recordset

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cConect As daAbertura
Dim nSequencia As Double


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Set SEQUENCIA_DIA_Alterar = SEQUENCIA_DIA_Critica(ADOConnection, cFields, 4)

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
Public Function SEQUENCIA_DIA_Critica(ByVal cCONECCAO As Connection, _
                                      ByVal cFields As Collection, _
                                      ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim sSQL As String
Dim sName As String

If nAction > 1 And nAction < 11 Then
   'Critica dos Campos dos SEQUENCIA_DIA
    If nAction = 1 And Len(Trim(cFields(1))) = 0 Then
       Err.Raise 50000, Err.Source, "Código da SEQUENCIA_DIA Inválido"
    End If
    If Len(Trim(cFields(2))) = 0 Then
       Err.Raise 50000, Err.Source, "identificação do cliente na SEQUENCIA_DIA Inválido"
    End If
    
    For nx = 2 To cFields.Count
        If IsNull(cFields(nx)) Then
            'sName = cFields(nx).Name
            cFields.Remove nx
            cFields.Add "NULL", , , nx - 1
        Else
            Select Case nx
                'Verifica campos alfanuméricos NULL
                Case 2 To 26, 28
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
                Case 0
                Rem
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
                 'Altera campos tipo data da inclusao
                Case 0
                    sSQL = "CONVERT(DATETIME, CONVERT(VARCHAR, GETDATE(), 103), 103)"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                 'Altera campos tipo hora da gravacao
                Case 0
                    sSQL = "CONVERT(DATETIME, CONVERT(VARCHAR, GETDATE(), 108), 108)"
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
                Case 6, 13
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

Set SEQUENCIA_DIA_Critica = SEQUENCIA_DIA_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE DADOS DO SEQUENCIA_DIA
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function SEQUENCIA_DIA_dados(ByVal cCONECCAO As Connection, _
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
    Case 1 'Consultar SEQUENCIA_DIA(filtrando por campos)
Rem ************************************************************************************************************
Rem ************************************************************************************************************
        
'    sSQL = "SELECT " & _
'           "SEQUENCIA_DIA.DATA, " & _
'           "SEQUENCIA_DIA.COD_CLIENTE, " & _
'           "SEQUENCIA_DIA.SEQUENCIA, " & _
'           "SEQUENCIA_DIA.EMPRESA "
'
'    sSQL = sSQL & "FROM SEQUENCIA_DIA "
'
'    sSQL = sSQL & " WHERE SEQUENCIA_DIA.COD_CLIENTE  = " & cFields(1)
'    sSQL = sSQL & " AND SEQUENCIA_DIA.EMPRESA  = " & cFields(2)

    sSQL = "SELECT " & _
           "COUNT(*) AS SEQUENCIA_DIA "

    sSQL = sSQL & "FROM MOV_ETIQ,ETIQUETA "
    
    sSQL = sSQL & " WHERE MOV_ETIQ.TIPO  = '" & cFields(1) & "'"
    sSQL = sSQL & " AND   ETIQUETA.ID_ETIQUETA = MOV_ETIQ.ID_ETIQUETA "
    sSQL = sSQL & " AND   CONVERT (date, GETDATE()) = CONVERT (date, ETIQUETA.DATA_CRIACAO) "
    
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 2 'Incluir SEQUENCIA_DIA
Rem ************************************************************************************************************
Rem ************************************************************************************************************

    sSQL = "INSERT INTO SEQUENCIA_DIA ( " & _
           "SEQUENCIA_DIA.DATA, " & _
           "SEQUENCIA_DIA.COD_CLIENTE, " & _
           "SEQUENCIA_DIA.SEQUENCIA, " & _
           "SEQUENCIA_DIA.EMPRESA "
           
    sSQL = sSQL & ") VALUES (" & _
                   cFields(1) & ", " & cFields(2) & ", " & _
                   cFields(3) & ", " & cFields(4) & ")"
                    
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 3 'Excluir a SEQUENCIA_DIA da tabela SEQUENCIA_DIAs
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    sSQL = "DELETE FROM SEQUENCIA_DIA "
    sSQL = sSQL & " where ID_SEQUENCIA_DIA  = " & cFields(1)
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 4 'Acrescentando SEQUENCIA_DIA DA DETERMINADA EMPRESA
Rem ************************************************************************************************************
Rem ************************************************************************************************************
        
    sSQL = "UPDATE SEQUENCIA_DIA SET " & _
           "SEQUENCIA_DIA.DATA = " & cFields(1) & _
           "SEQUENCIA_DIA.SEQUENCIA = " & cFields(2)

    sSQL = sSQL & " WHERE SEQUENCIA_DIA.COD_CLIENTE  = " & cFields(1)
    sSQL = sSQL & " AND SEQUENCIA_DIA.EMPRESA  = " & cFields(2)

End Select
Rem ************************************************************************************************************

Set rs = New ADODB.Recordset

cCONECCAO.CursorLocation = adUseClientBatch

rs.Open sSQL, cCONECCAO

Set SEQUENCIA_DIA_dados = rs
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


