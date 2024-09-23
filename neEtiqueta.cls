VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "neEtiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'**************** Ações Etiqueta
'Consultar Ações Etiqueta = 1
'Incluir Ações Etiqueta = 2
'Alterar Ações Etiqueta = 3
'Excluir Ações Etiqueta = 4
' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE NEGOCIOS DO Ações na tabela movEtiq
' ***********************************************************
' ***********************************************************
' ***********************************************************

Public Function Etiqueta_Consultar(ByVal sBanco As String, _
                                   ByVal c_Etiqueta As String) As ADODB.Recordset
                                   

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add IIf(Len(Trim(c_Etiqueta)) = 0, "", c_Etiqueta), "Etiqueta"

Set Etiqueta_Consultar = Etiqueta_Critica(ADOConnection, cFields, 1)

If Etiqueta_Consultar.RecordCount = 0 Then
   Err.Raise 50001, Err.Source, "Não há registros encontrados na tabela de Etiquetas!"
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
Public Function Etiqueta_Incluir(ByVal sBanco As String, _
                                 ByVal cFields As Collection) As ADODB.Recordset

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cConect As daAbertura
Dim nSequencia As Double


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Set Etiqueta_Incluir = Etiqueta_Critica(ADOConnection, cFields, 2)

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
Public Function Etiqueta_Excluir(ByVal sBanco As String, _
                                 ByVal cFields As Collection) As ADODB.Recordset

On Error GoTo Erro
Dim ADOConnection As ADODB.Connection
Dim cConect As daAbertura
Dim nSequencia As Double


Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

Set Etiqueta_Excluir = Etiqueta_Critica(ADOConnection, cFields, 14)

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
Public Function Etiqueta_Cons_Seq_Etiq(ByVal sBanco As String, sTipo As String) As ADODB.Recordset
                                   
Rem esta rotina foi criada para pegar quantos regitros foram gerados no dia de uma determinado tipo de etiqueta

On Error GoTo Erro

Dim ADOConnection As ADODB.Connection
Dim cFields As Collection
Dim cConect As daAbertura

Set ADOConnection = New ADODB.Connection
Set cConect = New daAbertura
Set cFields = New Collection
Set ADOConnection = cConect.Coneccao_OLEDB(sBanco, "A")

cFields.Add sTipo

Set Etiqueta_Cons_Seq_Etiq = Etiqueta_Critica(ADOConnection, cFields, 13)

If Etiqueta_Cons_Seq_Etiq.RecordCount = 0 Then
   nSequencia_Yamaha = 1
Else
   nSequencia_Yamaha = Etiqueta_Cons_Seq_Etiq.RecordCount
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
Public Function Etiqueta_Critica(ByVal cCONECCAO As Connection, _
                                ByVal cFields As Collection, _
                                ByVal nAction As Integer) As ADODB.Recordset

On Error GoTo Erro
Dim nx As Integer
Dim sSQL As String
Dim sName As String

'1Ccfields.Add sSequencia, "ID_ETIQUETA"
'2Ccfields.Add IIf(IsNull(cRec!id_cliente), "", Trim(cRec!id_cliente)), "id_cliente"
'3Ccfields.Add IIf(IsNull(cRec!Cod_Peca), "", "'" & Trim(cRec!Cod_Peca) & "'"), "ID_PECA"
'4Ccfields.Add "", "ID_BORDERO"
'5Ccfields.Add IIf(IsNull(cRec!Lote), "", "'" & Trim(cRec!Lote) & "'"), "Lote"
'6Ncfields.Add sQtdEetiq, "QTDE"
'7Ccfields.Add "", "TIPO_EMBALAGEM"
'8Dcfields.Add "", "DATA_CRIACAO"
'9Dcfields.Add "", "HORA_CRIACAO"
'10Ccfields.Add "AB", "Status"
'11Ccfields.Add IIf(IsNull(cRec!Cod_Cliente), "", "'" & Trim(cRec!Cod_Cliente) & "'"), "cod_no_cliente"
'12Ccfields.Add IIf(IsNull(cRec!Tipo_caixa), "", "'" & Trim(cRec!Tipo_caixa) & "'"), "Tipo_caixa"
'13Ncfields.Add IIf(IsNull(cRec!Peso), "", "'" & Trim(cRec!Peso) & "'"), "Peso"
'14Ccfields.Add " ", "REMESSA"
'15Ccfields.Add " ", "fatur"
'16Ccfields.Add " ", "xblnr"
'17Ncfields.Add "0", "posnr"
'18Ccfields.Add " ", "pallet"
'19Ccfields.Add " ", "conf_pallet"
'20Ccfields.Add " ", "Usuario"
'21Ccfields.Add "0", "tp_transporte"
'22Ncfields.Add "0", "tipo_padrao_cx"
'23Ccfields.Add " ", "tipo_frete"
'24Ccfields.Add " ", "Placa"
'25Ncfields.Add "0", "sequencia_placa"
'26Ncfields.Add "0", "id_coletor"
'27Ncfields.Add "PP04", "lgort"
'28Ncfields.Add "MATRICULA", "EMBALAGEM"



If nAction > 1 And nAction < 11 Then
   'Critica dos Campos dos Etiqueta
    If nAction = 1 And Len(Trim(cFields(1))) = 0 Then
       Err.Raise 50000, Err.Source, "Código da Etiqueta Inválido"
    End If
    If Len(Trim(cFields(2))) = 0 Then
       Err.Raise 50000, Err.Source, "identificação do cliente na Etiqueta Inválido"
    End If
    For nx = 2 To cFields.Count
        If IsNull(cFields(nx)) Then
            'sName = cFields(nx).Name
            cFields.Remove nx
            cFields.Add "NULL", , , nx - 1
        Else
            Select Case nx
                'Verifica campos alfanuméricos NULL
                Case 2 To 21, 23 To 26, 28
                 If Len(Trim(cFields(nx))) = 0 Then
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add "", , , nx - 1
                 End If
                
                'Verifica campos numéricos NULL
                Case 22
                    If Val(cFields(nx)) = 0 Then
                        'sName = cFields(nx).Name
                        cFields.Remove nx
                        cFields.Add "0", , , nx - 1
                    End If
            End Select
        End If
        
        If cFields(nx) <> "NULL" Then
            Select Case nx
                'Altera campos alfanuméricos
                Case 2 To 5, 7, 10 To 12, 14 To 16, 18 To 21, 23, 24, 27, 28
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
                Case 8
                    sSQL = "CONVERT(DATETIME, CONVERT(VARCHAR, GETDATE(), 103), 103)"
                    'sName = cFields(nx).Name
                    cFields.Remove nx
                    cFields.Add sSQL, , , nx - 1
                 'Altera campos tipo hora da gravacao
                Case 9
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

Set Etiqueta_Critica = Etiqueta_dados(cCONECCAO, cFields, nAction)

Exit Function

Erro:
Err.Raise Err.Number, Err.Source, Err.Description
End Function

' ***********************************************************
' ***********************************************************
' ***********************************************************
'   AQUI COMEÇA A CAMADA DE DADOS DO Etiqueta
' ***********************************************************
' ***********************************************************
' ***********************************************************
Public Function Etiqueta_dados(ByVal cCONECCAO As Connection, _
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
    Case 1 'Consultar Etiqueta(filtrando por campos)
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    bWere = False
    
    sSQL = "SELECT " & _
           "ETIQUETA.ID_ETIQUETA, " & _
           "ETIQUETA.ID_CLIENTE, " & _
           "ETIQUETA.ID_PECA, " & _
           "ETIQUETA.ID_BORDERO, " & _
           "ETIQUETA.LOTE, " & _
           "ETIQUETA.QTDE, " & _
           "ETIQUETA.TIPO_EMBALAGEM, " & _
           "ETIQUETA.DATA_CRIACAO, " & _
           "ETIQUETA.HORA_CRIACAO, " & _
           "ETIQUETA.STATUS, " & _
           "ETIQUETA.cod_no_cliente, " & _
           "ETIQUETA.tipo_caixa, " & _
           "ETIQUETA.peso, " & _
           "ETIQUETA.REMESSA, " & _
           "ETIQUETA.fatur, " & _
           "ETIQUETA.xblnr, " & _
           "ETIQUETA.posnr, " & _
           "ETIQUETA.pallet, " & _
           "ETIQUETA.conf_pallet, " & _
           "ETIQUETA.usuario, "
     sSQL = sSQL & _
           "ETIQUETA.tp_transporte, " & _
           "ETIQUETA.tipo_padrao_cx, " & _
           "ETIQUETA.tipo_frete, " & _
           "ETIQUETA.placa, " & _
           "ETIQUETA.sequencia_placa, " & _
           "ETIQUETA.id_coletor, " & _
           "ETIQUETA.lgort, " & _
           "ETIQUETA.embalagem "

    sSQL = sSQL & "FROM ETIQUETA "
    
    If Len(Trim(cFields(1))) > 0 Then ' ETIQUETA
       sSQL = sSQL & " and Etiqueta.ID_ETIQUETA  = " & cFields(1)
    End If
    
   
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 2 'Incluir Etiqueta
Rem ************************************************************************************************************
Rem ************************************************************************************************************

    sSQL = "INSERT INTO ETIQUETA ( " & _
           "ID_ETIQUETA, " & _
           "ID_CLIENTE, " & _
           "ID_PECA, " & _
           "ID_BORDERO, " & _
           "LOTE, " & _
           "QTDE, " & _
           "TIPO_EMBALAGEM, " & _
           "DATA_CRIACAO, " & _
           "HORA_CRIACAO, " & _
           "STATUS, " & _
           "cod_no_cliente, " & _
           "tipo_caixa, " & _
           "peso, " & _
           "REMESSA, " & _
           "fatur, " & _
           "xblnr, " & _
           "posnr, "
           
    sSQL = sSQL & _
           "pallet, " & _
           "conf_pallet, " & _
           "usuario, " & _
           "tp_transporte, " & _
           "tipo_padrao_cx, " & _
           "tipo_frete, " & _
           "placa, " & _
           "sequencia_placa, " & _
           "id_coletor, " & _
           "lgort "
'           "embalagem "
           
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
                   cFields(27) & ")"
                    
Rem                   cFields (27) & ", " & cFields(28) & ")"
                    
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 13 'Verificar quantas etiquetas foram emitidas da yamaha
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    sSQL = "SELECT COUNT(*) FROM MOV_ETIQ "
    sSQL = sSQL & " WHERE TIPO  = '" & cFields(1) & "'"
    sSQL = sSQL & " AND DATA_ETIQ = '" & Format(Now(), "YYYYMMDD")
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    Case 14 'Excluir a etiqueta da tabela etiquetas
Rem ************************************************************************************************************
Rem ************************************************************************************************************
    
    sSQL = "DELETE FROM ETIQUETA "
    sSQL = sSQL & " where ID_ETIQUETA  = " & cFields(1)
    sSQL = sSQL & " AND   (xblnr  = '' OR xblnr IS NULL)"


End Select
Rem ************************************************************************************************************

Set rs = New ADODB.Recordset

cCONECCAO.CursorLocation = adUseClientBatch

rs.Open sSQL, cCONECCAO

Set Etiqueta_dados = rs
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
