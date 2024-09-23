VERSION 5.00
Begin VB.Form frmOpcoes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opções"
   ClientHeight    =   3750
   ClientLeft      =   1980
   ClientTop       =   2595
   ClientWidth     =   4170
   Icon            =   "frmOpcoes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbo_impressora 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   270
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   210
      Width           =   3495
   End
   Begin VB.CommandButton cmdVisualizar 
      Caption         =   "Visuali&zar"
      Height          =   735
      Left            =   2520
      Picture         =   "frmOpcoes.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1530
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimi&r"
      Height          =   735
      Left            =   2520
      Picture         =   "frmOpcoes.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   690
      Width           =   1215
   End
   Begin VB.Frame fraOpcoes 
      Caption         =   "Impressão por:"
      Height          =   1575
      Left            =   240
      TabIndex        =   7
      Top             =   570
      Width           =   1935
      Begin VB.OptionButton optLote 
         Caption         =   "Lote da peça"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optPecas 
         Caption         =   "Código da peça"
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optGeral 
         Caption         =   "Todos"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox txtCodPeca 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   660
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   6
      Top             =   2325
      Width           =   1515
   End
   Begin VB.TextBox txtRegistro 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2325
      Width           =   855
   End
   Begin VB.TextBox txtQuantidade 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2685
      Width           =   855
   End
   Begin VB.CommandButton cmdPrimeiro 
      Caption         =   "Primeiro"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3210
      Width           =   975
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "Anterior"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   3210
      Width           =   975
   End
   Begin VB.CommandButton cmdProximo 
      Caption         =   "Próximo"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   3210
      Width           =   975
   End
   Begin VB.CommandButton cmdUltimo 
      Caption         =   "Último"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   3210
      Width           =   975
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   2370
      Width           =   495
   End
   Begin VB.Label lblRegistro 
      AutoSize        =   -1  'True
      Caption         =   "Registro"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2250
      TabIndex        =   12
      Top             =   2370
      Width           =   585
   End
   Begin VB.Label lblQuantidade 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade de etiquetas a imprimir"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   2730
      Width           =   2430
   End
End
Attribute VB_Name = "frmOpcoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "KERNEL32" _
        (ByVal dwMilliseconds As Long)
Public Peca As String
Public Lote As String

Public FlagGeral As Boolean
Public FlagPecas As Boolean
Public FlagLote As Boolean
Dim gNumeroArquivo As Integer
Dim ultimoTipoEtiqueta As String
Dim sEtiqueta As String
Dim oTela As frmExibicao12

Private Sub cmdAnterior_Click()
    dteEtiquetas.rsEtiquetas.MovePrevious
    If dteEtiquetas.rsEtiquetas.AbsolutePosition = 1 Then
        cmdPrimeiro.Enabled = False
        cmdAnterior.Enabled = False
    End If
    
    cmdProximo.Enabled = True
    cmdUltimo.Enabled = True
    
    MostraRegistroAtual
End Sub

Private Sub cmdImprimir_Click()

    Dim Vezes As Integer
    Dim nSequencial As Integer ' sequencial para impressao do codigo de barras
    Dim nx As Double
    Dim x As Printer
               
    nx = 0
    For Each x In Printers
       If InStr(1, UCase(x.DeviceName), UCase(Me.cbo_impressora.List(Me.cbo_impressora.ListIndex))) > 0 Then
          Set Printer = x
          nx = 1
          Exit For
       End If
    Next
    
    If nx = 0 Then
       MsgBox "Impressora da Etiqueta não encontrada, chame o responsável. o nome da impressora será - 'ETIQUETA FABRICA'"
       Exit Sub
    End If
    nx = 0
    
    'Comeca a imprimir do primeiro RecordSet
    dteEtiquetas.rsEtiquetas.MoveFirst
        
    While Not dteEtiquetas.rsEtiquetas.EOF
            
            Dim v As Integer
             
             If adMusashiDaAmazonia = 2 Then
'                If dteEtiquetas.rsEtiquetas.Fields("Tipo") = 7 Then
'                   Printer.Copies = Val(Me.txtQuantidade.Text)
'                Else
                   Printer.Copies = dteEtiquetas.rsEtiquetas.Fields("Qtd_Etiq")
'                End If
             End If

            For v = 0 To (Forms.Count - 1)

                If Forms(v).Name = "frmExibicao1" Or Forms(v).Name = "frmExibicao4" Or Forms(v).Name = "frmExibicao4MDA" Then
                   If objApplication.filial = adMusashiDaAmazonia Then
                      Printer.Orientation = 1
                   Else
                      Printer.Orientation = 2
                   End If
                   If Forms(v).Name = "frmExibicao4MDA" Then
                      frmExibicao4MDA.PrintForm
                   Else
                      frmAvulsoPadraoPonteiro.PrintForm
                   End If
                   Printer.Orientation = 2: Printer.EndDoc
                   Exit For
                End If
                
                If Forms(v).Name = "frmExibicao2" Then
                    If ultimoTipoEtiqueta = "F" Then
                       Call Imprime_Etiqueta_Fiat_FTP
                    ElseIf ultimoTipoEtiqueta = "Z" Then
                       Printer.Orientation = 1
                       frmExibicao2.nTamannhowidth = Me.Width
                       Rem retirado comentario "frmExibicao2.PrintForm" para esta emissão em 08/01/2020
                       frmExibicao2.PrintForm
                       Call Imprime_Etiqueta_Fiat_FTP
                       Call Imprime_Etiqueta_Fiat_FTP
                    Else
                       Printer.Orientation = 1
                       frmExibicao2.nTamannhowidth = Me.Width
                       frmExibicao2.PrintForm
                       oTela.Visible = True
                       oTela.Show
                       Printer.Orientation = 2
                       oTela.PrintForm
                       oTela.PrintForm
                       Unload oTela: Set oTela = Nothing
                       Printer.Orientation = 2: Printer.EndDoc
                    End If
                    Exit For
                End If
                If Forms(v).Name = "frmExibicao3" Then
'                    If frmExibicao3.PICT_CRISTAL.Visible = True Then
'                       Call Imprime_etiqueta_FORD_Cristal
'                    Else
                       Printer.Orientation = 2
                       frmExibicao3.PrintForm
                       frmExibicao3.PrintForm
                       Printer.Orientation = 2: Printer.EndDoc
'                    End If
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

                If Forms(v).Name = "frmExibicao7GM" Then
'                    Call Imprime_Etiqueta_GM(frmExibicao7Ref.DataToEncodeText2.Text)
                    Printer.Orientation = 2
                    Rem aqui marcos qdte de emissão de form
                    If Val(Me.txtQuantidade.Text) < 1 Then Me.txtQuantidade.Text = "1"
                    For Vezes = 0 To Val(Me.txtQuantidade.Text)
                        frmExibicao7GM.PrintForm
                        Printer.Orientation = 2: Printer.EndDoc
                    Next
                    Exit For
                    
'                    Printer.Copies = Val(Me.txtQuantidade.Text)
                End If
                
                If Forms(v).Name = "frmExibicao7GM2020" Or Forms(v).Name = "frmExibicao7GM2020_1" Then
                    Dim nTipo As Integer
                    
                    If Forms(v).Name = "frmExibicao7GM2020" Then nTipo = 1
                    
                    Printer.Orientation = 2
                    If Val(Me.txtQuantidade.Text) < 1 Then Me.txtQuantidade.Text = "1"
                    
                    For Vezes = 1 To Val(Me.txtQuantidade.Text)
                        If nTipo <> 1 Then
                           frmExibicao7GM2020_1.PrintForm
                        Else
                           frmExibicao7GM2020.PrintForm
                        End If
                        Printer.Orientation = 2: Printer.EndDoc
                    Next
                    Exit For
                    
'                    Printer.Copies = Val(Me.txtQuantidade.Text)
                End If

                If Forms(v).Name = "frmExibicao7GMNAC" Then
'                    Call Imprime_Etiqueta_GM(frmExibicao7Ref.DataToEncodeText2.Text)
                    Printer.Orientation = 2
                    frmExibicao7GMNAC.PrintForm
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
        'Next
        
        If dteEtiquetas.rsEtiquetas.AbsolutePosition <> dteEtiquetas.rsEtiquetas.RecordCount Then
            cmdProximo_Click
        Else
            dteEtiquetas.rsEtiquetas.MoveNext
        End If
        
    Wend
        
    If optGeral.Value = True Then
    
        MsgBox "Impressão concluída com sucesso! O aplicativo será encerrado!", vbOKOnly + vbInformation, "Tarefa Concluída"
        
        Rem abre o banco e seleciona os registros das etiquetas
        dteEtiquetas.rsEtiquetas.Close
        dteEtiquetas.rsEtiquetas.Source = "Select * from Etiquetas"
        dteEtiquetas.rsEtiquetas.Open
        dteEtiquetas.rsEtiquetas.MoveFirst
        
        Rem deleta todos os registros das etiquetas
        While Not dteEtiquetas.rsEtiquetas.EOF
            dteEtiquetas.rsEtiquetas.Delete
            dteEtiquetas.rsEtiquetas.Update
            dteEtiquetas.rsEtiquetas.MoveNext
        Wend
        dteEtiquetas.rsEtiquetas.Close
        
        If Dir("etiq.txt") = "etiq.txt" Then
            Close gNumeroArquivo
            Kill "etiq.txt"
        End If
        'Set objApplication = Nothing
        'End
        MDIEtiquetas.forcaSaida = True
        Unload MDIEtiquetas
        
    Else
        ' Usuário selecionou por código da peça - Deletar por código
        If optPecas.Value = True Then
            dteEtiquetas.rsEtiquetas.Close
            dteEtiquetas.rsEtiquetas.Source = "Select * from Etiquetas where Cod_Peca = '" & Peca & "'"
            dteEtiquetas.rsEtiquetas.Open
            dteEtiquetas.rsEtiquetas.MoveFirst
        
            While Not dteEtiquetas.rsEtiquetas.EOF
                dteEtiquetas.rsEtiquetas.Delete
                dteEtiquetas.rsEtiquetas.Update
                dteEtiquetas.rsEtiquetas.MoveNext
            Wend
        Else  'Usuário selecionou por lote - Deletar por lote
            dteEtiquetas.rsEtiquetas.Close
            dteEtiquetas.rsEtiquetas.Source = "Select * from Etiquetas where Lote = '" & Lote & "'"
            dteEtiquetas.rsEtiquetas.Open
            dteEtiquetas.rsEtiquetas.MoveFirst
        
            While Not dteEtiquetas.rsEtiquetas.EOF
                dteEtiquetas.rsEtiquetas.Delete
                dteEtiquetas.rsEtiquetas.Update
                dteEtiquetas.rsEtiquetas.MoveNext
            Wend
        End If
    End If
    
    If MDIEtiquetas.forcaSaida = False Then
        optGeral.SetFocus
        optGeral_Click
    End If
End Sub
Private Sub Imprime_Etiqueta_Fiat_FTP()

Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim nx As Double
Dim x As Printer
Dim rs As ADODB.Recordset
Dim cRec As ADODB.Recordset
Dim sData As String
Dim oTela As frmEscRelCristalReport

On Error GoTo Erro

Me.MousePointer = vbHourglass

Set cRec = New ADODB.Recordset

Set cRec = CCTempneMov_Etiq.Mov_Etiq_Consultar_FIAT_FTP(sBancoMusashi, _
                                                        dteEtiquetas.rsEtiquetas.Fields("id_etiqueta"))

If cRec.RecordCount = 0 Then
   MsgBox "Etiqueta não encontrada, anote a etiqueta e procure o responsável, etiq: " & dteEtiquetas.rsEtiquetas.Fields("id_etiqueta")
   Me.MousePointer = vbDefault
   Exit Sub
End If

nx = 0

On Error GoTo Erro
Set rs = New ADODB.Recordset

rs.Fields.Append "1_Num_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "2_Qtde_Emb", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "3_Classe_Func", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "4_Indicacao_Supl", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "5_Data_Fab_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "6_Cod_Emb", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "7_Vinculo", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "8_Lote_Sob_Desv", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "9_Qtde_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "10_Aplicacao", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "11_DUM", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "12_Embarque_Ctrl", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "13_Cod_Fornecedor", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "14_Num_Doc_Fis_BAM", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "15_Data", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "16_Pondo_Entrega", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "17_Denominacao", ADODB.DataTypeEnum.adChar, 80
rs.Fields.Append "18_Num_Desenho", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "19_Ctrl_Interno", ADODB.DataTypeEnum.adChar, 150
rs.Fields.Append "20_Ctrl_Oper_Log", ADODB.DataTypeEnum.adChar, 150
rs.Fields.Append "21_Codigo_Numero", ADODB.DataTypeEnum.adChar, 50
rs.Fields.Append "22_codigo_barras", ADODB.DataTypeEnum.adChar, 50

rs.Open
nx = 0

Rem aqui marcos 27/08/2020


cRec.MoveFirst
rs.AddNew
rs.Fields("1_Num_Lote").Value = IIf(IsNull(cRec.Fields("Num_Lote")), " ", Trim(cRec.Fields("Num_Lote")))
rs.Fields("2_Qtde_Emb").Value = IIf(IsNull(cRec.Fields("Qtde_Emb")), "0", Format(cRec.Fields("Qtde_Emb"), "00"))
rs.Fields("3_Classe_Func").Value = IIf(IsNull(cRec.Fields("Classe_Func")), " ", cRec.Fields("Classe_Func"))
rs.Fields("4_Indicacao_Supl").Value = IIf(IsNull(cRec.Fields("Indicacao_Supl")), " ", cRec.Fields("Indicacao_Supl"))
If Format(cRec.Fields("Data_Fab_Lote"), "DD/MM/YYYY") = "01/01/1900" Then
   rs.Fields("5_Data_Fab_Lote").Value = Format(Now(), "DD/MM/YYYY")
Else
   rs.Fields("5_Data_Fab_Lote").Value = IIf(IsNull(cRec.Fields("Data_Fab_Lote")), " ", Format(cRec.Fields("Data_Fab_Lote"), "DD/MM/YYYY"))
End If
rs.Fields("6_Cod_Emb").Value = IIf(IsNull(cRec.Fields("Cod_Emb")), " ", cRec.Fields("Cod_Emb"))
rs.Fields("7_Vinculo").Value = IIf(IsNull(cRec.Fields("Vinculo")), " ", cRec.Fields("Vinculo"))
rs.Fields("8_Lote_Sob_Desv").Value = IIf(IsNull(cRec.Fields("embalagem")), " ", cRec.Fields("embalagem"))
'rs.Fields("8_Lote_Sob_Desv").Value = IIf(IsNull(cRec.Fields("Lote_Sob_Desv")), " ", cRec.Fields("Lote_Sob_Desv"))
rs.Fields("9_Qtde_Lote").Value = IIf(IsNull(cRec.Fields("Qtde_Lote")), "0", cRec.Fields("Qtde_Lote"))
rs.Fields("10_Aplicacao").Value = IIf(IsNull(cRec.Fields("Aplicacao")), " ", Mid$(cRec.Fields("Aplicacao"), 1, 14))
If Format(cRec.Fields("DUM"), "DD/MM/YYYY") = "01/01/1900" Then
   rs.Fields("11_DUM").Value = " "
Else
   rs.Fields("11_DUM").Value = IIf(IsNull(cRec.Fields("DUM")), " ", Format(cRec.Fields("DUM"), "DD/MM/YYYY"))
End If
rs.Fields("12_Embarque_Ctrl").Value = IIf(IsNull(cRec.Fields("Embarque_Ctrl")), " ", cRec.Fields("Embarque_Ctrl"))
rs.Fields("13_Cod_Fornecedor").Value = "13093"
rs.Fields("14_Num_Doc_Fis_BAM").Value = IIf(IsNull(cRec.Fields("Num_Doc_Fis_BAM")), " ", cRec.Fields("Num_Doc_Fis_BAM"))
rs.Fields("15_Data").Value = IIf(IsNull(cRec.Fields("Data")), " ", cRec.Fields("Data"))
rs.Fields("16_Pondo_Entrega").Value = IIf(IsNull(cRec.Fields("Ponto_Entrega")), " ", Trim(cRec.Fields("Ponto_Entrega")))
rs.Fields("17_Denominacao").Value = IIf(IsNull(cRec.Fields("Denominacao")), " ", cRec.Fields("Denominacao"))
rs.Fields("18_Num_Desenho").Value = IIf(IsNull(cRec.Fields("Num_Desenho")), " ", Mid$(cRec.Fields("Num_Desenho"), 1, 20))
rs.Fields("19_Ctrl_Interno").Value = IIf(IsNull(cRec.Fields("Ctrl_Interno")), " ", cRec.Fields("Ctrl_Interno"))
rs.Fields("20_Ctrl_Oper_Log").Value = IIf(IsNull(cRec.Fields("Ctrl_Oper_Log")), " ", cRec.Fields("Ctrl_Oper_Log"))
rs.Fields("21_Codigo_Numero").Value = IIf(IsNull(cRec.Fields("Num_Desenho")), " ", cRec.Fields("Num_Desenho")) & Format(rs.Fields("2_Qtde_Emb").Value, "00000") & cRec.Fields("Cod_Emb") & "013093"
rs.Fields("22_codigo_barras").Value = IIf(IsNull(cRec.Fields("Num_Desenho")), " ", "*" & Mid$(cRec.Fields("Num_Desenho"), 1, 11) & Format(rs.Fields("2_Qtde_Emb").Value, "00000") & cRec.Fields("Cod_Emb") & "013093" & "*")
rs.Update

Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_FIAT.rpt")
CrystalReport1.Database.SetDataSource rs

rs.Clone

Rem mostrar a tela
'Set oTela = New frmEscRelCristalReport
'oTela.CRViewer1.ReportSource = CrystalReport1
'oTela.CRViewer1.ViewReport
'oTela.Show 0
Rem *************************

Rem nao mostrar a tela
CrystalReport1.PrintOutEx False
Rem *************************
'CrystalReport1.PrinterName = cbo_impressora.List(Me.cbo_impressora.ListIndex)
Rem VER AQUI MARCOS DEFAULT IMPRESSORA
Me.MousePointer = vbDefault

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub
Private Sub Imprime_Etiqueta_GM(ByVal sImagem_definicao As String)

Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim nx As Double
Dim x As Printer
Dim rs As ADODB.Recordset
Dim cRec As ADODB.Recordset
Dim sData As String
Dim oTela As frmEscRelCristalReport

On Error GoTo Erro

Me.MousePointer = vbHourglass

nx = 0

On Error GoTo Erro
Set rs = New ADODB.Recordset

rs.Fields.Append "Imagem_definicao", ADODB.DataTypeEnum.adChar, 200

rs.Open
nx = 0

rs.AddNew
rs.Fields("Imagem_definicao").Value = Trim(sImagem_definicao)
rs.Update

Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_GM.rpt")
CrystalReport1.Database.SetDataSource rs

rs.Clone

Rem mostrar a tela
Set oTela = New frmEscRelCristalReport
oTela.CRViewer1.ReportSource = CrystalReport1
oTela.CRViewer1.ViewReport
oTela.Show 0
Rem *************************

Rem nao mostrar a tela
'CrystalReport1.PrintOutEx False
Rem *************************

Me.MousePointer = vbDefault

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub


Private Sub cmdPrimeiro_Click()
    dteEtiquetas.rsEtiquetas.MoveFirst
    cmdPrimeiro.Enabled = False
    cmdAnterior.Enabled = False
    
    cmdProximo.Enabled = True
    cmdUltimo.Enabled = True
    
    MostraRegistroAtual
End Sub

Private Sub cmdProximo_Click()
    dteEtiquetas.rsEtiquetas.MoveNext
    If dteEtiquetas.rsEtiquetas.AbsolutePosition = dteEtiquetas.rsEtiquetas.RecordCount Then
        cmdProximo.Enabled = False
        cmdUltimo.Enabled = False
    End If

    cmdAnterior.Enabled = True
    cmdPrimeiro.Enabled = True
    
    MostraRegistroAtual
End Sub

Private Sub cmdUltimo_Click()
    dteEtiquetas.rsEtiquetas.MoveLast
    cmdUltimo.Enabled = False
    cmdProximo.Enabled = False
    
    cmdPrimeiro.Enabled = True
    cmdAnterior.Enabled = True
    
    MostraRegistroAtual
End Sub

Private Sub cmdVisualizar_Click()
    If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "1" Then
        frmAvulsoPadraoPonteiro.Show 'frmExibicao1.Show
        frmAvulsoPadraoPonteiro.Left = Me.Width
    ElseIf dteEtiquetas.rsEtiquetas.Fields("Tipo") = "2" Or _
           dteEtiquetas.rsEtiquetas.Fields("Tipo") = "F" Or _
           dteEtiquetas.rsEtiquetas.Fields("Tipo") = "Z" Then
        frmExibicao2.nTamannhowidth = Me.Width
        frmExibicao2.Show
    ElseIf dteEtiquetas.rsEtiquetas.Fields("Tipo") = "3" Then
        frmExibicao3.nTamWidth = Me.Width
        frmExibicao3.Show
    ElseIf dteEtiquetas.rsEtiquetas.Fields("Tipo") = "4" Or dteEtiquetas.rsEtiquetas.Fields("Tipo") = "8" Then
        frmAvulsoPadraoPonteiro.Show 'frmExibicao4.Show
        frmAvulsoPadraoPonteiro.Left = Me.Width
    ElseIf dteEtiquetas.rsEtiquetas.Fields("Tipo") = "A" Then
        frmExibicao6.Show
    ElseIf dteEtiquetas.rsEtiquetas.Fields("Tipo") = "M" Then
        frmExibicao4MDA.Show
    ElseIf dteEtiquetas.rsEtiquetas.Fields("Tipo") = "Y" Then
        If objApplication.filial = adMusashiDaAmazonia Then
           frmExibicao11.Show
        Else
           frmExibicao10.Show
        End If
    End If
End Sub

Private Sub Form_Load()
                                    
Dim x As Printer
Dim nx As Integer

For Each x In Printers
    If UCase(Mid$(x.DeviceName, 1, 8)) = "ETIQUETA" Then
       Me.cbo_impressora.AddItem x.DeviceName
    End If
Next

Rem verificar se ha impressoras cadastradas
If Me.cbo_impressora.ListCount = 0 Then
   MsgBox "Impressoras ETIQUETAS,não encontradas no sistema, Favor comunicar ao responsável para adiciona-las no sistema!"
   End
End If
Me.cbo_impressora.ListIndex = 0

Rem verificar a impressora padrão para ser usada neste relatório
For nx = 0 To Me.cbo_impressora.ListCount - 1
    If Trim(UCase(sImpressoraFabrica)) = Trim(UCase(Me.cbo_impressora.List(nx))) Then
       Me.cbo_impressora.ListIndex = nx
    End If
Next

'Habilita o flag geral, optGeral está selecionado
FlagGeral = True

Me.Left = 0
Me.Top = 0

If dteEtiquetas.rsEtiquetas.RecordCount = 1 Then
    cmdPrimeiro.Enabled = False
    cmdAnterior.Enabled = False
    cmdProximo.Enabled = False
    cmdUltimo.Enabled = False
End If

dteEtiquetas.rsEtiquetas.Close
dteEtiquetas.rsEtiquetas.Source = "Select * from Etiquetas order by Cod_Peca"
dteEtiquetas.rsEtiquetas.Open

If dteEtiquetas.rsEtiquetas.RecordCount > 0 Then
    dteEtiquetas.rsEtiquetas.MoveFirst
    MostraRegistroAtual
ElseIf dteEtiquetas.rsEtiquetas.RecordCount = 0 Then
    cmdPrimeiro.Enabled = False
    cmdAnterior.Enabled = False
    cmdProximo.Enabled = False
    cmdUltimo.Enabled = False
    cmdImprimir.Enabled = False
    cmdVisualizar.Enabled = False

    fraOpcoes.Enabled = False
    
    txtCodPeca.BackColor = vbButtonFace
    txtRegistro.BackColor = vbButtonFace
    txtQuantidade.BackColor = vbButtonFace
End If

End Sub

Public Sub MostraRegistroAtual()
    Dim sdata_aux As String
    Dim sdata_aux_Ano As String
    Dim qtdeConteiners As Integer
    Dim executaUnloadForm As Boolean
    Dim v As Integer
    Dim sSeqId As String * 18 ' usado para formatar o campo ID, da MVM
    Dim sQtdeCaixa As String
    Dim sdataAuxiliar As String
    Dim sTexto As String
    
    Rem PEGAR O NUMERO DA ETIQUETA
    sEtiqueta = dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA").Value
    
    If dteEtiquetas.rsEtiquetas.RecordCount = 1 Then
        cmdPrimeiro.Enabled = False
        cmdAnterior.Enabled = False
        cmdProximo.Enabled = False
        cmdUltimo.Enabled = False
    End If
    
    executaUnloadForm = False
    If Trim(ultimoTipoEtiqueta) <> "" Then
        If ultimoTipoEtiqueta <> dteEtiquetas.rsEtiquetas.Fields("Tipo").Value Then
            executaUnloadForm = True
        End If
    End If
    
    If executaUnloadForm = True Then
        'Fecha o form q estiver aberto
        For v = 0 To (Forms.Count - 1)
            If Forms(v).Name = "frmExibicao1" Or Forms(v).Name = "frmExibicao4" Then
                Unload frmAvulsoPadraoPonteiro
                Exit For
            End If
            If Forms(v).Name = "frmExibicao4MDA" Then
               Unload frmExibicao4MDA
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
            If Forms(v).Name = "frmExibicao6" Then
                Unload frmExibicao6
                Exit For
            End If
            If Forms(v).Name = "frmExibicao7UmProduto" Then
                Unload frmExibicao7UmProduto
                Exit For
            End If
            If Forms(v).Name = "frmExibicao7Ref" Then
                Unload frmExibicao7Ref
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
    End If
    
    ultimoTipoEtiqueta = dteEtiquetas.rsEtiquetas.Fields("Tipo").Value
    dteEtiquetas.rsEtiquetas.Fields("Tipo").Value = dteEtiquetas.rsEtiquetas.Fields("Tipo").Value
    
    'Atualiza o form frmOpcoes
    'Pode ser gravado nulo no caso da etiqueta N: 7 (Palete GM)
    
    Rem aqui txtCodPeca, tamanho = 7 para 8 (10/09/2010)
    Rem aqui txtCodPeca, tamanho = 8 para 9 (21/09/2012)

    If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("Cod_Peca")) Then
        txtCodPeca.Text = dteEtiquetas.rsEtiquetas.Fields("Cod_Peca")
    Else
        txtCodPeca.Text = ""
    End If
    txtQuantidade = dteEtiquetas.rsEtiquetas.Fields("Qtd_Etiq")
    txtRegistro = dteEtiquetas.rsEtiquetas.AbsolutePosition & "/" & dteEtiquetas.rsEtiquetas.RecordCount
    
    '---------------------------------------------------------------------------------------------------
    'Atualiza o form frmExibicao de acordo com o tipo
    'Se tipo 1 opcao padrão etiqueta pequena
    If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "1" Then
        frmAvulsoPadraoPonteiro.Show
        frmAvulsoPadraoPonteiro.Left = Me.Width
        If dteEtiquetas.rsEtiquetas.Fields("Cliente") <> "" Then
            frmAvulsoPadraoPonteiro.lblCliente.Caption = dteEtiquetas.rsEtiquetas.Fields("Cliente")
        End If
        
        If dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente") <> "" Then
            frmAvulsoPadraoPonteiro.lblCodCliente2.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Descr_Peca") <> "" Then
            frmAvulsoPadraoPonteiro.lblDescricao.Caption = dteEtiquetas.rsEtiquetas.Fields("Descr_Peca")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Lote") <> "" Then
            frmAvulsoPadraoPonteiro.lblLote2.Caption = dteEtiquetas.rsEtiquetas.Fields("Lote")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Peso") <> "" Then
            frmAvulsoPadraoPonteiro.lblPeso2.Caption = dteEtiquetas.rsEtiquetas.Fields("Peso")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa") <> "" Then
            frmAvulsoPadraoPonteiro.lblQtd2.Caption = dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Cod_Peca") <> "" Then
            frmAvulsoPadraoPonteiro.lblPeca.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Peca")
        End If
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")) Then
            frmAvulsoPadraoPonteiro.lblEmbalagem2.Caption = dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")
        End If
        Rem incluindo campos que vao fazer parte do novo padrao para a amd. 23-10-2012.
        
        frmAvulsoPadraoPonteiro.lbl_data.Caption = Format(dteEtiquetas.rsEtiquetas.Fields("data_etiq"), "dd/mm/yyyy")
        frmAvulsoPadraoPonteiro.LBL_MES.Caption = Pega_Mes(Mid$(Format(dteEtiquetas.rsEtiquetas.Fields("data_etiq"), "dd/mm/yyyy"), 4, 2))

        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA")) Then
            frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA"))
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
    If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "2" Or _
       dteEtiquetas.rsEtiquetas.Fields("Tipo") = "F" Or _
       dteEtiquetas.rsEtiquetas.Fields("Tipo") = "Z" Then
        frmExibicao2.nTamannhowidth = Me.Width
        frmExibicao2.Show
        'Mostra FIAT
        frmExibicao2.lblCod_Peca.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Peca")
        If dteEtiquetas.rsEtiquetas.Fields("Data_Expedicao") <> "" Then
            frmExibicao2.lblDataExpedicao2.Caption = dteEtiquetas.rsEtiquetas.Fields("Data_Expedicao")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor") <> "" Then
            frmExibicao2.lblCodFornec2.Caption = Format(dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor"), "000000")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Descr_Peca") <> "" Then
            frmExibicao2.lblDenominacao2.Caption = dteEtiquetas.rsEtiquetas.Fields("Descr_Peca")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Num_Doc_Fiscal") <> "" Then
            frmExibicao2.lblBam2.Caption = dteEtiquetas.rsEtiquetas.Fields("Num_Doc_Fiscal")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente") <> "" Then
            frmExibicao2.lblDesenho2.Caption = Format(dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente"), "00000000000")
        End If
        
        frmExibicao2.lblCodBarra.Caption = Format(dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente"), "00000000000") & Format(dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa"), "00000") _
                                      & dteEtiquetas.rsEtiquetas.Fields("Cod_Embalagem_pw") & Format(dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor"), "000000")
        frmExibicao2.lblCodBarra2.Caption = frmExibicao2.lblCodBarra.Caption
        'Adicionar o * de inicio e fim
        frmExibicao2.lblCodBarra.Caption = "*" & frmExibicao2.lblCodBarra.Caption & "*"
        frmExibicao2.lblCodBarraCp1.Caption = frmExibicao2.lblCodBarra.Caption
        frmExibicao2.lblCodBarraCp2.Caption = frmExibicao2.lblCodBarra.Caption
        
        If dteEtiquetas.rsEtiquetas.Fields("Data_Lote") <> "" Then
            frmExibicao2.lblDataProducao2.Caption = dteEtiquetas.rsEtiquetas.Fields("Data_Lote")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Cod_Embalagem") <> "" Then
            frmExibicao2.lblCodEmbalagem2.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Embalagem")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Lote") <> "" Then
            frmExibicao2.lblNumLote2.Caption = dteEtiquetas.rsEtiquetas.Fields("Lote")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Qtd_Lote") <> "" Then
            frmExibicao2.lblQtdLote2.Caption = dteEtiquetas.rsEtiquetas.Fields("Qtd_Lote")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa") <> "" Then
            frmExibicao2.lblQtdEmbalagem2.Caption = Format(dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa"), "00000")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Classe_Funcional") <> "" Then
            frmExibicao2.lblClasseFuncional2.Caption = dteEtiquetas.rsEtiquetas.Fields("Classe_Funcional")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Vinculo") <> "" Then
            frmExibicao2.lblVinculo2.Caption = dteEtiquetas.rsEtiquetas.Fields("Vinculo")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Ind_Suplementar") <> "" Then
            frmExibicao2.lblIndicacaoSuplementar2.Caption = dteEtiquetas.rsEtiquetas.Fields("Ind_Suplementar")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Embarque_Controlado") <> "" Then
            frmExibicao2.lblEmbarqueControlado2.Caption = dteEtiquetas.rsEtiquetas.Fields("Embarque_Controlado")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Desvio") <> "" Then
            frmExibicao2.lblLoteSobDesvio2.Caption = dteEtiquetas.rsEtiquetas.Fields("Desvio")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("DUM") <> "" Then
            frmExibicao2.lblDum2.Caption = dteEtiquetas.rsEtiquetas.Fields("DUM")
        End If
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")) Then
            frmExibicao2.lblEmbalagem2.Caption = dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")
        Else
            frmExibicao2.lblEmbalagem2.Caption = ""
        End If
    Rem acrescentado o pto de entrega - 09-09-2016
        If dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega") <> "" Then
            frmExibicao2.lblPontoEntrega2.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega"))
        End If
        
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA")) Then
            frmExibicao2.lblCodigoBarras.Caption = ""
            frmExibicao2.lblCodigoBarras.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA"))
            frmExibicao2.lblCodigoBarrasA.Caption = "*" & Trim(frmExibicao2.lblCodigoBarras.Caption) & "*"
            frmExibicao2.lblCodigoBarrasB.Caption = "*" & Trim(frmExibicao2.lblCodigoBarras.Caption) & "*"
'            frmExibicao2.lblCodigoBarrasC.Caption = "*" & frmExibicao2.lblCodigoBarras.Caption & "*"
'            frmExibicao2.lblCodigoBarrasD.Caption = "*" & frmExibicao2.lblCodigoBarras.Caption & "*"
         Else
            frmExibicao2.lblCodigoBarras.Caption = ""
            frmExibicao2.lblCodigoBarrasA.Caption = ""
            frmExibicao2.lblCodigoBarrasB.Caption = ""
         End If
         
         If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "2" Then ' aqui será preenchido os dados para a outra etiqueta tipo expotacao da fiat italiana

            Set oTela = New frmExibicao12
            oTela.slbl_01 = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Cliente")), " ", Trim(dteEtiquetas.rsEtiquetas.Fields("Cliente")))
            oTela.slbl_02 = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")), " ", Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")))
            oTela.slbl_03 = " "
            oTela.slbl_04 = "MUSASHI DO BRASIL LTDA"
            oTela.slbl_07 = " "
            oTela.slbl_08 = " " ' FALTA VER COM MAURO
            oTela.slbl_09 = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")), " ", Trim(dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")))
            oTela.slbl_10 = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Descr_Peca")), " ", Trim(dteEtiquetas.rsEtiquetas.Fields("Descr_Peca")))
            oTela.slbl_11 = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor")), " ", Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor")))
            oTela.slbl_12 = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Cod_Embalagem")), " ", Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Embalagem")))
            oTela.slbl_13 = Format(Now(), "DD/MM/YYYY")
            oTela.slbl_14 = " "
            oTela.slbl_15 = " "
            oTela.slbl_16 = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Lote")), " ", Trim(dteEtiquetas.rsEtiquetas.Fields("Lote")))
            oTela.ldl_usuario.Caption = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")), " ", dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM"))
            oTela.lbl_sequencial.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA"))
            oTela.lbl_barras1.Caption = "*" & oTela.lbl_sequencial.Caption & "*"
            oTela.lbl_barras2.Caption = "*" & oTela.lbl_sequencial.Caption & "*"
            oTela.Visible = False
            
         End If
    
         If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "Z" Then ' aqui será preenchido os dados para a outra etiqueta tipo expotacao da fiat italiana

            Set oTela = New frmExibicao12
            oTela.Visible = False
            oTela.slbl_01 = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Cliente")), " ", Trim(dteEtiquetas.rsEtiquetas.Fields("Cliente")))
            oTela.slbl_02 = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")), " ", Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")))
            oTela.slbl_03 = " "
            oTela.slbl_04 = "MUSASHI DO BRASIL LTDA"
            oTela.slbl_07 = " "
            oTela.slbl_08 = " " ' FALTA VER COM MAURO
            oTela.slbl_09 = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")), " ", Trim(dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")))
            oTela.slbl_10 = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Descr_Peca")), " ", Trim(dteEtiquetas.rsEtiquetas.Fields("Descr_Peca")))
            oTela.slbl_11 = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor")), " ", Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor")))
            oTela.slbl_12 = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Cod_Embalagem")), " ", Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Embalagem")))
            oTela.slbl_13 = Format(Now(), "DD/MM/YYYY")
            oTela.slbl_14 = " "
            oTela.slbl_15 = " "
            oTela.slbl_16 = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Lote")), " ", Trim(dteEtiquetas.rsEtiquetas.Fields("Lote")))
            oTela.ldl_usuario.Caption = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")), " ", dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM"))
            oTela.lbl_sequencial.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA"))
            oTela.lbl_barras1.Caption = "*" & oTela.lbl_sequencial.Caption & "*"
            oTela.lbl_barras2.Caption = "*" & oTela.lbl_sequencial.Caption & "*"
            oTela.Visible = False
         End If
    
    End If
    
    
    
    '---------------------------------------------------------------------------------------------------
    'Se tipo 3 opcao FORD
    If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "3" Then
        
        sdata_aux = Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 1, 2) & _
                    Pega_Mes(Val(Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 4, 2))) & _
                    Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 7, 4)
        frmExibicao3.lbl_Cliente.Caption = "MUSASHI DO BRASIL LTDA" 'Trim(dteEtiquetas.rsEtiquetas.Fields("Cliente"))
        frmExibicao3.lbl_Cod_Fornecedor.Caption = "CFEOA"  'Left(dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor"), 16)
        frmExibicao3.lbl_Cod_Fornecedor_Barras.Caption = "CFEOA" 'Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor"))
        frmExibicao3.lbl_qtd.Caption = Left(Trim(dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")), 16)
        frmExibicao3.lbl_qtd_barras.Caption = Left(Trim(dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")), 16)
        frmExibicao3.lbl_peso.Caption = Left(Trim(dteEtiquetas.rsEtiquetas.Fields("Peso")), 16)
        frmExibicao3.lbl_container.Caption = "KLT 4314 CFEOA" 'Left(Trim(dteEtiquetas.rsEtiquetas.Fields("Lote"), 16))
        frmExibicao3.lbl_lote.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("Lote"))
        frmExibicao3.lbl_data.Caption = sdata_aux
        frmExibicao3.lbl_Cod_cliente.Caption = Left(Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente")), 16)
        frmExibicao3.lbl_cod_cliente_1.Caption = Left(Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente")), 16)
        frmExibicao3.lbl_cod_cliente_Barras.Caption = Left(Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente")), 16)
        frmExibicao3.lbl_Cod_Peca.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Peca"))
        frmExibicao3.lbl_line_feed_loc2.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Peca")) & " " & Trim(dteEtiquetas.rsEtiquetas.Fields("Lote"))

        frmExibicao3.lbl_descr_peca.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("Descr_Peca"))
        frmExibicao3.lbl_id_etiqueta.Caption = Left(Trim(dteEtiquetas.rsEtiquetas.Fields("id_etiqueta")), 16)
        frmExibicao3.lbl_id_etiqueta_barra.Caption = Left(Trim(dteEtiquetas.rsEtiquetas.Fields("id_etiqueta")), 16)
        sdata_aux = Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 9, 2) & _
                    Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 4, 2) & _
                    Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 1, 2)
        frmExibicao3.DataToEncodeText.Text = Format(dteEtiquetas.rsEtiquetas.Fields("id_etiqueta"), "00000") & " (P)" & _
                                             Format(Trim(dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")), "0") & " (Q)" & _
                                             Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor")) & " (V)" & _
                                             sdata_aux & " (D)" & _
                                             Format(Trim(dteEtiquetas.rsEtiquetas.Fields("Serial")), "000") & " (S)"
        
        frmExibicao3.lbl_to.Caption = "FORD TAUBATE"
        frmExibicao3.lbl_cust.Caption = "FI05D"
        frmExibicao3.lbl_doc_code.Caption = "R3"
    
        frmExibicao3.PDF1.DataToEncode = frmExibicao3.DataToEncodeText.Text
'        frmExibicao3.PDF1.Height = frmExibicao3.ImageHeight.Text
'        frmExibicao3.PDF1.Width = frmExibicao3.ImageWidth.Text
'        frmExibicao3.PDF1.PDFColumns = frmExibicao3.PDFColumns.Text
'        frmExibicao3.PDF1.PDFErrorCorrectionLevel = frmExibicao3.PDFErrorCorrectionLevel.Text
'        frmExibicao3.PDF1.NarrowBarCM = frmExibicao3.NarrowBarWidth.Text
'        frmExibicao3.PDF1.TopMarginCM = frmExibicao3.TopMarginCM.Text
'        frmExibicao3.PDF1.LeftMarginCM = frmExibicao3.LeftMarginCM.Text
        frmExibicao3.nTamWidth = Me.Width
        frmExibicao3.Show


'        frmExibicao3.lblNumPecaA.Caption = "*P" & Trim(Left(dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente"), 16)) & "*"
'        frmExibicao3.lblNumPecaB.Caption = frmExibicao3.lblNumPecaA.Caption
'        frmExibicao3.lblNumPeca.Caption = Trim(frmExibicao3.lblNumPeca.Caption)
'        frmExibicao3.lblCod_Peca = dteEtiquetas.rsEtiquetas.Fields("Cod_Peca")
'        frmExibicao3.lblLote = dteEtiquetas.rsEtiquetas.Fields("Lote")
'        frmExibicao3.lblQtd.Caption = dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")
'        frmExibicao3.lblQtdA.Caption = "*Q" & dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa") & "*"
'        frmExibicao3.lblQtdB.Caption = frmExibicao3.lblQtdA.Caption
'        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor")) Then
'            frmExibicao3.lblNumFornec.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor")
'            frmExibicao3.lblNumFornecA.Caption = "*V" & dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor") & "*"
'            frmExibicao3.lblNumFornecB.Caption = frmExibicao3.lblNumFornecA.Caption
'        End If
'        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("Serial")) Then
'            frmExibicao3.lblNumSerial.Caption = dteEtiquetas.rsEtiquetas.Fields("Serial")
'            frmExibicao3.lblNumSerialA.Caption = "*S" & dteEtiquetas.rsEtiquetas.Fields("Serial") & "*"
'            frmExibicao3.lblNumSerialB.Caption = frmExibicao3.lblNumSerialA.Caption
'        End If
'        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("Cod_Util")) Then
'            frmExibicao3.lblCodUtil.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Util")
'        End If
'        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("Linha_Util")) Then
'            frmExibicao3.lblLinhaUtil.Caption = dteEtiquetas.rsEtiquetas.Fields("Linha_Util")
'        End If
'        frmExibicao3.lblSufixo.Caption = Trim(Right((dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente")), 4))
'        frmExibicao3.lblSufixoA.Caption = "*C" & frmExibicao3.lblSufixo.Caption & "*"
'        frmExibicao3.lblSufixoB.Caption = frmExibicao3.lblSufixoA.Caption
'        If (dteEtiquetas.rsEtiquetas.Fields("Desvio")) <> "" Then
'            frmExibicao3.lblDestino.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("Desvio"))
'        End If
'        frmExibicao3.lblDestinoA.Caption = "*D" & frmExibicao3.lblDestino.Caption & "*"
'        frmExibicao3.lblDestinoB.Caption = frmExibicao3.lblDestinoA.Caption
'
'        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")) Then
'            frmExibicao3.lblEmbalagem2.Caption = dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")
'        End If
'
'        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA")) Then
'            frmExibicao3.lblCodigoBarras.Caption = dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA")
'            frmExibicao3.lblCodigoBarrasA.Caption = "*" & frmExibicao3.lblCodigoBarras.Caption & "*"
'            frmExibicao3.lblCodigoBarrasB.Caption = "*" & frmExibicao3.lblCodigoBarras.Caption & "*"
'         Else
'            frmExibicao3.lblCodigoBarras.Caption = ""
'            frmExibicao3.lblCodigoBarrasA.Caption = ""
'            frmExibicao3.lblCodigoBarrasB.Caption = ""
'         End If
        
        
    End If
    
    '---------------------------------------------------------------------------------------------------
    'Se tipo 4 ou 8 opcao padrão etiqueta grande
    If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "4" Or _
       dteEtiquetas.rsEtiquetas.Fields("Tipo") = "8" Then
        'frmExibicao4.Show
        frmAvulsoPadraoPonteiro.Show
        frmAvulsoPadraoPonteiro.Left = Me.Width
        If dteEtiquetas.rsEtiquetas.Fields("Cliente") <> "" Then
            'frmExibicao4.lblCliente.Caption = dteEtiquetas.rsEtiquetas.Fields("Cliente")
            frmAvulsoPadraoPonteiro.lblCliente.Caption = dteEtiquetas.rsEtiquetas.Fields("Cliente")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente") <> "" Then
            'frmExibicao4.lblCodCliente2.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente")
            frmAvulsoPadraoPonteiro.lblCodCliente2.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Descr_Peca") <> "" Then
            'frmExibicao4.lblDescricao.Caption = dteEtiquetas.rsEtiquetas.Fields("Descr_Peca")
            frmAvulsoPadraoPonteiro.lblDescricao.Caption = dteEtiquetas.rsEtiquetas.Fields("Descr_Peca")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Lote") <> "" Then
            'frmExibicao4.lblLote2.Caption = dteEtiquetas.rsEtiquetas.Fields("Lote")
            frmAvulsoPadraoPonteiro.lblLote2.Caption = dteEtiquetas.rsEtiquetas.Fields("Lote")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Peso") <> "" Then
            'frmExibicao4.lblPeso2.Caption = dteEtiquetas.rsEtiquetas.Fields("Peso")
            frmAvulsoPadraoPonteiro.lblPeso2.Caption = dteEtiquetas.rsEtiquetas.Fields("Peso")
        End If
        
        If dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa") <> "" Then
            'frmExibicao4.lblQtd2.Caption = dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")
            frmAvulsoPadraoPonteiro.lblQtd2.Caption = dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")
        End If
        
        If dteEtiquetas.rsEtiquetas.Fields("Cod_Peca") <> "" Then
            'frmExibicao4.lblPeca.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Peca")
            frmAvulsoPadraoPonteiro.lblPeca.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Peca")
        End If
        
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")) Then
            frmAvulsoPadraoPonteiro.lblEmbalagem2.Caption = dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")
        End If
     
        frmAvulsoPadraoPonteiro.lbl_data.Caption = Format(dteEtiquetas.rsEtiquetas.Fields("data_etiq"), "dd/mm/yyyy")
        frmAvulsoPadraoPonteiro.LBL_MES.Caption = Pega_Mes(Mid$(Format(dteEtiquetas.rsEtiquetas.Fields("data_etiq"), "dd/mm/yyyy"), 4, 2))
        
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA")) Then
           frmAvulsoPadraoPonteiro.lbl_Seq_Milhar.Caption = Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA")), _
                                                            Len(Trim(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA"))) - 5, _
                                                            Len(Trim(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA"))))
           frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption = dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA")
           frmAvulsoPadraoPonteiro.lblCodigoBarrasA.Caption = "*" & frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption & "*"
           frmAvulsoPadraoPonteiro.lblCodigoBarrasB.Caption = "*" & frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption & "*"
        Else
           frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption = ""
           frmAvulsoPadraoPonteiro.lblCodigoBarrasA.Caption = ""
           frmAvulsoPadraoPonteiro.lblCodigoBarrasB.Caption = ""
        End If
        
        If (dteEtiquetas.rsEtiquetas.Fields("Tipo") = "8") Then
           frmAvulsoPadraoPonteiro.lblCodCliente1.Caption = "CODE:"
           frmAvulsoPadraoPonteiro.lblQtd1.Caption = "QTY.:"
           frmAvulsoPadraoPonteiro.lblPeso1.Caption = "WEIGHT:"
           frmAvulsoPadraoPonteiro.lblLote1.Caption = "LOT.:"
        Else
           If objApplication.filial <> adMusashiDaAmazonia Then
              If (dteEtiquetas.rsEtiquetas.Fields("indforjimport") = "X") Then
                 frmAvulsoPadraoPonteiro.lbl_kms.Visible = True
                 frmAvulsoPadraoPonteiro.lbl_tarja_kms.Visible = True
              Else
                 frmAvulsoPadraoPonteiro.lbl_kms.Visible = False
                 frmAvulsoPadraoPonteiro.lbl_tarja_kms.Visible = False
              End If
           End If
        End If
        
        frmAvulsoPadraoPonteiro.lblMsgProduto.Visible = False
'        frmAvulsoPadraoPonteiro.lbl_cli.Caption = dteEtiquetas.rsEtiquetas.Fields("id_cliente")
        If Val(dteEtiquetas.rsEtiquetas.Fields("id_cliente")) = 1 Or _
           Val(dteEtiquetas.rsEtiquetas.Fields("id_cliente")) = 2 Or _
           Val(dteEtiquetas.rsEtiquetas.Fields("id_cliente")) = 3 Then
           frmAvulsoPadraoPonteiro.lblMsgProduto.Visible = True
        End If
         
    End If
    
    
    '---------------------------------------------------------------------------------------------------
    'Se tipo 5 opcao GM etiqueta grande
    'MODIFICADO EM 23-08-2007 (MARCOS PEDROSA) ACRESCENTAR CAMPOS LAY-OUT NOVO
    If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "5" Then
        frmExibicao5.tWithOpcoes = Me.Width
        frmExibicao5.Show
        If dteEtiquetas.rsEtiquetas.Fields("Cliente") <> "" Then
            frmExibicao5.lblTo.Caption = dteEtiquetas.rsEtiquetas.Fields("Cliente")
        End If
        Rem incluido (marcos pedrosa) em 23-08-2007
        If dteEtiquetas.rsEtiquetas.Fields("MOTIVO_ALTERACAO_OUTROS") <> "" Then
            frmExibicao5.lblMotivo_alteracao_outros.Caption = Trim(dteEtiquetas.rsEtiquetas("MOTIVO_ALTERACAO_OUTROS").Value)
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Ind_Suplementar") <> "" Then
            frmExibicao5.lblPlant.Caption = dteEtiquetas.rsEtiquetas.Fields("Ind_Suplementar")
        End If
        Rem DEFINICAO ATE 23-08-2007 (MARCOS PEDROSA)
'''        If dteEtiquetas.rsEtiquetas.Fields("Cod_Util") <> "" Then
'''            frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & dteEtiquetas.rsEtiquetas.Fields("Cod_Util")
'''        End If
'''        If dteEtiquetas.rsEtiquetas.Fields("Desvio") <> "" Then
'''            frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & dteEtiquetas.rsEtiquetas.Fields("Desvio")
'''        End If
        Rem INCLUIDO MARCOS PEDROSA EM 23-08-2007
        If dteEtiquetas.rsEtiquetas.Fields("ind_suplementar") <> "" Then
            frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & dteEtiquetas.rsEtiquetas.Fields("ind_suplementar")
        End If
        Rem INCLUIDO MARCOS PEDROSA EM 23-08-2007
'        If dteEtiquetas.rsEtiquetas.Fields("Desvio") <> "" Then
'            frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & dteEtiquetas.rsEtiquetas.Fields("Desvio")
'        End If

        If dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente") <> "" Then
            frmExibicao5.lblPartNumber.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Descr_Peca") <> "" Then
            frmExibicao5.lblPeca.Caption = dteEtiquetas.rsEtiquetas.Fields("Descr_Peca")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa") <> 0 Then
            frmExibicao5.lblQtd.Caption = dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Modelo") <> "" Then
            frmExibicao5.lblMaterial.Caption = dteEtiquetas.rsEtiquetas.Fields("Modelo")
        End If
        frmExibicao5.lblReference.Caption = ""
'        If dteEtiquetas.rsEtiquetas.Fields("Cod_Embalagem_Pw") <> "" Then
'            frmExibicao5.lblReference.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Embalagem_Pw")
'        End If
        If dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega") <> "" Then
            frmExibicao5.lblLicenseA.Caption = "J1" & dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega") & Format(dteEtiquetas.rsEtiquetas.Fields("id_etiqueta"), "000000000")
'            frmExibicao5.lblLicenseA.Caption = frmExibicao5.lblLicense.Caption
            frmExibicao5.lblLicenseB.Caption = "J1" & dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega") & Format(dteEtiquetas.rsEtiquetas.Fields("id_etiqueta"), "000000000")
            frmExibicao5.DataToEncodeText.Text = frmExibicao5.lblLicenseA.Caption
            
        End If
        Rem comentado em 23-08-2007 (marcos pedrosa)
        If dteEtiquetas.rsEtiquetas.Fields("Cod_Embalagem") <> "" Then
            frmExibicao5.lblContainerType.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Embalagem")
        End If
        Rem INCLUIDO EM 23-08-2007
        If dteEtiquetas.rsEtiquetas.Fields("compl_peca1") <> "" Then
            frmExibicao5.lblContainerType.Caption = dteEtiquetas.rsEtiquetas.Fields("compl_peca1")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Peso") <> 0 Then
            frmExibicao5.lblgrossWeight.Caption = dteEtiquetas.rsEtiquetas.Fields("Peso")
        End If
'        If dteEtiquetas.rsEtiquetas.Fields("compl_peca2") <> "" Then
        sdata_aux = Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 1, 2) & _
                    Pega_Mes(Val(Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 4, 2)))
        frmExibicao5.lblRoute.Caption = sdata_aux
        frmExibicao5.lblRoute1.Caption = Mid$(dteEtiquetas.rsEtiquetas.Fields("data_etiq"), 7, 4)
'            frmExibicao5.lblRoute.Caption = dteEtiquetas.rsEtiquetas.Fields("compl_peca2")
'        End If
        If dteEtiquetas.rsEtiquetas.Fields("Lote") <> "" Then
            frmExibicao5.lblLot.Caption = dteEtiquetas.rsEtiquetas.Fields("Lote")
        End If
'        If dteEtiquetas.rsEtiquetas.Fields("id_up") <> "" Then
'            frmExibicao5.lblParts.Caption = dteEtiquetas.rsEtiquetas.Fields("id_up")
'        End If
        Rem comentado em 23-08-2007 (marcos pedrosa)
'''        If dteEtiquetas.rsEtiquetas.Fields("Dum") <> "" Then
'''            frmExibicao5.lblEng.Caption = dteEtiquetas.rsEtiquetas.Fields("Dum")
'''        End If
        Rem INCLUIDO EM 23-08-2007
        If dteEtiquetas.rsEtiquetas.Fields("data_lote") <> "" Then
            frmExibicao5.lblEng.Caption = Mid$(dteEtiquetas.rsEtiquetas.Fields("data_lote"), 1, 2) & _
                                          Pega_Mes(Val(Mid$(dteEtiquetas.rsEtiquetas.Fields("data_lote"), 4, 2))) & _
                                          Mid$(dteEtiquetas.rsEtiquetas.Fields("data_lote"), 7, 4)
        End If
        Rem INCLUIDO EM 23-08-2007
        If dteEtiquetas.rsEtiquetas.Fields("envio_lote") = "1" Then
            frmExibicao5.lblvalidade.Caption = "N"
        Else
            frmExibicao5.lblvalidade.Caption = ""
        End If
        
'        If dteEtiquetas.rsEtiquetas.Fields("id_mfg") <> "" Then
'            frmExibicao5.lblMfgDate.Caption = Mid$(dteEtiquetas.rsEtiquetas.Fields("id_mfg"), 1, 2) & _
'                                              Pega_Mes(Val(Mid$(dteEtiquetas.rsEtiquetas.Fields("id_mfg"), 4, 2))) & _
'                                              Mid$(dteEtiquetas.rsEtiquetas.Fields("id_mfg"), 7, 4)
'        End If
        If dteEtiquetas.rsEtiquetas.Fields("Cod_Peca") <> "" Then
            frmExibicao5.lblCodMSB.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Peca")
        End If
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")) Then
            frmExibicao5.lblEmbalagem2.Caption = dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")
        End If
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA")) Then
            frmExibicao5.lblCodigoBarras.Caption = dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA")
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
    If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "6" Or _
       dteEtiquetas.rsEtiquetas.Fields("Tipo") = "9" Then
        frmExibicao9.Show
        If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "9" Then
           frmExibicao9.PICT_CRISTAL.Visible = True
        Else
           frmExibicao9.PICT_CRISTAL.Visible = False
        End If
        frmExibicao9.Left = Me.Width
        If dteEtiquetas.rsEtiquetas.Fields("Cliente") <> "" Then
            frmExibicao9.lblCliente.Caption = dteEtiquetas.rsEtiquetas.Fields("Cliente")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente") <> "" Then
            frmExibicao9.lblCodBar_Cod_cliente.Caption = "*" & Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente")) & "*"
            frmExibicao9.lblCodBar_Cod_cliente1.Caption = "*" & Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente")) & "*"
            frmExibicao9.lbl_Cod_cliente.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente"))
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Descr_Peca") <> "" Then
            frmExibicao9.lbl_desc_peca.Caption = dteEtiquetas.rsEtiquetas.Fields("Descr_Peca")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("cod_fornecedor") <> "" Then
            frmExibicao9.lblCod_Fornecedor.Caption = dteEtiquetas.rsEtiquetas.Fields("cod_fornecedor")
        End If

        frmExibicao9.lbl_Fornecedor.Caption = "MUSASHI"
        If dteEtiquetas.rsEtiquetas.Fields("data_expedicao") <> "" Then
            frmExibicao9.lbl_data_expedicao.Caption = dteEtiquetas.rsEtiquetas.Fields("data_expedicao")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("qtd_caixa") <> "" Then
            frmExibicao9.lblCodBarQtd_caixa.Caption = "*" & dteEtiquetas.rsEtiquetas.Fields("qtd_caixa") & "*"
            frmExibicao9.lblqtd_caixa.Caption = dteEtiquetas.rsEtiquetas.Fields("qtd_caixa")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Lote") <> "" Then
            frmExibicao9.lblCodBar_lote.Caption = "*" & dteEtiquetas.rsEtiquetas.Fields("Lote") & "*"
            frmExibicao9.lblLote.Caption = dteEtiquetas.rsEtiquetas.Fields("Lote")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Cod_Peca") <> "" Then
            frmExibicao9.lbl_id_etiqueta.Caption = dteEtiquetas.rsEtiquetas.Fields("id_etiqueta")
            frmExibicao9.lblCodBar_Cod_Peca.Caption = "*" & dteEtiquetas.rsEtiquetas.Fields("id_etiqueta") & "*"
            frmExibicao9.lbl_Cod_Peca.Caption = dteEtiquetas.rsEtiquetas.Fields("cod_peca")
        End If
        
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")) Then
            frmExibicao9.lblCodFunc.Caption = dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")
        Else
            frmExibicao9.lblCodFunc.Caption = ""
        End If
        
        Rem preparar o "id" da etiqueta com sua formatacao conforme documento fornecedor(6)+dtEtiqueta(6-aammdd)+seq. do dia
        sSeqId = " "
        
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("desvio_aviso_mod")) Then
'            frmExibicao9.lblDesvio_Aviso_Mod.Caption = dteEtiquetas.rsEtiquetas.Fields("desvio_aviso_mod")
            sSeqId = Mid$(dteEtiquetas.rsEtiquetas.Fields("desvio_aviso_mod"), 1, 6) ' fornecedor
            sSeqId = Trim(sSeqId) & Mid$(dteEtiquetas.rsEtiquetas.Fields("desvio_aviso_mod"), 13, 2) ' (aa)mmdd
            sSeqId = Trim(sSeqId) & Mid$(dteEtiquetas.rsEtiquetas.Fields("desvio_aviso_mod"), 9, 2)  ' aa(mm)dd
            sSeqId = Trim(sSeqId) & Mid$(dteEtiquetas.rsEtiquetas.Fields("desvio_aviso_mod"), 7, 2) ' aamm(dd)
            sSeqId = Trim(sSeqId) & Format(Mid$(dteEtiquetas.rsEtiquetas.Fields("desvio_aviso_mod"), 15, 11), "000000") ' sequencial
            
            frmExibicao9.lblCodBar_Desvio_Aviso_Mod.Caption = "*" & sSeqId & "*"
            frmExibicao9.lblCodBar_Desvio_Aviso_Mod1.Caption = "*" & sSeqId & "*"
            frmExibicao9.lblDesvio_Aviso_Mod.Caption = sSeqId
            
        End If
         
    End If
'**********************************************************************************************************************
'**********************************************************************************************************************
'**********************************************************************************************************************
'-----------------------------------------------------------------------------------------------------
    
    'Se tipo 7 opcao Palete GM
    If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "7" Then
        If Mid$(dteEtiquetas.rsEtiquetas.Fields("cod_peca").Value, 1, 1) = "6" Then
           Set frmExibicao7Ref = frmExibicao7GM2020
        Else
           Set frmExibicao7Ref = frmExibicao7GM2020_1
        End If
'        Set frmExibicao7Ref = frmExibicao7GM
        frmExibicao7Ref.Show
        
        frmExibicao7Ref.lblPlant.Caption = "4J D30"
        
        ' Ajustado para emitir mais de um palet , liberando o campo quantidade
        Me.txtQuantidade.Locked = False
        Me.txtQuantidade.Enabled = True

        If dteEtiquetas.rsEtiquetas.Fields("cod_peca").Value = "2G01600" Then
           frmExibicao7Ref.lblCodMSB.Caption = "TAF1V22"
           frmExibicao7Ref.lblContainerType.Caption = "CX171203"
        ElseIf dteEtiquetas.rsEtiquetas.Fields("cod_peca").Value = "2G05601" Then
           frmExibicao7Ref.lblCodMSB.Caption = "TAF1V24"
           frmExibicao7Ref.lblContainerType.Caption = "CX151203"
        ElseIf dteEtiquetas.rsEtiquetas.Fields("cod_peca").Value = "2G01602" Then
           frmExibicao7Ref.lblCodMSB.Caption = "TAF1V11"
           frmExibicao7Ref.lblContainerType.Caption = "CX171203"
        ElseIf dteEtiquetas.rsEtiquetas.Fields("cod_peca").Value = "2G06600" Then
           frmExibicao7Ref.lblCodMSB.Caption = "TAF1V11"
           frmExibicao7Ref.lblContainerType.Caption = "CX171203"
        ElseIf dteEtiquetas.rsEtiquetas.Fields("cod_peca").Value = "6G04532" Then
           frmExibicao7Ref.lblCodMSB.Caption = "TAF4V02"
           frmExibicao7Ref.lblContainerType.Caption = "CX484022"
        Else
           frmExibicao7Ref.lblCodMSB.Caption = dteEtiquetas.rsEtiquetas.Fields("cod_peca").Value
           frmExibicao7Ref.lblContainerType.Caption = "CX"
        End If
        
        'Aqui terá algumas mudanças referentes ao codigo do produto e com quantidades fixas
        sQtdeCaixa = ""
        If IsNumeric(dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa").Value) Then
           frmExibicao7Ref.lblQtde1.Caption = Format(Val(dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa").Value), "000")
        Else
           frmExibicao7Ref.lblQtde1.Caption = "0"
        End If
        sQtdeCaixa = Trim(frmExibicao7Ref.lblQtde1.Caption)

'lblMaterial
        Rem INCLUIDO MARCOS PEDROSA EM 28-08-2007
        frmExibicao7Ref.lblMaterial.Caption = dteEtiquetas.rsEtiquetas.Fields("cod_cliente").Value
'lblCodigoProduto1
        frmExibicao7Ref.lblCodigoProduto1.Caption = dteEtiquetas.rsEtiquetas.Fields("cod_peca").Value
'lblRoute
        frmExibicao7Ref.lblRoute.Caption = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("compl_peca2").Value), " ", dteEtiquetas.rsEtiquetas.Fields("compl_peca2").Value)
'lblComplPeca1
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("COMPL_PECA1").Value) Then
            frmExibicao7Ref.lblComplPeca1.Caption = dteEtiquetas.rsEtiquetas.Fields("COMPL_PECA1").Value
        Else
            frmExibicao7Ref.lblComplPeca1.Caption = ""
        End If
        
'lblLicenseA/lbl_id_etiqueta

        sdataAuxiliar = Mid$(Format(Now(), "MMDDhhmms"), 1, 9)
        frmExibicao7Ref.lbl_id_etiqueta.Caption = "1J" & Replace(Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")), " ", "") & sdataAuxiliar
        
        sTexto = "03"
        frmExibicao7Ref.DataToEncodeText2.Text = "[)>" + Chr(30) + "06" + Chr(29) & _
                                                 "P" & Trim(dteEtiquetas.rsEtiquetas.Fields("cod_cliente").Value) & Chr(29) & _
                                                 "Q" & sQtdeCaixa & Chr(29) & _
                                                 "1J" & Replace(Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")), " ", "") & sdataAuxiliar & Chr(29) & _
                                                 "20L" & frmExibicao7Ref.lblCodMSB.Caption & Chr(29) & _
                                                 "21L" & Replace(frmExibicao7Ref.lblPlant.Caption, " ", "") & Chr(29) & _
                                                 "K" & "" & Chr(29) & _
                                                 "15K" & frmExibicao7Ref.lblComplPeca1.Caption & Chr(29) & _
                                                 "6D" & Format(Now(), "YYYYMMDD") & "011" & Chr(29) & _
                                                 "6D" & "000000" & "036" & Chr(29) & _
                                                 "B" & Trim(frmExibicao7Ref.lblContainerType.Caption) & Chr(29) & _
                                                 "7Q" & Trim(str(Int(VBA.CDbl(dteEtiquetas.rsEtiquetas.Fields("Peso"))))) & "GT" & Chr(30) & _
                                                 "" & Chr(4)
        
        If dteEtiquetas.rsEtiquetas.Fields("Descr_Peca") <> "" Then
            frmExibicao7Ref.lblDescricao.Caption = dteEtiquetas.rsEtiquetas.Fields("Descr_Peca")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Lote") <> "" Then
            frmExibicao7Ref.lblLote.Caption = dteEtiquetas.rsEtiquetas.Fields("Lote")
        End If


'*************************************************************************
Rem  AQUI NOVA IMAGENS A SEREM GERADAS************************************
'*************************************************************************
        
        Dim fso As New Scripting.FileSystemObject
        Dim ts As Scripting.TextStream
        Dim Imagem As String
        Dim bExiste As Boolean
        Dim nCont As Integer
        
        Imagem = sDirImagemEtiq & "\EtiqCodigo.txt"
        If Dir$(Imagem) <> "" Then Kill sDirImagemEtiq & "\EtiqCodigo.txt"
        
        Imagem = sDirImagemEtiq & "\DATA_MATRIX.jpg"
        If Dir$(Imagem) <> "" Then Kill sDirImagemEtiq & "\DATA_MATRIX.jpg"
        
        Imagem = sDirImagemEtiq & "\CODE_128.jpg"
        If Dir$(Imagem) <> "" Then Kill sDirImagemEtiq & "\CODE_128.jpg"
        
        Set ts = fso.OpenTextFile(sDirImagemEtiq & "\EtiqCodigo.txt", ForWriting, True)  'abre um arquivo para escrita , se não existir cria

        sTexto = sTexto & frmExibicao7Ref.DataToEncodeText2.Text
        ts.WriteLine sTexto '"03" & frmExibicao7Ref.DataToEncodeText2.Text

        ts.WriteLine "04" & "1J" & Replace(Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")), " ", "") & sdataAuxiliar
        
        ts.Close
        Set ts = Nothing

        Shell sDirImagemEtiq & "\JCodFactory.exe"

        bExiste = False
        Imagem = sDirImagemEtiq & "\DATA_MATRIX.jpg"
        nCont = 0

        Do While bExiste = False
           If Dir$(Imagem) <> "" Then bExiste = True
           Sleep 500
           nCont = nCont + 1
           If nCont > 4 Then
              MsgBox "Figura do código de Matrix com problemas de geração. Contacte o responsável"
              End
           End If
        Loop
'        Sleep 500
        frmExibicao7Ref.Image1.Picture = LoadPicture(Imagem)
        
        bExiste = False
        Imagem = sDirImagemEtiq & "\CODE_128.jpg"
        nCont = 0
        Do While bExiste = False
           If Dir$(Imagem) <> "" Then bExiste = True
           nCont = nCont + 1
           If nCont > 4000 Then
              MsgBox "Figura do código de barras com problemas de geração. Contacte o responsável"
              End
           End If
        Loop
        frmExibicao7Ref.Image2.Picture = LoadPicture(Imagem)
        

Rem  AQUI NOVA IMAGENS A SEREM GERADAS - TERMINO ************************************
        
'lblCodigoBarras cod musashi sequencial
        frmExibicao7Ref.lblCodigoBarras.Caption = Left(Trim(dteEtiquetas.rsEtiquetas.Fields("id_etiqueta")), 16)
        frmExibicao7Ref.lblCodigoBarrasA.Caption = "*" & frmExibicao7Ref.lblCodigoBarras.Caption & "*"
        frmExibicao7Ref.lblCodigoBarrasB.Caption = "*" & frmExibicao7Ref.lblCodigoBarras.Caption & "*"

' reconfigurando para emissao do codigo sem o formato do code128
        frmExibicao7Ref.lbl_id_etiqueta.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")) & " " & Mid$(Format(Now(), "MMDDhhmms"), 1, 9)

'lblShipmentDate
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("data_etiq")) Then
           sdata_aux = Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 1, 2) & _
                    Pega_Mes(Val(Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 4, 2)))

           sdata_aux_Ano = Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 7, 4)
        Else
           sdata_aux = Format(Now(), "DD") & _
                    Pega_Mes(Val(Format(Now(), "MM")))
           sdata_aux_Ano = Format(Now(), "YYYY")
        End If
        
        frmExibicao7Ref.lblShipmentDate.Caption = sdata_aux
        frmExibicao7Ref.lblShipmentAno.Caption = sdata_aux_Ano
        
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("data_etiq")) Then
            frmExibicao7Ref.lblshipdate.Caption = Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 1, 2) & _
                                              "/" & Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 4, 2) & _
                                              "/" & Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 7, 4)
        Else
            frmExibicao7Ref.lblshipdate.Caption = Format(Now(), "dd/MM/yyyy")
        End If
        
        frmExibicao7Ref.lblExpDate.Caption = "000000000"
        
        If dteEtiquetas.rsEtiquetas.Fields("Peso") <> 0 Then
            frmExibicao7Ref.lblgrossWeight.Caption = Trim(Val(dteEtiquetas.rsEtiquetas.Fields("Peso"))) & " KG"
        End If

        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")) Then
            frmExibicao7Ref.lblCodFunc.Caption = dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")
        Else
            frmExibicao7Ref.lblCodFunc.Caption = ""
        End If
        
    End If
    
    
'**********************************************************************************************************************
'**********************************************************************************************************************
'**********************************************************************************************************************
    'Se tipo M opcao padrão etiqueta grande nova mda 26/04/2019
    If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "M" Then
        frmExibicao4MDA.Show
        frmExibicao4MDA.Left = Me.Width
        If dteEtiquetas.rsEtiquetas.Fields("Cliente") <> "" Then
            frmExibicao4MDA.lblCliente.Caption = dteEtiquetas.rsEtiquetas.Fields("Cliente")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente") <> "" Then
            frmExibicao4MDA.lblCodCliente2.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente")
            frmExibicao4MDA.LBL_CODIGO_NOVO.Caption = "*" & Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente")) & "*"
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Descr_Peca") <> "" Then
            frmExibicao4MDA.lblDescricao.Caption = dteEtiquetas.rsEtiquetas.Fields("Descr_Peca")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Lote") <> "" Then
            frmExibicao4MDA.lblLote2.Caption = dteEtiquetas.rsEtiquetas.Fields("Lote")
            frmExibicao4MDA.LBL_LOTE_NOVO.Caption = "*" & Trim(dteEtiquetas.rsEtiquetas.Fields("Lote")) & "*"
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Peso") <> "" Then
            frmExibicao4MDA.lblPeso2.Caption = dteEtiquetas.rsEtiquetas.Fields("Peso")
        End If
        
        If dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa") <> "" Then
            frmExibicao4MDA.lblQtd2.Caption = dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")
            frmExibicao4MDA.LBL_QTDE_NOVO.Caption = "*" & dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa") & "*"
        End If
        
        If dteEtiquetas.rsEtiquetas.Fields("Cod_Peca") <> "" Then
            frmExibicao4MDA.lblPeca.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Peca")
        End If
        
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")) Then
            frmExibicao4MDA.lblEmbalagem2.Caption = dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")
        End If
     
        frmExibicao4MDA.lbl_data.Caption = Format(dteEtiquetas.rsEtiquetas.Fields("data_etiq"), "dd/mm/yyyy")
        frmExibicao4MDA.LBL_MES.Caption = Pega_Mes(Mid$(Format(dteEtiquetas.rsEtiquetas.Fields("data_etiq"), "dd/mm/yyyy"), 4, 2))
        
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA")) Then
           frmExibicao4MDA.lblCodigoBarras.Caption = dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA")
           frmExibicao4MDA.lblCodigoBarrasA.Caption = "*" & frmExibicao4MDA.lblCodigoBarras.Caption & "*"
           frmExibicao4MDA.lblCodigoBarrasB.Caption = "*" & frmExibicao4MDA.lblCodigoBarras.Caption & "*"
        Else
           frmExibicao4MDA.lblCodigoBarras.Caption = ""
           frmExibicao4MDA.lblCodigoBarrasA.Caption = ""
           frmExibicao4MDA.lblCodigoBarrasB.Caption = ""
        End If
        
        If objApplication.filial <> adMusashiDaAmazonia Then
           If (dteEtiquetas.rsEtiquetas.Fields("indforjimport") = "X") Then
              frmExibicao4MDA.lbl_kms.Visible = True
              frmExibicao4MDA.lbl_tarja_kms.Visible = True
           Else
              frmExibicao4MDA.lbl_kms.Visible = False
              frmExibicao4MDA.lbl_tarja_kms.Visible = False
           End If
        End If
        
        frmExibicao4MDA.lblMsgProduto.Visible = False
         
    End If
    
    '---------------------------------------------------------------------------------------------------
    
'**********************************************************************************************************************
'**********************************************************************************************************************
'**********************************************************************************************************************
'-----------------------------------------------------------------------------------------------------
    'Se tipo S opcao Palete GM
    If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "S" Then
        Set frmExibicao7Ref = frmExibicao7GMNAC
        
        frmExibicao7Ref.Show
        
        frmExibicao7Ref.lblPlant.Caption = "4J D30"

        If dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value = "4G03810A" Then
           frmExibicao7Ref.lblCodMSB.Caption = "TA-F1V05"
           frmExibicao7Ref.lblContainerType.Caption = "CX171203"
        ElseIf dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value = "6G03530" Then
           frmExibicao7Ref.lblCodMSB.Caption = "TA-F4V25"
           frmExibicao7Ref.lblContainerType.Caption = "CX484022"
        ElseIf dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value = "2G01602" Then
           frmExibicao7Ref.lblCodMSB.Caption = "TA-F1V05"
           frmExibicao7Ref.lblContainerType.Caption = "CX171203"
        Else
           frmExibicao7Ref.lblCodMSB.Caption = dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value
           frmExibicao7Ref.lblContainerType.Caption = dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value
        End If
'        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("IND_SUPLEMentar").Value) Then
'            frmExibicao7Ref.lblPlant.Caption = "72479 A215" ' dteEtiquetas.rsEtiquetas.Fields("IND_SUPLEMentar").Value
'        Else
'            frmExibicao7Ref.lblPlant.Caption = ""
'        End If
'quantity
        If (IsNumeric(dteEtiquetas.rsEtiquetas.Fields("QTDE_CAIXA1").Value) And IsNumeric(dteEtiquetas.rsEtiquetas.Fields("PECAS_CAIXA1").Value)) Then
            frmExibicao7Ref.lblQtde1.Caption = "168" ' dteEtiquetas.rsEtiquetas.Fields("QTDE_CAIXA1").Value
        Else
            frmExibicao7Ref.lblQtde1.Caption = "168"
        End If
'lblMaterial
        Rem INCLUIDO MARCOS PEDROSA EM 28-08-2007
'        If dteEtiquetas.rsEtiquetas.Fields("Modelo") <> "" Then
            frmExibicao7Ref.lblMaterial.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("cod_cliente").Value)
'        End If
'lblCodigoProduto1
        frmExibicao7Ref.lblCodigoProduto1.Caption = dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value
'lblComplPeca1
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("COMPL_PECA1").Value) Then
            frmExibicao7Ref.lblComplPeca1.Caption = dteEtiquetas.rsEtiquetas.Fields("COMPL_PECA1").Value
        Else
            frmExibicao7Ref.lblComplPeca1.Caption = ""
        End If
'lblLicenseA/lbl_id_etiqueta
        If IsNull(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")) Then
           dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega") = "4J D30"
        End If
        frmExibicao7Ref.lbl_id_etiqueta.Caption = "1J" & Replace(Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")), " ", "") & Mid$(Format(Now(), "MMDDhhmms"), 1, 9)
        frmExibicao7Ref.DataToEncodeText.Text = "1J" & Replace(Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")), " ", "") & Mid$(Format(Now(), "MMDDhhmms"), 1, 9)

        frmExibicao7Ref.DataToEncodeText2.Text = "[)>" + Chr(30) + "06" + Chr(29) & _
                                                 "P" & Trim(dteEtiquetas.rsEtiquetas.Fields("cod_cliente").Value) & Chr(29) & _
                                                 "Q" & "168" & Chr(29) & _
                                                 "1J" & Replace(Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")), " ", "") & Mid$(Format(Now(), "MMDDhhmms"), 1, 9) & Chr(29) & _
                                                 "20L" & frmExibicao7Ref.lblCodMSB.Caption & Chr(29) & _
                                                 "21L" & Replace(frmExibicao7Ref.lblPlant.Caption, " ", "") & Chr(29) & _
                                                 "K" & "" & Chr(29) & _
                                                 "15K" & frmExibicao7Ref.lblComplPeca1.Caption & Chr(29) & _
                                                 "B" & Trim(frmExibicao7Ref.lblContainerType.Caption) & Chr(29) & _
                                                 "7Q" & Trim(str(Int(VBA.CDbl(dteEtiquetas.rsEtiquetas.Fields("Peso"))))) & "GT" & Chr(29) & _
                                                 "2S" & "" & Chr(30) & _
                                                 "" & Chr(4)

        frmExibicao7Ref.PDF1.DataToEncode = frmExibicao7Ref.DataToEncodeText2.Text

'        Dim fso As New Scripting.FileSystemObject
'        Dim ts As Scripting.TextStream



'        Set ts = fso.OpenTextFile("C:\Arquivos de programas\Etiquetas\EtiqCodigo.txt", ForWriting, True)  'abre um arquivo para escrita , se não existir cria
'        Set ts = fso.OpenTextFile(App.Path & "\EtiqCodigo.txt", ForWriting, True)  'abre um arquivo para escrita , se não existir cria
        
'        ts.Write frmExibicao7Ref.DataToEncodeText2.Text
'        ts.Close
'        Set ts = Nothing

'        Shell App.Path & "\curl\I386\curl.exe http://192.168.33.10:3000/barcodes/index"
'        Shell App.Path & "\curl\I386\curl.exe http://192.168.33.10/barcodes/index"

        frmExibicao7Ref.Text1.Text = frmExibicao7Ref.DataToEncodeText2.Text
        
'lblCodigoBarras cod musashi sequencial
        frmExibicao7Ref.lblCodigoBarras.Caption = Left(Trim(dteEtiquetas.rsEtiquetas.Fields("id_etiqueta")), 16)
        frmExibicao7Ref.lblCodigoBarrasMusashiA.Caption = "*" & frmExibicao7Ref.lblCodigoBarras.Caption & "*"

' reconfigurando para emissao do codigo sem o formato do code128
        frmExibicao7Ref.lbl_id_etiqueta.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")) & " " & Mid$(Format(Now(), "MMDDhhmms"), 1, 9)


'lblShipmentDate
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("data_etiq")) Then
           sdata_aux = Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 1, 2) & _
                    Pega_Mes(Val(Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 4, 2))) & _
                    Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 7, 4)
        Else
           sdata_aux = Format(Now(), "DD") & _
                    Pega_Mes(Val(Format(Now(), "MM"))) & _
                    Format(Now(), "YYYY")
        End If
        
        frmExibicao7Ref.lblShipmentDate.Caption = sdata_aux
        
        If dteEtiquetas.rsEtiquetas.Fields("Peso") <> 0 Then
            frmExibicao7Ref.lblgrossWeight.Caption = dteEtiquetas.rsEtiquetas.Fields("Peso")
        End If

        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")) Then
            frmExibicao7Ref.lblCodFunc.Caption = dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")
        Else
            frmExibicao7Ref.lblCodFunc.Caption = ""
        End If
        
    End If
    
'**********************************************************************************************************************
'**********************************************************************************************************************
'**********************************************************************************************************************
'-----------------------------------------------------------------------------------------------------
    'Se tipo S opcao NOVA ETIQUETA DA HONDA 09/05/2019 COM QRCODE
    If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "H" Then
        Set frmExibicao7Ref = frmExibicao7GMNAC
        
        frmExibicao7Ref.Show
        
        frmExibicao7Ref.lblPlant.Caption = "4J D30"

        If dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value = "4G03810A" Then
           frmExibicao7Ref.lblCodMSB.Caption = "TA-F1V05"
           frmExibicao7Ref.lblContainerType.Caption = "CX171203"
        ElseIf dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value = "6G03530" Then
           frmExibicao7Ref.lblCodMSB.Caption = "TA-F4V25"
           frmExibicao7Ref.lblContainerType.Caption = "CX484022"
        ElseIf dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value = "2G01602" Then
           frmExibicao7Ref.lblCodMSB.Caption = "TA-F1V05"
           frmExibicao7Ref.lblContainerType.Caption = "CX171203"
        Else
           frmExibicao7Ref.lblCodMSB.Caption = dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value
           frmExibicao7Ref.lblContainerType.Caption = dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value
        End If
'        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("IND_SUPLEMentar").Value) Then
'            frmExibicao7Ref.lblPlant.Caption = "72479 A215" ' dteEtiquetas.rsEtiquetas.Fields("IND_SUPLEMentar").Value
'        Else
'            frmExibicao7Ref.lblPlant.Caption = ""
'        End If
'quantity
        If (IsNumeric(dteEtiquetas.rsEtiquetas.Fields("QTDE_CAIXA1").Value) And IsNumeric(dteEtiquetas.rsEtiquetas.Fields("PECAS_CAIXA1").Value)) Then
            frmExibicao7Ref.lblQtde1.Caption = "168" ' dteEtiquetas.rsEtiquetas.Fields("QTDE_CAIXA1").Value
        Else
            frmExibicao7Ref.lblQtde1.Caption = "168"
        End If
'lblMaterial
        Rem INCLUIDO MARCOS PEDROSA EM 28-08-2007
'        If dteEtiquetas.rsEtiquetas.Fields("Modelo") <> "" Then
            frmExibicao7Ref.lblMaterial.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("cod_cliente").Value)
'        End If
'lblCodigoProduto1
        frmExibicao7Ref.lblCodigoProduto1.Caption = dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value
'lblComplPeca1
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("COMPL_PECA1").Value) Then
            frmExibicao7Ref.lblComplPeca1.Caption = dteEtiquetas.rsEtiquetas.Fields("COMPL_PECA1").Value
        Else
            frmExibicao7Ref.lblComplPeca1.Caption = ""
        End If
'lblLicenseA/lbl_id_etiqueta
        If IsNull(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")) Then
           dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega") = "4J D30"
        End If
        frmExibicao7Ref.lbl_id_etiqueta.Caption = "1J" & Replace(Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")), " ", "") & Mid$(Format(Now(), "MMDDhhmms"), 1, 9)
        frmExibicao7Ref.DataToEncodeText.Text = "1J" & Replace(Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")), " ", "") & Mid$(Format(Now(), "MMDDhhmms"), 1, 9)

        frmExibicao7Ref.DataToEncodeText2.Text = "[)>" + Chr(30) + "06" + Chr(29) & _
                                                 "P" & Trim(dteEtiquetas.rsEtiquetas.Fields("cod_cliente").Value) & Chr(29) & _
                                                 "Q" & "168" & Chr(29) & _
                                                 "1J" & Replace(Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")), " ", "") & Mid$(Format(Now(), "MMDDhhmms"), 1, 9) & Chr(29) & _
                                                 "20L" & frmExibicao7Ref.lblCodMSB.Caption & Chr(29) & _
                                                 "21L" & Replace(frmExibicao7Ref.lblPlant.Caption, " ", "") & Chr(29) & _
                                                 "K" & "" & Chr(29) & _
                                                 "15K" & frmExibicao7Ref.lblComplPeca1.Caption & Chr(29) & _
                                                 "B" & Trim(frmExibicao7Ref.lblContainerType.Caption) & Chr(29) & _
                                                 "7Q" & Trim(str(Int(VBA.CDbl(dteEtiquetas.rsEtiquetas.Fields("Peso"))))) & "GT" & Chr(29) & _
                                                 "2S" & "" & Chr(30) & _
                                                 "" & Chr(4)

        frmExibicao7Ref.PDF1.DataToEncode = frmExibicao7Ref.DataToEncodeText2.Text

'        Dim fso As New Scripting.FileSystemObject
'        Dim ts As Scripting.TextStream



'        Set ts = fso.OpenTextFile("C:\Arquivos de programas\Etiquetas\EtiqCodigo.txt", ForWriting, True)  'abre um arquivo para escrita , se não existir cria
'        Set ts = fso.OpenTextFile(App.Path & "\EtiqCodigo.txt", ForWriting, True)  'abre um arquivo para escrita , se não existir cria
        
'        ts.Write frmExibicao7Ref.DataToEncodeText2.Text
'        ts.Close
'        Set ts = Nothing

'        Shell App.Path & "\curl\I386\curl.exe http://192.168.33.10:3000/barcodes/index"
'        Shell App.Path & "\curl\I386\curl.exe http://192.168.33.10/barcodes/index"

        frmExibicao7Ref.Text1.Text = frmExibicao7Ref.DataToEncodeText2.Text
        
'lblCodigoBarras cod musashi sequencial
        frmExibicao7Ref.lblCodigoBarras.Caption = Left(Trim(dteEtiquetas.rsEtiquetas.Fields("id_etiqueta")), 16)
        frmExibicao7Ref.lblCodigoBarrasMusashiA.Caption = "*" & frmExibicao7Ref.lblCodigoBarras.Caption & "*"

' reconfigurando para emissao do codigo sem o formato do code128
        frmExibicao7Ref.lbl_id_etiqueta.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")) & " " & Mid$(Format(Now(), "MMDDhhmms"), 1, 9)


'lblShipmentDate
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("data_etiq")) Then
           sdata_aux = Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 1, 2) & _
                    Pega_Mes(Val(Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 4, 2))) & _
                    Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 7, 4)
        Else
           sdata_aux = Format(Now(), "DD") & _
                    Pega_Mes(Val(Format(Now(), "MM"))) & _
                    Format(Now(), "YYYY")
        End If
        
        frmExibicao7Ref.lblShipmentDate.Caption = sdata_aux
        
        If dteEtiquetas.rsEtiquetas.Fields("Peso") <> 0 Then
            frmExibicao7Ref.lblgrossWeight.Caption = dteEtiquetas.rsEtiquetas.Fields("Peso")
        End If

        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")) Then
            frmExibicao7Ref.lblCodFunc.Caption = dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")
        Else
            frmExibicao7Ref.lblCodFunc.Caption = ""
        End If
        
    End If

'**********************************************************************************************************************
'**********************************************************************************************************************
'**********************************************************************************************************************
'-----------------------------------------------------------------------------------------------------
    '-------------------------------------------------------------------
    'Se tipo A - Identificação de alterações do produto e processo
    If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "A" Then
        
        frmExibicao6.Show
        
        With frmExibicao6
            .lblDesenho.Caption = Trim(dteEtiquetas.rsEtiquetas("COD_CLIENTE").Value & "")
            .lblDesvio.Caption = Trim(dteEtiquetas.rsEtiquetas("DESVIO_AVISO_MOD").Value & "")
            .lblData1.Caption = Format(Date, "dd/mm/yyyy")
            .lblData2.Caption = .lblData1.Caption
            .lblNotaFiscal1.Caption = Trim(dteEtiquetas.rsEtiquetas("NUM_DOC_FISCAL").Value & "")
            .lblNotaFiscal2.Caption = .lblNotaFiscal1.Caption
            
            .lblOptDefinitiva.Visible = False
            .lblOptProvisoria.Visible = False
            .lblOptLoteUnico.Visible = False
            
            Select Case Trim(dteEtiquetas.rsEtiquetas("TIPO_ALTERACAO").Value)
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
            Select Case Trim(dteEtiquetas.rsEtiquetas("MOTIVO_ALTERACAO").Value)
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
            
            .lblMotivoAlteracaoOutros.Caption = Trim(dteEtiquetas.rsEtiquetas("MOTIVO_ALTERACAO_OUTROS").Value & "")
            
            .lblOptPrimeiroEnvio.Visible = False
            .lblOptLoteIntermediario.Visible = False
            .lblOptUltimoLote.Visible = False
            Select Case Trim(dteEtiquetas.rsEtiquetas("ENVIO_LOTE").Value)
            Case "1"
                .lblOptPrimeiroEnvio.Visible = True
            Case "2"
                .lblOptLoteIntermediario.Visible = True
            Case "3"
                .lblOptUltimoLote.Visible = True
            End Select
            
            .lblNumAm.Caption = Trim(dteEtiquetas.rsEtiquetas("NUM_AM").Value & "")
            
        End With
        
        frmOpcoes.txtQuantidade.Text = dteEtiquetas.rsEtiquetas("QTD_ETIQ").Value
        
    End If
    
    'Se tipo X opcao padrão etiqueta cliente SHOWA
    If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "x" Then
        'frmExibicao44.Show
        frmAvulsoPadraoPonteiro.Show
        frmAvulsoPadraoPonteiro.Left = Me.Width
        If dteEtiquetas.rsEtiquetas.Fields("Cliente") <> "" Then
            'frmExibicao4.lblCliente.Caption = dteEtiquetas.rsEtiquetas.Fields("Cliente")
            frmAvulsoPadraoPonteiro.lblCliente.Caption = dteEtiquetas.rsEtiquetas.Fields("Cliente")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente") <> "" Then
            'frmExibicao4.lblCodCliente2.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente")
            frmAvulsoPadraoPonteiro.lblCodCliente2.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Descr_Peca") <> "" Then
            'frmExibicao4.lblDescricao.Caption = dteEtiquetas.rsEtiquetas.Fields("Descr_Peca")
            frmAvulsoPadraoPonteiro.lblDescricao.Caption = dteEtiquetas.rsEtiquetas.Fields("Descr_Peca")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Lote") <> "" Then
            'frmExibicao4.lblLote2.Caption = dteEtiquetas.rsEtiquetas.Fields("Lote")
            frmAvulsoPadraoPonteiro.lblLote2.Caption = dteEtiquetas.rsEtiquetas.Fields("Lote")
        End If
        If dteEtiquetas.rsEtiquetas.Fields("Peso") <> "" Then
            'frmExibicao4.lblPeso2.Caption = dteEtiquetas.rsEtiquetas.Fields("Peso")
            frmAvulsoPadraoPonteiro.lblPeso2.Caption = dteEtiquetas.rsEtiquetas.Fields("Peso")
        End If
        
        If dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa") <> "" Then
            'frmExibicao4.lblQtd2.Caption = dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")
            frmAvulsoPadraoPonteiro.lblQtd2.Caption = dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")
        End If
        
        If dteEtiquetas.rsEtiquetas.Fields("Cod_Peca") <> "" Then
            'frmExibicao4.lblPeca.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Peca")
            frmAvulsoPadraoPonteiro.lblPeca.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Peca")
        End If
        
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")) Then
            frmAvulsoPadraoPonteiro.lblEmbalagem2.Caption = dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")
        End If
     
        frmAvulsoPadraoPonteiro.lbl_data.Caption = Format(dteEtiquetas.rsEtiquetas.Fields("data_etiq"), "dd/mm/yyyy")
        frmAvulsoPadraoPonteiro.LBL_MES.Caption = Pega_Mes(Mid$(Format(dteEtiquetas.rsEtiquetas.Fields("data_etiq"), "dd/mm/yyyy"), 4, 2))
        
        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA")) Then
           frmAvulsoPadraoPonteiro.lbl_Seq_Milhar.Caption = Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA")), _
                                                            Len(Trim(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA"))) - 3, _
                                                            Len(Trim(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA"))))
           frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption = dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA")
           frmAvulsoPadraoPonteiro.lblCodigoBarrasA.Caption = "*" & frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption & "*"
           frmAvulsoPadraoPonteiro.lblCodigoBarrasB.Caption = "*" & frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption & "*"
        Else
           frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption = ""
           frmAvulsoPadraoPonteiro.lblCodigoBarrasA.Caption = ""
           frmAvulsoPadraoPonteiro.lblCodigoBarrasB.Caption = ""
        End If
        
        If (dteEtiquetas.rsEtiquetas.Fields("Tipo") = "8") Then
           frmAvulsoPadraoPonteiro.lblCodCliente1.Caption = "CODE:"
           frmAvulsoPadraoPonteiro.lblQtd1.Caption = "QTY.:"
           frmAvulsoPadraoPonteiro.lblPeso1.Caption = "WEIGHT:"
           frmAvulsoPadraoPonteiro.lblLote1.Caption = "LOT.:"
        Else
           If objApplication.filial <> adMusashiDaAmazonia Then
              If (dteEtiquetas.rsEtiquetas.Fields("indforjimport") = "X") Then
                 frmAvulsoPadraoPonteiro.lbl_kms.Visible = True
                 frmAvulsoPadraoPonteiro.lbl_tarja_kms.Visible = True
              Else
                 frmAvulsoPadraoPonteiro.lbl_kms.Visible = False
                 frmAvulsoPadraoPonteiro.lbl_tarja_kms.Visible = False
              End If
           End If
        End If
         
         
    End If
    
    '-------------------------------------------------------------------
    'Se tipo A - Identificação de alterações do produto e processo
    If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "Y" Then
        If objApplication.filial = adMusashiDaAmazonia Then
           frmExibicao11.Show
           With frmExibicao11
               .lbl_LPN_COD.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor") & Format(Now(), "DDMMYY") & Format(dteEtiquetas.rsEtiquetas.Fields("Sequencia_Dia"), "000000")
                Call CodeRefresh(.lbl_LPN_COD.Caption)
               .lbl_LPN_COD_BARRA.Caption = sCodigo128
               .lbl_LPN_COD_B.Caption = .lbl_LPN_COD.Caption
               .lbl_CODIGO_NUM.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente"))
               .lbl_SUPPLIER_COD.Caption = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor")), " ", dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor"))
               .lbl_USER_COD.Caption = "9219"
               .lbl_YAMAHA_COD_BARRAS.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente") & "-" & _
                                                .lbl_SUPPLIER_COD.Caption & "-" & _
                                                "9219"
                Call CodeRefresh(.lbl_YAMAHA_COD_BARRAS.Caption)
               .lbl_YAMAHA_BARRAS.Caption = sCodigo128
               
               .lbl_NOME_DESCRICAO.Caption = dteEtiquetas.rsEtiquetas.Fields("descr_peca")
   '            If objApplication.filial = adMusashiDaAmazonia Then
                  .lbl_FORNECEDOR_NOME.Caption = "MUSASHI DO BRASIL LTDA"
   '            Else
   '               .lbl_FORNECEDOR_NOME.Caption = ""
   '            End If
               .lbl_QTDE_NUM.Caption = dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")
               .lbl_NF_NUM.Caption = ""
               .LBL_MES.Caption = Pega_Mes(Mid$(Format(dteEtiquetas.rsEtiquetas.Fields("data_etiq"), "dd/mm/yyyy"), 4, 2))
               .lbl_ANO.Caption = Format(dteEtiquetas.rsEtiquetas.Fields("data_etiq"), "YY")
               .lbl_QTDE_BARRAS.Caption = dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")
               .lbl_QTDE_NUM1.Caption = dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")
               .lbl_COD_MUSASHI_BARRAS.Caption = "*" & Trim(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA")) & "*"
               .lbl_COD_MUSASHI_NUM.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA"))
               
           End With
        Else
           frmExibicao10.Show
           With frmExibicao10
               .lbl_LPN_COD.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor") & Format(Now(), "DDMMYY") & Format(dteEtiquetas.rsEtiquetas.Fields("Sequencia_Dia"), "000000")
                Call CodeRefresh(.lbl_LPN_COD.Caption)
               .lbl_LPN_COD_BARRA.Caption = sCodigo128
               .lbl_LPN_COD_B.Caption = .lbl_LPN_COD.Caption
               .lbl_CODIGO_NUM.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente"))
               .lbl_SUPPLIER_COD.Caption = "5859" 'IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor")), " ", dteEtiquetas.rsEtiquetas.Fields("Cod_Fornecedor"))
               .lbl_QA.Caption = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Cod_util")), " ", dteEtiquetas.rsEtiquetas.Fields("cod_util"))
               .lbl_USER_COD.Caption = "9219"
               .lbl_YAMAHA_COD_BARRAS.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Cliente")) & "-" & _
                                                .lbl_SUPPLIER_COD.Caption & "-" & _
                                                "9219"
                Call CodeRefresh(.lbl_YAMAHA_COD_BARRAS.Caption)
               .lbl_YAMAHA_BARRAS.Caption = sCodigo128
               
               .lbl_NOME_DESCRICAO.Caption = dteEtiquetas.rsEtiquetas.Fields("descr_peca")
   '            If objApplication.filial = adMusashiDaAmazonia Then
                  .lbl_FORNECEDOR_NOME.Caption = "MUSASHI DO BRASIL LTDA"
   '            Else
   '               .lbl_FORNECEDOR_NOME.Caption = ""
   '            End If
               .lbl_QTDE_NUM.Caption = dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")
               .lbl_NF_NUM.Caption = ""
               .LBL_MES.Caption = Pega_Mes(Mid$(Format(dteEtiquetas.rsEtiquetas.Fields("data_etiq"), "dd/mm/yyyy"), 4, 2))
               .lbl_ANO.Caption = Format(dteEtiquetas.rsEtiquetas.Fields("data_etiq"), "YY")
               .lbl_QTDE_BARRAS.Caption = dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")
               .lbl_QTDE_NUM1.Caption = dteEtiquetas.rsEtiquetas.Fields("Qtd_Caixa")
               .lbl_COD_MUSASHI_BARRAS.Caption = "*" & Trim(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA")) & "*"
               .lbl_COD_MUSASHI_NUM.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("ID_ETIQUETA"))
               If dteEtiquetas.rsEtiquetas.Fields("Lote") <> "" Then
                  .lblLot.Caption = dteEtiquetas.rsEtiquetas.Fields("Lote")
               Else
                  .lblLot.Caption = ""
               End If
               Rem aqui marcos 31/01/2024, e onde tiver mais "9219" form10
               If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("Cod_Peca")) Then
                   .lblCodMSB.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Peca")
                   .lblCodMSB_Letra = Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Peca")), Len(Trim(dteEtiquetas.rsEtiquetas.Fields("Cod_Peca"))), 1)
               Else
                   .lblCodMSB.Caption = ""
               End If
           End With
        End If
        
        frmOpcoes.txtQuantidade.Text = dteEtiquetas.rsEtiquetas("QTD_ETIQ").Value
        
    End If
    
    
    '-------------------------------------------------------------------
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Cancel = 0 Then
   Call Fechar_Form_Etiqueta
End If
End Sub

Private Sub optGeral_Click()
    
    dteEtiquetas.rsEtiquetas.Close
    dteEtiquetas.rsEtiquetas.Source = "Select * from Etiquetas"
    dteEtiquetas.rsEtiquetas.Open
    
    If dteEtiquetas.rsEtiquetas.RecordCount = 0 Then
        MsgBox "Impressão concluída com sucesso! O aplicativo será encerrado!", vbInformation + vbOKOnly, "Tarefa Concluída"
        dteEtiquetas.rsEtiquetas.Close
        If Dir("etiq.txt") = "etiq.txt" Then
            Close gNumeroArquivo
            Kill "etiq.txt"
        End If
        'Set objApplication = Nothing
        'End
        MDIEtiquetas.forcaSaida = True
        Unload MDIEtiquetas
    End If
    
    dteEtiquetas.rsEtiquetas.MoveFirst
    MostraRegistroAtual
    
    If dteEtiquetas.rsEtiquetas.RecordCount = 1 Then
        cmdProximo.Enabled = False
        cmdUltimo.Enabled = False
        cmdAnterior.Enabled = False
        cmdProximo.Enabled = False
    Else
        cmdProximo.Enabled = True
        cmdUltimo.Enabled = True
        cmdAnterior.Enabled = False
        cmdPrimeiro.Enabled = False
    End If
    
    'Habilita os flags
    FlagGeral = True
    FlagPecas = False
    FlagLote = False
    
End Sub

Private Sub optLote_Click()
    Lote = Empty
    Lote = InputBox("Lote", "Informe o lote da Peça")
    If Trim(Lote) = "" Then
        optGeral.SetFocus
        Exit Sub
    Else
        Lote = Trim(UCase(Lote))
    End If
    
    dteEtiquetas.rsEtiquetas.Close
    dteEtiquetas.rsEtiquetas.Source = "Select * from Etiquetas where Lote = '" & Lote & "'"
    dteEtiquetas.rsEtiquetas.Open
    
    If dteEtiquetas.rsEtiquetas.RecordCount = 0 Then
        MsgBox "O lote " & "'" & Lote & "'" & " não foi encontrado! Tente novamente", vbInformation + vbOKOnly, "ATENÇÃO !!!"
        optGeral.SetFocus
        Exit Sub
    Else
        dteEtiquetas.rsEtiquetas.MoveFirst
        MostraRegistroAtual
    End If
    
    'Habilita os flags
    FlagGeral = False
    FlagPecas = False
    FlagLote = True

End Sub

Private Sub optPecas_Click()
    
    Peca = Empty
    Peca = InputBox("Código", "Informe o código da Peça")
    If Trim(Peca) = "" Then
        optGeral.SetFocus
        Exit Sub
    Else
        Peca = Trim(UCase(Peca))
    End If
    
    dteEtiquetas.rsEtiquetas.Close
    dteEtiquetas.rsEtiquetas.Source = "Select * from Etiquetas where Cod_Peca = '" & Peca & "'"
    dteEtiquetas.rsEtiquetas.Open
    
    If dteEtiquetas.rsEtiquetas.RecordCount = 0 Then
        MsgBox "A peça " & "'" & Peca & "'" & " não foi encontrada! Tente novamente", vbInformation + vbOKOnly, "ATENÇÃO !!!"
        optGeral.SetFocus
        Exit Sub
    Else
        dteEtiquetas.rsEtiquetas.MoveFirst
        MostraRegistroAtual
    End If
        
    'Habilita os flags
    FlagGeral = False
    FlagPecas = True
    FlagLote = False
    
    
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
        
Unload frmOpcoes

End Function


Private Sub Imprime_Etiqueta_GM_P()

Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim nx As Double
Dim x As Printer
Dim rs As ADODB.Recordset
Dim cRec As ADODB.Recordset
Dim sData As String
Dim oTela As frmEscRelCristalReport

On Error GoTo Erro

Me.MousePointer = vbHourglass

'Set cRec = New ADODB.Recordset
'
'Set cRec = CCTempneMov_Etiq.Mov_Etiq_Consultar_GM_P(sBancoMusashi, _
'                                                    Me.txtsequencial.Text)
'
'If cRec.RecordCount = 0 Then
'   MsgBox "Etiqueta não encontrada, anote a etiqueta e procure o responsável, etiq: " & dteEtiquetas.rsEtiquetas.Fields("id_etiqueta")
'   Me.MousePointer = vbDefault
'   Exit Sub
'End If

nx = 0

On Error GoTo Erro
Set rs = New ADODB.Recordset


rs.Fields.Append "1_DESTINO", ADODB.DataTypeEnum.adChar, 40
rs.Fields.Append "2_FORNECEDOR", ADODB.DataTypeEnum.adChar, 40
rs.Fields.Append "3_PARTNAME", ADODB.DataTypeEnum.adChar, 60
rs.Fields.Append "4_PARTNUMBER", ADODB.DataTypeEnum.adChar, 10
rs.Fields.Append "5_QUANTITY", ADODB.DataTypeEnum.adChar, 7
rs.Fields.Append "6_REFERENCE", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "7_CONTAINER", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "8_GROSSHEIGHTP", ADODB.DataTypeEnum.adChar, 6
rs.Fields.Append "9_GROSSHEIGHTU", ADODB.DataTypeEnum.adChar, 3
rs.Fields.Append "10_MATERIALCODE", ADODB.DataTypeEnum.adChar, 15
rs.Fields.Append "11_PLDOCSTR", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "12_EXPDATE", ADODB.DataTypeEnum.adChar, 10
rs.Fields.Append "13_SHIPMENTDATE", ADODB.DataTypeEnum.adChar, 10
rs.Fields.Append "14_MUSASHI", ADODB.DataTypeEnum.adChar, 20

rs.Open
nx = 0

rs.AddNew
rs.Fields("1_DESTINO").Value = "GENERAL MOTORS DO BRASIL"
rs.Fields("2_FORNECEDOR").Value = "MUSASHI DO BRASIL S/A"
rs.Fields("3_PARTNAME").Value = frmExibicao7GM2020.lblDescricao.Caption
rs.Fields("4_PARTNUMBER").Value = Trim(frmExibicao7GM2020.lblMaterial.Caption)
rs.Fields("5_QUANTITY").Value = frmExibicao7GM2020.lblQtde1.Caption
rs.Fields("6_REFERENCE").Value = " "
rs.Fields("7_CONTAINER").Value = frmExibicao7GM2020.lblContainerType.Caption
rs.Fields("8_GROSSHEIGHTP").Value = frmExibicao7GM2020.lblgrossWeight
rs.Fields("9_GROSSHEIGHTU").Value = "KG"
rs.Fields("10_MATERIALCODE").Value = frmExibicao7GM2020.lblCodMSB.Caption
rs.Fields("11_PLDOCSTR").Value = frmExibicao7GM2020.lblPlant.Caption & Trim(frmExibicao7GM2020.lblRoute.Caption)
rs.Fields("12_EXPDATE").Value = frmExibicao7GM2020.lblExpDate.Caption
rs.Fields("13_SHIPMENTDATE").Value = frmExibicao7GM2020.lblshipdate.Caption
rs.Fields("14_MUSASHI").Value = " "
rs.Update

Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_GM_P.rpt")
CrystalReport1.Database.SetDataSource rs

rs.Clone

Rem mostrar a tela
'Set oTela = New frmEscRelCristalReport
'oTela.CRViewer1.ReportSource = CrystalReport1
'oTela.CRViewer1.ViewReport
'oTela.Show 0
Rem *************************

Rem nao mostrar a tela
CrystalReport1.PrintOutEx False
Rem *************************

Me.MousePointer = vbDefault

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub





'**********************************************************************************************************************
'**********************************************************************************************************************
'**********************************************************************************************************************
'-----------------------------------------------------------------------------------------------------
''''''
''''''    'Se tipo 7 opcao Palete GM
''''''    If dteEtiquetas.rsEtiquetas.Fields("Tipo") = "7" Then
''''''        Set frmExibicao7Ref = frmExibicao7GM2020
'''''''        Set frmExibicao7Ref = frmExibicao7GM
''''''        frmExibicao7Ref.Show
''''''
''''''        frmExibicao7Ref.lblPlant.Caption = "4J D30"
''''''
''''''        ' Ajustado para emitir mais de um palet , liberando o campo quantidade
''''''        Me.txtQuantidade.Locked = False
''''''        Me.txtQuantidade.Enabled = True
''''''
''''''        If dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value = "2G01600" Then
''''''           frmExibicao7Ref.lblCodMSB.Caption = "TA-F1V05"
''''''           frmExibicao7Ref.lblContainerType.Caption = "CX171203"
''''''        ElseIf dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value = "6G03530" Or _
''''''               dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value = "6G04532" Then
''''''           frmExibicao7Ref.lblCodMSB.Caption = "TA-F4V25"
''''''           frmExibicao7Ref.lblContainerType.Caption = "CX484022"
''''''        ElseIf dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value = "2G01602" Then
''''''           frmExibicao7Ref.lblCodMSB.Caption = "TA-F1V05"
''''''           frmExibicao7Ref.lblContainerType.Caption = "CX171203"
''''''        End If
'''''''        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("IND_SUPLEMentar").Value) Then
'''''''            frmExibicao7Ref.lblPlant.Caption = "72479 A215" ' dteEtiquetas.rsEtiquetas.Fields("IND_SUPLEMentar").Value
'''''''        Else
'''''''            frmExibicao7Ref.lblPlant.Caption = ""
'''''''        End If
'''''''quantity
'''''''        If (IsNumeric(dteEtiquetas.rsEtiquetas.Fields("QTDE_CAIXA1").Value) And IsNumeric(dteEtiquetas.rsEtiquetas.Fields("PECAS_CAIXA1").Value)) Then
'''''''            frmExibicao7Ref.lblQtde1.Caption = 168 ' dteEtiquetas.rsEtiquetas.Fields("QTDE_CAIXA1").Value
'''''''            dteEtiquetas.rsEtiquetas.Fields("QTDE_CAIXA1").Value = 168
'''''''        Else
'''''''            frmExibicao7Ref.lblQtde1.Caption = ""
'''''''        End If
''''''        'Aqui terá algumas mudanças referentes ao codigo do produto e com quantidades fixas
''''''        sQtdeCaixa = ""
''''''        If dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value = "XXXXXX" Then ' "2G01600" Then
''''''            frmExibicao7Ref.lblQtde1.Caption = "1440"
''''''            sQtdeCaixa = "1440" 'dteEtiquetas.rsEtiquetas.Fields("QTDE_CAIXA1").Value = 1440, campo não suporta o tamanho só ate 3 caractere
''''''        ElseIf dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value = "6G03531" Or _
''''''               dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value = "6G04532" Then
''''''            frmExibicao7Ref.lblQtde1.Caption = "168"
''''''            sQtdeCaixa = "168" 'dteEtiquetas.rsEtiquetas.Fields("QTDE_CAIXA1").Value = 168
''''''        ElseIf dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value = "2G01602" Or _
''''''               dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value = "2G01600" Or _
''''''               dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value = "2G05601" Then
''''''            Dim oTela1 As frmEtiquetaDigitaQtde
''''''
''''''            Set oTela1 = New frmEtiquetaDigitaQtde
''''''            oTela1.Show 1
''''''            frmExibicao7Ref.lblQtde1.Caption = Trim(oTela1.txt_qtde.Text)
''''''            sQtdeCaixa = Trim(oTela1.txt_qtde.Text) ' dteEtiquetas.rsEtiquetas.Fields("QTDE_CAIXA1").Value = Trim(oTela1.txt_qtde.Text)
''''''            Unload oTela1: Set oTela1 = Nothing
''''''        Else
''''''           If (IsNumeric(dteEtiquetas.rsEtiquetas.Fields("QTDE_CAIXA1").Value) And IsNumeric(dteEtiquetas.rsEtiquetas.Fields("PECAS_CAIXA1").Value)) Then
''''''               frmExibicao7Ref.lblQtde1.Caption = dteEtiquetas.rsEtiquetas.Fields("QTDE_CAIXA1").Value
''''''           Else
''''''               frmExibicao7Ref.lblQtde1.Caption = ""
''''''           End If
''''''        End If
''''''
'''''''lblMaterial
''''''        Rem INCLUIDO MARCOS PEDROSA EM 28-08-2007
''''''        frmExibicao7Ref.lblMaterial.Caption = dteEtiquetas.rsEtiquetas.Fields("cod_cliente").Value
'''''''lblCodigoProduto1
''''''        frmExibicao7Ref.lblCodigoProduto1.Caption = dteEtiquetas.rsEtiquetas.Fields("CODIGO_PRODUTO1").Value
'''''''lblRote
''''''        frmExibicao7Ref.lblRote.Caption = IIf(IsNull(dteEtiquetas.rsEtiquetas.Fields("Embarque_Controlado").Value), "AAAAA", dteEtiquetas.rsEtiquetas.Fields("Embarque_Controlado").Value)
'''''''lblComplPeca1
''''''        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("COMPL_PECA1").Value) Then
''''''            frmExibicao7Ref.lblComplPeca1.Caption = dteEtiquetas.rsEtiquetas.Fields("COMPL_PECA1").Value
''''''        Else
''''''            frmExibicao7Ref.lblComplPeca1.Caption = ""
''''''        End If
''''''
'''''''lblLicenseA/lbl_id_etiqueta
''''''        frmExibicao7Ref.lbl_id_etiqueta.Caption = "1J" & Replace(Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")), " ", "") & Mid$(Format(Now(), "MMDDhhmms"), 1, 9)
''''''        frmExibicao7Ref.DataToEncodeText.Text = frmExibicao7Ref.lbl_id_etiqueta.Caption  '"1J" & Replace(Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")), " ", "") & Mid$(Format(Now(), "MMDDhhmms"), 1, 9)
''''''
''''''        frmExibicao7Ref.DataToEncodeText2.Text = "[)>" + Chr(30) + "06" + Chr(29) & _
''''''                                                 "P" & Trim(dteEtiquetas.rsEtiquetas.Fields("cod_cliente").Value) & Chr(29) & _
''''''                                                 "Q" & sQtdeCaixa & Chr(29) & _
''''''                                                 "1J" & Replace(Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")), " ", "") & Mid$(Format(Now(), "MMDDhhmms"), 1, 9) & Chr(29) & _
''''''                                                 "20L" & frmExibicao7Ref.lblCodMSB.Caption & Chr(29) & _
''''''                                                 "21L" & Replace(frmExibicao7Ref.lblPlant.Caption, " ", "") & Chr(29) & _
''''''                                                 "K" & "" & Chr(29) & _
''''''                                                 "15K" & frmExibicao7Ref.lblComplPeca1.Caption & Chr(29) & _
''''''                                                 "6D" & Format(Now(), "YYYYMMDD") & "011" & Chr(29) & _
''''''                                                 "6D" & "000000" & "000" & Chr(29) & _
''''''                                                 "6D" & "094" & Chr(29) & _
''''''                                                 "B" & Trim(frmExibicao7Ref.lblContainerType.Caption) & Chr(29) & _
''''''                                                 "7Q" & Trim(str(Int(VBA.CDbl(dteEtiquetas.rsEtiquetas.Fields("Peso"))))) & "GT" & Chr(29) & _
''''''                                                 "2S" & "" & Chr(30) & _
''''''                                                 "" & Chr(4)
''''''Rem marcos, por 0000000 em
''''''Rem           6D" & Format(VBA.DateAdd("m", 36, Now()), "YYYYMMDD") & "036" & Chr(29) & _
''''''
''''''        frmExibicao7Ref.PDF1.DataToEncode = frmExibicao7Ref.DataToEncodeText2.Text
''''''
''''''        Dim fso As New Scripting.FileSystemObject
''''''        Dim ts As Scripting.TextStream
''''''        Dim Imagem As String
''''''        'Pega o caminho para a imagem com o Nome do TextBox'
''''''        Imagem = App.Path & "\barcode\output.jpg"
''''''        If Dir$(Imagem) <> "" Then Kill App.Path & "\barcode\output.jpg"
''''''
''''''        Imagem = App.Path & "\barcode\EtiqCodigo.txt"
''''''        If Dir$(Imagem) <> "" Then Kill App.Path & "\EtiqCodigo.txt"
''''''
''''''        Set ts = fso.OpenTextFile(App.Path & "\EtiqCodigo.txt", ForWriting, True)  'abre um arquivo para escrita , se não existir cria
''''''
''''''        ts.Write frmExibicao7Ref.DataToEncodeText2.Text
''''''        ts.Close
''''''        Set ts = Nothing
''''''
''''''        Shell App.Path & "\curl\I386\curl.exe http://localhost:4000/barcodes/index"
'''''''        Shell App.Path & "\DataMatrix.exe"
''''''
''''''        frmExibicao7Ref.Text1.Text = frmExibicao7Ref.DataToEncodeText2.Text
''''''
'''''''lblCodigoBarras cod musashi sequencial
''''''        frmExibicao7Ref.lblCodigoBarras.Caption = Left(Trim(dteEtiquetas.rsEtiquetas.Fields("id_etiqueta")), 16)
''''''        frmExibicao7Ref.lblCodigoBarrasMusashiA.Caption = "*" & frmExibicao7Ref.lblCodigoBarras.Caption & "*"
''''''
''''''' reconfigurando para emissao do codigo sem o formato do code128
''''''        frmExibicao7Ref.lbl_id_etiqueta.Caption = Trim(dteEtiquetas.rsEtiquetas.Fields("Pto_Entrega")) & " " & Mid$(Format(Now(), "MMDDhhmms"), 1, 9)
''''''
''''''
'''''''lblShipmentDate
''''''        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("data_etiq")) Then
''''''           sdata_aux = Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 1, 2) & _
''''''                    Pega_Mes(Val(Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 4, 2)))
''''''
''''''           sdata_aux_Ano = Mid$(Trim(dteEtiquetas.rsEtiquetas.Fields("data_etiq")), 7, 4)
''''''        Else
''''''           sdata_aux = Format(Now(), "DD") & _
''''''                    Pega_Mes(Val(Format(Now(), "MM")))
''''''           sdata_aux_Ano = Format(Now(), "YYYY")
''''''        End If
''''''
''''''        frmExibicao7Ref.lblShipmentDate.Caption = sdata_aux
''''''        frmExibicao7Ref.lblShipmentAno.Caption = sdata_aux_Ano
''''''        frmExibicao7Ref.lblExpDate.Caption = "000000000" 'sdata_aux & "/" & sdata_aux_Ano
''''''
'''''''lblContainerType
'''''''        If dteEtiquetas.rsEtiquetas.Fields("Cod_Embalagem") <> "" Then
'''''''            frmExibicao7Ref.lblContainerType.Caption = dteEtiquetas.rsEtiquetas.Fields("Cod_Embalagem")
'''''''        Else
'''''''            frmExibicao7Ref.lblContainerType.Caption = ""
'''''''        End If
'''''''lblgrossWeight
''''''        If dteEtiquetas.rsEtiquetas.Fields("Peso") <> 0 Then
''''''            frmExibicao7Ref.lblgrossWeight.Caption = dteEtiquetas.rsEtiquetas.Fields("Peso")
''''''        End If
''''''
''''''        If Not IsNull(dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")) Then
''''''            frmExibicao7Ref.lblCodFunc.Caption = dteEtiquetas.rsEtiquetas.Fields("EMBALAGEM")
''''''        Else
''''''            frmExibicao7Ref.lblCodFunc.Caption = ""
''''''        End If
''''''
''''''    End If
''''''

