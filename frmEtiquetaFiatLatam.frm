VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEtiquetaFiatLatam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Etiquetas Fiat Latam"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9585
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
      Left            =   1350
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   5580
      Width           =   4695
   End
   Begin VB.TextBox TXT_EMBALAGEM 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1350
      MaxLength       =   20
      TabIndex        =   21
      ToolTipText     =   "Digie a sequencial da etiqueta"
      Top             =   5220
      Width           =   1515
   End
   Begin VB.CommandButton cmd_Impressao_Master 
      Enabled         =   0   'False
      Height          =   735
      Left            =   2940
      Picture         =   "frmEtiquetaFiatLatam.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Imprimir"
      Top             =   5010
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   3705
      Left            =   90
      TabIndex        =   11
      Top             =   1260
      Width           =   9285
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   3315
         Left            =   120
         TabIndex        =   12
         Top             =   180
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   5847
         _Version        =   393216
         Cols            =   37
         ForeColorFixed  =   16711680
         BackColorSel    =   65535
         ForeColorSel    =   65535
         AllowBigSelection=   0   'False
         HighLight       =   0
         GridLines       =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"frmEtiquetaFiatLatam.frx":030A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdfechar 
      Height          =   735
      Left            =   8430
      Picture         =   "frmEtiquetaFiatLatam.frx":0313
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Fecha a tela de etiqueta"
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmd_Impressao 
      Enabled         =   0   'False
      Height          =   735
      Left            =   7410
      Picture         =   "frmEtiquetaFiatLatam.frx":0755
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprimir"
      Top             =   5040
      Width           =   975
   End
   Begin VB.Frame frm_filtro 
      Caption         =   "Filtro da etiqueta"
      Height          =   1065
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   9285
      Begin VB.TextBox txtsequencial 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   4320
         MaxLength       =   16
         TabIndex        =   7
         ToolTipText     =   "Digie a sequencial da etiqueta"
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmd_limpar 
         Height          =   495
         Left            =   8610
         Picture         =   "frmEtiquetaFiatLatam.frx":0A5F
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Limpar tela para nova etiqueta"
         Top             =   270
         Width           =   555
      End
      Begin VB.CommandButton btoConfirma 
         Height          =   495
         Left            =   8040
         Picture         =   "frmEtiquetaFiatLatam.frx":0EA1
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Confirma dados do filtro"
         Top             =   270
         Width           =   555
      End
      Begin VB.Frame Frame1 
         Height          =   765
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   3345
         Begin VB.OptionButton Opt_Pallet 
            Caption         =   "Pelo Pallet"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   300
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton Opt_etiqueta 
            Caption         =   "Pelo Nº da Etiqueta"
            Height          =   255
            Left            =   1500
            TabIndex        =   3
            Top             =   270
            Width           =   1755
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Digitação:"
         Height          =   195
         Left            =   3570
         TabIndex        =   8
         Top             =   510
         Width           =   720
      End
   End
   Begin VB.CommandButton cmd_Visualizar 
      Caption         =   "Visualizar"
      Enabled         =   0   'False
      Height          =   735
      Left            =   6390
      TabIndex        =   0
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Impressora:"
      ForeColor       =   &H80000007&
      Height          =   195
      Left            =   180
      TabIndex        =   23
      Top             =   5610
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "NF.Embalagem:"
      Height          =   195
      Left            =   180
      TabIndex        =   20
      Top             =   5280
      Width           =   1125
   End
   Begin VB.Label lbl_qtd 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9180
      TabIndex        =   18
      Top             =   6780
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Qtde.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   8130
      TabIndex        =   17
      Top             =   6750
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label lbl_peca 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9180
      TabIndex        =   16
      Top             =   6525
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label lbl_produto 
      AutoSize        =   -1  'True
      Caption         =   "Peça.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   8130
      TabIndex        =   15
      Top             =   6510
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lbl_sequencia 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9180
      TabIndex        =   14
      Top             =   6270
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label label33 
      AutoSize        =   -1  'True
      Caption         =   "Sequência.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   8130
      TabIndex        =   13
      Top             =   6270
      Visible         =   0   'False
      Width           =   1035
   End
End
Attribute VB_Name = "frmEtiquetaFiatLatam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável para MDIapp
Public cRec As ADODB.Recordset 'conterá os dados do registro corrente
Public bAtivo As Boolean
Public bTelaImp As Boolean
Public bJafoi As Boolean
Public nLogin As Integer ' conterá o codigo do usuario, quando confirmar a senha
Public nTipo As Integer ' conterá o tipo do usuario, quando confirmar a senha
Public nMatricula As Double ' conterá a matricula do usuario, quando confirmar a senha

Private Sub btoConfirma_Click()

Dim nx As Integer
Dim sData As String
Dim sNota As String
Dim sLote As String

On Error GoTo Erro

Me.MousePointer = vbHourglass
Me.cmd_Impressao.Enabled = False
Me.cmd_Visualizar.Enabled = False
Me.cmd_Impressao_Master.Enabled = False

If Me.Opt_Pallet.Value = True Then
   sData = "1"
Else
   sData = "2"
End If
Set cRec = New ADODB.Recordset

Set cRec = CCTempneMov_Etiq.Mov_Etiq_Consultar_Filtro_FIAT_LATAM(sBancoMusashi, _
                                                                 Me.txtsequencial.Text, _
                                                                 sData)

If cRec.RecordCount > 0 Then
   Call carrega_Grid
   Me.cmd_Impressao.Enabled = True
   Me.cmd_Visualizar.Enabled = True
   Me.cmd_Impressao_Master.Enabled = True
   Grid1.col = 0
   Me.lbl_produto.Visible = True
   Me.lbl_qtd.Visible = True
   Me.lbl_sequencia.Visible = True
   Me.lbl_peca.Visible = True
   Me.label33.Visible = True
   Me.Label5.Visible = True
   Rem criticas a serem realizadas no arquivo caso seja por pallet
   If cRec.RecordCount > 0 And Me.Opt_Pallet.Value = True Then
      cRec.MoveFirst
      If IsNull(cRec!Num_Doc_Fiscal) Then
         MsgBox "Pallet sem Nota fiscal. Verifique na tela."
         Me.MousePointer = vbDefault
         Me.cmd_Impressao.Enabled = False
         Me.cmd_Visualizar.Enabled = False
         Me.cmd_Impressao_Master.Enabled = False
         Exit Sub
      End If
      sNota = Trim(cRec!Num_Doc_Fiscal)
      sLote = Trim(cRec!Num_Lote)
      While Not cRec.EOF
            If sNota <> Trim(cRec!Num_Doc_Fiscal) Then
               MsgBox "Existem Notas fiscais diferentes dentro de um mesmo Pallet. Verifique na tela."
               Me.cmd_Impressao.Enabled = False
               Me.cmd_Visualizar.Enabled = False
               Me.cmd_Impressao_Master.Enabled = False
            End If
            If sLote <> Trim(cRec!Num_Lote) Then
               MsgBox "Existem Lotes diferentes dentro mesmo Pallet. Verifique na tela."
'               Me.cmd_Impressao.Enabled = False
'               Me.cmd_Visualizar.Enabled = False
            End If
            cRec.MoveNext
      Wend
   End If
   
End If

Me.MousePointer = vbDefault

Exit Sub

Erro:

MsgBox Err.Description
Me.MousePointer = vbDefault
If Err.Number = 50000 Then
   cmd_limpar_Click
End If
If Err.Number = 50001 Then
   Call Limpar_Grid
   Me.txtsequencial.SetFocus
End If
End Sub

Private Sub cmd_Impressao_Click()

Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim nx As Double
Dim x As Printer
Dim rs As ADODB.Recordset
Dim sTexto As String
Dim nqtde As Double

nx = 0

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

Printer.Orientation = 2

On Error GoTo Erro
Set rs = New ADODB.Recordset

rs.Fields.Append "1_Peso_Bruto", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "2_ODM", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "3_Data_Producao", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "4_Data_validade", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "5_Data_Expedicao", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "6_Cod_Emb", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "7_Quantidade", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "8_DOCA", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "9_Ponto_Entrega", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "10_Control_Log_Qual", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "11_Cod_Fornecedor", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "12_Num_Doc_Fiscal", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "13_Lote_Sob_Desv", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "14_Qtde_emb", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "15_Num_Sheda_Serial", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "16_Id_Inter_Fornecedor", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "17_Embarque_Ctrl", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "18_Indicacao_Supl", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "19_Classe_Func", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "20_Dados_Transporte", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "21_Qtde_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "22_Num_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "23_Razao_Social", ADODB.DataTypeEnum.adChar, 30
rs.Fields.Append "24_Cod_Barras", ADODB.DataTypeEnum.adChar, 200
rs.Fields.Append "25_Desenho_Chrysler", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "26_Descricao_Produto", ADODB.DataTypeEnum.adChar, 40
rs.Fields.Append "27_Num_Desenho", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "28_Destino", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "29_Cod_Destino", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "30_Vinculo", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "31_Restricoes", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "32_QR_Codes", ADODB.DataTypeEnum.adChar, 200
rs.Fields.Append "33_LogoMarca", ADODB.DataTypeEnum.adChar, 100
rs.Fields.Append "34_DUM", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "35_PDF_417", ADODB.DataTypeEnum.adChar, 200
rs.Fields.Append "36_Incoterms", ADODB.DataTypeEnum.adChar, 20

rs.Open

Me.MousePointer = vbHourglass

cRec.MoveFirst
nx = 0

If Grid1.Rows > 0 Then
   
   rs.AddNew
 
   Grid1.col = 0: rs.Fields("1_Peso_Bruto").Value = Trim(Grid1.Text)
   Grid1.col = 1: rs.Fields("2_ODM").Value = Trim(Grid1.Text)
   Grid1.col = 2: rs.Fields("3_Data_Producao").Value = Trim(Grid1.Text)
   Grid1.col = 3: rs.Fields("4_Data_validade").Value = Trim(Grid1.Text)
   Grid1.col = 4
   If Format(Trim(Grid1.Text), "DD/MM/YYYY") = "01/01/1900" Then
      rs.Fields("5_Data_Expedicao").Value = Format(Now(), "DD/MM/YYYY")
   Else
      rs.Fields("5_Data_Expedicao").Value = Trim(Grid1.Text)
   End If
   Grid1.col = 5: rs.Fields("6_Cod_Emb").Value = Trim(Grid1.Text)
   Grid1.col = 6: rs.Fields("7_Quantidade").Value = Trim(Grid1.Text)
   Grid1.col = 7: rs.Fields("8_DOCA").Value = Trim(Grid1.Text)
   Grid1.col = 8: rs.Fields("9_Ponto_Entrega").Value = Trim(Grid1.Text)
   Grid1.col = 9: rs.Fields("10_Control_Log_Qual").Value = Trim(Grid1.Text)
   Grid1.col = 10
   If Format(Trim(Grid1.Text), "DD/MM/YYYY") = "01/01/1900" Then
      rs.Fields("11_Cod_Fornecedor").Value = " "
   Else
      rs.Fields("11_Cod_Fornecedor").Value = Trim(Grid1.Text)
   End If
   Grid1.col = 11: rs.Fields("12_Num_Doc_Fiscal").Value = Trim(Grid1.Text)
   Grid1.col = 12: rs.Fields("13_Lote_Sob_Desv").Value = Trim(Grid1.Text)
   Grid1.col = 13: rs.Fields("14_Qtde_emb").Value = Trim(Grid1.Text)
   Grid1.col = 14: rs.Fields("15_Num_Sheda_Serial").Value = Trim(Grid1.Text)
   Grid1.col = 15: rs.Fields("16_Id_Inter_Fornecedor").Value = Trim(Grid1.Text)
   Grid1.col = 16: rs.Fields("17_Embarque_Ctrl").Value = Trim(Grid1.Text)
   Grid1.col = 17: rs.Fields("18_Indicacao_Supl").Value = Trim(Grid1.Text)
   Grid1.col = 18: rs.Fields("19_Classe_Func").Value = Trim(Grid1.Text)
   Grid1.col = 19: rs.Fields("20_Dados_Transporte").Value = Trim(Grid1.Text)
   Grid1.col = 20: rs.Fields("21_Qtde_Lote").Value = Trim(Grid1.Text)
   Grid1.col = 21: rs.Fields("22_Num_Lote").Value = Trim(Grid1.Text)
   Grid1.col = 22: rs.Fields("23_Razao_Social").Value = Trim(Grid1.Text)
   Rem ver aqui marcos
Rem construir o codigo de barras (24_Cod_Barras)
   Grid1.col = 26: sTexto = Trim(Grid1.Text) '"27_Num_Desenho"
   sTexto = sTexto & "0013093" ' codigo fornecedor da Musashi no cliente
   Grid1.col = 6: sTexto = sTexto & Format(Trim(Grid1.Text), "00000") '"7_Quantidade"
   Grid1.col = 5: sTexto = sTexto & Format(Trim(Grid1.Text), "0000") ' "6_Cod_Emb"
   rs.Fields("24_Cod_Barras").Value = "*" & sTexto & "*"
   
   Grid1.col = 24: rs.Fields("25_Desenho_Chrysler").Value = Trim(Grid1.Text)
   Grid1.col = 25: rs.Fields("26_Descricao_Produto").Value = Mid$(Trim(Grid1.Text), 1, 40)
   Grid1.col = 26: rs.Fields("27_Num_Desenho").Value = Trim(Grid1.Text)
   Grid1.col = 27: rs.Fields("28_Destino").Value = Mid$(Trim(Grid1.Text), 1, 20)
   Grid1.col = 28: rs.Fields("29_Cod_Destino").Value = Trim(Grid1.Text)
   Grid1.col = 29: rs.Fields("30_Vinculo").Value = Trim(Grid1.Text)
   Grid1.col = 30: rs.Fields("31_Restricoes").Value = Trim(Grid1.Text)
   
   Rem Formação dos dados referente a etiqueta QR-CODE
'    1(11)       2(7)    3(10)      4(8)     5(3)6(5)  7(4)8(4) 9(3)10(9)     11(4)12(8)    13(10)      14(1)
'   "00055267989 0013093 0000000001 09112016 G09 00002 CPG 0001 000 000330999 0000 00000000 00160920160 0
   sTexto = ""
   Grid1.col = 26: sTexto = sTexto & Format(Trim(Grid1.Text), "00000000000")  '"27_Num_Desenho" 1
   sTexto = sTexto & "0013093"  ' codigo fornecedor da Musashi no cliente 2
   sTexto = sTexto & "0000000001"  ' id interno fornecedor nao se repete durante um ano 3 - VER COM MAURO
   Grid1.col = 2: sTexto = sTexto & Format(Trim(Grid1.Text), "DDMMYYYY") '"3_Data_Producao"
   Grid1.col = 8: sTexto = sTexto & Trim(Grid1.Text) '"9_Ponto_Entrega" 5 - VER COM MAURO
   Grid1.col = 6: sTexto = sTexto & Format(Trim(Grid1.Text), "00000") '"7_Quantidade" 6
   Grid1.col = 5: sTexto = sTexto & Format(Trim(Grid1.Text), "0000") ' "6_Cod_Emb" 7 - VER COM MAURO
   sTexto = sTexto & "0001"  ' Qtde de embalagem enviada 8
   sTexto = sTexto & "000"  ' cod de destino 9 - VER COM MAURO
   Grid1.col = 11: sTexto = sTexto & Format(Trim(Grid1.Text), "000000000") '"12_Num_Doc_Fiscal" 10
   
   Grid1.col = 1:
   If Len(Trim(Mid$(Trim(Grid1.Text), 1, 4))) > 0 Then
      sTexto = sTexto & Format(Mid$(Trim(Grid1.Text), 1, 4), "0000") '"2_ODM" 11
   Else
      sTexto = sTexto & "0000" '"2_ODM" 11
   End If
   sTexto = sTexto & "00000000"  ' Data de Validade '12
   Grid1.col = 21: sTexto = sTexto & Format(Trim(Grid1.Text), "0000000000") '"22_Num_Lote" 13
   sTexto = sTexto & "0"  ' Tipo da etiqueta 14
   Grid1.col = 31: rs.Fields("32_QR_Codes").Value = sTexto
'********************************************************************************************************
'********************************************************************************************************

   Grid1.col = 32: rs.Fields("33_LogoMarca").Value = Trim(Grid1.Text)
   Grid1.col = 33: rs.Fields("34_DUM").Value = Trim(Grid1.Text)
   
   Rem Formação dos dados referente a etiqueta PDF_417
   'sTexto = "[)><0x1e>06<0x1d>P9999999999<0x1d>Q00002<0x1d>V0013093<0x1d>3S00000<0x1d>1P0000055267989<0x1e><0x04>"
   
   sTexto = "[)><0x1e>06<0x1d>P9999999999<0x1d>"
   Grid1.col = 6: sTexto = sTexto & "Q" & Format(Trim(Grid1.Text), "000000") & "<0x1d>" '"7_Quantidade"
   sTexto = sTexto & "V" & "0013093" & "<0x1d>" ' codigo fornecedor da Musashi no cliente
   sTexto = sTexto & "3S" & "0000000000" & "<0x1d>" ' Numero serial
   Grid1.col = 26: sTexto = sTexto & "1P" & Format(Trim(Grid1.Text), "00000000000") & "<0x1e><0x04>" '"27_Num_Desenho"
   Grid1.col = 34: rs.Fields("35_PDF_417").Value = sTexto
'********************************************************************************************************
'********************************************************************************************************

   Grid1.col = 35: rs.Fields("36_Incoterms").Value = Trim(Grid1.Text)
   
   rs.Update
   
End If

Me.MousePointer = vbHourglass

Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_FIAT_LATAM.rpt")

CrystalReport1.Database.SetDataSource rs

rs.Clone

CrystalReport1.PrintOutEx False

Me.MousePointer = vbDefault

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

Private Sub cmd_Impressao_Master_Click()

Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim oTela As frmEscRelCristalReport
Dim nx As Double
Dim x As Printer
Dim rs As ADODB.Recordset
Dim sTexto As String
Dim nqtde As Double

If Len(Trim(Me.TXT_EMBALAGEM.Text)) = 0 Then
   MsgBox "Digite a Embalagem, para emitir o Relatório de Etiqueta Master!"
   Me.TXT_EMBALAGEM.SetFocus
   Exit Sub
End If

Set oTela = New frmEscRelCristalReport

On Error GoTo Erro

Set rs = New ADODB.Recordset

rs.Fields.Append "1_Peso_Bruto", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "2_ODM", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "3_Data_Producao", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "4_Data_validade", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "5_Data_Expedicao", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "6_Cod_Emb", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "7_Quantidade", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "8_DOCA", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "9_Ponto_Entrega", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "10_Control_Log_Qual", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "11_Cod_Fornecedor", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "12_Num_Doc_Fiscal", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "13_Lote_Sob_Desv", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "14_Qtde_emb", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "15_Num_Sheda_Serial", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "16_Id_Inter_Fornecedor", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "17_Embarque_Ctrl", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "18_Indicacao_Supl", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "19_Classe_Func", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "20_Dados_Transporte", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "21_Qtde_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "22_Num_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "23_Razao_Social", ADODB.DataTypeEnum.adChar, 30
rs.Fields.Append "24_Cod_Barras", ADODB.DataTypeEnum.adChar, 200
rs.Fields.Append "25_Desenho_Chrysler", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "26_Descricao_Produto", ADODB.DataTypeEnum.adChar, 40
rs.Fields.Append "27_Num_Desenho", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "28_Destino", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "29_Cod_Destino", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "30_Vinculo", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "31_Restricoes", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "32_QR_Codes", ADODB.DataTypeEnum.adChar, 200
rs.Fields.Append "33_LogoMarca", ADODB.DataTypeEnum.adChar, 100
rs.Fields.Append "34_DUM", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "35_PDF_417", ADODB.DataTypeEnum.adChar, 200
rs.Fields.Append "36_Incoterms", ADODB.DataTypeEnum.adChar, 20

rs.Open

Me.MousePointer = vbHourglass

nx = 0

Grid1.Row = 1
cRec.MoveFirst
bJafoi = True

nqtde = 0
While Not cRec.EOF
    nqtde = nqtde + cRec!quantidade
    cRec.MoveNext
Wend

For nx = 1 To 25
   rs.AddNew

   rs.Fields("1_Peso_Bruto").Value = str(nx)
   Grid1.col = 1: rs.Fields("2_ODM").Value = Trim(Grid1.Text)
   Grid1.col = 2: rs.Fields("3_Data_Producao").Value = Trim(Grid1.Text)
   Grid1.col = 3: rs.Fields("4_Data_validade").Value = Trim(Grid1.Text)
   Grid1.col = 4
   If Format(Trim(Grid1.Text), "DD/MM/YYYY") = "01/01/1900" Then
      rs.Fields("5_Data_Expedicao").Value = " "
   Else
      rs.Fields("5_Data_Expedicao").Value = Trim(Grid1.Text)
   End If
   Grid1.col = 5: rs.Fields("6_Cod_Emb").Value = Trim(Grid1.Text)
   rs.Fields("7_Quantidade").Value = nqtde
   Grid1.col = 7: rs.Fields("8_DOCA").Value = Trim(Grid1.Text)
   Grid1.col = 8: rs.Fields("9_Ponto_Entrega").Value = Trim(Grid1.Text)
   Grid1.col = 9: rs.Fields("10_Control_Log_Qual").Value = Trim(Grid1.Text)
   Grid1.col = 10
   If Format(Trim(Grid1.Text), "DD/MM/YYYY") = "01/01/1900" Then
      rs.Fields("11_Cod_Fornecedor").Value = " "
   Else
      rs.Fields("11_Cod_Fornecedor").Value = Trim(Grid1.Text)
   End If
   Grid1.col = 11: rs.Fields("12_Num_Doc_Fiscal").Value = Trim(Grid1.Text)
   Grid1.col = 12: rs.Fields("13_Lote_Sob_Desv").Value = Trim(Grid1.Text)
   rs.Fields("14_Qtde_emb").Value = cRec.RecordCount
   Grid1.col = 14: rs.Fields("15_Num_Sheda_Serial").Value = Trim(Grid1.Text)
   Grid1.col = 15: rs.Fields("16_Id_Inter_Fornecedor").Value = Trim(Grid1.Text)
   Grid1.col = 16: rs.Fields("17_Embarque_Ctrl").Value = Trim(Grid1.Text)
   Grid1.col = 17: rs.Fields("18_Indicacao_Supl").Value = Trim(Grid1.Text)
   Grid1.col = 18: rs.Fields("19_Classe_Func").Value = Trim(Grid1.Text)
   Grid1.col = 19: rs.Fields("20_Dados_Transporte").Value = Trim(Grid1.Text)
   Grid1.col = 20: rs.Fields("21_Qtde_Lote").Value = Trim(Grid1.Text)
   Grid1.col = 21: rs.Fields("22_Num_Lote").Value = Trim(Grid1.Text)
   Grid1.col = 22: rs.Fields("23_Razao_Social").Value = Trim(Grid1.Text)
   Grid1.col = 23: rs.Fields("24_Cod_Barras").Value = Trim(Grid1.Text)
   Grid1.col = 24: rs.Fields("25_Desenho_Chrysler").Value = Trim(Grid1.Text)
   Grid1.col = 25: rs.Fields("26_Descricao_Produto").Value = Mid$(Trim(Grid1.Text), 1, 40)
   Grid1.col = 26: rs.Fields("27_Num_Desenho").Value = Trim(Grid1.Text)
   Grid1.col = 27: rs.Fields("28_Destino").Value = Mid$(Trim(Grid1.Text), 1, 20)
   Grid1.col = 28: rs.Fields("29_Cod_Destino").Value = Trim(Grid1.Text)
   Grid1.col = 29: rs.Fields("30_Vinculo").Value = Trim(Grid1.Text)
   Grid1.col = 30: rs.Fields("31_Restricoes").Value = Trim(Grid1.Text)
   Grid1.col = 31: rs.Fields("32_QR_Codes").Value = Trim(Grid1.Text)
   rs.Fields("33_LogoMarca").Value = Me.TXT_EMBALAGEM.Text
   Grid1.col = 33: rs.Fields("34_DUM").Value = Trim(Grid1.Text)
'    1           2       3   4         5         6     7     8    9   10        11   12
'   "00055267989 0013093 00  000000235 000540403 00271 620";'   "003309990 0000 00000000 0 16092016 0";
   sTexto = ""
   Grid1.col = 26: sTexto = sTexto & Format(Trim(Grid1.Text), "00000000000")  '"27_Num_Desenho" 1
   sTexto = sTexto & "0013093"  ' codigo fornecedor da Musashi no cliente 2
   sTexto = sTexto & "01"  ' id interno 3
   Grid1.col = 11: sTexto = sTexto & Format(Trim(Grid1.Text), "000000000") '"12_Num_Doc_Fiscal" 4
   Grid1.col = 11: sTexto = sTexto & Trim(Me.TXT_EMBALAGEM.Text) 'NF embalagem, digitado  5
   Grid1.col = 6: sTexto = sTexto & Format(nqtde, "00000") '"7_Quantidade" 6
   Grid1.col = 5: sTexto = sTexto & Format(Trim(Grid1.Text), "0000") ' "6_Cod_Emb" 7
   sTexto = sTexto & Format(cRec.RecordCount, "0000") '"14_Qtde_emb" 8
   sTexto = sTexto & "000"  ' cod de destino 9
   sTexto = sTexto & "0"  ' Tipo da etiqueta 14
   
   Grid1.col = 31: rs.Fields("32_QR_Codes").Value = sTexto
   
   Grid1.col = 34: rs.Fields("35_PDF_417").Value = Trim(Grid1.Text)
   Grid1.col = 35: rs.Fields("36_Incoterms").Value = Trim(Grid1.Text)

   
   rs.Update

Next

Me.MousePointer = vbHourglass

Set oTela = New frmEscRelCristalReport

Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_FIAT_LATAM_MASTER1.rpt")

CrystalReport1.Database.SetDataSource rs

rs.Clone
oTela.CRViewer1.ReportSource = CrystalReport1
oTela.CRViewer1.ViewReport

oTela.Show 0

Me.MousePointer = vbDefault

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

Private Sub cmd_limpar_Click()
Call Fechar_Form_Etiqueta

Call Limpar_Grid

Me.txtsequencial.Text = ""

Me.lbl_produto.Visible = False
Me.lbl_qtd.Visible = False
Me.lbl_sequencia.Visible = False
Me.lbl_peca.Visible = False
Me.label33.Visible = False
Me.Label5.Visible = False
Me.txtsequencial.Enabled = True

Me.cmd_Impressao.Enabled = False
Me.cmd_Visualizar.Enabled = False
Me.cmd_Impressao_Master.Enabled = False

End Sub

Private Sub cmdfechar_Click()
Unload Me
End Sub

Private Sub cmd_visualizar_Click()

Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim oTela As frmEscRelCristalReport
Dim nx As Double
Dim x As Printer
Dim rs As ADODB.Recordset
Dim sTexto As String
Dim nqtde As Double

Set oTela = New frmEscRelCristalReport

On Error GoTo Erro

Set rs = New ADODB.Recordset

rs.Fields.Append "1_Peso_Bruto", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "2_ODM", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "3_Data_Producao", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "4_Data_validade", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "5_Data_Expedicao", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "6_Cod_Emb", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "7_Quantidade", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "8_DOCA", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "9_Ponto_Entrega", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "10_Control_Log_Qual", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "11_Cod_Fornecedor", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "12_Num_Doc_Fiscal", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "13_Lote_Sob_Desv", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "14_Qtde_emb", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "15_Num_Sheda_Serial", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "16_Id_Inter_Fornecedor", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "17_Embarque_Ctrl", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "18_Indicacao_Supl", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "19_Classe_Func", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "20_Dados_Transporte", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "21_Qtde_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "22_Num_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "23_Razao_Social", ADODB.DataTypeEnum.adChar, 30
rs.Fields.Append "24_Cod_Barras", ADODB.DataTypeEnum.adChar, 200
rs.Fields.Append "25_Desenho_Chrysler", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "26_Descricao_Produto", ADODB.DataTypeEnum.adChar, 40
rs.Fields.Append "27_Num_Desenho", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "28_Destino", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "29_Cod_Destino", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "30_Vinculo", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "31_Restricoes", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "32_QR_Codes", ADODB.DataTypeEnum.adChar, 200
rs.Fields.Append "33_LogoMarca", ADODB.DataTypeEnum.adChar, 100
rs.Fields.Append "34_DUM", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "35_PDF_417", ADODB.DataTypeEnum.adChar, 200
rs.Fields.Append "36_Incoterms", ADODB.DataTypeEnum.adChar, 20

rs.Open

Me.MousePointer = vbHourglass

nx = 0

If Grid1.Rows > 0 Then
   nqtde = 0
   rs.AddNew

   Grid1.col = 0: rs.Fields("1_Peso_Bruto").Value = Trim(Grid1.Text)
   Grid1.col = 1: rs.Fields("2_ODM").Value = Trim(Grid1.Text)
   Grid1.col = 2: rs.Fields("3_Data_Producao").Value = Trim(Grid1.Text)
   Grid1.col = 3: rs.Fields("4_Data_validade").Value = Trim(Grid1.Text)
   Grid1.col = 4
   If Format(Trim(Grid1.Text), "DD/MM/YYYY") = "01/01/1900" Then
      rs.Fields("5_Data_Expedicao").Value = " "
   Else
      rs.Fields("5_Data_Expedicao").Value = Trim(Grid1.Text)
   End If
   Grid1.col = 5: rs.Fields("6_Cod_Emb").Value = Trim(Grid1.Text)
   Grid1.col = 6: rs.Fields("7_Quantidade").Value = Trim(Grid1.Text)
   Grid1.col = 7: rs.Fields("8_DOCA").Value = Trim(Grid1.Text)
   Grid1.col = 8: rs.Fields("9_Ponto_Entrega").Value = Trim(Grid1.Text)
   Grid1.col = 9: rs.Fields("10_Control_Log_Qual").Value = Trim(Grid1.Text)
   Grid1.col = 10
   If Format(Trim(Grid1.Text), "DD/MM/YYYY") = "01/01/1900" Then
      rs.Fields("11_Cod_Fornecedor").Value = " "
   Else
      rs.Fields("11_Cod_Fornecedor").Value = Trim(Grid1.Text)
   End If
   Grid1.col = 11: rs.Fields("12_Num_Doc_Fiscal").Value = Trim(Grid1.Text)
   Grid1.col = 12: rs.Fields("13_Lote_Sob_Desv").Value = Trim(Grid1.Text)
   Grid1.col = 13: rs.Fields("14_Qtde_emb").Value = Trim(Grid1.Text)
   Grid1.col = 14: rs.Fields("15_Num_Sheda_Serial").Value = Trim(Grid1.Text)
   Grid1.col = 15: rs.Fields("16_Id_Inter_Fornecedor").Value = Trim(Grid1.Text)
   Grid1.col = 16: rs.Fields("17_Embarque_Ctrl").Value = Trim(Grid1.Text)
   Grid1.col = 17: rs.Fields("18_Indicacao_Supl").Value = Trim(Grid1.Text)
   Grid1.col = 18: rs.Fields("19_Classe_Func").Value = Trim(Grid1.Text)
   Grid1.col = 19: rs.Fields("20_Dados_Transporte").Value = Trim(Grid1.Text)
   Grid1.col = 20: rs.Fields("21_Qtde_Lote").Value = Trim(Grid1.Text)
   Grid1.col = 21: rs.Fields("22_Num_Lote").Value = Trim(Grid1.Text)
   Grid1.col = 22: rs.Fields("23_Razao_Social").Value = Trim(Grid1.Text)
   Rem ver aqui marcos
Rem construir o codigo de barras (24_Cod_Barras)
   Grid1.col = 26: sTexto = Trim(Grid1.Text) '"27_Num_Desenho"
   sTexto = sTexto & "0013093" ' codigo fornecedor da Musashi no cliente
   Grid1.col = 6: sTexto = sTexto & Format(Trim(Grid1.Text), "00000") '"7_Quantidade"
   Grid1.col = 5: sTexto = sTexto & Format(Trim(Grid1.Text), "0000") ' "6_Cod_Emb"
   rs.Fields("24_Cod_Barras").Value = "*" & sTexto & "*"
   
   Grid1.col = 24: rs.Fields("25_Desenho_Chrysler").Value = Trim(Grid1.Text)
   Grid1.col = 25: rs.Fields("26_Descricao_Produto").Value = Mid$(Trim(Grid1.Text), 1, 40)
   Grid1.col = 26: rs.Fields("27_Num_Desenho").Value = Trim(Grid1.Text)
   Grid1.col = 27: rs.Fields("28_Destino").Value = Mid$(Trim(Grid1.Text), 1, 20)
   Grid1.col = 28: rs.Fields("29_Cod_Destino").Value = Trim(Grid1.Text)
   Grid1.col = 29: rs.Fields("30_Vinculo").Value = Trim(Grid1.Text)
   Grid1.col = 30: rs.Fields("31_Restricoes").Value = Trim(Grid1.Text)
   
   Rem Formação dos dados referente a etiqueta QR-CODE
'    1(11)       2(7)    3(10)      4(8)     5(3)6(5)  7(4)8(4) 9(3)10(9)     11(4)12(8)    13(10)      14(1)
'   "00055267989 0013093 0000000001 09112016 G09 00002 CPG 0001 000 000330999 0000 00000000 00160920160 0
   sTexto = ""
   Grid1.col = 26: sTexto = sTexto & Format(Trim(Grid1.Text), "00000000000")  '"27_Num_Desenho" 1
   sTexto = sTexto & "0013093"  ' codigo fornecedor da Musashi no cliente 2
   sTexto = sTexto & "0000000001"  ' id interno fornecedor nao se repete durante um ano 3 - VER COM MAURO
   Grid1.col = 2: sTexto = sTexto & Format(Trim(Grid1.Text), "DDMMYYYY") '"3_Data_Producao"
   Grid1.col = 8: sTexto = sTexto & Trim(Grid1.Text) '"9_Ponto_Entrega" 5 - VER COM MAURO
   Grid1.col = 6: sTexto = sTexto & Format(Trim(Grid1.Text), "00000") '"7_Quantidade" 6
   Grid1.col = 5: sTexto = sTexto & Format(Trim(Grid1.Text), "0000") ' "6_Cod_Emb" 7 - VER COM MAURO
   sTexto = sTexto & "0001"  ' Qtde de embalagem enviada 8
   sTexto = sTexto & "000"  ' cod de destino 9 - VER COM MAURO
   Grid1.col = 11: sTexto = sTexto & Format(Trim(Grid1.Text), "000000000") '"12_Num_Doc_Fiscal" 10
   
   Grid1.col = 1:
   If Len(Trim(Mid$(Trim(Grid1.Text), 1, 4))) > 0 Then
      sTexto = sTexto & Format(Mid$(Trim(Grid1.Text), 1, 4), "0000") '"2_ODM" 11
   Else
      sTexto = sTexto & "0000" '"2_ODM" 11
   End If
   sTexto = sTexto & "00000000"  ' Data de Validade '12
   Grid1.col = 21: sTexto = sTexto & Format(Trim(Grid1.Text), "0000000000") '"22_Num_Lote" 13
   sTexto = sTexto & "0"  ' Tipo da etiqueta 14
   Grid1.col = 31: rs.Fields("32_QR_Codes").Value = sTexto
'********************************************************************************************************
'********************************************************************************************************

   Grid1.col = 32: rs.Fields("33_LogoMarca").Value = Trim(Grid1.Text)
   Grid1.col = 33: rs.Fields("34_DUM").Value = Trim(Grid1.Text)
   
   Rem Formação dos dados referente a etiqueta PDF_417
   'sTexto = "[)><0x1e>06<0x1d>P9999999999<0x1d>Q00002<0x1d>V0013093<0x1d>3S00000<0x1d>1P0000055267989<0x1e><0x04>"
   
   sTexto = "[)><0x1e>06<0x1d>P9999999999<0x1d>"
   Grid1.col = 6: sTexto = sTexto & "Q" & Format(Trim(Grid1.Text), "000000") & "<0x1d>" '"7_Quantidade"
   sTexto = sTexto & "V" & "0013093" & "<0x1d>" ' codigo fornecedor da Musashi no cliente
   sTexto = sTexto & "3S" & "0000000000" & "<0x1d>" ' Numero serial
   Grid1.col = 26: sTexto = sTexto & "1P" & Format(Trim(Grid1.Text), "00000000000") & "<0x1e><0x04>" '"27_Num_Desenho"
   Grid1.col = 34: rs.Fields("35_PDF_417").Value = sTexto
'********************************************************************************************************
'********************************************************************************************************

   Grid1.col = 35: rs.Fields("36_Incoterms").Value = Trim(Grid1.Text)
   
   rs.Update

'   For nx = 1 To Grid1.Rows - 1
'       Grid1.Col = 1: nqtde = nqtde + VBA.CDbl(Trim(Grid1.Text))
'       Grid1.Row = nx
'   Next
'
  
End If

Me.MousePointer = vbHourglass

Set oTela = New frmEscRelCristalReport

If Me.Opt_Pallet.Value = True Then
   Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_FIAT_LATAM.rpt")
Else
   Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_FIAT_LATAM.rpt")
End If

CrystalReport1.Database.SetDataSource rs

rs.Clone
oTela.CRViewer1.ReportSource = CrystalReport1
oTela.CRViewer1.ViewReport

oTela.Show 0

Me.MousePointer = vbDefault

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()

'If bTelaImp Then
'   cmd_limpar_Click
'   bTelaImp = False
'End If

If bAtivo Then Exit Sub

bAtivo = True
bJafoi = False
Call Limpar_Grid
Me.txtsequencial.Text = ""
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' para funcionar , tem que mudar o keyPreviwe=true
If KeyCode = 13 Then
   If Me.ActiveControl.TabIndex > -1 Then
      SendKeys "{TAB}"
   End If
ElseIf KeyCode = 27 Then
   If Me.ActiveControl.TabIndex < 1 Then
      If 6 = MsgBox("Deseja realmente sair deste módulo?", 32 + 4) Then
         Unload Me
      End If
   Else
       SendKeys "+{TAB}" ' retornar campo
   End If
End If

End Sub

Private Sub Form_Load()
Dim x As Printer
Dim nx As Integer

nx = 0
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
Me.Left = 0
Me.Top = 0
bTelaImp = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
bAtivo = False
'Fecha o form q estiver aberto
Call Fechar_Form_Etiqueta
End Sub

Private Sub carrega_Grid()
Dim nx As Double
Dim nLinhas As Double
Dim nqtde As Double

Call Limpar_Grid

Grid1.Row = 1
cRec.MoveFirst
bJafoi = True

For nx = 1 To cRec.RecordCount
    Grid1.col = 0: Grid1.Text = IIf(IsNull(cRec.Fields("Peso_Bruto")), " ", cRec.Fields("Peso_Bruto"))
    Grid1.col = 1: Grid1.Text = IIf(IsNull(cRec.Fields("ODM")), " ", cRec.Fields("ODM"))
    Grid1.col = 2: Grid1.Text = IIf(IsNull(cRec.Fields("Data_Producao")), " ", cRec.Fields("Data_Producao"))
    Grid1.col = 3: Grid1.Text = IIf(IsNull(cRec.Fields("Data_validade")), " ", cRec.Fields("Data_validade"))
    Grid1.col = 4: Grid1.Text = IIf(IsNull(cRec.Fields("Data_Expedicao")), " ", cRec.Fields("Data_Expedicao"))
    Grid1.col = 5: Grid1.Text = IIf(IsNull(cRec.Fields("Cod_Emb")), "0000", Format(cRec.Fields("Cod_Emb"), "0000"))
    Grid1.col = 6: Grid1.Text = IIf(IsNull(cRec.Fields("Quantidade")), " ", Format(cRec.Fields("Quantidade"), "00000"))
    Grid1.col = 7: Grid1.Text = IIf(IsNull(cRec.Fields("DOCA")), " ", cRec.Fields("DOCA"))
    Grid1.col = 8: Grid1.Text = IIf(IsNull(cRec.Fields("Ponto_Entrega")), " ", cRec.Fields("Ponto_Entrega"))
    Grid1.col = 9: Grid1.Text = IIf(IsNull(cRec.Fields("Control_Log_Qual")), " ", cRec.Fields("Control_Log_Qual"))
    Grid1.col = 10: Grid1.Text = IIf(IsNull(cRec.Fields("Cod_Fornecedor")), " ", Format(cRec.Fields("Cod_Fornecedor"), "0000000"))
    Grid1.col = 11: Grid1.Text = IIf(IsNull(cRec.Fields("Num_Doc_Fiscal")), " ", cRec.Fields("Num_Doc_Fiscal"))
    Grid1.col = 12: Grid1.Text = IIf(IsNull(cRec.Fields("Lote_Sob_Desv")), " ", cRec.Fields("Lote_Sob_Desv"))
    Grid1.col = 13: Grid1.Text = IIf(IsNull(cRec.Fields("Qtde_emb")), "0000", Format(cRec.Fields("Qtde_emb"), "0000"))
    Grid1.col = 14: Grid1.Text = IIf(IsNull(cRec.Fields("Num_Sheda_Serial")), " ", cRec.Fields("Num_Sheda_Serial"))
    Grid1.col = 15: Grid1.Text = IIf(IsNull(cRec.Fields("Id_Inter_Fornecedor")), " ", cRec.Fields("Id_Inter_Fornecedor"))
    Grid1.col = 16: Grid1.Text = IIf(IsNull(cRec.Fields("Embarque_Ctrl")), " ", cRec.Fields("Embarque_Ctrl"))
    Grid1.col = 17: Grid1.Text = IIf(IsNull(cRec.Fields("Indicacao_Supl")), " ", cRec.Fields("Indicacao_Supl"))
    Grid1.col = 18: Grid1.Text = IIf(IsNull(cRec.Fields("Classe_Func")), " ", cRec.Fields("Classe_Func"))
    Grid1.col = 19: Grid1.Text = IIf(IsNull(cRec.Fields("Dados_Transporte")), " ", cRec.Fields("Dados_Transporte"))
    Grid1.col = 20: Grid1.Text = IIf(IsNull(cRec.Fields("Qtde_Lote")), "00000", Format(cRec.Fields("Qtde_Lote"), "00000"))
    Grid1.col = 21: Grid1.Text = IIf(IsNull(cRec.Fields("Num_Lote")), " ", cRec.Fields("Num_Lote"))
    Grid1.col = 22: Grid1.Text = IIf(IsNull(cRec.Fields("Razao_Social")), " ", cRec.Fields("Razao_Social"))
    Grid1.col = 23: Grid1.Text = IIf(IsNull(cRec.Fields("Cod_Barras")), " ", cRec.Fields("Cod_Barras"))
    Grid1.col = 24: Grid1.Text = IIf(IsNull(cRec.Fields("Desenho_Chrysler")), " ", cRec.Fields("Desenho_Chrysler"))
    Grid1.col = 25: Grid1.Text = IIf(IsNull(cRec.Fields("Descricao_Produto")), " ", cRec.Fields("Descricao_Produto"))
    Grid1.col = 26: Grid1.Text = IIf(IsNull(cRec.Fields("Num_Desenho")), " ", Format(cRec.Fields("Num_Desenho"), "00000000000"))
    Grid1.col = 27: Grid1.Text = IIf(IsNull(cRec.Fields("Destino")), " ", cRec.Fields("Destino"))
    Grid1.col = 28: Grid1.Text = IIf(IsNull(cRec.Fields("Cod_Destino")), " ", cRec.Fields("Cod_Destino"))
    Grid1.col = 29: Grid1.Text = IIf(IsNull(cRec.Fields("Vinculo")), " ", cRec.Fields("Vinculo"))
    Grid1.col = 30: Grid1.Text = IIf(IsNull(cRec.Fields("Restricoes")), " ", cRec.Fields("Restricoes"))
    Grid1.col = 31: Grid1.Text = IIf(IsNull(cRec.Fields("QR_Codes")), " ", cRec.Fields("QR_Codes"))
    Grid1.col = 32: Grid1.Text = IIf(IsNull(cRec.Fields("LogoMarca")), " ", cRec.Fields("LogoMarca"))
    Grid1.col = 33: Grid1.Text = IIf(IsNull(cRec.Fields("DUM")), " ", cRec.Fields("DUM"))
    Grid1.col = 34: Grid1.Text = IIf(IsNull(cRec.Fields("PDF_417")), " ", cRec.Fields("PDF_417"))
    Grid1.col = 35: Grid1.Text = IIf(IsNull(cRec.Fields("Incoterms")), " ", cRec.Fields("Incoterms"))
    Grid1.col = 36: Grid1.Text = IIf(IsNull(cRec.Fields("ID_ETIQUETA")), " ", cRec.Fields("ID_ETIQUETA"))
    
'    Grid1.Col = 19: Me.lbl_sequencia.Caption = Me.Grid1.Text
'    Grid1.Col = 18: Me.lbl_peca.Caption = Me.Grid1.Text
'    Grid1.Col = 1: nqtde = nqtde + VBA.CDbl(Me.Grid1.Text)
    Grid1.col = 0
    
    cRec.MoveNext
    
    If Not cRec.EOF Then
       Grid1.Rows = Grid1.Rows + 1
       Grid1.Row = Grid1.Row + 1
    End If
Next

Me.lbl_qtd.Caption = nqtde

bJafoi = False
End Sub

Private Sub Limpar_Grid()
Dim nx As Double
Dim nLinhas As Double
Dim nLinhas1 As Double

Grid1.Clear
nLinhas = Grid1.Rows

If Grid1.Rows > 2 Then
   For nx = Grid1.Rows To nLinhas1 - 2 Step -1
       If nx > 2 Then Grid1.RemoveItem (nx)
   Next
End If

    Grid1.Row = 0
    Grid1.col = 0: Grid1.ColWidth(0) = 1000: Grid1.Text = "Peso_Bruto"
    Grid1.col = 1: Grid1.ColWidth(0) = 1000: Grid1.Text = "ODM"
    Grid1.col = 2: Grid1.ColWidth(0) = 1000: Grid1.Text = "Data_Producao"
    Grid1.col = 3: Grid1.ColWidth(0) = 1000: Grid1.Text = "Data_validade"
    Grid1.col = 4: Grid1.ColWidth(0) = 1000: Grid1.Text = "Data_Expedicao"
    Grid1.col = 5: Grid1.ColWidth(0) = 1000: Grid1.Text = "Cod_Emb"
    Grid1.col = 6: Grid1.ColWidth(0) = 1000: Grid1.Text = "Quantidade"
    Grid1.col = 7: Grid1.ColWidth(0) = 1000: Grid1.Text = "DOCA"
    Grid1.col = 8: Grid1.ColWidth(0) = 1000: Grid1.Text = "Ponto_Entrega"
    Grid1.col = 9: Grid1.ColWidth(0) = 1000: Grid1.Text = "Control_Log_Qual"
    Grid1.col = 10: Grid1.ColWidth(0) = 1000: Grid1.Text = "Cod_Fornecedor"
    Grid1.col = 11: Grid1.ColWidth(0) = 1000: Grid1.Text = "Num_Doc_Fiscal"
    Grid1.col = 12: Grid1.ColWidth(0) = 1000: Grid1.Text = "Lote_Sob_Desv"
    Grid1.col = 13: Grid1.ColWidth(0) = 1000: Grid1.Text = "Qtde_emb"
    Grid1.col = 14: Grid1.ColWidth(0) = 1000: Grid1.Text = "Num_Sheda_Serial"
    Grid1.col = 15: Grid1.ColWidth(0) = 1000: Grid1.Text = "Id_Inter_Fornecedor"
    Grid1.col = 16: Grid1.ColWidth(0) = 1000: Grid1.Text = "Embarque_Ctrl"
    Grid1.col = 17: Grid1.ColWidth(0) = 1000: Grid1.Text = "Indicacao_Supl"
    Grid1.col = 18: Grid1.ColWidth(0) = 1000: Grid1.Text = "Classe_Func"
    Grid1.col = 19: Grid1.ColWidth(0) = 1000: Grid1.Text = "Dados_Transporte"
    Grid1.col = 20: Grid1.ColWidth(0) = 1000: Grid1.Text = "Qtde_Lote"
    Grid1.col = 21: Grid1.ColWidth(0) = 1000: Grid1.Text = "Num_Lote"
    Grid1.col = 22: Grid1.ColWidth(0) = 1000: Grid1.Text = "Razao_Social"
    Grid1.col = 23: Grid1.ColWidth(0) = 1000: Grid1.Text = "Cod_Barras"
    Grid1.col = 24: Grid1.ColWidth(0) = 1000: Grid1.Text = "Desenho_Chrysler"
    Grid1.col = 25: Grid1.ColWidth(0) = 1000: Grid1.Text = "Descricao_Produto"
    Grid1.col = 26: Grid1.ColWidth(0) = 1000: Grid1.Text = "Num_Desenho"
    Grid1.col = 27: Grid1.ColWidth(0) = 1000: Grid1.Text = "Destino"
    Grid1.col = 28: Grid1.ColWidth(0) = 1000: Grid1.Text = "Cod_Destino"
    Grid1.col = 29: Grid1.ColWidth(0) = 1000: Grid1.Text = "Vinculo"
    Grid1.col = 30: Grid1.ColWidth(0) = 1000: Grid1.Text = "Restricoes"
    Grid1.col = 31: Grid1.ColWidth(0) = 1000: Grid1.Text = "QR_Codes"
    Grid1.col = 32: Grid1.ColWidth(0) = 1000: Grid1.Text = "LogoMarca"
    Grid1.col = 33: Grid1.ColWidth(0) = 1000: Grid1.Text = "DUM"
    Grid1.col = 34: Grid1.ColWidth(0) = 1000: Grid1.Text = "PDF_417"
    Grid1.col = 35: Grid1.ColWidth(0) = 1000: Grid1.Text = "Incoterms"
    Grid1.col = 36: Grid1.ColWidth(0) = 1000: Grid1.Text = "Etiqueta"

Grid1.Row = 0

Grid1.HighLight = False
Grid1.ColAlignment(0) = flexAlignLeftCenter
Grid1.ColAlignment(1) = flexAlignLeftCenter
Grid1.ColAlignment(2) = flexAlignLeftCenter
Grid1.ColAlignment(3) = flexAlignLeftCenter
Grid1.ColAlignment(4) = flexAlignLeftCenter
Grid1.ColAlignment(5) = flexAlignLeftCenter
Grid1.ColAlignment(7) = flexAlignLeftCenter
Grid1.ColAlignment(8) = flexAlignLeftCenter
Grid1.ColAlignment(9) = flexAlignLeftCenter
Grid1.ColAlignment(10) = flexAlignLeftCenter
Grid1.ColAlignment(11) = flexAlignLeftCenter
Grid1.ColAlignment(12) = flexAlignLeftCenter
Grid1.ColAlignment(13) = flexAlignLeftCenter
Grid1.ColAlignment(14) = flexAlignLeftCenter
Grid1.ColAlignment(15) = flexAlignLeftCenter
Grid1.ColAlignment(16) = flexAlignLeftCenter
Grid1.ColAlignment(17) = flexAlignLeftCenter
Grid1.ColAlignment(18) = flexAlignLeftCenter
Grid1.ColAlignment(19) = flexAlignLeftCenter
Grid1.ColAlignment(20) = flexAlignLeftCenter
Grid1.ColAlignment(21) = flexAlignLeftCenter
Grid1.ColAlignment(22) = flexAlignLeftCenter
Grid1.ColAlignment(23) = flexAlignLeftCenter
Grid1.ColAlignment(24) = flexAlignLeftCenter
Grid1.ColAlignment(25) = flexAlignLeftCenter
Grid1.ColAlignment(26) = flexAlignLeftCenter
Grid1.ColAlignment(27) = flexAlignLeftCenter
Grid1.ColAlignment(28) = flexAlignLeftCenter
Grid1.ColAlignment(29) = flexAlignLeftCenter
Grid1.ColAlignment(30) = flexAlignLeftCenter
Grid1.ColAlignment(31) = flexAlignLeftCenter
Grid1.ColAlignment(32) = flexAlignLeftCenter
Grid1.ColAlignment(33) = flexAlignLeftCenter
Grid1.ColAlignment(34) = flexAlignLeftCenter
Grid1.ColAlignment(35) = flexAlignLeftCenter
Grid1.ColAlignment(36) = flexAlignLeftCenter

End Sub

Private Sub txtsequencial_Change()
If Len(Trim(txtsequencial.Text)) = 11 Then
   Call btoConfirma_Click
End If

End Sub

Private Sub txtsequencial_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
   If MsgBox("Deseja sair deste módulo?", vbQuestion + vbYesNo, "ATENÇÃO !!!") = vbNo Then
      Me.txtsequencial.Text = ""
      Me.txtsequencial.SetFocus
   Else
      Unload Me
   End If
End If

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


