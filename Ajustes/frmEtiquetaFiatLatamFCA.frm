VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEtiquetaFiatLatamFCA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Etiquetas Fiat Latam"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9495
   Begin VB.TextBox txt_Transportadora 
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
      Height          =   315
      Left            =   1260
      MaxLength       =   11
      TabIndex        =   22
      ToolTipText     =   "Digite a Transportadora"
      Top             =   5160
      Width           =   1365
   End
   Begin VB.CommandButton cmd_Visualizar 
      Caption         =   "Visualizar"
      Enabled         =   0   'False
      Height          =   735
      Left            =   6390
      TabIndex        =   14
      Top             =   5040
      Width           =   975
   End
   Begin VB.Frame frm_filtro 
      Caption         =   "Filtro da etiqueta"
      Height          =   1065
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   9285
      Begin VB.Frame Frame1 
         Height          =   765
         Left            =   90
         TabIndex        =   10
         Top             =   180
         Width           =   3345
         Begin VB.OptionButton Opt_etiqueta 
            Caption         =   "Pelo Nº da Etiqueta"
            Height          =   255
            Left            =   1500
            TabIndex        =   12
            Top             =   270
            Width           =   1755
         End
         Begin VB.OptionButton Opt_Pallet 
            Caption         =   "Pelo Pallet"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   300
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.CommandButton btoConfirma 
         Height          =   495
         Left            =   8040
         Picture         =   "frmEtiquetaFiatLatamFCA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Confirma dados do filtro"
         Top             =   270
         Width           =   555
      End
      Begin VB.CommandButton cmd_limpar 
         Height          =   495
         Left            =   8610
         Picture         =   "frmEtiquetaFiatLatamFCA.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Limpar tela para nova etiqueta"
         Top             =   270
         Width           =   555
      End
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Digitação:"
         Height          =   195
         Left            =   3570
         TabIndex        =   13
         Top             =   510
         Width           =   720
      End
   End
   Begin VB.CommandButton cmd_Impressao 
      Enabled         =   0   'False
      Height          =   735
      Left            =   7410
      Picture         =   "frmEtiquetaFiatLatamFCA.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Imprimir"
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton cmdfechar 
      Height          =   735
      Left            =   8430
      Picture         =   "frmEtiquetaFiatLatamFCA.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Fecha a tela de etiqueta"
      Top             =   5040
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Height          =   3705
      Left            =   90
      TabIndex        =   2
      Top             =   1260
      Width           =   9285
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   3315
         Left            =   120
         TabIndex        =   3
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
         FormatString    =   $"frmEtiquetaFiatLatamFCA.frx":0E98
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
   Begin VB.CommandButton cmd_Impressao_Master 
      Enabled         =   0   'False
      Height          =   735
      Left            =   3060
      Picture         =   "frmEtiquetaFiatLatamFCA.frx":0EA1
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir"
      Top             =   6330
      Width           =   975
   End
   Begin VB.TextBox TXT_EMBALAGEM 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1470
      MaxLength       =   20
      TabIndex        =   0
      ToolTipText     =   "Digie a sequencial da etiqueta"
      Top             =   6570
      Width           =   1515
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Transportadora:"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   5190
      Width           =   1125
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
      TabIndex        =   21
      Top             =   6270
      Visible         =   0   'False
      Width           =   1035
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
      TabIndex        =   20
      Top             =   6270
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
      TabIndex        =   19
      Top             =   6510
      Visible         =   0   'False
      Width           =   570
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
      TabIndex        =   18
      Top             =   6525
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
      TabIndex        =   16
      Top             =   6780
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "NF.Embalagem:"
      Height          =   195
      Left            =   300
      TabIndex        =   15
      Top             =   6630
      Width           =   1125
   End
End
Attribute VB_Name = "frmEtiquetaFiatLatamFCA"
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
Private nSequencia_Escolhida As Double
Private Declare Sub Sleep Lib "KERNEL32" _
        (ByVal dwMilliseconds As Long)
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
Dim nx As Double
Dim x As Printer

On Error GoTo Erro

nx = 0
For Each x In Printers
   If InStr(1, UCase(x.DeviceName), "ETIQUETA FABRICA") > 0 Then
      Set Printer = x
      nx = 1
      Exit For
   End If
Next

If nx = 0 Then
   MsgBox "impressora da Etiqueta não encontrada, chame o responsável. o nome da impressora será - 'ETIQUETA FABRICA'"
   Close
End If
nx = 0

If Grid1.Rows > 0 Then
   For nx = 1 To Grid1.Rows
       nSequencia_Escolhida = nx
       Call cmd_visualizar_Click
       frmEtiquetaFiatLatamFCAEtiq.PrintForm
       If nx = 4 Then GoTo saida
   Next
End If

saida:

Printer.EndDoc
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

Dim nx As Double
Dim x As Printer
Dim rs As ADODB.Recordset
Dim sTexto As String
Dim nqtde As Double

On Error GoTo Erro

If Len(Trim(txt_Transportadora.Text)) = 0 Then
   MsgBox "Digite a transportadora para poder emitir a etiqueta!"
   Me.txt_Transportadora.SetFocus
   Exit Sub
End If


Set rs = New ADODB.Recordset

rs.Fields.Append "01_Peso_Bruto", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "02_ODM", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "03_Data_Producao", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "04_Data_validade", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "05_Data_Expedicao", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "06_Cod_Emb", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "07_Quantidade", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "08_DOCA", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "09_Ponto_Entrega", ADODB.DataTypeEnum.adChar, 20
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
rs.Fields.Append "ID_ETIQUETA", ADODB.DataTypeEnum.adChar, 20

rs.Open

Me.MousePointer = vbHourglass

nx = 0

If Grid1.Rows > 0 Then
   nqtde = 0
   rs.AddNew

   Grid1.col = 0: rs.Fields("01_Peso_Bruto").Value = Trim(Grid1.Text)
   Grid1.col = 1: rs.Fields("02_ODM").Value = "0000" ' Trim(Grid1.Text)
   Grid1.col = 2: rs.Fields("03_Data_Producao").Value = Trim(Grid1.Text)
   Grid1.col = 3: rs.Fields("04_Data_validade").Value = Trim(Grid1.Text)
   Grid1.col = 4
   If Format(Trim(Grid1.Text), "DD/MM/YYYY") = "01/01/1900" Then
      rs.Fields("05_Data_Expedicao").Value = " "
   Else
      rs.Fields("05_Data_Expedicao").Value = Trim(Grid1.Text)
   End If
   Grid1.col = 5: rs.Fields("06_Cod_Emb").Value = "4322" 'Trim(Grid1.Text)
   Grid1.col = 6: rs.Fields("07_Quantidade").Value = Trim(Grid1.Text)
   Grid1.col = 7: rs.Fields("08_DOCA").Value = Trim(Grid1.Text)
   Grid1.col = 8: rs.Fields("09_Ponto_Entrega").Value = Trim(Grid1.Text)
   Grid1.col = 9: rs.Fields("10_Control_Log_Qual").Value = "OK" 'Trim(Grid1.Text)
   Grid1.col = 10
   If Format(Trim(Grid1.Text), "DD/MM/YYYY") = "01/01/1900" Then
      rs.Fields("11_Cod_Fornecedor").Value = " "
   Else
      rs.Fields("11_Cod_Fornecedor").Value = Trim(Grid1.Text)
   End If
   Grid1.col = 11: rs.Fields("12_Num_Doc_Fiscal").Value = Trim(Grid1.Text)
   Grid1.col = 12: rs.Fields("13_Lote_Sob_Desv").Value = Trim(Grid1.Text)
   Grid1.col = 13: rs.Fields("14_Qtde_emb").Value = "0001" 'Trim(Grid1.Text)
   Grid1.col = 14: rs.Fields("15_Num_Sheda_Serial").Value = Trim(Grid1.Text)
   Grid1.col = 15: rs.Fields("16_Id_Inter_Fornecedor").Value = Trim(Grid1.Text)
   Grid1.col = 16: rs.Fields("17_Embarque_Ctrl").Value = Trim(Grid1.Text)
   Grid1.col = 17: rs.Fields("18_Indicacao_Supl").Value = Trim(Grid1.Text)
   Grid1.col = 18: rs.Fields("19_Classe_Func").Value = Trim(Grid1.Text)
   Grid1.col = 19: rs.Fields("20_Dados_Transporte").Value = Me.txt_Transportadora.Text 'Trim(Grid1.Text)
   Grid1.col = 20: rs.Fields("21_Qtde_Lote").Value = Trim(Grid1.Text)
   Grid1.col = 21: rs.Fields("22_Num_Lote").Value = Trim(Grid1.Text)
   Grid1.col = 22: rs.Fields("23_Razao_Social").Value = Trim(Grid1.Text)
   Rem ver aqui marcos
Rem construir o codigo de barras (24_Cod_Barras)
   Grid1.col = 26: sTexto = Trim(Grid1.Text) '"27_Num_Desenho"
   sTexto = sTexto & "0013093" ' codigo fornecedor da Musashi no cliente
   Grid1.col = 6: sTexto = sTexto & Format(Trim(Grid1.Text), "00000") '"07_Quantidade"
   Grid1.col = 5: sTexto = sTexto & Format(Trim(Grid1.Text), "0000") ' "06_Cod_Emb"
   rs.Fields("24_Cod_Barras").Value = "*" & sTexto & "*"
   
   Grid1.col = 24: rs.Fields("25_Desenho_Chrysler").Value = "9999999999" 'Trim(Grid1.Text)
   Grid1.col = 25: rs.Fields("26_Descricao_Produto").Value = Mid$(Trim(Grid1.Text), 1, 40)
   Grid1.col = 26: rs.Fields("27_Num_Desenho").Value = Trim(Grid1.Text)
   Grid1.col = 27: rs.Fields("28_Destino").Value = Mid$(Trim(Grid1.Text), 1, 20)
   Grid1.col = 28: rs.Fields("29_Cod_Destino").Value = "FPTBE" 'Trim(Grid1.Text)
   Grid1.col = 29: rs.Fields("30_Vinculo").Value = "196" 'Trim(Grid1.Text)
   Grid1.col = 30: rs.Fields("31_Restricoes").Value = Trim(Grid1.Text)
   
   Rem Formação dos dados referente a etiqueta QR-CODE
'    1(11)       2(7)    3(10)      4(8)     5(3)6(5)  7(4)8(4) 9(3)10(9)     11(4)12(8)    13(10)      14(1)
'   "00055267989 0013093 0000000001 09112016 G09 00002 CPG 0001 000 000330999 0000 00000000 00160920160 0
   sTexto = ""
   Grid1.col = 26: sTexto = sTexto & Format(Trim(Grid1.Text), "00000000000")  '"27_Num_Desenho" 1
   sTexto = sTexto & "0013093"  ' codigo fornecedor da Musashi no cliente 2
   sTexto = sTexto & "0000000001"  ' id interno fornecedor nao se repete durante um ano 3 - VER COM MAURO
   Grid1.col = 2: sTexto = sTexto & Format(Trim(Grid1.Text), "DDMMYYYY") '"03_Data_Producao"
   Grid1.col = 8: sTexto = sTexto & Trim(Grid1.Text) '"09_Ponto_Entrega"
   Grid1.col = 6: sTexto = sTexto & Format(Trim(Grid1.Text), "00000") '"07_Quantidade" 6
   Grid1.col = 5: sTexto = sTexto & Format(Trim(Grid1.Text), "0000") ' "06_Cod_Emb" 7 - VER COM MAURO
   sTexto = sTexto & "0001"  ' Qtde de embalagem enviada 8
   sTexto = sTexto & "000"  ' cod de destino 9 - VER COM MAURO
   Grid1.col = 11: sTexto = sTexto & Format(Trim(Grid1.Text), "000000000") '"12_Num_Doc_Fiscal" 10
   
   Grid1.col = 1:
   sTexto = sTexto & Format(Mid$(Trim(Grid1.Text), 1, 4), "0000") '"02_ODM" 11
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
   Grid1.col = 6: sTexto = sTexto & "Q" & Format(Trim(Grid1.Text), "000000") & "<0x1d>" '"07_Quantidade"
   sTexto = sTexto & "V" & "0013093" & "<0x1d>" ' codigo fornecedor da Musashi no cliente
   sTexto = sTexto & "3S" & "0000000000" & "<0x1d>" ' Numero serial
   Grid1.col = 26: sTexto = sTexto & "1P" & Format(Trim(Grid1.Text), "00000000000") & "<0x1e><0x04>" '"27_Num_Desenho"
   Grid1.col = 34: rs.Fields("35_PDF_417").Value = sTexto
'********************************************************************************************************
'********************************************************************************************************

   Grid1.col = 35: rs.Fields("36_Incoterms").Value = Trim(Grid1.Text)
   Grid1.col = 36: rs.Fields("ID_ETIQUETA").Value = Trim(Grid1.Text)
  
   rs.Update
   
   rs.MoveFirst
   
   frmEtiquetaFiatLatamFCAEtiq.Show
   Dim sNumDesenho As String
   
   With frmEtiquetaFiatLatamFCAEtiq
       .lbl_01_Peso_Bruto.Caption = Trim(rs.Fields("01_Peso_Bruto").Value)
       .lbl_02_ODM.Caption = Trim(rs.Fields("02_ODM").Value)
       .lbl_03_Data_Producao.Caption = Trim(rs.Fields("03_Data_Producao").Value)
       .lbl_04_Data_validade.Caption = Trim(rs.Fields("04_Data_validade").Value)
       
       .lbl_05_Data_Expedicao.Caption = Trim(rs.Fields("05_Data_Expedicao").Value)
       .lbl_06_Cod_Emb.Caption = Trim(rs.Fields("06_Cod_Emb").Value)
       .lbl_07_Quantidade.Caption = Trim(rs.Fields("07_Quantidade").Value)
       .lbl_08_DOCA.Caption = Trim(rs.Fields("08_DOCA").Value)
       .lbl_09_Ponto_Entrega.Caption = Trim(rs.Fields("09_Ponto_Entrega").Value)
       
       .lbl_10_Control_Log_Qual.Caption = Trim(rs.Fields("10_Control_Log_Qual").Value)
       .lbl_11_Cod_Fornecedor.Caption = Trim(rs.Fields("11_Cod_Fornecedor").Value)
       .lbl_12_Num_Doc_Fiscal.Caption = Trim(rs.Fields("12_Num_Doc_Fiscal").Value)
       .lbl_13_Lote_Sob_Desv.Caption = Trim(rs.Fields("13_Lote_Sob_Desv").Value)
       .lbl_14_Qtde_emb.Caption = Trim(rs.Fields("14_Qtde_emb").Value)

       .lbl_15_Num_Sheda_Serial.Caption = Trim(rs.Fields("15_Num_Sheda_Serial").Value)
       .lbl_16_Id_Inter_Fornecedor.Caption = Trim(rs.Fields("16_Id_Inter_Fornecedor").Value)
       .lbl_17_Embarque_Ctrl.Caption = Trim(rs.Fields("17_Embarque_Ctrl").Value)
       .lbl_18_Indicacao_Supl.Caption = Trim(rs.Fields("18_Indicacao_Supl").Value)
       .lbl_19_Classe_Func.Caption = Trim(rs.Fields("19_Classe_Func").Value)

       .lbl_20_Dados_Transporte.Caption = Trim(rs.Fields("20_Dados_Transporte").Value)
       .lbl_21_Qtde_Lote.Caption = Trim(rs.Fields("21_Qtde_Lote").Value)
       .lbl_22_Num_Lote.Caption = Trim(rs.Fields("22_Num_Lote").Value)
       .lbl_23_Razao_Social.Caption = Trim(rs.Fields("23_Razao_Social").Value)
'       .lbl_24_Codigo_Barra_A.Caption = Trim(rs.Fields("24_Cod_Barras").Value)
'       .lbl_24_Codigo_Barra_B.Caption = Trim(rs.Fields("24_Cod_Barras").Value)

       .lbl_25_Desenho_Chrysler.Caption = Trim(rs.Fields("25_Desenho_Chrysler").Value)
       .lbl_26_Descricao_Produto.Caption = Trim(rs.Fields("26_Descricao_Produto").Value)
       .lbl_27A_Num_Desenho.Caption = Mid$(Format(rs.Fields("27_Num_Desenho").Value, "00000000000"), 1, 6)
       .lbl_27B_Num_Desenho.Caption = Mid$(Format(rs.Fields("27_Num_Desenho").Value, "00000000000"), 7, 11)
       sNumDesenho = .lbl_27A_Num_Desenho.Caption & .lbl_27B_Num_Desenho.Caption
       .lbl_28_Destino.Caption = Trim(rs.Fields("28_Destino").Value)
       .lbl_29_Cod_Destino.Caption = Trim(rs.Fields("29_Cod_Destino").Value)
       
       .lbl_30_Vinculo.Caption = Trim(rs.Fields("30_Vinculo").Value)
       .lbl_31_Restricoes.Caption = Trim(rs.Fields("31_Restricoes").Value)
       
       '.lbl_32_QR_Codes.Caption = Trim(rs.Fields("32_QR_Codes").Value)
       '.lbl_33_LogoMarca.Caption = Trim(rs.Fields("33_LogoMarca").Value)
       .lbl_34_DUM.Caption = Trim(rs.Fields("34_DUM").Value)
       '.lbl_35_PDF_417.Caption = Trim(rs.Fields("35_PDF_417").Value)
       .lbl_36_Incoterms.Caption = Trim(rs.Fields("36_Incoterms").Value)

Rem Codigos especiais(24)- Código de Barras
       .lbl_24_Codigo_Barra_A.Caption = sNumDesenho & _
                                        Format(rs.Fields("11_Cod_Fornecedor").Value, "0000000") & _
                                        Format(rs.Fields("07_Quantidade").Value, "000000") & _
                                        Mid(Trim(rs.Fields("06_Cod_Emb").Value), 1, 4)
'       .lbl_24_Codigo_Barra_A.Caption = "*" & _
'                                        sNumDesenho & _
'                                        Format(rs.Fields("11_Cod_Fornecedor").Value, "0000000") & _
'                                        Format(rs.Fields("07_Quantidade").Value, "000000") & _
'                                        Mid(Trim(rs.Fields("06_Cod_Emb").Value), 1, 4) & _
'                                        "*"
       .lbl_24_Codigo_Barra_B.Caption = .lbl_24_Codigo_Barra_A.Caption
       .lbl_24_Codigo_Barra_C.Caption = .lbl_24_Codigo_Barra_A.Caption
'       .lbl_24_Codigo_Barra_D.Caption = .lbl_24_Codigo_Barra_A.Caption


Rem Codigos especiais(32)QR Code.
'I - Desenho (11 dígitos). Desenho FCA correspondente ao material (Ex: 00518975620).
'II - Fornecedor (7 dígitos) - Fornecedor de origem do material.
'III - Índice “ID interno do fornecedor” (10 dígitos) - Serial gerado pelo sistema do fornecedor (Nunca repete
'durante 1 ano).
'IV - Data de produção (8 Dígitos) – DDMMAAAA.
'V - Ponto de entrega do material (3 dígitos) - Galpão de entrega. (Ex: 89 = CDC).
'VI - Quantidade (5 dígitos).
'VII - Embalagem (4 dígitos) - Código da embalagem cadastrado na planta de destino (Ex: 403 = Caçamba
'de 2000 kg).
'VIII - Quantidade de embalagem (4 dígitos) - Quantidade de embalagens enviada.
'IX – Código de destino\Planta (3 dígitos) - Planta de destino do material (Ex: 161 = FIASA).
'X - Nota fiscal (9 dígitos). Nota fiscal do material.
'XI - ODM – (4 dígitos).
'XII - Data de validade (8 Dígitos) – DDMMAAAA.
'XIII - Numero do lote – 10 Dígitos.
'XIV - Código do tipo de etiqueta (1 digito) - Especifica o tipo de etiqueta. - Inserir o código “1” para etiqueta LATAM.

        Dim fso As New Scripting.FileSystemObject
        Dim ts As Scripting.TextStream
        Dim Imagem As String
        
'        sDirImagemEtiq = "X:\Transfere"
'        sDirImagemEtiq = "C:\Users\consu\Documents\CompartilhadaXP\Transfere"
        
        'Pega o caminho para a imagem com o Nome do TextBox'
        Imagem = sDirImagemEtiq & "\QR_CODE.jpg"
        If Dir$(Imagem) <> "" Then Kill sDirImagemEtiq & "\QR_CODE.jpg"

        Imagem = sDirImagemEtiq & "\EtiqCodigo.txt"
        If Dir$(Imagem) <> "" Then Kill sDirImagemEtiq & "\EtiqCodigo.txt"

        .sTextoQrCode = "01" & _
                        Format(rs.Fields("27_Num_Desenho").Value, "00000000000") & _
                        Format(rs.Fields("11_Cod_Fornecedor").Value, "0000000") & _
                        Format(rs.Fields("ID_ETIQUETA").Value, "0000000000") & _
                        Replace(rs.Fields("05_Data_Expedicao").Value, "/", "") & _
                        Format(rs.Fields("09_Ponto_Entrega").Value, "000") & _
                        Format(rs.Fields("07_Quantidade").Value, "00000") & _
                        Trim(rs.Fields("06_Cod_Emb").Value) & _
                        Format(rs.Fields("14_Qtde_emb").Value, "0000") & _
                        rs.Fields("29_Cod_Destino").Value & _
                        Format(rs.Fields("12_Num_Doc_Fiscal").Value, "000000000") & _
                        Format(rs.Fields("02_ODM").Value, "0000") & _
                        "          " & _
                        Format(rs.Fields("22_Num_Lote").Value, "0000000000") & _
                        "1"


        Set ts = fso.OpenTextFile(sDirImagemEtiq & "\EtiqCodigo.txt", ForWriting, True)  'abre um arquivo para escrita , se não existir cria

        ts.Write .sTextoQrCode
        ts.Close
        Set ts = Nothing

        Shell sDirImagemEtiq & "\JCodFactory.exe"

        Dim bExiste As Boolean
        bExiste = False
        Imagem = sDirImagemEtiq & "\QR_CODE.jpg"
        Do While bExiste = False
           If Dir$(Imagem) <> "" Then bExiste = True
        Loop
        Sleep 500
        
        
        .IMG_32_QRCODE.Picture = LoadPicture(Imagem)
        
        Imagem = sDirImagemEtiq & "\EtiqCodigo.txt"
        If Dir$(Imagem) <> "" Then Kill sDirImagemEtiq & "\EtiqCodigo.txt"
        Imagem = sDirImagemEtiq & "\PDF_417.jpg"
        If Dir$(Imagem) <> "" Then Kill sDirImagemEtiq & "\PDF_417.jpg"

        .sTextoPDF417 = "02" & _
                        "[)>" + Chr(30) + "06" + Chr(29) & _
                        "P" & Trim(rs.Fields("11_Cod_Fornecedor").Value) & Chr(29) & _
                        "Q" & "1" & Chr(29) & _
                        "V" & "       " & Chr(29) & _
                        "3S" & Format(rs.Fields("ID_ETIQUETA").Value, "0000000000") & Chr(29) & _
                        "1P" & "99999999999" & Chr(29) & _
                        "1B" & "   " & "" & Chr(30) & _
                        "" & Chr(4)

        Set ts = fso.OpenTextFile(sDirImagemEtiq & "\EtiqCodigo.txt", ForWriting, True)  'abre um arquivo para escrita , se não existir cria

        ts.Write .sTextoPDF417
        ts.Close
        Set ts = Nothing

        Shell sDirImagemEtiq & "\JCodFactory.exe"

        bExiste = False
        Do While bExiste = False
           If Dir$(Imagem) <> "" Then bExiste = True
        Loop
        Sleep 500

'        MsgBox "pdf11"
        
        .IMG_32_PDF417.Picture = LoadPicture(Imagem)

'***********************************************************************************

        Set ts = fso.OpenTextFile(sDirImagemEtiq & "\EtiqCodigo.txt", ForWriting, True)  'abre um arquivo para escrita , se não existir cria

        ts.Write .sTextoQrCode
        ts.Close
        Set ts = Nothing

        Shell sDirImagemEtiq & "\JCodFactory.exe"

        Dim bExiste As Boolean
        bExiste = False
        Imagem = sDirImagemEtiq & "\QR_CODE.jpg"
        Do While bExiste = False
           If Dir$(Imagem) <> "" Then bExiste = True
        Loop
        Sleep 500
        
        
        .IMG_32_QRCODE.Picture = LoadPicture(Imagem)
        
        Imagem = sDirImagemEtiq & "\EtiqCodigo.txt"
        If Dir$(Imagem) <> "" Then Kill sDirImagemEtiq & "\EtiqCodigo.txt"
        Imagem = sDirImagemEtiq & "\PDF_417.jpg"
        If Dir$(Imagem) <> "" Then Kill sDirImagemEtiq & "\PDF_417.jpg"

        .sTextoPDF417 = "02" & _
                        "[)>" + Chr(30) + "06" + Chr(29) & _
                        "P" & Trim(rs.Fields("11_Cod_Fornecedor").Value) & Chr(29) & _
                        "Q" & "1" & Chr(29) & _
                        "V" & "       " & Chr(29) & _
                        "3S" & Format(rs.Fields("ID_ETIQUETA").Value, "0000000000") & Chr(29) & _
                        "1P" & "99999999999" & Chr(29) & _
                        "1B" & "   " & "" & Chr(30) & _
                        "" & Chr(4)

        Set ts = fso.OpenTextFile(sDirImagemEtiq & "\EtiqCodigo.txt", ForWriting, True)  'abre um arquivo para escrita , se não existir cria

        ts.Write .sTextoPDF417
        ts.Close
        Set ts = Nothing

        Shell sDirImagemEtiq & "\JCodFactory.exe"

        bExiste = False
        Do While bExiste = False
           If Dir$(Imagem) <> "" Then bExiste = True
        Loop
        Sleep 500

'        MsgBox "pdf11"
        
        .IMG_32_PDF417.Picture = LoadPicture(Imagem)

 End With
 
frmEtiquetaFiatLatamFCAEtiq.Left = Me.Width
frmEtiquetaFiatLatamFCAEtiq.Show
DoEvents

'   Printer.Orientation = 2
'   frmExibicaoHondaNova.PrintForm
'   Printer.EndDoc

End If

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




