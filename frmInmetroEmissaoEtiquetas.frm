VERSION 5.00
Begin VB.Form frmInmetroEmissaoEtiquetas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão Etiquetas Inmetro"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   7185
   Begin VB.TextBox TXT_QTDE_ETIQ 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   5220
      MaxLength       =   12
      TabIndex        =   26
      Text            =   "1"
      ToolTipText     =   "Digite a Qtde de Etiquetas a ser impressas.Digite números Pares."
      Top             =   5100
      Width           =   615
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Visualizar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6030
      Picture         =   "frmInmetroEmissaoEtiquetas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5100
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Caption         =   "Confirmar Peça Musashi"
      Height          =   705
      Left            =   60
      TabIndex        =   15
      Top             =   60
      Width           =   7005
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   405
         Left            =   3690
         TabIndex        =   33
         Top             =   180
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   405
         Left            =   3240
         TabIndex        =   32
         Top             =   180
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdOk 
         Height          =   405
         Left            =   5730
         Picture         =   "frmInmetroEmissaoEtiquetas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Confirmar existência código digitado"
         Top             =   210
         Width           =   435
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Height          =   405
         Left            =   6210
         Picture         =   "frmInmetroEmissaoEtiquetas.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Limpar campos"
         Top             =   210
         Width           =   435
      End
      Begin VB.CommandButton cmd_pesquisa 
         Caption         =   "..."
         Height          =   255
         Left            =   2190
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   300
         Width           =   255
      End
      Begin VB.TextBox TXT_COD_PECA_MUSASHI 
         Height          =   315
         Left            =   870
         MaxLength       =   15
         TabIndex        =   19
         Top             =   270
         Width           =   1605
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   300
         Width           =   540
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "DADOS"
      Enabled         =   0   'False
      Height          =   4245
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   7005
      Begin VB.TextBox TXT_REGISTRO_INMETRO 
         BackColor       =   &H8000000F&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   5070
         MaxLength       =   12
         TabIndex        =   28
         Top             =   630
         Width           =   1605
      End
      Begin VB.TextBox TXT_NOME_CLIENTE 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1020
         MaxLength       =   42
         TabIndex        =   24
         Top             =   270
         Width           =   5655
      End
      Begin VB.TextBox TXT_COD_PECA_CLIENTE 
         BackColor       =   &H80000004&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1020
         MaxLength       =   12
         TabIndex        =   12
         Top             =   630
         Width           =   1605
      End
      Begin VB.Frame Frame2 
         Caption         =   "Descriçao da Peça"
         Height          =   1365
         Left            =   120
         TabIndex        =   5
         Top             =   2790
         Width           =   6765
         Begin VB.TextBox TXT_DESC_PECA1 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   900
            MaxLength       =   42
            TabIndex        =   8
            Top             =   270
            Width           =   5655
         End
         Begin VB.TextBox TXT_DESC_PECA2 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   900
            MaxLength       =   42
            TabIndex        =   7
            Top             =   610
            Width           =   5655
         End
         Begin VB.TextBox TXT_DESC_PECA3 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   900
            MaxLength       =   42
            TabIndex        =   6
            Top             =   930
            Width           =   5655
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Linha 3:"
            Height          =   195
            Left            =   150
            TabIndex        =   11
            Top             =   975
            Width           =   570
         End
         Begin VB.Label lblSenha 
            AutoSize        =   -1  'True
            Caption         =   "Linha 2:"
            Height          =   195
            Left            =   150
            TabIndex        =   10
            Top             =   615
            Width           =   570
         End
         Begin VB.Label lblLogin 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Linha 1:"
            Height          =   195
            Left            =   150
            TabIndex        =   9
            Top             =   270
            Width           =   570
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Modelo da Peça"
         Height          =   1695
         Left            =   120
         TabIndex        =   1
         Top             =   1050
         Width           =   6765
         Begin VB.TextBox TXT_DESC_MOD4 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   900
            MaxLength       =   42
            TabIndex        =   30
            Top             =   1290
            Width           =   5655
         End
         Begin VB.TextBox TXT_DESC_MOD1 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   900
            MaxLength       =   42
            TabIndex        =   4
            Top             =   270
            Width           =   5655
         End
         Begin VB.TextBox TXT_DESC_MOD2 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   900
            MaxLength       =   42
            TabIndex        =   3
            Top             =   615
            Width           =   5655
         End
         Begin VB.TextBox TXT_DESC_MOD3 
            BackColor       =   &H8000000F&
            Height          =   285
            Left            =   900
            MaxLength       =   42
            TabIndex        =   2
            Top             =   945
            Width           =   5655
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Linha 4:"
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   1320
            Width           =   570
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Linha 3:"
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   975
            Width           =   570
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Linha 2:"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   615
            Width           =   570
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Linha 1:"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   270
            Width           =   570
         End
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nº Registro:"
         Height          =   195
         Left            =   4140
         TabIndex        =   29
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Peça:"
         Height          =   195
         Left            =   180
         TabIndex        =   14
         Top             =   660
         Width           =   420
      End
      Begin VB.Label lblNome 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   300
         Width           =   525
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Digite Qtde de Etiquetas para Impressão :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   90
      TabIndex        =   27
      Top             =   5130
      Width           =   5085
   End
End
Attribute VB_Name = "frmInmetroEmissaoEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável para MDIapp
Public Flag_ativo As Boolean 'Conterá true se o form ja foi ativado
Public cRec As ADODB.Recordset 'conterá os dados do registro corrente
Public Confirma_Mudanca As Boolean 'Servirá para confirmar as mudanças de alteracoes dos campos na tela
Public cTipo_Movimentacao As Integer  'Se for 1=Inclusão;2=Alteração

Private Sub cmd_pesquisa_Click()
Dim oTela As frmPesquisarInmCadPeca

Set oTela = New frmPesquisarInmCadPeca

    oTela.Show 1
    If oTela.ccodigo_pesquisa = "" Then
        Me.TXT_COD_PECA_MUSASHI.Text = ""
        Me.TXT_COD_PECA_MUSASHI.BackColor = &H80000005
        Me.TXT_COD_PECA_MUSASHI.Enabled = True
        Me.TXT_COD_PECA_MUSASHI.SetFocus
    Else
        Me.TXT_COD_PECA_MUSASHI.Text = oTela.ccodigo_pesquisa
        Call cmdOk_Click
    End If
    Unload oTela: Set oTela = Nothing

End Sub

Private Sub cmdCancelar_Click()
Call Limpar_campos

Me.TXT_QTDE_ETIQ.Enabled = False
Me.cmdImprimir.Enabled = False

Me.TXT_COD_PECA_MUSASHI.Text = ""
Me.TXT_COD_PECA_MUSASHI.Enabled = True
Me.TXT_COD_PECA_MUSASHI.BackColor = &H80000005
Me.TXT_COD_PECA_MUSASHI.SetFocus

End Sub

Private Sub cmdImprimir_Click()
If Len(Trim(Me.TXT_QTDE_ETIQ.Text)) = 0 Then
   MsgBox "Digite a quantidade de etiquetas a serem impressas."
   Me.TXT_QTDE_ETIQ.Text = ""
   Exit Sub
End If

If Val(Me.TXT_QTDE_ETIQ.Text) = 0 Then
   MsgBox "Digite a quantidade de etiquetas a serem impressas."
   Me.TXT_QTDE_ETIQ.Text = ""
   Exit Sub
End If

frmExibicaoInmetroYamaha.Show
frmExibicaoInmetroYamaha.Left = Me.Width

frmExibicaoInmetroYamaha.nQtde_Etiquetas = Me.TXT_QTDE_ETIQ.Text

frmExibicaoInmetroYamaha.LBL_COD_PECA_CLIENTE.Caption = Me.TXT_COD_PECA_CLIENTE.Text
frmExibicaoInmetroYamaha.LBL_REGISTRO_INMETRO.Caption = Me.TXT_REGISTRO_INMETRO.Text

frmExibicaoInmetroYamaha.LBL_DESC_MOD1.Caption = Me.TXT_DESC_MOD1.Text
frmExibicaoInmetroYamaha.LBL_DESC_MOD2.Caption = Me.TXT_DESC_MOD2.Text
frmExibicaoInmetroYamaha.LBL_DESC_MOD3.Caption = Me.TXT_DESC_MOD3.Text

frmExibicaoInmetroYamaha.LBL_DESC_PECA1.Caption = Me.TXT_DESC_PECA1.Text
frmExibicaoInmetroYamaha.LBL_DESC_PECA2.Caption = Me.TXT_DESC_PECA2.Text
frmExibicaoInmetroYamaha.LBL_DESC_PECA3.Caption = Me.TXT_DESC_PECA3.Text

frmExibicaoInmetroYamaha.LBL_COD_PECA_CLIENTE1.Caption = Me.TXT_COD_PECA_CLIENTE.Text
frmExibicaoInmetroYamaha.LBL_REGISTRO_INMETRO1.Caption = Me.TXT_REGISTRO_INMETRO.Text

frmExibicaoInmetroYamaha.LBL_DESC_MOD4.Caption = Me.TXT_DESC_MOD1.Text
frmExibicaoInmetroYamaha.LBL_DESC_MOD5.Caption = Me.TXT_DESC_MOD2.Text
frmExibicaoInmetroYamaha.LBL_DESC_MOD6.Caption = Me.TXT_DESC_MOD3.Text

frmExibicaoInmetroYamaha.LBL_DESC_MOD7.Caption = Me.TXT_DESC_MOD4.Text
frmExibicaoInmetroYamaha.LBL_DESC_MOD8.Caption = Me.TXT_DESC_MOD4.Text

frmExibicaoInmetroYamaha.LBL_DESC_PECA4.Caption = Me.TXT_DESC_PECA1.Text
frmExibicaoInmetroYamaha.LBL_DESC_PECA5.Caption = Me.TXT_DESC_PECA2.Text
frmExibicaoInmetroYamaha.LBL_DESC_PECA6.Caption = Me.TXT_DESC_PECA3.Text
End Sub

Private Sub cmdOk_Click()
On Error GoTo Erro
Me.cmdOk.Default = False
Me.MousePointer = vbHourglass

If Trim(Len(Me.TXT_COD_PECA_MUSASHI.Text)) = 0 Then
   MsgBox "Digite o Código da Peça da Musashi", , Me.Caption
   Me.MousePointer = vbDefault
   Me.TXT_COD_PECA_MUSASHI.SetFocus
   Exit Sub
End If

Set cRec = CCTempneInmetroCadPeca.INM_CAD_PECA_Cons_Impressao(sBancoMusashi, Trim(Me.TXT_COD_PECA_MUSASHI.Text))

Call Carregar_campos

Me.TXT_QTDE_ETIQ.Enabled = True
Me.cmdImprimir.Enabled = True

Me.TXT_COD_PECA_MUSASHI.BackColor = &H8000000F
Me.TXT_COD_PECA_MUSASHI.Enabled = False

Set cRec = Nothing
Me.MousePointer = vbDefault
Exit Sub

Erro:
Set cRec = Nothing
MsgBox Err.Description
Me.MousePointer = vbDefault
Me.TXT_COD_PECA_MUSASHI.Enabled = True
Me.TXT_COD_PECA_MUSASHI.SetFocus
Me.cmdOk.Default = True

End Sub

Private Sub Command1_Click()

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
rs.Fields("3_PARTNAME").Value = "3_PARTNAME"
rs.Fields("4_PARTNUMBER").Value = "4_PA"
rs.Fields("5_QUANTITY").Value = "5_QUAN"
rs.Fields("6_REFERENCE").Value = " "
rs.Fields("7_CONTAINER").Value = "7_CONTAINER"
rs.Fields("8_GROSSHEIGHTP").Value = "8_GROS"
rs.Fields("9_GROSSHEIGHTU").Value = "KG"
rs.Fields("10_MATERIALCODE").Value = "10_MATERIALCODE"
rs.Fields("11_PLDOCSTR").Value = "11_PLDOCSTR"
rs.Fields("12_EXPDATE").Value = "12_EXPDATE"
rs.Fields("13_SHIPMENTDATE").Value = "13_SHIP"
rs.Fields("14_MUSASHI").Value = " "
rs.Update

Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_GM_P.rpt")

CrystalReport1.ParameterFields(1).AddCurrentValue App.Path & "\DATA_MATRIX.JPG"
CrystalReport1.ParameterFields(1).DiscreteOrRangeKind = crDiscreteValue

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

Private Sub Command2_Click()
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


rs.Fields.Append "1_DESTINO", ADODB.DataTypeEnum.adVariant



'rs.Fields("1_DESTINO").Value = "GENERAL MOTORS DO BRASIL"

rs.Open


rs.AddNew

        Dim fso As New Scripting.FileSystemObject
        Dim ts As Scripting.TextStream
        Dim Imagem As String
        Dim bExiste As Boolean

'        Imagem = sDirImagemEtiq & "\EtiqCodigo.txt"
'        If Dir$(Imagem) <> "" Then Kill sDirImagemEtiq & "\EtiqCodigo.txt"
        sDirImagemEtiq = App.Path  ' "C:\Arquivos de programas\Etiquetas"
        Imagem = sDirImagemEtiq & "\DATA_MATRIX.jpg"
      

        bExiste = False
        Imagem = sDirImagemEtiq & "\DATA_MATRIX.jpg"
        Do While bExiste = False
           If Dir$(Imagem) <> "" Then bExiste = True
        Loop

        rs.Fields("1_DESTINO").Value = LoadPicture(Imagem)


rs.Update

Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_imagem.rpt")

CrystalReport1.ParameterFields(1).AddCurrentValue LoadPicture(Imagem)  '"C:\Arquivos de programas\Etiquetas\DATA_MATRIX.JPG"
CrystalReport1.ParameterFields(1).DiscreteOrRangeKind = crDiscreteValue

'CrystalReport1.Database.SetDataSource rs

'CrystalReport1.Sections(1).AddPictureObject Imagem, 1, 1 '  (LoadPicture(Imagem))

'rs.Clone

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

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Flag_ativo = False
End Sub
Private Sub Form_Activate()

If Flag_ativo = True Then
   Exit Sub
End If
Me.Top = 0
Me.Left = 0
Flag_ativo = True

Call Limpar_campos

End Sub

Private Function Limpar_campos()

Me.TXT_NOME_CLIENTE.Text = ""
Me.TXT_COD_PECA_CLIENTE.Text = ""
Me.TXT_REGISTRO_INMETRO.Text = ""

Me.TXT_DESC_MOD1.Text = ""
Me.TXT_DESC_MOD2.Text = ""
Me.TXT_DESC_MOD3.Text = ""
Me.TXT_DESC_MOD4.Text = ""

Me.TXT_DESC_PECA1.Text = ""
Me.TXT_DESC_PECA2.Text = ""
Me.TXT_DESC_PECA3.Text = ""

End Function
Private Function Carregar_campos()
Me.TXT_NOME_CLIENTE.Text = cRec!NOME_CLIENTE
Me.TXT_COD_PECA_CLIENTE.Text = cRec!COD_PECA_CLIENTE
Me.TXT_REGISTRO_INMETRO.Text = cRec!REGISTRO_INMETRO

Me.TXT_DESC_MOD1.Text = cRec!DESC_MOD1
Me.TXT_DESC_MOD2.Text = IIf(IsNull(cRec!DESC_MOD2), "", cRec!DESC_MOD2)
Me.TXT_DESC_MOD3.Text = IIf(IsNull(cRec!DESC_MOD3), "", cRec!DESC_MOD3)
Me.TXT_DESC_MOD4.Text = IIf(IsNull(cRec!DESC_MOD4), "", cRec!DESC_MOD4)

Me.TXT_DESC_PECA1.Text = IIf(IsNull(cRec!DESC_PECA1), "", cRec!DESC_PECA1)
Me.TXT_DESC_PECA2.Text = IIf(IsNull(cRec!DESC_PECA2), "", cRec!DESC_PECA2)
Me.TXT_DESC_PECA3.Text = "Data de Fabricaçao: " & Format(Now(), "MM/yyyy") '  IIf(IsNull(cRec!DESC_PECA3), "", cRec!DESC_PECA3)

End Function


Private Sub TXT_COD_PECA_MUSASHI_GotFocus()
Me.cmdOk.Default = True
End Sub
