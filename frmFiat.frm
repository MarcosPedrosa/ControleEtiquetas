VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFiat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Infomações da etiqueta - Modelo FIAT"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5190
   Icon            =   "frmFiat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5190
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
      Left            =   180
      TabIndex        =   21
      Text            =   "Combo1"
      Top             =   5400
      Width           =   3735
   End
   Begin VB.TextBox txtEmbalagem 
      Height          =   285
      Left            =   240
      MaxLength       =   5
      TabIndex        =   19
      Top             =   4440
      Width           =   975
   End
   Begin MSMask.MaskEdBox mskDum 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      Height          =   300
      Left            =   2760
      TabIndex        =   18
      Top             =   3840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   529
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtLoteDesvio 
      Height          =   285
      Left            =   1560
      TabIndex        =   17
      Top             =   3840
      Width           =   1095
   End
   Begin VB.TextBox txtEmbarque 
      Height          =   285
      Left            =   240
      TabIndex        =   16
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox txtIndicacao 
      Height          =   285
      Left            =   2520
      TabIndex        =   15
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtVinculo 
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox txtClasseFuncional 
      Height          =   285
      Left            =   240
      TabIndex        =   13
      Top             =   3120
      Width           =   735
   End
   Begin VB.TextBox txtQtdEmbalagem 
      Height          =   300
      Left            =   2520
      TabIndex        =   12
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtQtdEtiquetas 
      Height          =   300
      Left            =   3360
      MaxLength       =   3
      TabIndex        =   20
      Text            =   "1"
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   855
      Left            =   4080
      Picture         =   "frmFiat.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Nova Etiqueta"
      Height          =   855
      Left            =   4080
      Picture         =   "frmFiat.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   855
      Left            =   4080
      Picture         =   "frmFiat.frx":0EEE
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtNumLote 
      Height          =   300
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   735
   End
   Begin MSMask.MaskEdBox mskDataFabricacao 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   8
      Top             =   1680
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   529
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtCodFiat 
      Height          =   300
      Left            =   2880
      MaxLength       =   3
      TabIndex        =   9
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtQtdLote 
      Height          =   300
      Left            =   1200
      TabIndex        =   11
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtNumFiat 
      Height          =   300
      Left            =   240
      MaxLength       =   11
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox txtDescrPeca 
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox txtDocFiscal 
      Height          =   300
      Left            =   2040
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin MSMask.MaskEdBox maskDataExpedicao 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   529
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   10
      Mask            =   "99/99/9999"
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Impressora"
      Height          =   195
      Left            =   180
      TabIndex        =   38
      Top             =   5130
      Width           =   765
   End
   Begin VB.Label lblEmbalagem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Embalagem"
      Height          =   195
      Left            =   240
      TabIndex        =   37
      Top             =   4200
      Width           =   825
   End
   Begin VB.Label lblDum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DUM"
      Height          =   195
      Left            =   3120
      TabIndex        =   36
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label lblLoteDesvio 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lote sob Desvio"
      Height          =   195
      Left            =   1560
      TabIndex        =   35
      Top             =   3600
      Width           =   1155
   End
   Begin VB.Label lblEmbarque 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Emb. controlado"
      Height          =   195
      Left            =   240
      TabIndex        =   34
      Top             =   3600
      Width           =   1155
   End
   Begin VB.Label lblIndicacao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ind. Suplementar"
      Height          =   195
      Left            =   2520
      TabIndex        =   33
      Top             =   2880
      Width           =   1200
   End
   Begin VB.Label lblVinculo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vínculo"
      Height          =   195
      Left            =   1320
      TabIndex        =   32
      Top             =   2880
      Width           =   555
   End
   Begin VB.Label lblClasseFuncional 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Classe Func."
      Height          =   195
      Left            =   240
      TabIndex        =   31
      Top             =   2880
      Width           =   915
   End
   Begin VB.Label lblQtdEmbalagem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qtd. da Embalag."
      Height          =   195
      Left            =   2520
      TabIndex        =   30
      Top             =   2160
      Width           =   1230
   End
   Begin VB.Label lblQtdEtiquetas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade de etiquetas a imprimir"
      Height          =   195
      Left            =   840
      TabIndex        =   29
      Top             =   4860
      Width           =   2430
   End
   Begin VB.Label lblLote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do lote"
      Height          =   195
      Left            =   240
      TabIndex        =   28
      Top             =   2160
      Width           =   705
   End
   Begin VB.Label lblDataFabricacao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data de prod."
      Height          =   195
      Left            =   1800
      TabIndex        =   27
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblCodEmbalagem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód. Embalag."
      Height          =   195
      Left            =   2880
      TabIndex        =   26
      Top             =   1440
      Width           =   1035
   End
   Begin VB.Label lblQtdLote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qtd. do Lote"
      Height          =   195
      Left            =   1200
      TabIndex        =   22
      Top             =   2160
      Width           =   885
   End
   Begin VB.Label lblNumeroPeca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do desenho FIAT"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label lblPeca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição da peça"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1350
   End
   Begin VB.Label lblDocFiscal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Doc. Fiscal(BAM)"
      Height          =   195
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   1680
   End
   Begin VB.Label lblDataExpedicao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data de expedição"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1350
   End
End
Attribute VB_Name = "frmFiat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer    'contador de etiquetas

Private Sub cmdImprimir_Click()
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
    
    On Error GoTo Erro
    
    If Not IsDate(maskDataExpedicao.Text) Then
        MsgBox "Informe uma data válida", vbInformation + vbOKOnly, "ATENÇÃO!!!"
        maskDataExpedicao.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(mskDataFabricacao.Text) Then
        MsgBox "Informe uma data válida", vbInformation + vbOKOnly, "ATENÇÃO!!!"
        mskDataFabricacao.SetFocus
        Exit Sub
    End If
    
    If Trim(txtDocFiscal.Text) = Empty Then
        MsgBox "Informe o número do documento fiscal", vbInformation + vbOKOnly, "ATENÇÃO!!!"
        txtDocFiscal.SetFocus
        Exit Sub
    End If
    
    If Trim(txtDescrPeca.Text) = Empty Then
        MsgBox "Informe a descrição da peça", vbInformation + vbOKOnly, "ATENÇÃO!!!"
        txtDescrPeca.SetFocus
        Exit Sub
    End If
    
    If Trim(txtNumFiat.Text) = Empty Then
        MsgBox "Informe o número do desenho FIAT", vbOKOnly + vbInformation, "ATENÇÃO!!!"
        txtNumFiat.SetFocus
        Exit Sub
    End If
    
'    If Trim(txtQtd.Text) = Empty Then
'        MsgBox "Informe a quantidade de peças", vbOKOnly + vbInformation, "ATENÇÃO!!!"
'        txtQtd.SetFocus
'        Exit Sub
'    End If
    
    If Trim(txtCodFiat.Text) = Empty Then
        MsgBox "Informe o código da embalagem FIAT", vbOKOnly + vbInformation, "ATENÇÃO!!!"
        txtCodFiat.SetFocus
        Exit Sub
    End If

    If Trim(txtNumLote.Text) = Empty Then
        MsgBox "Informe o número do lote", vbOKOnly + vbInformation, "ATENÇÃO!!!"
        txtNumLote.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtQtdEtiquetas.Text) Then
        MsgBox "Digite a quantidade de etiquetas a serem impressas.", vbExclamation, objApplication.tituloPrograma
        txtQtdEtiquetas.SetFocus
        Exit Sub
    End If
    
    'Printer.Orientation = 2
    'vbPRORPortrait = 1
    'vbPRORLandscape = 2
    'Printer.Orientation = vbPRORLandscape 'PASSOU A IMPRIMIR ERRADO!
    
    Printer.Orientation = vbPRORPortrait
    
    Printer.Copies = CInt(txtQtdEtiquetas.Text)
    frmExibicao2.PrintForm
    Printer.Orientation = 2: Printer.EndDoc
    
    MsgBox "Impressão concluída com sucesso! O aplicativo será encerrado!", vbOKOnly + vbInformation, "Tarefa Concluída"
    MDIEtiquetas.forcaSaida = True
    Unload MDIEtiquetas
    
    Exit Sub
Erro:
    MsgBox "Ocorreu um erro na impressão da etiqueta." & vbNewLine & "Err. N: " & Err.Number & vbNewLine & "Desc: " & Err.Description, vbCritical, objApplication.tituloPrograma
    
End Sub

Private Sub cmdNovo_Click()
    frmFiat.maskDataExpedicao.Text = "__/__/____"
    frmFiat.txtDocFiscal.Text = ""
    frmFiat.txtDescrPeca.Text = ""
    frmFiat.txtCodFiat.Text = ""
    frmFiat.txtNumFiat.Text = ""
    frmFiat.txtNumLote.Text = ""
    frmFiat.txtQtdEmbalagem.Text = ""
    frmFiat.mskDataFabricacao.Text = "__/__/____"
    frmFiat.txtQtdEtiquetas.Text = "1"
    frmFiat.txtQtdLote.Text = ""
    frmFiat.txtClasseFuncional.Text = ""
    frmFiat.txtVinculo.Text = ""
    frmFiat.txtIndicacao.Text = ""
    frmFiat.txtEmbarque.Text = ""
    frmFiat.txtLoteDesvio.Text = ""
    frmFiat.mskDum.Text = "__/__/____"
    frmFiat.maskDataExpedicao.SetFocus
End Sub


Private Sub cmdSair_Click()
    Unload Me
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
'Me.Height = 5385
'Me.Width = 5325
frmExibicao2.nTamannhowidth = Me.Width
frmExibicao2.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmExibicao2
End Sub

Private Sub maskDataExpedicao_Change()
    frmExibicao2.lblDataExpedicao2.Caption = maskDataExpedicao.Text
End Sub

Private Sub maskDataFabricacao_Change()
    'frmExibicao.lblDataFabricacao2.Caption = mskDataFabricacao.Text
End Sub


Private Sub maskDataExpedicao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtDocFiscal.SetFocus
    End If
End Sub

Private Sub mskDataFabricacao_Change()
    frmExibicao2.lblDataProducao2.Caption = mskDataFabricacao.Text
End Sub

Private Sub mskDataFabricacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtCodFiat.SetFocus
    End If
End Sub

Private Sub mskDum_Change()
    frmExibicao2.lblDum2.Caption = mskDum.Text
End Sub

Private Sub mskDum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtEmbalagem.SetFocus
    End If
End Sub

Private Sub txtClasseFuncional_Change()
    frmExibicao2.lblClasseFuncional2.Caption = txtClasseFuncional.Text
End Sub

Private Sub txtClasseFuncional_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtVinculo.SetFocus
    End If
End Sub

Private Sub txtCodFiat_Change()
    frmExibicao2.lblCodEmbalagem2.Caption = txtCodFiat.Text
    frmExibicao2.lblCodBarra.Caption = txtNumFiat.Text & txtQtdEmbalagem.Text & txtCodFiat.Text
    frmExibicao2.lblCodBarraCp1.Caption = txtNumFiat.Text & txtQtdEmbalagem.Text & txtCodFiat.Text
    frmExibicao2.lblCodBarraCp2.Caption = txtNumFiat.Text & txtQtdEmbalagem.Text & txtCodFiat.Text
    frmExibicao2.lblCodBarra2.Caption = txtNumFiat.Text & txtQtdEmbalagem.Text & txtCodFiat.Text
End Sub

Private Sub txtCodFiat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtNumLote.SetFocus
    End If
End Sub


Private Sub txtDescrPeca_Change()
    frmExibicao2.lblDenominacao2.Caption = txtDescrPeca.Text
End Sub
Private Sub txtDescrPeca_KeyPress(KeyAscii As Integer)
    Dim char As String * 1
    
    If KeyAscii = vbKeyReturn Then
        txtNumFiat.SetFocus
    End If
    
    char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(char))
End Sub


Private Sub txtDocFiscal_Change()
    frmExibicao2.lblBam2.Caption = txtDocFiscal.Text
End Sub

Private Sub txtDocFiscal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDescrPeca.SetFocus
    End If
End Sub

Private Sub txtEmbalagem_Change()
    frmExibicao2.lblEmbalagem2.Caption = txtEmbalagem.Text
End Sub

Private Sub txtEmbalagem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtQtdEtiquetas.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtEmbarque_Change()
    frmExibicao2.lblEmbarqueControlado2.Caption = txtEmbarque.Text
End Sub

Private Sub txtEmbarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtLoteDesvio.SetFocus
    End If
End Sub

Private Sub txtIndicacao_Change()
    frmExibicao2.lblIndicacaoSuplementar2.Caption = txtIndicacao.Text
End Sub

Private Sub txtIndicacao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtEmbarque.SetFocus
    End If
End Sub

Private Sub txtLoteDesvio_Change()
    frmExibicao2.lblLoteSobDesvio2.Caption = txtLoteDesvio.Text
End Sub

Private Sub txtLoteDesvio_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        mskDum.SetFocus
    End If
End Sub

Private Sub txtNumFiat_Change()
    frmExibicao2.lblDesenho2.Caption = txtNumFiat.Text
    frmExibicao2.lblCodBarra.Caption = txtNumFiat.Text & txtQtdEmbalagem.Text & txtCodFiat.Text
    frmExibicao2.lblCodBarraCp1.Caption = txtNumFiat.Text & txtQtdEmbalagem.Text & txtCodFiat.Text
    frmExibicao2.lblCodBarraCp2.Caption = txtNumFiat.Text & txtQtdEmbalagem.Text & txtCodFiat.Text
    frmExibicao2.lblCodBarra2.Caption = txtNumFiat.Text & txtQtdEmbalagem.Text & txtCodFiat.Text
End Sub
Private Sub txtNumFiat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        mskDataFabricacao.SetFocus
    End If
End Sub


Private Sub txtNumLote_Change()
    frmExibicao2.lblNumLote2.Caption = txtNumLote.Text
End Sub

Private Sub txtNumLote_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtQtdLote.SetFocus
    End If
End Sub

Private Sub txtQtd_Change()
'    frmExibicao.lblQtd2.Caption = txtQtd
'    frmExibicao.lblCodBarra.Caption = txtNumFiat.Text & txtQtd.Text & txtCodFiat.Text
'    frmExibicao.lblCodBarraCp1.Caption = txtNumFiat.Text & txtQtd.Text & txtCodFiat.Text
'    frmExibicao.lblCodBarraCp2.Caption = txtNumFiat.Text & txtQtd.Text & txtCodFiat.Text
'    frmExibicao.lblCodBarra2.Caption = txtNumFiat.Text & txtQtd.Text & txtCodFiat.Text
End Sub


Private Sub txtQtd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCodFiat.SetFocus
    End If
    
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub


Private Sub txtQtdEmbalagem_Change()
    frmExibicao2.lblQtdEmbalagem2.Caption = txtQtdEmbalagem.Text
End Sub

Private Sub txtQtdEmbalagem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtClasseFuncional.SetFocus
    End If
End Sub

Private Sub txtQtdEtiquetas_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        cmdImprimir.SetFocus
    End If
    
    KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    
End Sub



Public Sub Limpa_Etiqueta()
    'Procedimento para limpar os campos etiqueta
'''    frmExibicao.lblDataExpedicao2.Caption = ""
'''    frmExibicao.lblDocFiscal2.Caption = ""
'''    frmExibicao.lblPeca2.Caption = ""
'''    frmExibicao.lblDesenho2.Caption = ""
'''    frmExibicao.lblCodBarra.Caption = ""
'''    frmExibicao.lblCodBarra2.Caption = ""
'''    frmExibicao.lblCodBarraCp1.Caption = ""
'''    frmExibicao.lblCodBarraCp2.Caption = ""
'''    frmExibicao.lblQtd2.Caption = ""
'''    frmExibicao.lblCodEmbalagem2.Caption = ""
'''    frmExibicao.lblDataFabricacao2.Caption = ""
'''    frmExibicao.lblNumLote2.Caption = ""
End Sub

Private Sub txtQtdEtiquetas_LostFocus()
    
    If Not IsNumeric(txtQtdEtiquetas.Text) Then
        MsgBox "Digite um número válido.", vbExclamation, objApplication.tituloPrograma
        txtQtdEtiquetas.SetFocus
    End If
    
End Sub

Private Sub txtQtdLote_Change()
    frmExibicao2.lblQtdLote2.Caption = txtQtdLote.Text
End Sub

Private Sub txtQtdLote_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtQtdEmbalagem.SetFocus
    End If
End Sub

Private Sub txtVinculo_Change()
    frmExibicao2.lblVinculo2.Caption = txtVinculo.Text
End Sub

Private Sub txtVinculo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtIndicacao.SetFocus
    End If
End Sub
