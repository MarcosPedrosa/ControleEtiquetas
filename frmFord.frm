VERSION 5.00
Begin VB.Form frmFord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Infomações da etiqueta - Modelo FORD"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   Icon            =   "frmFord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5115
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
      Left            =   240
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   3630
      Width           =   3555
   End
   Begin VB.TextBox txtEmbalagem 
      Height          =   300
      Left            =   1440
      MaxLength       =   5
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtDestino 
      Height          =   300
      Left            =   240
      MaxLength       =   5
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtCodUtil 
      Height          =   300
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox TxtSufixo 
      Height          =   300
      Left            =   2640
      MaxLength       =   5
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtCodigoMsb 
      Height          =   300
      Left            =   240
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtNumSerial 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   300
      Left            =   2640
      MaxLength       =   10
      TabIndex        =   15
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtQtdEtiquetas 
      Height          =   300
      Left            =   3240
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "1"
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   855
      Left            =   3960
      Picture         =   "frmFord.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Nova Etiqueta"
      Height          =   855
      Left            =   3960
      Picture         =   "frmFord.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   855
      Left            =   3960
      Picture         =   "frmFord.frx":0EEE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtLote 
      Height          =   300
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtLinhaUtil 
      Height          =   300
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtQtd 
      Height          =   300
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtNumFornecedor 
      Height          =   300
      Left            =   240
      MaxLength       =   8
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtNumPeca 
      Height          =   300
      Left            =   240
      MaxLength       =   15
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Impressora"
      Height          =   195
      Left            =   240
      TabIndex        =   28
      Top             =   3360
      Width           =   765
   End
   Begin VB.Label lblEmbalagem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Embalagem"
      Height          =   195
      Left            =   1440
      TabIndex        =   27
      Top             =   2280
      Width           =   825
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Destino"
      Height          =   195
      Left            =   240
      TabIndex        =   26
      Top             =   2280
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cod. Utilização"
      Height          =   195
      Left            =   1440
      TabIndex        =   25
      Top             =   1560
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sufixo"
      Height          =   195
      Left            =   2640
      TabIndex        =   24
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cod. MSB"
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   840
      Width           =   720
   End
   Begin VB.Label lblQtdEmbalagem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Serial"
      Height          =   195
      Left            =   2640
      TabIndex        =   22
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label lblQtdEtiquetas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade de etiquetas a imprimir"
      Height          =   195
      Left            =   600
      TabIndex        =   21
      Top             =   3150
      Width           =   2430
   End
   Begin VB.Label lblLote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do lote"
      Height          =   195
      Left            =   1440
      TabIndex        =   20
      Top             =   840
      Width           =   705
   End
   Begin VB.Label lblLinhaUtilizacao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lin. utilização"
      Height          =   195
      Left            =   2640
      TabIndex        =   19
      Top             =   1560
      Width           =   960
   End
   Begin VB.Label lblQtdLote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde"
      Height          =   195
      Left            =   2640
      TabIndex        =   18
      Top             =   840
      Width           =   345
   End
   Begin VB.Label lblNumeroPeca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº fornecedor"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   1560
      Width           =   990
   End
   Begin VB.Label lblDocFiscal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº da peça"
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   810
   End
End
Attribute VB_Name = "frmFord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim char As String * 1
Dim i As Integer    'contador de etiquetas

Private Sub cmdImprimir_Click()
    On Error GoTo Erro
'''    If Not IsDate(mskDataFabricacao.Text) Then
'''        MsgBox "Informe uma data válida", vbInformation + vbOKOnly, "ATENÇÃO!!!"
'''        mskDataFabricacao.SetFocus
'''        Exit Sub
'''    End If
'''
'''    'Virou txtNumeroPeca
'''    If Trim(txtDocFiscal.Text) = Empty Then
'''        MsgBox "Informe o número do documento fiscal", vbInformation + vbOKOnly, "ATENÇÃO!!!"
'''        txtDocFiscal.SetFocus
'''        Exit Sub
'''    End If
'''
'''    If Trim(txtDescrPeca.Text) = Empty Then
'''        MsgBox "Informe a descrição da peça", vbInformation + vbOKOnly, "ATENÇÃO!!!"
'''        txtDescrPeca.SetFocus
'''        Exit Sub
'''    End If
'''
'''    If Trim(txtNumFiat.Text) = Empty Then
'''        MsgBox "Informe o número do desenho FIAT", vbOKOnly + vbInformation, "ATENÇÃO!!!"
'''        txtNumFiat.SetFocus
'''        Exit Sub
'''    End If
'''
''''    If Trim(txtQtd.Text) = Empty Then
''''        MsgBox "Informe a quantidade de peças", vbOKOnly + vbInformation, "ATENÇÃO!!!"
''''        txtQtd.SetFocus
''''        Exit Sub
''''    End If
'''
'''    If Trim(txtCodFiat.Text) = Empty Then
'''        MsgBox "Informe o código da embalagem FIAT", vbOKOnly + vbInformation, "ATENÇÃO!!!"
'''        txtCodFiat.SetFocus
'''        Exit Sub
'''    End If
'''
'''    If Trim(txtNumLote.Text) = Empty Then
'''        MsgBox "Informe o número do lote", vbOKOnly + vbInformation, "ATENÇÃO!!!"
'''        txtNumLote.SetFocus
'''        Exit Sub
'''    End If
'''
'''    If Trim(txtQtdEtiquetas.Text) = Empty Then
'''        MsgBox "Selecion quantas etiquetas deseja imprimir", vbExclamation, objApplication.tituloPrograma
'''        txtQtdEtiquetas.SetFocus
'''        Exit Sub
'''    End If
'''
'''    For I = 1 To txtQtdEtiquetas.Text Step 1
'''        Printer.Orientation = 2
'''        frmExibicao3ant.PrintForm
'''    Next
'''
'''    MsgBox "Impressão concluída com sucesso! O aplicativo será encerrado!", vbOKOnly + vbInformation, "Tarefa Concluída"
'''    MDIEtiquetas.forcaSaida = True
'''    Unload MDIEtiquetas
'''
    
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
    
    Printer.Orientation = 2
    Printer.Copies = CInt(txtQtdEtiquetas.Text)
    
    frmExibicao3ant.PrintForm
    Printer.Orientation = 2: Printer.EndDoc
    
    MsgBox "Impressão concluída com sucesso! O aplicativo será encerrado!", vbOKOnly + vbInformation, "Tarefa Concluída"
    MDIEtiquetas.forcaSaida = True
    Unload MDIEtiquetas
    
    
    Exit Sub
Erro:
    MsgBox "Ocorreu um erro na impressão da etiqueta." & vbNewLine & "Err. N: " & Err.Number & vbNewLine & "Desc: " & Err.Description, vbCritical, objApplication.tituloPrograma
End Sub

Private Sub cmdNovo_Click()
    Limpa_Etiqueta
    txtNumPeca.SetFocus
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
frmExibicao3ant.Show

txtNumSerial.Text = CStr(objEtiquetaControlador.getProximaEtiquetaFordAvulso)
    
End Sub

Public Sub Limpa_Etiqueta()
    txtNumPeca.Text = ""
    TxtSufixo.Text = ""
    txtCodigoMsb.Text = ""
    txtLote.Text = ""
    txtQtd.Text = ""
    txtNumFornecedor.Text = ""
    txtCodUtil.Text = ""
    txtLinhaUtil.Text = ""
    txtDestino.Text = ""
    
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

Private Sub Form_Unload(Cancel As Integer)
    Unload frmExibicao3ant
End Sub

Private Sub txtCodigoMsb_Change()
    
    frmExibicao3ant.lblCod_Peca.Caption = txtCodigoMsb.Text
    
End Sub

Private Sub txtCodigoMsb_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtLote.SetFocus
    End If
    
    char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(char))
    
End Sub

Private Sub txtCodUtil_Change()
    
    frmExibicao3ant.lblCodUtil = txtCodUtil.Text
    
End Sub

Private Sub txtCodUtil_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtLinhaUtil.SetFocus
    End If
    
    char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(char))
    
End Sub

Private Sub txtDestino_Change()
    
    frmExibicao3ant.lblDestino.Caption = txtDestino.Text
    
    frmExibicao3ant.lblDestinoA.Caption = "*D" & frmExibicao3ant.lblDestino.Caption & "*"
    frmExibicao3ant.lblDestinoB.Caption = frmExibicao3ant.lblDestinoA.Caption
    
End Sub

Private Sub txtDestino_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtEmbalagem.SetFocus
    End If
    
    char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(char))
    
End Sub

Private Sub txtEmbalagem_Change()
    frmExibicao3ant.lblEmbalagem2.Caption = txtEmbalagem.Text
End Sub

Private Sub txtEmbalagem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtQtdEtiquetas.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtLinhaUtil_Change()
    
    frmExibicao3ant.lblLinhaUtil = txtLinhaUtil.Text
    
End Sub

Private Sub txtLinhaUtil_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtDestino.SetFocus
    End If
    
    char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(char))
    
End Sub

Private Sub txtLote_Change()
    
    frmExibicao3ant.lblLote.Caption = txtLote.Text
    
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtQtd.SetFocus
    End If
    
    char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(char))
    
End Sub

Private Sub txtNumFornecedor_Change()
    
    frmExibicao3ant.lblNumFornec = txtNumFornecedor.Text
    If Trim(txtNumFornecedor.Text) <> "" Then
        frmExibicao3ant.lblNumFornecA.Caption = "*V" & frmExibicao3ant.lblNumFornec.Caption & "*"
        frmExibicao3ant.lblNumFornecB.Caption = frmExibicao3ant.lblNumFornecA.Caption
    Else
        frmExibicao3ant.lblNumFornecA.Caption = ""
        frmExibicao3ant.lblNumFornecB.Caption = ""
    End If
    
End Sub

Private Sub txtNumFornecedor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtCodUtil.SetFocus
    End If
    
    char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(char))
    
End Sub

Private Sub txtNumPeca_Change()
    
    frmExibicao3ant.lblNumPeca.Caption = Left(txtNumPeca.Text, 16)
    frmExibicao3ant.lblNumPecaA.Caption = "*P" & Trim(frmExibicao3ant.lblNumPeca.Caption) & "*"
    frmExibicao3ant.lblNumPecaB.Caption = frmExibicao3ant.lblNumPecaA.Caption
    
    frmExibicao3ant.lblNumPeca.Caption = Trim(frmExibicao3ant.lblNumPeca.Caption)
    
End Sub

Private Sub txtNumPeca_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        TxtSufixo.SetFocus
    End If
    
    char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(char))
    
End Sub

Private Sub txtNumSerial_Change()
    
    frmExibicao3ant.lblNumSerial.Caption = txtNumSerial.Text
    
    frmExibicao3ant.lblNumSerialA.Caption = "*S" & frmExibicao3ant.lblNumSerial.Caption & "*"
    frmExibicao3ant.lblNumSerialB.Caption = frmExibicao3ant.lblNumSerialA.Caption
    
End Sub

Private Sub txtQtd_Change()
    
    frmExibicao3ant.lblQtd.Caption = txtQtd.Text
    frmExibicao3ant.lblQtdA.Caption = "*Q" & frmExibicao3ant.lblQtd.Caption & "*"
    frmExibicao3ant.lblQtdB.Caption = frmExibicao3ant.lblQtdA.Caption
    
End Sub

Private Sub txtQtd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtNumFornecedor.SetFocus
    End If
    
    KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    
End Sub

Private Sub txtQtdEtiquetas_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtQtdEtiquetas_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        cmdImprimir.SetFocus
    End If
    KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    
End Sub

Private Sub txtQtdEtiquetas_LostFocus()
    If Not IsNumeric(txtQtdEtiquetas.Text) Then
        MsgBox "Digite um número válido.", vbExclamation, objApplication.tituloPrograma
        txtQtdEtiquetas.SetFocus
    End If
End Sub

Private Sub TxtSufixo_Change()
    
    frmExibicao3ant.lblSufixo.Caption = Trim(Right(Trim(TxtSufixo.Text), 5))
    frmExibicao3ant.lblSufixoA.Caption = "*C" & frmExibicao3ant.lblSufixo.Caption & "*"
    frmExibicao3ant.lblSufixoB.Caption = frmExibicao3ant.lblSufixoA.Caption
    
End Sub

Private Sub TxtSufixo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtCodigoMsb.SetFocus
    End If
    
    char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(char))
    
End Sub
