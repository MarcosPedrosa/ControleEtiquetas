VERSION 5.00
Begin VB.Form frmEtiquetaAvulsoPadrao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Etiqueta avulso"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5160
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
      Left            =   2040
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   2520
      Width           =   2925
   End
   Begin VB.ComboBox cbxTipoEmbarque 
      Height          =   315
      ItemData        =   "frmEtiquetaAvulsoPadrao.frx":0000
      Left            =   2040
      List            =   "frmEtiquetaAvulsoPadrao.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtCodigoFuncionario 
      Height          =   285
      Left            =   2040
      MaxLength       =   4
      TabIndex        =   5
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtQuantidade 
      Height          =   285
      Left            =   2040
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Left            =   2040
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   1
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtPeca 
      Height          =   285
      Left            =   2040
      MaxLength       =   9
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   855
      Left            =   3960
      Picture         =   "frmEtiquetaAvulsoPadrao.frx":0014
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1620
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   855
      Left            =   3960
      Picture         =   "frmEtiquetaAvulsoPadrao.frx":0456
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IMPRESSORA:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   15
      Top             =   2520
      Width           =   1155
   End
   Begin VB.Label lblCodigoFuncionario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CÓD. FUNCIONÁRIO:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   14
      Top             =   2160
      Width           =   1785
   End
   Begin VB.Label lblTipoEmbarque 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO EMBARQUE:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   1470
   End
   Begin VB.Label lblQuantidade 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTIDADE:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Label lblLote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOTE:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   525
   End
   Begin VB.Label lblCliente 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTE:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   840
   End
   Begin VB.Label lblPeca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUTO/MATERIAL:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1785
   End
End
Attribute VB_Name = "frmEtiquetaAvulsoPadrao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbxTipoEmbarque_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtCodigoFuncionario.SetFocus
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim objPecaAvulso As PecaAvulso
    Dim qtdePecas As Long
    Dim qtdePecasRestante As Long
    Dim qtdePecasNaCaixa As Integer
    Dim pesoCaixaCompleta As Double
    Dim qtdeEtiquetas As Integer
    Dim nx As Double
    Dim x As Printer
    
    On Error GoTo Erro
           
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
    
    If Len(txtPeca.Text) = 0 Then
        MsgBox "Preencha a peça.", vbExclamation, objApplication.tituloPrograma
        txtPeca.SetFocus
        Exit Sub
    End If
    If Len(txtCliente.Text) = 0 Then
        MsgBox "Preencha o cliente.", vbExclamation, objApplication.tituloPrograma
        txtCliente.SetFocus
        Exit Sub
    End If
    If Len(txtLote.Text) = 0 Then
        MsgBox "Preencha o lote.", vbExclamation, objApplication.tituloPrograma
        txtLote.SetFocus
        Exit Sub
    End If
    If Len(txtQuantidade.Text) = 0 Then
        MsgBox "Preencha a quantidade.", vbExclamation, objApplication.tituloPrograma
        txtQuantidade.SetFocus
        Exit Sub
    End If
    If Len(cbxTipoEmbarque.Text) = 0 Then
        MsgBox "Preencha o tipo do embarque.", vbExclamation, objApplication.tituloPrograma
        cbxTipoEmbarque.SetFocus
        Exit Sub
    End If
    If Len(txtCodigoFuncionario.Text) = 0 Then
        MsgBox "Preencha o código do funcionário.", vbExclamation, objApplication.tituloPrograma
        txtCodigoFuncionario.SetFocus
        Exit Sub
    End If
    If Len(txtCliente.Text) = 0 Then
        MsgBox "Preencha o cliente.", vbExclamation, objApplication.tituloPrograma
        txtCliente.SetFocus
        Exit Sub
    End If
    
    Set objPecaAvulso = objPecaAvulsoControlador.getPecaAvulso(txtPeca.Text, txtCliente.Text)
    If objPecaAvulso Is Nothing Then
        MsgBox "Esta peça associada a este cliente não existe.", vbExclamation, objApplication.tituloPrograma
        txtPeca.SetFocus
        Exit Sub
    End If
    
    frmAvulsoPadraoPonteiro.lblCliente.Caption = objPecaAvulso.itens.Item("CLIENTE")
    frmAvulsoPadraoPonteiro.lblDescricao.Caption = objPecaAvulso.itens.Item("DESCR_PECA")
    frmAvulsoPadraoPonteiro.lblCodCliente2.Caption = objPecaAvulso.itens.Item("COD_CLIENTE")
    
    'FALTA PESO E QTDE
    qtdePecas = CLng(txtQuantidade.Text)
    qtdePecasRestante = qtdePecas
    
    If cbxTipoEmbarque.Text = "A" Then
        qtdePecasNaCaixa = objPecaAvulso.itens.Item("QTDE_CAIXA_A")
        pesoCaixaCompleta = CDbl(objPecaAvulso.itens.Item("PESO_CAIXA_A"))
    ElseIf cbxTipoEmbarque.Text = "R" Then
        qtdePecasNaCaixa = objPecaAvulso.itens.Item("QTDE_CAIXA_R")
        pesoCaixaCompleta = CDbl(objPecaAvulso.itens.Item("PESO_CAIXA_R"))
    Else
        MsgBox "Tipo de embarque inválido.", vbExclamation, objApplication.tituloPrograma
        Exit Sub
    End If
    
    If qtdePecas <= qtdePecasNaCaixa Then 'Imprimirá apenas uma etiqueta
        frmAvulsoPadraoPonteiro.lblQtd2.Caption = qtdePecas
        frmAvulsoPadraoPonteiro.lblPeso2.Caption = Round(qtdePecas * pesoCaixaCompleta / qtdePecasNaCaixa, 2)
        If frmAvulsoPadraoPonteiro.lblPeso2.Caption = "0" Then
            frmAvulsoPadraoPonteiro.lblPeso2.Caption = "1"
        End If
        
        Printer.Copies = 1
        frmAvulsoPadraoPonteiro.PrintForm
        Printer.Orientation = 2: Printer.EndDoc
        DoEvents
        '-------------------------
        'IMPRESSÃO = QTDEETIQUETAS
        '-------------------------
        
    Else 'Imprimirá mais de uma etiqueta
        qtdeEtiquetas = Int(qtdePecas / qtdePecasNaCaixa)
        
        frmAvulsoPadraoPonteiro.lblQtd2.Caption = qtdePecasNaCaixa
        frmAvulsoPadraoPonteiro.lblPeso2.Caption = Round(pesoCaixaCompleta, 2)
        
        Printer.Copies = qtdeEtiquetas
        frmAvulsoPadraoPonteiro.PrintForm
        Printer.Orientation = 2: Printer.EndDoc
        DoEvents
        '----------------------------
        'IMPRESSÃO = QTDEETIQUETAS
        'Nº DE CÓPIAS = qtdeEtiquetas
        '----------------------------
        
        
        If qtdePecas Mod qtdePecasNaCaixa <> 0 Then 'Existe uma etiqueta com quantidade restante
            qtdePecasRestante = qtdePecas - (qtdeEtiquetas * qtdePecasNaCaixa)
            frmAvulsoPadraoPonteiro.lblQtd2.Caption = qtdePecasRestante
            frmAvulsoPadraoPonteiro.lblPeso2.Caption = Round(qtdePecasRestante * pesoCaixaCompleta / qtdePecasNaCaixa, 2)
            If frmAvulsoPadraoPonteiro.lblPeso2.Caption = "0" Then
                frmAvulsoPadraoPonteiro.lblPeso2.Caption = "1"
            End If
            
            Printer.Copies = 1
            frmAvulsoPadraoPonteiro.PrintForm
            Printer.Orientation = 2: Printer.EndDoc
            DoEvents
            '-------------------------
            'IMPRESSÃO = QTDEETIQUETAS
            '-------------------------
        End If
        
    End If
    
    cbxTipoEmbarque.ListIndex = -1
    
    MsgBox "Impressão concluída.", vbInformation, objApplication.tituloPrograma
    txtPeca.SetFocus
    
    Exit Sub
Erro:
    MsgErro "Ocorreu um erro ao imprimir a etiqueta."
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

Me.Top = 0
Me.Left = 0
Me.Height = 3870
Me.Width = 5190

frmAvulsoPadraoPonteiro.Show
frmAvulsoPadraoPonteiro.Left = Me.Width

frmAvulsoPadraoPonteiro.lblCliente = ""
frmAvulsoPadraoPonteiro.lblPeca = ""
frmAvulsoPadraoPonteiro.lblDescricao = ""
frmAvulsoPadraoPonteiro.lblCodCliente2 = ""
frmAvulsoPadraoPonteiro.lblQtd2 = ""
frmAvulsoPadraoPonteiro.lblPeso2 = ""
frmAvulsoPadraoPonteiro.lblLote2 = ""
    
End Sub

Private Sub txtCliente_Change()
    'Este eh o código do cliente e não a descrição
    '''''''''''''frmAvulsoPadraoPonteiro.lblCliente.Caption = txtCliente.Text
End Sub

Private Sub txtCliente_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtLote.SetFocus
    End If
End Sub

Private Sub txtCliente_LostFocus()
    txtCliente.Text = Format(UCase(Trim(txtCliente.Text)), "0000000000")
End Sub

Private Sub txtCodigoFuncionario_Change()
    frmAvulsoPadraoPonteiro.lblEmbalagem2.Caption = txtCodigoFuncionario.Text
End Sub

Private Sub txtCodigoFuncionario_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtCodigoFuncionario_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdImprimir.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtCodigoFuncionario_LostFocus()
    txtCodigoFuncionario.Text = Format(Trim(txtCodigoFuncionario.Text), "0000")
End Sub

Private Sub txtLote_Change()
    frmAvulsoPadraoPonteiro.lblLote2.Caption = txtLote.Text
End Sub

Private Sub txtLote_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtQuantidade.SetFocus
    End If
End Sub

Private Sub txtLote_LostFocus()
    txtLote.Text = UCase(Trim(txtLote.Text))
End Sub

Private Sub txtPeca_Change()
    frmAvulsoPadraoPonteiro.lblPeca.Caption = txtPeca.Text
End Sub

Private Sub txtPeca_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtPeca_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtCliente.SetFocus
    End If
End Sub

Private Sub txtPeca_LostFocus()
    txtPeca.Text = UCase(Trim(txtPeca.Text))
End Sub

Private Sub txtQuantidade_GotFocus()
    SendKeys "{HOME}+{END}"
End Sub

Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cbxTipoEmbarque.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtQuantidade_LostFocus()
    txtQuantidade.Text = Trim(txtQuantidade.Text)
End Sub
