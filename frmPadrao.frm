VERSION 5.00
Begin VB.Form frmPadrao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informações da etiqueta - Modelo padrão"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "frmPadrao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4170
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
      Left            =   120
      TabIndex        =   21
      Text            =   "Combo1"
      Top             =   3750
      Width           =   3735
   End
   Begin VB.TextBox txtEmbalagem 
      Height          =   285
      Left            =   120
      MaxLength       =   5
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   855
      Left            =   4050
      Picture         =   "frmPadrao.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Nova Etiqueta"
      Height          =   855
      Left            =   4050
      Picture         =   "frmPadrao.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2130
      Width           =   975
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   855
      Left            =   4050
      Picture         =   "frmPadrao.frx":1116
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3210
      Width           =   975
   End
   Begin VB.TextBox txtQtdEtiquetas 
      Height          =   300
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "1"
      Top             =   3150
      Width           =   495
   End
   Begin VB.TextBox txtLote 
      Height          =   285
      Left            =   120
      MaxLength       =   18
      TabIndex        =   7
      Top             =   2160
      Width           =   3735
   End
   Begin VB.TextBox txtPeso 
      Height          =   285
      Left            =   2640
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtQtd 
      Height          =   285
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtCodCliente 
      Height          =   285
      Left            =   120
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtDescricao 
      Height          =   285
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox txtCodPeca 
      Height          =   285
      Left            =   120
      MaxLength       =   7
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   120
      MaxLength       =   50
      TabIndex        =   1
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Impressora"
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   3480
      Width           =   765
   End
   Begin VB.Label lblEmbalagem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Embalagem"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   2520
      Width           =   825
   End
   Begin VB.Label lblQtdEtiquetas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade de etiquetas a imprimir"
      Height          =   195
      Left            =   90
      TabIndex        =   19
      Top             =   3180
      Width           =   2430
   End
   Begin VB.Label lblLote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lote"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   315
   End
   Begin VB.Label lblPeso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Peso"
      Height          =   195
      Left            =   2640
      TabIndex        =   17
      Top             =   1320
      Width           =   360
   End
   Begin VB.Label lblQuantidade 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qtd."
      Height          =   195
      Left            =   1920
      TabIndex        =   16
      Top             =   1320
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód. Cliente"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblDescricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      Height          =   195
      Left            =   1200
      TabIndex        =   14
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lblCodPeca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód. MSB"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lblCliente 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmPadrao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim char As String * 1


Private Sub cmdImprimir_Click()
    Dim Vezes As Integer
    Dim v As Integer
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
            
'''    For v = 0 To (Forms.Count - 1)
'''        If Forms(v).Name = "frmExibicao1" Or Forms(v).Name = "frmExibicao4" Then
'''            Printer.Orientation = 2
'''            frmAvulsoPadraoPonteiro.PrintForm
'''            Printer.EndDoc
'''            Exit For
'''        End If
'''        If Forms(v).Name = "frmExibicao2" Then
'''            Printer.Orientation = 1
'''            frmExibicao2.PrintForm
'''            Printer.EndDoc
'''            Exit For
'''        End If
'''        If Forms(v).Name = "frmExibicao3" Then
'''            Printer.Orientation = 2
'''            frmExibicao3.PrintForm
'''            Printer.Orientation = 2 : Printer.EndDoc
'''            Exit For
'''        End If
'''        If Forms(v).Name = "frmExibicao5" Then
'''            Printer.Orientation = 2
'''            frmExibicao5.PrintForm
'''            Printer.Orientation = 2 : Printer.EndDoc
'''            Exit For
'''        End If
'''    Next
    
    If Not IsNumeric(txtQtdEtiquetas.Text) Then
        MsgBox "Digite a quantidade de etiquetas a serem impressas.", vbExclamation, objApplication.tituloPrograma
        txtQtdEtiquetas.SetFocus
        Exit Sub
    End If
    
'    Printer.Copies = CInt(txtQtdEtiquetas.Text)
    Vezes = 0
    While CInt(txtQtdEtiquetas.Text) > Vezes
          Printer.Orientation = 1
          frmAvulsoPadraoPonteiro.Height = 4785
          frmAvulsoPadraoPonteiro.PrintForm
          Printer.Orientation = 2: Printer.EndDoc
          Vezes = Vezes + 1
          If Vezes < CInt(txtQtdEtiquetas.Text) Then Set Printer = x
    Wend
    
    MsgBox "Impressão concluída com sucesso! O aplicativo será encerrado!", vbOKOnly + vbInformation, "Tarefa Concluída"
    MDIEtiquetas.forcaSaida = True
    Unload MDIEtiquetas
    
    Exit Sub
Erro:
    MsgBox "Ocorreu um erro na impressão da etiqueta." & vbNewLine & "Err. N: " & Err.Number & vbNewLine & "Desc: " & Err.Description, vbCritical, objApplication.tituloPrograma
End Sub

Private Sub cmdNovo_Click()
    
    'Para forçar o evento on_change, porque alguns campos podem estar em branco e jogando vazio ("") não ocorrerá o on_change
    txtCliente.Text = " "
    txtCodPeca.Text = " "
    txtCodCliente.Text = " "
    txtQtd.Text = " "
    txtDescricao.Text = " "
    txtPeso.Text = " "
    txtLote.Text = " "
    
    txtCliente.Text = ""
    txtCodPeca.Text = ""
    txtCodCliente.Text = ""
    txtQtd.Text = ""
    txtDescricao.Text = ""
    txtPeso.Text = ""
    txtLote.Text = ""
    txtQtdEtiquetas.Text = "1"
    
    txtCliente.SetFocus
    
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
    
Me.Top = 0
Me.Left = 0
frmAvulsoPadraoPonteiro.Show
frmAvulsoPadraoPonteiro.Left = Me.Width
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmAvulsoPadraoPonteiro
End Sub

Private Sub txtCliente_Change()
    
    frmAvulsoPadraoPonteiro.lblCliente.Caption = txtCliente.Text
    
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtCodPeca.SetFocus
    End If
    
    char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(char))
End Sub


Private Sub txtCodCliente_Change()
    
    frmAvulsoPadraoPonteiro.lblCodCliente2.Caption = txtCodCliente.Text
    
End Sub

Private Sub txtCodCliente_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtQtd.SetFocus
    End If
    
    char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(char))
    
End Sub


Private Sub txtCodPeca_Change()
    
    frmAvulsoPadraoPonteiro.lblPeca.Caption = txtCodPeca.Text
    
End Sub


Private Sub txtCodPeca_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtDescricao.SetFocus
    End If
    
    char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(char))
End Sub


Private Sub txtDescricao_Change()
    
    frmAvulsoPadraoPonteiro.lblDescricao.Caption = txtDescricao.Text
    
End Sub

Private Sub txtDescricao_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtCodCliente.SetFocus
    End If
    
    char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(char))
    
End Sub


Private Sub txtEmbalagem_Change()
    frmAvulsoPadraoPonteiro.lblEmbalagem2.Caption = txtEmbalagem.Text
End Sub

Private Sub txtEmbalagem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtQtdEtiquetas.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtLote_Change()
    
    frmAvulsoPadraoPonteiro.lblLote2.Caption = txtLote.Text
    
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtEmbalagem.SetFocus
    End If
    
    char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(char))
    
End Sub


Private Sub txtPeso_Change()
    
    frmAvulsoPadraoPonteiro.lblPeso2.Caption = txtPeso.Text
    
End Sub

Private Sub txtPeso_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtLote.SetFocus
    End If
    
    KeyAscii = KeyAsciiCriticaValor(KeyAscii)
    
End Sub

Private Sub txtPeso_LostFocus()
    
    If Trim(txtPeso.Text) <> "" Then
        If Not IsNumeric(txtPeso.Text) Then
            MsgBox "Digite um número válido", vbExclamation, objApplication.tituloPrograma
            txtPeso.SetFocus
            SendKeys "{home}+{end}"
        End If
    End If
    
End Sub

Private Sub txtQtd_Change()
    
    frmAvulsoPadraoPonteiro.lblQtd2.Caption = txtQtd.Text
    
End Sub

Private Sub txtQtd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtPeso.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
    
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
    If Trim(txtQtdEtiquetas.Text) <> "" Then
        If Not IsNumeric(txtQtdEtiquetas.Text) Then
            MsgBox "Digite um número válido", vbExclamation, objApplication.tituloPrograma
            txtQtdEtiquetas.SetFocus
            SendKeys "{home}+{end}"
        End If
    End If
End Sub
