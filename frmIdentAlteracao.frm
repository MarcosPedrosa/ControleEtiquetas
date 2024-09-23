VERSION 5.00
Begin VB.Form frmIdentAlteracao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Identificação de Alterações do produto e processo"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6030
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
      TabIndex        =   26
      Text            =   "Combo1"
      Top             =   4740
      Width           =   4425
   End
   Begin VB.Frame fraEnvioLote 
      Height          =   615
      Left            =   120
      TabIndex        =   25
      Top             =   3360
      Width           =   4455
      Begin VB.OptionButton optPrimeiroEnvio 
         Caption         =   "1º Envio"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton optLoteIntermediario 
         Caption         =   "Lote intermediário"
         Height          =   255
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton optUltimoLote 
         Caption         =   "Último lote"
         Height          =   255
         Left            =   3120
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame fraMotivoAlteracao 
      Caption         =   "Motivo da alteração"
      Height          =   1095
      Left            =   120
      TabIndex        =   24
      Top             =   2160
      Width           =   4455
      Begin VB.OptionButton optOutros 
         Caption         =   "Outros"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   750
         Width           =   1095
      End
      Begin VB.OptionButton optProdutoNovo 
         Caption         =   "Produto novo"
         Height          =   195
         Left            =   2400
         TabIndex        =   10
         Top             =   510
         Width           =   1455
      End
      Begin VB.OptionButton optDesvio 
         Caption         =   "Desvio"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   270
         Width           =   975
      End
      Begin VB.OptionButton optReparoRetrabalho 
         Caption         =   "Reparo retrabalho"
         Height          =   195
         Left            =   2400
         TabIndex        =   9
         Top             =   270
         Width           =   1575
      End
      Begin VB.OptionButton optMaterialSelecionado 
         Caption         =   "Material selecionado"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   510
         Width           =   1815
      End
   End
   Begin VB.Frame fraTipoAlteracao 
      Caption         =   "Tipo de alteração"
      Height          =   615
      Left            =   120
      TabIndex        =   23
      Top             =   1440
      Width           =   4455
      Begin VB.OptionButton optLoteUnico 
         Caption         =   "Lote único"
         Height          =   255
         Left            =   3120
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optProvisoria 
         Caption         =   "Provisória"
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton optDefinitiva 
         Caption         =   "Definitiva"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.TextBox txtDesenho 
      Height          =   300
      Left            =   240
      MaxLength       =   15
      TabIndex        =   0
      Text            =   "93380063"
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   4800
      Picture         =   "frmIdentAlteracao.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Nova Etiqueta"
      Height          =   855
      Left            =   4800
      Picture         =   "frmIdentAlteracao.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sai&r"
      Height          =   855
      Left            =   4800
      Picture         =   "frmIdentAlteracao.frx":0CD4
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtQtdEtiquetas 
      Height          =   300
      Left            =   3480
      MaxLength       =   3
      TabIndex        =   14
      Text            =   "1"
      Top             =   4200
      Width           =   495
   End
   Begin VB.TextBox txtNotaFiscal 
      Height          =   300
      Left            =   240
      MaxLength       =   12
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtDesvio 
      Height          =   300
      Left            =   2040
      TabIndex        =   1
      Text            =   "Engrenagem planetária"
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Impressora"
      Height          =   195
      Left            =   180
      TabIndex        =   27
      Top             =   4470
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desvio/Aviso de mod."
      Height          =   195
      Left            =   2040
      TabIndex        =   22
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desenho"
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   120
      Width           =   645
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade de etiquetas a imprimir"
      Height          =   195
      Left            =   840
      TabIndex        =   20
      Top             =   4260
      Width           =   2430
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota Fiscal"
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   720
      Width           =   795
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2640
      TabIndex        =   18
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "frmIdentAlteracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
    
    Printer.Orientation = 2
    Printer.Copies = CInt(txtQtdEtiquetas.Text)
    
    frmExibicao6.PrintForm
    Printer.Orientation = 2: Printer.EndDoc
    
    MsgBox "Impressão concluída com sucesso! O aplicativo será encerrado!", vbOKOnly + vbInformation, "Tarefa Concluída"
    MDIEtiquetas.forcaSaida = True
    Unload MDIEtiquetas
    
End Sub

Private Sub cmdNovo_Click()
    
    txtDesenho.Text = ""
    txtDesvio.Text = ""
    txtNotaFiscal.Text = ""
    
    optProvisoria.Value = True
    optDesvio.Value = True
    optLoteIntermediario.Value = True
    
    txtQtdEtiquetas.Text = "1"
    
    txtDesenho.SetFocus
    
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

frmExibicao6.Show

'Setando os valores padrões
txtDesenho.Text = "581885710"
txtDesvio.Text = "PWT 040 de 15/04/2005"
txtNotaFiscal.Text = "137633"

optProvisoria.Value = True
optDesvio.Value = True
optLoteIntermediario.Value = True

txtQtdEtiquetas.Text = "1"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmExibicao6
End Sub


Private Sub optDefinitiva_Click()
    Call acaoFraTipoAlteracao
End Sub

Private Sub optDefinitiva_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call focaFraMotivoAlteracao
    End If
    
End Sub



Private Sub optDesvio_Click()
    Call acaoFraMotivoAlteracao
End Sub

Private Sub optDesvio_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call focaFraEnvioLote
    End If
    
End Sub

Private Sub optLoteIntermediario_Click()
    Call acaoFraEnvioLote
End Sub

Private Sub optLoteIntermediario_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtQtdEtiquetas.SetFocus
    End If
    
End Sub

Private Sub optLoteUnico_Click()
    Call acaoFraTipoAlteracao
End Sub

Private Sub optLoteUnico_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call focaFraMotivoAlteracao
    End If
    
End Sub


Private Sub optMaterialSelecionado_Click()
    Call acaoFraMotivoAlteracao
End Sub

Private Sub optMaterialSelecionado_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call focaFraEnvioLote
    End If
    
End Sub

Private Sub optOutros_Click()
    Call acaoFraMotivoAlteracao
End Sub

Private Sub optOutros_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call focaFraEnvioLote
    End If
    
End Sub


Private Sub optPrimeiroEnvio_Click()
    Call acaoFraEnvioLote
End Sub

Private Sub optPrimeiroEnvio_KeyPress(KeyAscii As Integer)
        
    If KeyAscii = vbKeyReturn Then
        txtQtdEtiquetas.SetFocus
    End If
    
End Sub

Private Sub optProdutoNovo_Click()
    Call acaoFraMotivoAlteracao
End Sub

Private Sub optProdutoNovo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call focaFraEnvioLote
    End If
    
End Sub

Private Sub optProvisoria_Click()
    Call acaoFraTipoAlteracao
End Sub

Private Sub optProvisoria_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call focaFraMotivoAlteracao
    End If
    
End Sub

Private Sub optReparoRetrabalho_Click()
    Call acaoFraMotivoAlteracao
End Sub

Private Sub optReparoRetrabalho_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call focaFraEnvioLote
    End If
    
End Sub

Private Sub optUltimoLote_Click()
    Call acaoFraEnvioLote
End Sub

Private Sub optUltimoLote_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtQtdEtiquetas.SetFocus
    End If
    
End Sub

Private Sub txtDesenho_Change()
    frmExibicao6.lblDesenho.Caption = txtDesenho.Text
End Sub

Private Sub txtDesenho_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtDesenho_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtDesvio.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
    
End Sub

Private Sub txtDesvio_Change()
    frmExibicao6.lblDesvio.Caption = txtDesvio.Text
End Sub

Private Sub txtDesvio_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtDesvio_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        txtNotaFiscal.SetFocus
    End If
    
End Sub

Private Sub txtNotaFiscal_Change()
    frmExibicao6.lblNotaFiscal1.Caption = txtNotaFiscal.Text
    frmExibicao6.lblNotaFiscal2.Caption = frmExibicao6.lblNotaFiscal1.Caption
End Sub

Private Sub txtNotaFiscal_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtNotaFiscal_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call focaFraTipoAlteracao
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
    
End Sub

Private Sub focaFraTipoAlteracao()

    If optDefinitiva.Value = True Then
        optDefinitiva.SetFocus
    ElseIf optProvisoria.Value = True Then
        optProvisoria.SetFocus
    ElseIf optLoteUnico.Value = True Then
        optLoteUnico.SetFocus
    End If
    
End Sub

Private Sub focaFraMotivoAlteracao()
    
    If optDesvio.Value = True Then
        optDesvio.SetFocus
    ElseIf optMaterialSelecionado.Value = True Then
        optMaterialSelecionado.SetFocus
    ElseIf optOutros.Value = True Then
        optOutros.SetFocus
    ElseIf optReparoRetrabalho.Value = True Then
        optReparoRetrabalho.SetFocus
    ElseIf optProdutoNovo.Value = True Then
        optProdutoNovo.SetFocus
    Else
        optDesvio.SetFocus
    End If
    
End Sub

Private Sub focaFraEnvioLote()
    
    If optPrimeiroEnvio.Value = True Then
        optPrimeiroEnvio.SetFocus
    ElseIf optLoteIntermediario.Value = True Then
        optLoteIntermediario.SetFocus
    ElseIf optUltimoLote.Value = True Then
        optUltimoLote.SetFocus
    End If
    
End Sub

Private Sub txtQtdEtiquetas_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtQtdEtiquetas_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        cmdImprimir.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
     
    
End Sub

Private Sub txtQtdEtiquetas_LostFocus()
    If Not IsNumeric(txtQtdEtiquetas) Then
        MsgBox "Digite um número válido.", vbExclamation, objApplication.tituloPrograma
        txtQtdEtiquetas.SetFocus
    End If
End Sub

Private Sub acaoFraTipoAlteracao()
    
    frmExibicao6.lblOptDefinitiva.Visible = False
    frmExibicao6.lblOptProvisoria.Visible = False
    frmExibicao6.lblOptLoteUnico.Visible = False
    If optDefinitiva.Value = True Then
        frmExibicao6.lblOptDefinitiva.Visible = True
    ElseIf optProvisoria.Value = True Then
        frmExibicao6.lblOptProvisoria.Visible = True
    ElseIf optLoteUnico.Value = True Then
        frmExibicao6.lblOptLoteUnico.Visible = True
    End If
    
End Sub

Private Sub acaoFraMotivoAlteracao()
    
    frmExibicao6.lblOptDesvio.Visible = False
    frmExibicao6.lblOptMaterialSelecionado.Visible = False
    frmExibicao6.lblOptOutros.Visible = False
    frmExibicao6.lblOptReparoRetrabaho.Visible = False
    frmExibicao6.lblOptProdutoNovo.Visible = False
    If optDesvio.Value = True Then
        frmExibicao6.lblOptDesvio.Visible = True
    ElseIf optMaterialSelecionado.Value = True Then
        frmExibicao6.lblOptMaterialSelecionado.Visible = True
    ElseIf optOutros.Value = True Then
        frmExibicao6.lblOptOutros.Visible = True
    ElseIf optReparoRetrabalho.Value = True Then
        frmExibicao6.lblOptReparoRetrabaho.Visible = True
    ElseIf optProdutoNovo.Value = True Then
        frmExibicao6.lblOptProdutoNovo.Visible = True
    End If
    
End Sub

Private Sub acaoFraEnvioLote()
    
    frmExibicao6.lblOptPrimeiroEnvio.Visible = False
    frmExibicao6.lblOptLoteIntermediario.Visible = False
    frmExibicao6.lblOptUltimoLote.Visible = False
    If optPrimeiroEnvio.Value = True Then
        frmExibicao6.lblOptPrimeiroEnvio.Visible = True
    ElseIf optLoteIntermediario.Value = True Then
        frmExibicao6.lblOptLoteIntermediario.Visible = True
    ElseIf optUltimoLote.Value = True Then
        frmExibicao6.lblOptUltimoLote.Visible = True
    End If
    
End Sub
