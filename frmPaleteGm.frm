VERSION 5.00
Begin VB.Form frmPaleteGm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Infomações da etiqueta - Modelo GM"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6180
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
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   3900
      Width           =   4725
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Left            =   2520
      MaxLength       =   9
      TabIndex        =   20
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame fraProdutos 
      Caption         =   "PRODUTOS"
      Height          =   2055
      Left            =   0
      TabIndex        =   33
      Top             =   840
      Width           =   4935
      Begin VB.TextBox txtComplPeca2 
         Height          =   300
         Left            =   3720
         MaxLength       =   12
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtComplPeca3 
         Height          =   300
         Left            =   3720
         MaxLength       =   12
         TabIndex        =   13
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtComplPeca4 
         Height          =   300
         Left            =   3720
         MaxLength       =   12
         TabIndex        =   17
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtComplPeca1 
         Height          =   300
         Left            =   3720
         MaxLength       =   12
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtCodigoProduto4 
         Height          =   300
         Left            =   120
         MaxLength       =   15
         TabIndex        =   14
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtPecasPorCaixa4 
         Height          =   300
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   16
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtQtdCaixas4 
         Height          =   300
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   15
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtCodigoProduto3 
         Height          =   300
         Left            =   120
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtPecasPorCaixa3 
         Height          =   300
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   12
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtQtdCaixas3 
         Height          =   300
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox txtCodigoProduto2 
         Height          =   300
         Left            =   120
         MaxLength       =   15
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtPecasPorCaixa2 
         Height          =   300
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtQtdCaixas2 
         Height          =   300
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtCodigoProduto1 
         Height          =   300
         Left            =   120
         MaxLength       =   15
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtPecasPorCaixa1 
         Height          =   300
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox txtQtdCaixas1 
         Height          =   300
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento"
         Height          =   195
         Left            =   3720
         TabIndex        =   39
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Produto"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Peça por caixa"
         Height          =   195
         Left            =   2520
         TabIndex        =   37
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde de caixas"
         Height          =   195
         Left            =   1320
         TabIndex        =   36
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2520
         TabIndex        =   35
         Top             =   240
         Width           =   45
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   2520
         TabIndex        =   34
         Top             =   840
         Width           =   45
      End
   End
   Begin VB.TextBox txtLicense 
      Height          =   300
      Left            =   1440
      MaxLength       =   21
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtCodMSB 
      Height          =   300
      Left            =   120
      MaxLength       =   8
      TabIndex        =   18
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtPlant 
      Height          =   300
      Left            =   120
      MaxLength       =   15
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   5085
      Picture         =   "frmPaleteGm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Nova Etiqueta"
      Height          =   855
      Left            =   5085
      Picture         =   "frmPaleteGm.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sai&r"
      Height          =   855
      Left            =   5085
      Picture         =   "frmPaleteGm.frx":0CD4
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtQtdEtiquetas 
      Height          =   300
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   21
      Text            =   "1"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtPeso 
      Height          =   300
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   19
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Impressora"
      Height          =   195
      Left            =   120
      TabIndex        =   41
      Top             =   3630
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      Height          =   195
      Left            =   2520
      TabIndex        =   40
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Peso"
      Height          =   195
      Left            =   1320
      TabIndex        =   32
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Licença"
      Height          =   195
      Left            =   1440
      TabIndex        =   31
      Top             =   120
      Width           =   570
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código Musashi"
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   3000
      Width           =   1125
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2640
      TabIndex        =   29
      Top             =   120
      Width           =   45
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fábrica/Doca"
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   990
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Etiquetas"
      Height          =   195
      Left            =   3720
      TabIndex        =   27
      Top             =   3000
      Width           =   660
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2520
      TabIndex        =   26
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "frmPaleteGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprimir_Click()
    
    Dim x As Printer
    Dim nx As Double
               
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
    
    frmExibicao7Ref.PrintForm
    Printer.Orientation = 2: Printer.EndDoc
    
    MsgBox "Impressão concluída com sucesso! O aplicativo será encerrado!", vbOKOnly + vbInformation, "Tarefa Concluída"
    MDIEtiquetas.forcaSaida = True
    Unload MDIEtiquetas
    
End Sub

Private Sub cmdNovo_Click()
    
    txtPlant.Text = "4J"
    txtLicense.Text = "UN 897803870"
    
    txtCodigoProduto1.Text = ""
    txtQtdCaixas1.Text = ""
    txtPecasPorCaixa1.Text = ""
    txtComplPeca1.Text = ""
    
    txtCodigoProduto2.Text = ""
    txtQtdCaixas2.Text = ""
    txtPecasPorCaixa2.Text = ""
    txtComplPeca2.Text = ""
    
    txtCodigoProduto3.Text = ""
    txtQtdCaixas3.Text = ""
    txtPecasPorCaixa3.Text = ""
    txtComplPeca3.Text = ""
    
    txtCodigoProduto4.Text = ""
    txtQtdCaixas4.Text = ""
    txtPecasPorCaixa4.Text = ""
    txtComplPeca4.Text = ""
    
    txtCodMSB.Text = ""
    txtPeso.Text = ""
    
    txtQtdEtiquetas.Text = "1"
    
    txtPlant.SetFocus
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

frmExibicao7Ref.Show

'txtPlant.Text = "4J"
'txtLicense.Text = "UN 897803870"

txtQtdEtiquetas.Text = "1"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmExibicao7Ref
End Sub

Private Sub txtCodigoProduto1_Change()
    frmExibicao7Ref.lblCodigoProduto1.Caption = txtCodigoProduto1.Text
End Sub

Private Sub txtCodigoProduto1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtQtdCaixas1.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtCodigoProduto2_Change()
    frmExibicao7Ref.lblCodigoProduto2.Caption = txtCodigoProduto2.Text
End Sub

Private Sub txtCodigoProduto2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtQtdCaixas2.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtCodigoProduto3_Change()
    frmExibicao7Ref.lblCodigoProduto3.Caption = txtCodigoProduto3.Text
End Sub

Private Sub txtCodigoProduto3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtQtdCaixas3.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtCodigoProduto4_Change()
    frmExibicao7Ref.lblCodigoProduto4.Caption = txtCodigoProduto4.Text
End Sub

Private Sub txtCodigoProduto4_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtQtdCaixas4.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtCodMSB_Change()
    frmExibicao7Ref.lblCodMSB.Caption = txtCodMSB.Text
End Sub

Private Sub txtCodMSB_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtCodMSB_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtPeso.SetFocus
    Else
        '''KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtContainerType_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtContainerType_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
    End If
End Sub

Private Sub txtEmbalagem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtQtdEtiquetas.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtEng_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtGrossWeight_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtGrossWeight_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
    End If
End Sub

Private Sub txtCodMSB_LostFocus()
    txtCodMSB.Text = UCase(txtCodMSB.Text)
End Sub

Private Sub txtComplPeca1_Change()
    frmExibicao7Ref.lblComplPeca1.Caption = txtComplPeca1.Text
End Sub

Private Sub txtComplPeca1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If txtCodigoProduto2.Enabled Then
            txtCodigoProduto2.SetFocus
        Else
            txtCodMSB.SetFocus
        End If
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtComplPeca2_Change()
    frmExibicao7Ref.lblComplPeca2.Caption = txtComplPeca2.Text
End Sub

Private Sub txtComplPeca2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtCodigoProduto3.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtComplPeca3_Change()
    frmExibicao7Ref.lblComplPeca3.Caption = txtComplPeca3.Text
End Sub

Private Sub txtComplPeca3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtCodigoProduto4.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtComplPeca4_Change()
    frmExibicao7Ref.lblComplPeca4.Caption = txtComplPeca4.Text
End Sub

Private Sub txtComplPeca4_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtCodMSB.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtData_Change()
    frmExibicao7Ref.lblShipmentDate.Caption = txtData.Text
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtQtdEtiquetas.SetFocus
    End If
End Sub

Private Sub txtData_LostFocus()
    txtData.Text = UCase(txtData.Text)
End Sub

Private Sub txtLicense_Change()
    frmExibicao7Ref.lblLicense = txtLicense.Text
    
    frmExibicao7Ref.lblLicenseA.Caption = "*" & frmExibicao7Ref.lblLicense & "*"
    frmExibicao7Ref.lblLicenseB.Caption = "*" & frmExibicao7Ref.lblLicense & "*"
End Sub

Private Sub txtLicense_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtLicense_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtCodigoProduto1.SetFocus
    End If
End Sub

Private Sub txtLot_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtMaterial_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtMfgDate_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtPartNumber_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtPeca_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtPeca_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
    End If
End Sub

Private Sub txtPecasPorCaixa1_Change()
    Call calculaQtdeProduto(txtQtdCaixas1.Text, txtPecasPorCaixa1.Text, frmExibicao7Ref.lblQtde1)
End Sub

Private Sub txtPecasPorCaixa1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtComplPeca1.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtPecasPorCaixa2_Change()
    Call calculaQtdeProduto(txtQtdCaixas2.Text, txtPecasPorCaixa2.Text, frmExibicao7Ref.lblQtde2)
End Sub

Private Sub txtPecasPorCaixa2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtComplPeca2.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtPecasPorCaixa3_Change()
    Call calculaQtdeProduto(txtQtdCaixas3.Text, txtPecasPorCaixa3.Text, frmExibicao7Ref.lblQtde3)
End Sub

Private Sub txtPecasPorCaixa3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtComplPeca3.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtPecasPorCaixa4_Change()
    Call calculaQtdeProduto(txtQtdCaixas4.Text, txtPecasPorCaixa4.Text, frmExibicao7Ref.lblQtde4)
End Sub

Private Sub txtPecasPorCaixa4_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtComplPeca4.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtPeso_Change()
    frmExibicao7Ref.lblPeso.Caption = txtPeso.Text
End Sub

Private Sub txtPeso_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtData.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtPlant_Change()
    frmExibicao7Ref.lblPlant.Caption = txtPlant.Text
End Sub

Private Sub txtPlant_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtQtd_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtPlant_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtLicense.SetFocus
    End If
End Sub

Private Sub txtPlant_LostFocus()
    txtPlant.Text = UCase(txtPlant.Text)
End Sub

Private Sub txtQtdCaixas1_Change()
    Call calculaQtdeProduto(txtQtdCaixas1.Text, txtPecasPorCaixa1.Text, frmExibicao7Ref.lblQtde1)
    If frmExibicao7Ref.Name = "frmExibicao7UmProduto" Then
        frmExibicao7Ref.lblQtdeContainers.Caption = txtQtdCaixas1.Text
    ElseIf frmExibicao7Ref.Name = "frmExibicao7VariosProdutos" Then
        frmExibicao7Ref.lblQtdeContainers.Caption = getQtdeContainers(txtQtdCaixas1.Text, _
                                                                      txtQtdCaixas2.Text, _
                                                                      txtQtdCaixas3.Text, _
                                                                      txtQtdCaixas4.Text)
    End If
End Sub

Private Sub txtQtdCaixas1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtPecasPorCaixa1.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtQtdCaixas2_Change()
    Call calculaQtdeProduto(txtQtdCaixas2.Text, txtPecasPorCaixa2.Text, frmExibicao7Ref.lblQtde2)
    If frmExibicao7Ref.Name = "frmExibicao7VariosProdutos" Then
        frmExibicao7Ref.lblQtdeContainers.Caption = getQtdeContainers(txtQtdCaixas1.Text, _
                                                                      txtQtdCaixas2.Text, _
                                                                      txtQtdCaixas3.Text, _
                                                                      txtQtdCaixas4.Text)
    End If
End Sub

Private Sub txtQtdCaixas2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtPecasPorCaixa2.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtQtdCaixas3_Change()
    Call calculaQtdeProduto(txtQtdCaixas3.Text, txtPecasPorCaixa3.Text, frmExibicao7Ref.lblQtde3)
    If frmExibicao7Ref.Name = "frmExibicao7VariosProdutos" Then
        frmExibicao7Ref.lblQtdeContainers.Caption = getQtdeContainers(txtQtdCaixas1.Text, _
                                                                      txtQtdCaixas2.Text, _
                                                                      txtQtdCaixas3.Text, _
                                                                      txtQtdCaixas4.Text)
    End If
End Sub

Private Sub txtQtdCaixas3_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtPecasPorCaixa3.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtQtdCaixas4_Change()
    Call calculaQtdeProduto(txtQtdCaixas4.Text, txtPecasPorCaixa4.Text, frmExibicao7Ref.lblQtde4)
    If frmExibicao7Ref.Name = "frmExibicao7VariosProdutos" Then
        frmExibicao7Ref.lblQtdeContainers.Caption = getQtdeContainers(txtQtdCaixas1.Text, _
                                                                      txtQtdCaixas2.Text, _
                                                                      txtQtdCaixas3.Text, _
                                                                      txtQtdCaixas4.Text)
    End If
End Sub

Private Sub txtQtdCaixas4_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtPecasPorCaixa4.SetFocus
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

Private Sub txtReference_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtReference_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtLicense.SetFocus
    End If
End Sub

Private Sub txtRoute_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub calculaQtdeProduto(ByVal qtdeCaixas As String, _
                               ByVal pecasPorCaixa As String, _
                               ByRef labelDestino As Label)
    If IsNumeric(qtdeCaixas) And IsNumeric(pecasPorCaixa) Then
        labelDestino.Caption = qtdeCaixas & " X " & pecasPorCaixa & " PC"
        If frmExibicao7Ref.Name = "frmExibicao7UmProduto" Then
            frmExibicao7UmProduto.lblQtdeTot1 = CLng(qtdeCaixas) * CLng(pecasPorCaixa)
        End If
    Else
        labelDestino.Caption = ""
    End If
End Sub

Private Function getQtdeContainers(ByVal qtde1 As String, _
                                   ByVal qtde2 As String, _
                                   ByVal qtde3 As String, _
                                   ByVal qtde4 As String) As Long
    Dim qtdeContainers As Long
    qtdeContainers = 0
    
    
    If (IsNumeric(qtde1)) Then
        qtdeContainers = qtdeContainers + CInt(qtde1)
    End If
    If (IsNumeric(qtde2)) Then
        qtdeContainers = qtdeContainers + CInt(qtde2)
    End If
    If (IsNumeric(qtde3)) Then
        qtdeContainers = qtdeContainers + CInt(qtde3)
    End If
    If (IsNumeric(qtde4)) Then
        qtdeContainers = qtdeContainers + CInt(qtde4)
    End If
    
    getQtdeContainers = qtdeContainers
    
End Function

