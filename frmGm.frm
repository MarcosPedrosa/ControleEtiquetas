VERSION 5.00
Begin VB.Form frmGm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Infomações da etiqueta - Modelo GM"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5280
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
      Left            =   150
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   4410
      Width           =   3825
   End
   Begin VB.TextBox txtEmbalagem 
      Height          =   300
      Left            =   120
      MaxLength       =   5
      TabIndex        =   14
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtEng 
      Height          =   300
      Left            =   120
      MaxLength       =   10
      TabIndex        =   11
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtMfgDate 
      Height          =   300
      Left            =   2760
      MaxLength       =   10
      TabIndex        =   13
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtLot 
      Height          =   300
      Left            =   1440
      MaxLength       =   13
      TabIndex        =   12
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtContainerType 
      Height          =   300
      Left            =   120
      MaxLength       =   23
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtRoute 
      Height          =   300
      Left            =   2760
      MaxLength       =   23
      TabIndex        =   10
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtLicense 
      Height          =   300
      Left            =   120
      MaxLength       =   21
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtCodMSB 
      Height          =   300
      Left            =   1440
      MaxLength       =   22
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtPeca 
      Height          =   300
      Left            =   1440
      MaxLength       =   33
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox txtMaterial 
      Height          =   300
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtReference 
      Height          =   300
      Left            =   2760
      MaxLength       =   9
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtQtd 
      Height          =   300
      Left            =   120
      MaxLength       =   10
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtPlant 
      Height          =   300
      Left            =   2760
      MaxLength       =   15
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   4110
      Picture         =   "frmGm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Nova Etiqueta"
      Height          =   855
      Left            =   4125
      Picture         =   "frmGm.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sai&r"
      Height          =   855
      Left            =   4125
      Picture         =   "frmGm.frx":0CD4
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtQtdEtiquetas 
      Height          =   300
      Left            =   3480
      MaxLength       =   3
      TabIndex        =   15
      Text            =   "1"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtPartNumber 
      Height          =   300
      Left            =   120
      MaxLength       =   12
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtGrossWeight 
      Height          =   300
      Left            =   1440
      MaxLength       =   23
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Impressora"
      Height          =   195
      Left            =   150
      TabIndex        =   44
      Top             =   4140
      Width           =   765
   End
   Begin VB.Label lblEmbalagem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Embalagem"
      Height          =   195
      Left            =   120
      TabIndex        =   43
      Top             =   3120
      Width           =   825
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lote"
      Height          =   195
      Left            =   1440
      TabIndex        =   42
      Top             =   2520
      Width           =   315
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data revisão"
      Height          =   195
      Left            =   120
      TabIndex        =   41
      Top             =   2520
      Width           =   900
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data lote"
      Height          =   195
      Left            =   2760
      TabIndex        =   40
      Top             =   2520
      Width           =   645
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2520
      TabIndex        =   39
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2520
      TabIndex        =   38
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Peso"
      Height          =   195
      Left            =   1440
      TabIndex        =   37
      Top             =   1920
      Width           =   360
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo container"
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   1920
      Width           =   1020
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rota"
      Height          =   195
      Left            =   2760
      TabIndex        =   35
      Top             =   1920
      Width           =   345
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2520
      TabIndex        =   34
      Top             =   1920
      Width           =   45
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2520
      TabIndex        =   33
      Top             =   1920
      Width           =   45
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Licença"
      Height          =   195
      Left            =   120
      TabIndex        =   32
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2520
      TabIndex        =   31
      Top             =   1320
      Width           =   45
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código Musashi"
      Height          =   195
      Left            =   1440
      TabIndex        =   30
      Top             =   1320
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
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      Height          =   195
      Left            =   1320
      TabIndex        =   28
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2520
      TabIndex        =   27
      Top             =   720
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Referência"
      Height          =   195
      Left            =   2775
      TabIndex        =   26
      Top             =   720
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde caixa"
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   720
      Width           =   765
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modelo"
      Height          =   195
      Left            =   1440
      TabIndex        =   24
      Top             =   720
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fábrica/Doca"
      Height          =   195
      Left            =   2760
      TabIndex        =   23
      Top             =   1320
      Width           =   990
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade de etiquetas a imprimir"
      Height          =   195
      Left            =   720
      TabIndex        =   22
      Top             =   3900
      Width           =   2430
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Peça"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2520
      TabIndex        =   20
      Top             =   120
      Width           =   45
   End
End
Attribute VB_Name = "frmGm"
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
    nx = 0
    
    Printer.Orientation = 2
    Printer.Copies = CInt(txtQtdEtiquetas.Text)
    
    frmExibicao5.PrintForm
    Printer.Orientation = 2: Printer.EndDoc
    
    MsgBox "Impressão concluída com sucesso! O aplicativo será encerrado!", vbOKOnly + vbInformation, "Tarefa Concluída"
    MDIEtiquetas.forcaSaida = True
    Unload MDIEtiquetas
    
End Sub

Private Sub cmdNovo_Click()
    
    txtPlant.Text = "4J"
    txtLicense.Text = "UN 897803870"
    
    txtPartNumber.Text = ""
    txtPeca.Text = ""
    txtQtd.Text = ""
    txtMaterial.Text = ""
    txtReference.Text = ""
    txtCodMSB.Text = ""
    txtContainerType.Text = ""
    txtGrossWeight.Text = ""
    txtRoute.Text = ""
    txtEng.Text = ""
    txtLot.Text = ""
    txtMfgDate.Text = ""
    
    txtQtdEtiquetas.Text = "1"
    
    txtPartNumber.SetFocus
    
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
frmExibicao5.tWithOpcoes = Me.Width
frmExibicao5.Show

txtPlant.Text = "4J"
txtLicense.Text = "UN 897803870"

txtQtdEtiquetas.Text = "1"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload frmExibicao5
End Sub

Private Sub txtCodMSB_Change()
    frmExibicao5.lblCodMSB.Caption = txtCodMSB.Text
End Sub

Private Sub txtCodMSB_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtCodMSB_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtPlant.SetFocus
    End If
End Sub

Private Sub txtContainerType_Change()
    frmExibicao5.lblContainerType.Caption = txtContainerType.Text
End Sub

Private Sub txtContainerType_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtContainerType_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtGrossWeight.SetFocus
    End If
End Sub

Private Sub txtEmbalagem_Change()
    frmExibicao5.lblEmbalagem2.Caption = txtEmbalagem.Text
End Sub

Private Sub txtEmbalagem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtQtdEtiquetas.SetFocus
    Else
        KeyAscii = KeyAsciiCriticaNumero(KeyAscii)
    End If
End Sub

Private Sub txtEng_Change()
    frmExibicao5.lblEng.Caption = txtEng.Text
End Sub

Private Sub txtEng_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtEng_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtLot.SetFocus
    End If
End Sub

Private Sub txtGrossWeight_Change()
    frmExibicao5.lblgrossWeight.Caption = txtGrossWeight
End Sub

Private Sub txtGrossWeight_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtGrossWeight_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtRoute.SetFocus
    End If
End Sub

'Private Sub txtLicense_Change()
'    frmExibicao5.lblLicense = txtLicense.Text
'
'    frmExibicao5.lblLicenseA.Caption = "*" & frmExibicao5.lblLicense & "*"
'    frmExibicao5.lblLicenseB.Caption = "*" & frmExibicao5.lblLicense & "*"
'End Sub

Private Sub txtLicense_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtLicense_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtCodMSB.SetFocus
    End If
End Sub

Private Sub txtLot_Change()
    frmExibicao5.lblLot.Caption = txtLot.Text
End Sub

Private Sub txtLot_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtLot_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtMfgDate.SetFocus
    End If
End Sub

Private Sub txtMaterial_Change()
    frmExibicao5.lblMaterial.Caption = txtMaterial.Text
End Sub

Private Sub txtMaterial_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtMaterial_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtReference.SetFocus
    End If
End Sub

Private Sub txtMfgDate_Change()
    frmExibicao5.lblMfgDate.Caption = txtMfgDate.Text
End Sub

Private Sub txtMfgDate_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtMfgDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtEmbalagem.SetFocus
    End If
End Sub

Private Sub txtPartNumber_Change()
    frmExibicao5.lblPartNumber.Caption = txtPartNumber.Text
End Sub

Private Sub txtPartNumber_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtPartNumber_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtPeca.SetFocus
    End If
End Sub

Private Sub txtPeca_Change()
    frmExibicao5.lblPeca.Caption = txtPeca.Text
End Sub

Private Sub txtPeca_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtPeca_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtQtd.SetFocus
    End If
End Sub

Private Sub txtPlant_Change()
    frmExibicao5.lblPlant.Caption = txtPlant.Text
End Sub

Private Sub txtPlant_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtPlant_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtContainerType.SetFocus
    End If
End Sub

Private Sub txtQtd_Change()
    frmExibicao5.lblQtd.Caption = txtQtd.Text
End Sub

Private Sub txtQtd_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtQtd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtMaterial.SetFocus
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

Private Sub txtReference_Change()
    frmExibicao5.lblReference.Caption = txtReference.Text
End Sub

Private Sub txtReference_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtReference_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtLicense.SetFocus
    End If
End Sub

Private Sub txtRoute_Change()
    frmExibicao5.lblRoute.Caption = txtRoute.Text
End Sub

Private Sub txtRoute_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtRoute_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txtEng.SetFocus
    End If
End Sub
