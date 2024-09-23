VERSION 5.00
Begin VB.Form frmExibicao41 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Etiqueta SHOWA"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_Fechar 
      BackColor       =   &H000000FF&
      Caption         =   "Fechar"
      Height          =   285
      Left            =   60
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   4230
      Width           =   705
   End
   Begin VB.CommandButton cmd_imprime 
      BackColor       =   &H0080FF80&
      Caption         =   "Imprime"
      Height          =   285
      Left            =   7710
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   4230
      Width           =   885
   End
   Begin VB.Label lbl_Cod_Barra_Cnpj_Pedido3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "11956981000161000023"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   720
      TabIndex        =   32
      Top             =   1050
      Width           =   3300
   End
   Begin VB.Label lbl_notificacao 
      AutoSize        =   -1  'True
      Caption         =   "Serão impressas xxxx etiquetas de xx produtos diferentes,confirme a impressão."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   840
      TabIndex        =   30
      Top             =   4260
      Width           =   6765
   End
   Begin VB.Label lbl_Cod_Barra_Nf_Serie_DtEmissao3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "221002a1120112"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5880
      TabIndex        =   29
      Top             =   1080
      Width           =   2310
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6600
      Picture         =   "frmExibicao41.frx":0000
      Stretch         =   -1  'True
      Top             =   3750
      Width           =   1515
   End
   Begin VB.Label lblPeso2 
      BackStyle       =   0  'Transparent
      Caption         =   "3000/2500"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   750
      TabIndex        =   27
      Top             =   3930
      Width           =   2490
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PESO BRUTO/LIQ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   750
      TabIndex        =   26
      Top             =   3780
      Width           =   1320
   End
   Begin VB.Label lbl_volume 
      BackStyle       =   0  'Transparent
      Caption         =   "75/100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5730
      TabIndex        =   25
      Top             =   3510
      Width           =   960
   End
   Begin VB.Label lbl_Dt_Fab_Validade 
      BackStyle       =   0  'Transparent
      Caption         =   "XX-XX-XX / XX-XX-XX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   750
      TabIndex        =   24
      Top             =   3510
      Width           =   3720
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VOLUME"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5730
      TabIndex        =   23
      Top             =   3360
      Width           =   690
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA FAB/VALIDADE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   750
      TabIndex        =   22
      Top             =   3360
      Width           =   1665
   End
   Begin VB.Label lblCodigoBarras1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1234567890123456789012345"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   720
      TabIndex        =   21
      Top             =   3000
      Width           =   6375
   End
   Begin VB.Label lblCodigoBarras 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1234567890123456789012345"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   720
      TabIndex        =   20
      Top             =   2760
      Width           =   6375
   End
   Begin VB.Label lblLote2 
      BackStyle       =   0  'Transparent
      Caption         =   "103347/12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5310
      TabIndex        =   19
      Top             =   2460
      Width           =   2940
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4740
      TabIndex        =   18
      Top             =   2520
      Width           =   420
   End
   Begin VB.Label lblQtd2 
      BackStyle       =   0  'Transparent
      Caption         =   "279000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   17
      Top             =   2460
      Width           =   1560
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QTDE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   750
      TabIndex        =   16
      Top             =   2520
      Width           =   420
   End
   Begin VB.Label lblDescricao 
      BackStyle       =   0  'Transparent
      Caption         =   "MOLA DIANTEIRA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   720
      TabIndex        =   15
      Top             =   2160
      Width           =   7560
   End
   Begin VB.Label lblCodigoBarrasB 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1234567890123456789012345"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   690
      TabIndex        =   14
      Top             =   1800
      Width           =   7125
   End
   Begin VB.Label lblCodigoBarrasA 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1234567890123456789012345"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   690
      TabIndex        =   13
      Top             =   1620
      Width           =   7125
   End
   Begin VB.Label lblPeca 
      BackStyle       =   0  'Transparent
      Caption         =   "HKRE2-380-00-BR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   12
      Top             =   1290
      Width           =   6840
   End
   Begin VB.Label lbl_Cod_Barra_Nf_Serie_DtEmissao2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "221002a1120112"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5880
      TabIndex        =   11
      Top             =   930
      Width           =   2310
   End
   Begin VB.Label lbl_Cod_Barra_Nf_Serie_DtEmissao1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "221002a1120112"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5880
      TabIndex        =   10
      Top             =   720
      Width           =   2310
   End
   Begin VB.Label lbl_Cod_Barra_Cnpj_Pedido2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "11956981000161000023"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   720
      TabIndex        =   9
      Top             =   930
      Width           =   3300
   End
   Begin VB.Label lbl_data_emissao 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "12/01/12"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7140
      TabIndex        =   8
      Top             =   480
      Width           =   945
   End
   Begin VB.Label lbl_pedido 
      BackStyle       =   0  'Transparent
      Caption         =   "334701"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6150
      TabIndex        =   7
      Top             =   480
      Width           =   945
   End
   Begin VB.Label lbl_Nota_Fiscal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "221002"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4920
      TabIndex        =   6
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA EMISSÃO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   7200
      TabIndex        =   5
      Top             =   330
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PEDIDO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   6150
      TabIndex        =   4
      Top             =   330
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTA FISCAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   5190
      TabIndex        =   3
      Top             =   330
      Width           =   780
   End
   Begin VB.Label lblCliente 
      BackStyle       =   0  'Transparent
      Caption         =   "V e M DO BRASIL S.A."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   3420
   End
   Begin VB.Label lbl_Cod_Barra_Cnpj_Pedido1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "11956981000161000023"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   720
      TabIndex        =   1
      Top             =   720
      Width           =   3300
   End
   Begin VB.Label lbl_fornecedor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FORNECEDOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   720
      TabIndex        =   0
      Top             =   330
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   3885
      Left            =   510
      Top             =   300
      Width           =   7905
   End
End
Attribute VB_Name = "frmExibicao41"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private nQtde_Etiquetas As Integer
Private nQtde_Imprime As Integer

Rem arquivo para movimentacao de etiquetas do cliente showa

Private Sub cmd_Fechar_Click()
Close #11: End
End Sub

Private Sub cmd_imprime_Click()
Dim nx As Integer
Dim x As Printer
Dim y As Integer
Dim nRet As Integer

Me.cmd_imprime.Visible = False

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
   End:
End If
nx = 0

Printer.Orientation = 2
Me.cmd_imprime.Visible = False
Me.cmd_Fechar.Visible = False
Me.lbl_notificacao.Visible = False

For y = 1 To gUltimoRegistro41
     Get #gNumeroArquivo41, y, Arq_Mov_Showa
     nQtde_Etiquetas = Val(Arq_Mov_Showa.FVolume)
     
     For nx = 1 To nQtde_Etiquetas
     
         Me.lblCliente.Caption = Arq_Mov_Showa.FNome_Forenecedor
         Me.lbl_Nota_Fiscal.Caption = Format(Arq_Mov_Showa.FNum_nf, "######000")
         Me.lbl_pedido.Caption = Format(Arq_Mov_Showa.FNum_Pedido, "###000")
         Me.lbl_data_emissao.Caption = Mid$(Arq_Mov_Showa.FDt_Emissao, 1, 2) & "/" & Mid$(Arq_Mov_Showa.FDt_Emissao, 3, 2) & "/" & Mid$(Arq_Mov_Showa.FDt_Emissao, 5, 2)
         Me.lbl_Cod_Barra_Nf_Serie_DtEmissao1.Caption = "*" & Arq_Mov_Showa.FCod_Barra_NF & "*"
         Me.lbl_Cod_Barra_Nf_Serie_DtEmissao2.Caption = "*" & Arq_Mov_Showa.FCod_Barra_NF & "*"
         Me.lbl_Cod_Barra_Nf_Serie_DtEmissao3.Caption = "*" & Arq_Mov_Showa.FCod_Barra_NF & "*"
         Me.lbl_Cod_Barra_Cnpj_Pedido1.Caption = "*" & Arq_Mov_Showa.FCod_Barra_CNPJ_PED & "*"
         Me.lbl_Cod_Barra_Cnpj_Pedido2.Caption = "*" & Arq_Mov_Showa.FCod_Barra_CNPJ_PED & "*"
         Me.lbl_Cod_Barra_Cnpj_Pedido3.Caption = "*" & Arq_Mov_Showa.FCod_Barra_CNPJ_PED & "*"
         Me.lblPeca.Caption = Arq_Mov_Showa.FCod_Barra_MATERIAL
         Me.lblCodigoBarrasA.Caption = "*" & Arq_Mov_Showa.FCod_Barra_MATERIAL & "*"
         Me.lblCodigoBarrasB.Caption = "*" & Arq_Mov_Showa.FCod_Barra_MATERIAL & "*"
         Me.lblDescricao.Caption = Arq_Mov_Showa.FDescricao_Produto
         Me.lblQtd2.Caption = Arq_Mov_Showa.FQuantidade
         Me.lblLote2.Caption = Arq_Mov_Showa.FLote
         Me.lblCodigoBarras.Caption = "*" & Arq_Mov_Showa.FCod_Barra_Qtde_Lote & "*"
         Me.lblCodigoBarras1.Caption = "*" & Arq_Mov_Showa.FCod_Barra_Qtde_Lote & "*"
         Me.lblPeso2.Caption = Arq_Mov_Showa.FPeso_bruto & "/" & _
                               Arq_Mov_Showa.FPeso_liquido
         Me.lbl_volume.Caption = Format(nx, "00") & "/" & Arq_Mov_Showa.FVolume
         nQtde_Etiquetas = Val(Arq_Mov_Showa.FVolume)
         frmExibicao41.PrintForm
     Next
     
'     nRet = MsgBox("Confirma SAIDA?", vbQuestion & vbYesNo, Me.Caption)
'     'Se confirmou
'     If nRet = 6 Then GoTo SAIDA

Next


saida:

Printer.Orientation = 2: Printer.EndDoc
Close gNumeroArquivo41
Kill objApplication.caminhoImportacao & "\etiqshowa.txt"

MsgBox "Impressão concluída com sucesso! ", vbOKOnly + vbInformation, "Tarefa Concluída"

End:

End Sub

Private Sub Form_Load()
Me.Top = 0
Dim y As Integer
Dim x As Integer

'Close 11
'Open "C:\Sistemas\mussashi\Desenvolvimento\Fontes\MD_Etiquetas\etiqshowa.txt" For Random Access Read Write As #11 Len = Len(Arq_Mov_Showa)
nQtde_Etiquetas = 0

x = gUltimoRegistro41

nQtde_Imprime = 0

For y = 1 To x
    Get #gNumeroArquivo41, y, Arq_Mov_Showa
     
    nQtde_Etiquetas = nQtde_Etiquetas + Val(Arq_Mov_Showa.FVolume)
    
    If y = 1 Then ' so preencher uma etiqueta na tela
       Me.lblCliente.Caption = Arq_Mov_Showa.FNome_Forenecedor
       Me.lbl_Nota_Fiscal.Caption = Format(Arq_Mov_Showa.FNum_nf, "######000")
       Me.lbl_pedido.Caption = Format(Arq_Mov_Showa.FNum_Pedido, "###000")
       Me.lbl_data_emissao.Caption = Mid$(Arq_Mov_Showa.FDt_Emissao, 1, 2) & "/" & Mid$(Arq_Mov_Showa.FDt_Emissao, 3, 2) & "/" & Mid$(Arq_Mov_Showa.FDt_Emissao, 5, 2)
       Me.lbl_Cod_Barra_Nf_Serie_DtEmissao1.Caption = Arq_Mov_Showa.FCod_Barra_NF
       Me.lbl_Cod_Barra_Nf_Serie_DtEmissao2.Caption = Arq_Mov_Showa.FCod_Barra_NF
       Me.lbl_Cod_Barra_Nf_Serie_DtEmissao3.Caption = Arq_Mov_Showa.FCod_Barra_NF
       Me.lbl_Cod_Barra_Cnpj_Pedido1.Caption = Arq_Mov_Showa.FCod_Barra_CNPJ_PED
       Me.lbl_Cod_Barra_Cnpj_Pedido2.Caption = Arq_Mov_Showa.FCod_Barra_CNPJ_PED
       Me.lbl_Cod_Barra_Cnpj_Pedido3.Caption = Arq_Mov_Showa.FCod_Barra_CNPJ_PED
       Me.lblPeca.Caption = Arq_Mov_Showa.FCod_Barra_MATERIAL
       Me.lblCodigoBarrasA.Caption = Arq_Mov_Showa.FCod_Barra_MATERIAL
       Me.lblCodigoBarrasB.Caption = Arq_Mov_Showa.FCod_Barra_MATERIAL
       Me.lblDescricao.Caption = Arq_Mov_Showa.FDescricao_Produto
       Me.lblQtd2.Caption = Arq_Mov_Showa.FQuantidade
       Me.lblLote2.Caption = Arq_Mov_Showa.FLote
       Me.lblCodigoBarras.Caption = Arq_Mov_Showa.FCod_Barra_Qtde_Lote
       Me.lblCodigoBarras1.Caption = Arq_Mov_Showa.FCod_Barra_Qtde_Lote
       Me.lblPeso2.Caption = Arq_Mov_Showa.FPeso_bruto & "/" & _
                             Arq_Mov_Showa.FPeso_liquido
       Me.lbl_volume.Caption = "01/" & Arq_Mov_Showa.FVolume
    End If
     
Next

Me.lbl_notificacao.Caption = "Serão impressas " & Format(nQtde_Etiquetas, "000") & " etiquetas de " & Format(x, "00") & " produtos diferentes,confirme a impressão?"

End Sub

