VERSION 5.00
Begin VB.Form frmExibicaoHondaNova 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prévia de impressão - HONDA (Nova)"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_mome_impressora 
      Height          =   345
      Left            =   1650
      MaxLength       =   50
      TabIndex        =   42
      Text            =   "Name Impressora"
      Top             =   6960
      Width           =   5685
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   240
      TabIndex        =   28
      Top             =   3870
      Width           =   1155
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   60
         ScaleHeight     =   47
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   60
         TabIndex        =   29
         Top             =   150
         Width           =   900
      End
   End
   Begin VB.TextBox Text2 
      Height          =   435
      Left            =   1650
      MaxLength       =   700
      MultiLine       =   -1  'True
      TabIndex        =   27
      Top             =   6360
      Width           =   5685
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   6360
      ScaleHeight     =   116
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   120
      TabIndex        =   25
      Top             =   1290
      Width           =   1800
   End
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   1680
      MaxLength       =   700
      MultiLine       =   -1  'True
      TabIndex        =   24
      Top             =   5820
      Width           =   5685
   End
   Begin VB.CommandButton cmd_imprime 
      Caption         =   "&Imprime"
      Height          =   255
      Left            =   5610
      TabIndex        =   23
      ToolTipText     =   "Imprime esta etiqueta"
      Top             =   5010
      Width           =   1335
   End
   Begin VB.Label lbl_Seq_Milhar 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2122"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6930
      TabIndex        =   41
      Top             =   4440
      Width           =   1095
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "LOTE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4770
      TabIndex        =   40
      Top             =   3540
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PEÇA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2310
      TabIndex        =   39
      Top             =   3540
      Width           =   540
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DTE.. :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6450
      TabIndex        =   38
      Top             =   3960
      Width           =   630
   End
   Begin VB.Label lbl_data_etiq 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7110
      TabIndex        =   37
      Top             =   3990
      Width           =   1245
   End
   Begin VB.Label lbl_embalador 
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7110
      TabIndex        =   36
      Top             =   3660
      Width           =   1245
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "EMB. :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6450
      TabIndex        =   35
      Top             =   3630
      Width           =   615
   End
   Begin VB.Label lbl_cod_barras 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2610
      TabIndex        =   34
      Top             =   5070
      Width           =   2880
   End
   Begin VB.Label lbl_cod_barras1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "00001864"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1440
      TabIndex        =   33
      Top             =   4500
      Width           =   5250
   End
   Begin VB.Label lbl_cod_barras2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "00001864"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1440
      TabIndex        =   32
      Top             =   4740
      Width           =   5250
   End
   Begin VB.Label lbl_lote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0014869256"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3690
      TabIndex        =   31
      Top             =   3750
      Width           =   2700
   End
   Begin VB.Label lbl_peca 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2H70660"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1575
      TabIndex        =   30
      Top             =   3750
      Width           =   1965
   End
   Begin VB.Label lbl_Codigo_Musashi 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BR999990"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   4
      Top             =   5010
      Width           =   900
   End
   Begin VB.Label lbl_Cod_Qrcode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "BR999990"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6960
      TabIndex        =   26
      Top             =   3120
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DESTINO"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4260
      TabIndex        =   22
      Top             =   2820
      Width           =   615
   End
   Begin VB.Label Lbl_Destino 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CD6"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4260
      TabIndex        =   21
      Top             =   3000
      Width           =   465
   End
   Begin VB.Label lbl_Fornecedor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MUSASHI DO BRASIL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   20
      Top             =   2940
      Width           =   2940
   End
   Begin VB.Label lbl_Embalagem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IK33"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4260
      TabIndex        =   19
      Top             =   2220
      Width           =   495
   End
   Begin VB.Label Lbl_tipo_peca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PC"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3240
      TabIndex        =   18
      Top             =   2130
      Width           =   465
   End
   Begin VB.Label lbl_Sequencial 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "001/036"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7050
      TabIndex        =   17
      Top             =   540
      Width           =   885
   End
   Begin VB.Label lbl_Data_Entrega 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "23/08/2018"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3765
      TabIndex        =   16
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label lbl_Nota_Fiscal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000109158-1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      TabIndex        =   15
      Top             =   510
      Width           =   1905
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NOTA FISCAL"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1230
      TabIndex        =   14
      Top             =   270
      Width           =   960
   End
   Begin VB.Label lblCodFunc 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "14490"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1590
      TabIndex        =   13
      Top             =   5460
      Width           =   450
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTIDADE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   270
      TabIndex        =   12
      Top             =   1920
      Width           =   945
   End
   Begin VB.Label lbl_Empresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   180
      TabIndex        =   11
      Top             =   510
      Width           =   555
   End
   Begin VB.Label lblFrom2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   10
      Top             =   270
      Width           =   690
   End
   Begin VB.Label lblTo2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA ENTREGA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3810
      TabIndex        =   9
      Top             =   270
      Width           =   1110
   End
   Begin VB.Label lblitem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   8
      Top             =   870
      Width           =   330
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   8370
      X2              =   150
      Y1              =   3540
      Y2              =   3555
   End
   Begin VB.Label lbl_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "72479  A215"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   210
      TabIndex        =   7
      Top             =   1020
      Width           =   2610
   End
   Begin VB.Label lbl_Descricao_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PARAF.FLANGE 10X235.5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   6
      Top             =   1590
      Width           =   2670
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INFORMAÇÕES ADICIONAIS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   270
      TabIndex        =   5
      Top             =   3330
      Width           =   1965
   End
   Begin VB.Label lbl_Quantidade 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "999999999"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   240
      TabIndex        =   3
      Top             =   2100
      Width           =   2160
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "EMBALAGEM"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4230
      TabIndex        =   2
      Top             =   2070
      Width           =   930
   End
   Begin VB.Label lbl_data 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "dd/mm/yyyy"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   5460
      Width           =   900
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000007&
      BorderWidth     =   2
      Height          =   5145
      Left            =   120
      Top             =   180
      Width           =   8295
   End
   Begin VB.Label lblLicense2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "FORNECEDOR"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   240
      TabIndex        =   0
      Top             =   2730
      Width           =   1005
   End
End
Attribute VB_Name = "frmExibicaoHondaNova"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sCodigo As String
Private cQrCode As ClsQrCode
Public sCodQrcode1 As String
Public sCodQrcode2 As String

Private Sub cmd_imprime_Click()
Dim x As Printer
Dim nx As Integer

nx = 0
For Each x In Printers
   If InStr(1, UCase(x.DeviceName), UCase(Me.txt_mome_impressora.Text)) > 0 Then
      Set Printer = x
      nx = 1
      Exit For
   End If
Next

If nx = 0 Then
   MsgBox "Impressora da Etiqueta não encontrada, chame o responsável. o nome da impressora - " & UCase(Me.txt_mome_impressora.Text)
   Exit Sub
End If
nx = 0

Me.cmd_imprime.Visible = False
Printer.Orientation = 2
Me.PrintForm
Printer.Orientation = 2: Printer.EndDoc
Me.cmd_imprime.Visible = True
Unload Me
End Sub

Private Sub Command2_Click()
    Call Calcular_Qrcode
End Sub
Public Function Calcular_Qrcode()

    Set cQrCode = New ClsQrCode

    Picture1.Picture = cQrCode.GetPictureQrCode(Me.Text1.Text, Picture1.ScaleWidth, Picture1.ScaleHeight)
    If Picture1.Picture Is Nothing Then MsgBox "Error!"
    Picture1.Picture = cQrCode.GetPictureQrCode(Me.Text1.Text, 120, 120, "UTF-8", "L", vbBlack, vbWhite, 3)

    Picture2.Picture = cQrCode.GetPictureQrCode(Me.Text2.Text, Picture2.ScaleWidth, Picture2.ScaleHeight)
    If Picture2.Picture Is Nothing Then MsgBox "Error!"
    Picture2.Picture = cQrCode.GetPictureQrCode(Me.Text2.Text, 70, 70, "UTF-8", "L", vbBlack, vbWhite, 3)
End Function

Private Sub Form_Load()

    'Limpar Campos
    
    lbl_Empresa.Caption = ""
    lbl_Nota_Fiscal.Caption = ""
    lbl_Data_Entrega.Caption = ""
    lbl_Sequencial.Caption = ""
    lbl_Item.Caption = ""
    lbl_Descricao_Item.Caption = ""
    lbl_Quantidade.Caption = ""

    lbl_Embalagem.Caption = ""
    lbl_Fornecedor.Caption = ""
    Lbl_Destino.Caption = ""
    lbl_Codigo_Musashi.Caption = ""
    lbl_Cod_Qrcode.Caption = ""
    lbl_lote.Caption = ""
    lbl_peca.Caption = ""
    lbl_cod_barras.Caption = ""
    lbl_cod_barras1.Caption = ""
    lbl_cod_barras2.Caption = ""
    lbl_Seq_Milhar.Caption = ""
    
    Me.Top = 0
    Me.Left = frmEtiquetaHondaQrcode.Width
    Me.lbl_data.Caption = Format(Now(), "dd/mm/yyyy")
    Me.Text1.Text = ""
    Me.Text2.Text = ""
'****************************************************************
    
End Sub
