VERSION 5.00
Begin VB.Form frmExibicao10 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prévia de Impressão - YAMAHA"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   6900
      TabIndex        =   65
      Top             =   5880
      Width           =   1305
   End
   Begin VB.TextBox DataToEncodeText 
      Height          =   390
      Left            =   8250
      MaxLength       =   640
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmExibicao10.frx":0000
      Top             =   6270
      Width           =   2250
   End
   Begin VB.Label lblCodMSB_Letra 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   29.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3960
      TabIndex        =   67
      Top             =   4440
      Width           =   450
   End
   Begin VB.Label lbl_QA 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "QA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7710
      TabIndex        =   66
      Top             =   210
      Width           =   540
   End
   Begin VB.Label lbl_COD_MUSASHI_NUM 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "00001234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6090
      TabIndex        =   64
      Top             =   5070
      Width           =   885
   End
   Begin VB.Label lbl_COD_MUSASHI_BARRAS 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00001234"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   29.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5130
      TabIndex        =   63
      Top             =   4560
      Width           =   2760
   End
   Begin VB.Label lbl_MUSASHI 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "COD DA PEÇA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2880
      TabIndex        =   62
      Top             =   4500
      Width           =   1035
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   2820
      X2              =   2820
      Y1              =   5310
      Y2              =   4500
   End
   Begin VB.Label lbl_QTDE_NUM1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1125
      TabIndex        =   61
      Top             =   5070
      Width           =   345
   End
   Begin VB.Label lbl_QTDE_BARRAS 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   30
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   660
      TabIndex        =   60
      Top             =   4500
      Width           =   1350
   End
   Begin VB.Label lbl_ANO 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   29.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   6960
      TabIndex        =   59
      Top             =   3750
      Width           =   780
   End
   Begin VB.Label lbl_MES 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "JUL"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   29.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   5520
      TabIndex        =   58
      Top             =   3750
      Width           =   1260
   End
   Begin VB.Label lbl_FIFO 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "FIFO:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4800
      TabIndex        =   57
      Top             =   3810
      Width           =   510
   End
   Begin VB.Label lbl_NF_NUM 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "000065423"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2610
      TabIndex        =   56
      Top             =   4110
      Width           =   1350
   End
   Begin VB.Label lbl_NF 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NF/INVOICE:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1890
      TabIndex        =   55
      Top             =   3810
      Width           =   1125
   End
   Begin VB.Label lbl_QTDE_NUM 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   690
      TabIndex        =   54
      Top             =   4110
      Width           =   450
   End
   Begin VB.Label lbl_QTDE 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "QTDE:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   53
      Top             =   3810
      Width           =   585
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   8595
      X2              =   330
      Y1              =   4500
      Y2              =   4500
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   4740
      X2              =   4740
      Y1              =   4500
      Y2              =   3810
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   1830
      X2              =   1830
      Y1              =   4500
      Y2              =   3810
   End
   Begin VB.Label lbl_FORNECEDOR_NOME 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NISSIN BRAKE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4890
      TabIndex        =   52
      Top             =   3390
      Width           =   1800
   End
   Begin VB.Label lbl_FORNECEDOR 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "FORNECEDOR:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4800
      TabIndex        =   51
      Top             =   3090
      Width           =   1395
   End
   Begin VB.Label lbl_NOME_DESCRICAO 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PAINEL TRAS 18D (PRETO)"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   50
      Top             =   3420
      Width           =   4335
   End
   Begin VB.Label lbl_NOME 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NOME:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   49
      Top             =   3090
      Width           =   630
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   8595
      X2              =   330
      Y1              =   3810
      Y2              =   3810
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4740
      X2              =   4740
      Y1              =   3780
      Y2              =   3090
   End
   Begin VB.Label lbl_YAMAHA_COD_BARRAS 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "5089060613000001"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3510
      TabIndex        =   48
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lbl_YAMAHA_BARRAS 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5089-0606130-00001"
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   30
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2100
      TabIndex        =   47
      Top             =   2250
      Width           =   4500
   End
   Begin VB.Label lbl_USER_COD 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "9219"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7020
      TabIndex        =   46
      Top             =   1860
      Width           =   600
   End
   Begin VB.Label lbl_USER 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "USER:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   6060
      TabIndex        =   45
      Top             =   1560
      Width           =   570
   End
   Begin VB.Label lbl_SUPPLIER_COD 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "5859"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   44
      Top             =   1860
      Width           =   600
   End
   Begin VB.Label lbl_SUPPLIER 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3870
      TabIndex        =   43
      Top             =   1560
      Width           =   990
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   6000
      X2              =   6000
      Y1              =   2250
      Y2              =   1560
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   3810
      X2              =   3810
      Y1              =   2250
      Y2              =   1560
   End
   Begin VB.Label lbl_CODIGO_NUM 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "18DF5320103380"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   420
      TabIndex        =   42
      Top             =   1860
      Width           =   2145
   End
   Begin VB.Label lbl_CODIGO 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   41
      Top             =   1560
      Width           =   840
   End
   Begin VB.Label lbl_LPN_COD_B 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "5089060613000001"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3720
      TabIndex        =   40
      Top             =   1350
      Width           =   1695
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   7410
      X2              =   7410
      Y1              =   810
      Y2              =   120
   End
   Begin VB.Label lbl_LPN_COD 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "5089060613000001"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   480
      TabIndex        =   39
      Top             =   390
      Width           =   2400
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   11280
      Picture         =   "frmExibicao10.frx":0004
      Stretch         =   -1  'True
      Top             =   1380
      Width           =   2085
   End
   Begin VB.Label lblMusashi2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MUSASHI DO BRASIL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11205
      TabIndex        =   38
      Top             =   6270
      Width           =   1995
   End
   Begin VB.Label lblPartNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00001234"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1140
      TabIndex        =   37
      Top             =   7125
      Width           =   2640
   End
   Begin VB.Label lblLicenseB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UN 123456789 A2B4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1050
      TabIndex        =   36
      Top             =   8670
      Width           =   4080
   End
   Begin VB.Label lblQtd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
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
      Left            =   690
      TabIndex        =   35
      Top             =   6570
      Width           =   510
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   8595
      X2              =   330
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblMaterial 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "02W C32"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3750
      TabIndex        =   34
      Top             =   6630
      Width           =   2325
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   8595
      X2              =   330
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label lblQtd2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY :"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9060
      TabIndex        =   33
      Top             =   1530
      Width           =   555
   End
   Begin VB.Label lblMaterial2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MATERIAL HANDLING CODE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8730
      TabIndex        =   32
      Top             =   1890
      Width           =   1365
   End
   Begin VB.Label lblReference 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "F734"
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
      Left            =   6360
      TabIndex        =   31
      Top             =   7290
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lbl_LPN 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "LPN:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   360
      TabIndex        =   30
      Top             =   105
      Width           =   435
   End
   Begin VB.Label lblEndereco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRAÇA MOTOGEAR 111"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11205
      TabIndex        =   29
      Top             =   6480
      Width           =   2235
   End
   Begin VB.Label lblIgarassu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IGARASSU"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11205
      TabIndex        =   28
      Top             =   6690
      Width           =   1005
   End
   Begin VB.Label lblTelefone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "81 35436000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11205
      TabIndex        =   27
      Top             =   6885
      Width           =   1110
   End
   Begin VB.Label lblBrasil 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MADE IN BRAZIL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   11205
      TabIndex        =   26
      Top             =   7095
      Width           =   1530
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   6270
      X2              =   6270
      Y1              =   810
      Y2              =   120
   End
   Begin VB.Label lblTo2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TO:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9090
      TabIndex        =   25
      Top             =   2190
      Width           =   165
   End
   Begin VB.Label lblPlant2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PLANT/DOCK"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10365
      TabIndex        =   24
      Top             =   3420
      Width           =   630
   End
   Begin VB.Label lblPeca 
      BackStyle       =   0  'Transparent
      Caption         =   "Engrenagem"
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
      Left            =   5460
      TabIndex        =   23
      Top             =   6240
      Width           =   2265
   End
   Begin VB.Label lblParts2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PARTS GOOD UP TO:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9090
      TabIndex        =   22
      Top             =   4980
      Width           =   1020
   End
   Begin VB.Label lblLote2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "LOT NO:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9090
      TabIndex        =   21
      Top             =   4725
      Width           =   405
   End
   Begin VB.Label lblEng 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "29MAR1998"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9690
      TabIndex        =   20
      Top             =   4440
      Width           =   1065
   End
   Begin VB.Label lblMfgDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "01APR1999"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9690
      TabIndex        =   19
      Top             =   5190
      Width           =   1065
   End
   Begin VB.Label lblParts 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "01MAY2005"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   9720
      TabIndex        =   18
      Top             =   4890
      Width           =   1065
   End
   Begin VB.Label lblLot 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0987654321"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2880
      TabIndex        =   17
      Top             =   4980
      Width           =   1050
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   8595
      X2              =   330
      Y1              =   3090
      Y2              =   3090
   End
   Begin VB.Label lblTo 
      BackStyle       =   0  'Transparent
      Caption         =   "GENERAL MOTORS DO"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10680
      TabIndex        =   16
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label lblPlant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "B1-RRD-22W"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10635
      TabIndex        =   15
      Top             =   3540
      Width           =   2250
   End
   Begin VB.Label lblCodMSB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2G01740"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2880
      TabIndex        =   14
      Top             =   4710
      Width           =   960
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   8610
      X2              =   330
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lbl_LPN_COD_BARRA 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5089060613000001"
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   27.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2970
      TabIndex        =   13
      Top             =   810
      Width           =   3210
   End
   Begin VB.Label lblRoute 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "05JUN"
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
      Left            =   6300
      TabIndex        =   12
      Top             =   7995
      Width           =   1020
   End
   Begin VB.Label lblgrossWeight 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "18 KG"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6300
      TabIndex        =   11
      Top             =   8895
      Width           =   465
   End
   Begin VB.Label lblContainerType 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "KLT3214"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6300
      TabIndex        =   10
      Top             =   8550
      Width           =   675
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   5205
      Left            =   330
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label lblEmbalagem2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   7920
      TabIndex        =   9
      Top             =   5550
      Width           =   45
   End
   Begin VB.Label lblCodigoBarras 
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
      Left            =   1650
      TabIndex        =   8
      Top             =   9270
      Width           =   2880
   End
   Begin VB.Label lblCodigoBarrasA 
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
      Left            =   660
      TabIndex        =   7
      Top             =   9660
      Width           =   5250
   End
   Begin VB.Label lblCodigoBarrasB 
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
      Left            =   660
      TabIndex        =   6
      Top             =   9900
      Width           =   5250
   End
   Begin VB.Label lblMotivo_alteracao_outros 
      BackStyle       =   0  'Transparent
      Caption         =   "SAO CAETANO DO SUL,WW"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10650
      TabIndex        =   5
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label lblvalidade 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4890
      TabIndex        =   4
      Top             =   6060
      Width           =   360
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BRASIL LTDA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10650
      TabIndex        =   3
      Top             =   2820
      Width           =   2775
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
      Left            =   300
      TabIndex        =   2
      Top             =   5370
      Width           =   900
   End
   Begin VB.Label lblRoute1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "2017"
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
      Left            =   7470
      TabIndex        =   1
      Top             =   8070
      Width           =   540
   End
End
Attribute VB_Name = "frmExibicao10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim v As Integer
Dim sData As String
Dim nx As Double

Dim x As Printer
               
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
   Exit Sub
End If
nx = 0


Me.Command1.Visible = False
Printer.Orientation = 2
Me.PrintForm
Printer.Orientation = 2: Printer.EndDoc
Me.Command1.Visible = True

End Sub

Private Sub DataToEncodeText_Change()
Call CodeRefresh(DataToEncodeText.Text)
'Me.lblLicenseA.Caption = sCodigo
'Me.lblLicense.Caption = sCodigo

End Sub

Private Sub Form_Load()
Dim v As Integer

Me.lbl_LPN_COD.Caption = ""
Me.lbl_LPN_COD_BARRA.Caption = ""
Me.lbl_LPN_COD_B.Caption = ""
Me.lbl_CODIGO_NUM.Caption = ""
Me.lbl_SUPPLIER_COD.Caption = ""
Me.lbl_USER_COD.Caption = ""
Me.lbl_YAMAHA_BARRAS.Caption = ""
Me.lbl_YAMAHA_COD_BARRAS.Caption = ""
Me.lbl_NOME_DESCRICAO.Caption = ""
Me.lbl_FORNECEDOR_NOME.Caption = ""
Me.lbl_QTDE_NUM.Caption = ""
Me.lbl_NF_NUM.Caption = ""
Me.LBL_MES.Caption = ""
Me.lbl_ANO.Caption = ""
Me.lbl_QTDE_BARRAS.Caption = ""
Me.lbl_QTDE_NUM1.Caption = ""
Me.lbl_QA.Caption = ""
Me.lblCodMSB_Letra.Caption = ""

'Me.Height = 6630
'Me.Width = 9780
Me.Top = 0
Me.Left = frmOpcoes.Width + 100

'For v = 0 To (Forms.Count - 1)
'    If Forms(v).Name = "frmExibicao10" Then
'        Me.Left = frmExibicao10.Width
'        Exit For
'    End If
'Next

Me.lbl_data.Caption = Format(Now(), "dd/mm/yyyy")

End Sub
