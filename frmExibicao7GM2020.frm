VERSION 5.00
Begin VB.Form frmExibicao7GM2020 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prévia de impressão - GM"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox DataToEncodeText3 
      Height          =   345
      Left            =   780
      TabIndex        =   40
      Text            =   "Text1"
      Top             =   6450
      Width           =   6465
   End
   Begin VB.TextBox DataToEncodeText2 
      Height          =   345
      Left            =   750
      TabIndex        =   39
      Text            =   "Text1"
      Top             =   6030
      Width           =   6465
   End
   Begin VB.Label lblLote 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
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
      Left            =   3870
      TabIndex        =   47
      Top             =   4560
      Width           =   150
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NUM. LOTE :"
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
      Left            =   4080
      TabIndex        =   46
      Top             =   4410
      Width           =   600
   End
   Begin VB.Label lblCodigoBarrasA 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "123456789hgf"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   45
      Top             =   4920
      Width           =   5175
   End
   Begin VB.Label lblCodigoBarrasB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "123456789hgf"
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
      Left            =   0
      TabIndex        =   44
      Top             =   5070
      Width           =   5175
   End
   Begin VB.Label lblCodigoBarras 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "12345"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2310
      TabIndex        =   43
      Top             =   5340
      Width           =   540
   End
   Begin VB.Label lblshipdate 
      Caption         =   "Label5"
      Height          =   165
      Left            =   3990
      TabIndex        =   42
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label lblDescricao 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   195
      Left            =   4080
      TabIndex        =   41
      Top             =   7080
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   735
      Left            =   270
      Stretch         =   -1  'True
      Top             =   3270
      Width           =   4755
   End
   Begin VB.Label lblRoute 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
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
      Left            =   5760
      TabIndex        =   38
      Top             =   4860
      Width           =   90
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "STK ="
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
      Left            =   5250
      TabIndex        =   37
      Top             =   4860
      Width           =   465
   End
   Begin VB.Label lblExpDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6030
      TabIndex        =   36
      Top             =   5160
      Width           =   165
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "EXP. DATE :"
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
      Left            =   5250
      TabIndex        =   35
      Top             =   5280
      Width           =   585
   End
   Begin VB.Label lblShipmentAno 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "2006"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7740
      TabIndex        =   34
      Top             =   3450
      Width           =   660
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   3150
      X2              =   3150
      Y1              =   2400
      Y2              =   1680
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "REFERENCE :"
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
      Left            =   1590
      TabIndex        =   33
      Top             =   1680
      Width           =   660
   End
   Begin VB.Label lblLicense2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "LICENSE PLATE(1J)"
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
      Left            =   180
      TabIndex        =   32
      Top             =   3060
      Width           =   945
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "SAO JOSE DOS CAMPOS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2700
      TabIndex        =   31
      Top             =   930
      Width           =   2505
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
      TabIndex        =   30
      Top             =   5610
      Width           =   900
   End
   Begin VB.Label lblMaterial 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "02W C32"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   38.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   960
      TabIndex        =   29
      Top             =   2310
      Width           =   2595
   End
   Begin VB.Label lblTo 
      BackStyle       =   0  'Transparent
      Caption         =   "GENERAL MOTORS DO"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2700
      TabIndex        =   28
      Top             =   450
      Width           =   2385
   End
   Begin VB.Label lblTo1 
      BackStyle       =   0  'Transparent
      Caption         =   "BRASIL LTDA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2700
      TabIndex        =   27
      Top             =   690
      Width           =   2775
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MATERIAL HANDLING CODE :"
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
      Left            =   3180
      TabIndex        =   26
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NUMBER"
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
      Left            =   180
      TabIndex        =   25
      Top             =   2550
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PART"
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
      Left            =   180
      TabIndex        =   24
      Top             =   2430
      Width           =   270
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000007&
      BorderWidth     =   2
      Height          =   5325
      Left            =   150
      Top             =   300
      Width           =   8475
   End
   Begin VB.Label lblComplPeca1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G1155"
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
      Left            =   1605
      TabIndex        =   23
      Top             =   1800
      Width           =   1380
   End
   Begin VB.Label lblQtde1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   180
      TabIndex        =   22
      Top             =   1710
      Width           =   660
   End
   Begin VB.Label lblCodigoProduto1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   960
      TabIndex        =   21
      Top             =   4380
      Width           =   2550
   End
   Begin VB.Label lblContainer2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "FIFO DATE:"
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
      Left            =   5250
      TabIndex        =   20
      Top             =   2460
      Width           =   555
   End
   Begin VB.Label lblShipmentDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "12/JAN"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   5370
      TabIndex        =   19
      Top             =   2490
      Width           =   2820
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód. MUSASHI"
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
      Left            =   210
      TabIndex        =   18
      Top             =   4440
      Width           =   705
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5190
      X2              =   120
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   1530
      X2              =   1530
      Y1              =   2430
      Y2              =   1680
   End
   Begin VB.Label lblCodMSB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2G01740"
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
      Left            =   3180
      TabIndex        =   17
      Top             =   1800
      Width           =   1890
   End
   Begin VB.Label lblPlant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "72479  A21"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3450
      TabIndex        =   16
      Top             =   1110
      Width           =   1635
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   8580
      X2              =   120
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Label lblPlant2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PLANT/DOCK:"
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
      Left            =   2685
      TabIndex        =   15
      Top             =   1230
      Width           =   690
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
      Left            =   2685
      TabIndex        =   14
      Top             =   285
      Width           =   165
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   2640
      X2              =   2640
      Y1              =   1650
      Y2              =   300
   End
   Begin VB.Label lblBrasil 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MADE IN BRAZIL"
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
      Left            =   195
      TabIndex        =   13
      Top             =   1365
      Width           =   1350
   End
   Begin VB.Label lblTelefone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "81 35436000"
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
      Left            =   195
      TabIndex        =   12
      Top             =   1125
      Width           =   945
   End
   Begin VB.Label lblIgarassu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Novelino 111 - IGARASSU"
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
      Left            =   195
      TabIndex        =   11
      Top             =   900
      Width           =   2040
   End
   Begin VB.Label lblEndereco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Av. Antonio Vicente "
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
      Left            =   195
      TabIndex        =   10
      Top             =   690
      Width           =   1590
   End
   Begin VB.Label lblFrom2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "FROM:"
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
      Left            =   180
      TabIndex        =   9
      Top             =   285
      Width           =   315
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   5160
      X2              =   150
      Y1              =   3060
      Y2              =   3060
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   8580
      X2              =   120
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   5190
      X2              =   5190
      Y1              =   1680
      Y2              =   5610
   End
   Begin VB.Label lblMusashi2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MUSASHI DO BRASIL"
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
      Left            =   195
      TabIndex        =   8
      Top             =   480
      Width           =   1755
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY :"
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
      Left            =   180
      TabIndex        =   7
      Top             =   1680
      Width           =   570
   End
   Begin VB.Label lblGross2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "GROSS WEIGHT:"
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
      Left            =   7350
      TabIndex        =   6
      Top             =   3930
      Width           =   795
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CONTAINER TYPE:"
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
      Left            =   5220
      TabIndex        =   5
      Top             =   3900
      Width           =   900
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
      Left            =   5340
      TabIndex        =   4
      Top             =   4080
      Width           =   675
   End
   Begin VB.Label lblgrossWeight 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "528,23 KG"
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
      Left            =   7350
      TabIndex        =   3
      Top             =   4140
      Width           =   780
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
      TabIndex        =   2
      Top             =   5610
      Width           =   450
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " DELIVERY NOTE or PUS OR INVOICE NUMBER :"
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
      Left            =   5190
      TabIndex        =   1
      Top             =   4410
      Width           =   2310
   End
   Begin VB.Label lbl_id_etiqueta 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "6JUN"
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
      Left            =   2175
      TabIndex        =   0
      Top             =   3990
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   1515
      Left            =   5970
      Stretch         =   -1  'True
      Top             =   510
      Width           =   1935
   End
End
Attribute VB_Name = "frmExibicao7GM2020"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sCodigo As String
Private Sub Form_Load()

    Me.lblPlant.Caption = ""
    'me.blQtdeContainers.Caption = ""
    Me.lblCodigoProduto1.Caption = ""
    Me.lblQtde1.Caption = ""
    Me.lblComplPeca1.Caption = ""
    
    Me.lblDescricao.Caption = ""
    Me.DataToEncodeText2.Text = ""
    Me.DataToEncodeText3.Text = ""
    Me.lblCodigoBarras.Caption = ""
    Me.lblCodigoBarrasA.Caption = ""
    Me.lblCodigoBarrasB.Caption = ""
    Me.lblCodMSB.Caption = ""
    Me.lblshipdate.Caption = ""
    Me.lblRoute.Caption = ""
    Me.lblExpDate.Caption = ""
    Me.lblLote.Caption = ""

    Me.Top = 0
    Me.Left = frmOpcoes.Width
    Me.lbl_data.Caption = Format(Now(), "dd/mm/yyyy")
    
End Sub
