VERSION 5.00
Begin VB.Form frmExibicao5 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prévia de Impressão - GM"
   ClientHeight    =   5595
   ClientLeft      =   1680
   ClientTop       =   2385
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExibicao5.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   8730
   Begin VB.TextBox DataToEncodeText 
      Height          =   390
      Left            =   900
      MaxLength       =   640
      MultiLine       =   -1  'True
      TabIndex        =   48
      Text            =   "frmExibicao5.frx":030A
      Top             =   6180
      Width           =   2250
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
      Left            =   7110
      TabIndex        =   47
      Top             =   3330
      Width           =   540
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   5940
      X2              =   5940
      Y1              =   1650
      Y2              =   2370
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
      TabIndex        =   46
      Top             =   5550
      Width           =   900
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
      Left            =   2970
      TabIndex        =   45
      Top             =   480
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
      TabIndex        =   44
      Top             =   6030
      Width           =   360
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
      Left            =   2970
      TabIndex        =   43
      Top             =   780
      Width           =   2775
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
      Left            =   390
      TabIndex        =   42
      Top             =   5130
      Width           =   5250
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
      Left            =   390
      TabIndex        =   41
      Top             =   4890
      Width           =   5250
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
      Left            =   1380
      TabIndex        =   40
      Top             =   4500
      Width           =   2880
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
      TabIndex        =   39
      Top             =   5520
      Width           =   45
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   5385
      Left            =   270
      Top             =   90
      Width           =   8295
   End
   Begin VB.Label lblRota2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SHIPMENT DATE :"
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
      Left            =   6030
      TabIndex        =   38
      Top             =   3105
      Width           =   855
   End
   Begin VB.Label lblGross2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "GROSS WEIGHT :"
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
      Left            =   6030
      TabIndex        =   37
      Top             =   3990
      Width           =   825
   End
   Begin VB.Label lblContainer2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CONTAINER TYPE :"
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
      Left            =   6030
      TabIndex        =   36
      Top             =   3630
      Width           =   930
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
      Left            =   6030
      TabIndex        =   35
      Top             =   3780
      Width           =   675
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
      Left            =   6030
      TabIndex        =   34
      Top             =   4125
      Width           =   465
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
      Left            =   6030
      TabIndex        =   33
      Top             =   3225
      Width           =   1020
   End
   Begin VB.Label lblLicenseA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "123456"
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
      Left            =   810
      TabIndex        =   32
      Top             =   3330
      Width           =   990
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
      Left            =   330
      TabIndex        =   31
      Top             =   3120
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cód. MUSASHI :"
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
      Left            =   420
      TabIndex        =   30
      Top             =   4470
      Width           =   765
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5940
      X2              =   240
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label lblPart2B 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NUMBER :"
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
      Left            =   330
      TabIndex        =   29
      Top             =   2580
      Width           =   495
   End
   Begin VB.Label lblCodMSB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2G01740"
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
      Left            =   4590
      TabIndex        =   28
      Top             =   4380
      Width           =   1215
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
      Left            =   2955
      TabIndex        =   27
      Top             =   1200
      Width           =   2250
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
      Left            =   2970
      TabIndex        =   26
      Top             =   180
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   8570
      X2              =   270
      Y1              =   4380
      Y2              =   4380
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
      Height          =   300
      Left            =   9690
      TabIndex        =   25
      Top             =   4620
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
      TabIndex        =   24
      Top             =   4860
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
      TabIndex        =   23
      Top             =   5160
      Width           =   1065
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
      TabIndex        =   22
      Top             =   4410
      Width           =   1065
   End
   Begin VB.Label lblMfg 
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
      Left            =   6060
      TabIndex        =   21
      Top             =   5220
      Width           =   585
   End
   Begin VB.Label lblEng2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DELIVERY NOTE or PUS or INVOICE NUMBER :"
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
      Left            =   6060
      TabIndex        =   20
      Top             =   4425
      Width           =   2235
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
      TabIndex        =   19
      Top             =   4695
      Width           =   405
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
      TabIndex        =   18
      Top             =   4950
      Width           =   1020
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
      TabIndex        =   17
      Top             =   6210
      Width           =   2265
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
      Left            =   2685
      TabIndex        =   16
      Top             =   1080
      Width           =   630
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
      Left            =   2760
      TabIndex        =   15
      Top             =   210
      Width           =   165
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   2640
      X2              =   2640
      Y1              =   1680
      Y2              =   180
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
      Left            =   405
      TabIndex        =   14
      Top             =   1185
      Width           =   1530
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
      Left            =   405
      TabIndex        =   13
      Top             =   975
      Width           =   1110
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
      Left            =   405
      TabIndex        =   12
      Top             =   780
      Width           =   1005
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
      Left            =   405
      TabIndex        =   11
      Top             =   570
      Width           =   2235
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
      Left            =   420
      TabIndex        =   10
      Top             =   165
      Width           =   315
   End
   Begin VB.Label lblReference2 
      Alignment       =   2  'Center
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
      Left            =   5970
      TabIndex        =   9
      Top             =   2400
      Width           =   660
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
      Left            =   6090
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   1050
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
      Left            =   3390
      TabIndex        =   7
      Top             =   1680
      Width           =   1365
   End
   Begin VB.Label lblPart2A 
      AutoSize        =   -1  'True
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
      Height          =   165
      Left            =   330
      TabIndex        =   6
      Top             =   2430
      Width           =   270
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
      Left            =   330
      TabIndex        =   5
      Top             =   1680
      Width           =   555
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   8540
      X2              =   240
      Y1              =   3120
      Y2              =   3120
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
      Left            =   3480
      TabIndex        =   4
      Top             =   1860
      Width           =   2325
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   8540
      X2              =   240
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   5940
      X2              =   5940
      Y1              =   2400
      Y2              =   5460
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
      Left            =   420
      TabIndex        =   3
      Top             =   1800
      Width           =   510
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
      Left            =   780
      TabIndex        =   2
      Top             =   3900
      Width           =   4080
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   3330
      X2              =   3330
      Y1              =   1680
      Y2              =   2400
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
      Left            =   870
      TabIndex        =   1
      Top             =   2355
      Width           =   2640
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
      Left            =   405
      TabIndex        =   0
      Top             =   360
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   6360
      Picture         =   "frmExibicao5.frx":030E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   2085
   End
End
Attribute VB_Name = "frmExibicao5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sCodigo As String
Public tWithOpcoes As Integer
Private Sub DataToEncodeText_Change()
Call CodeRefresh(DataToEncodeText.Text)
Me.lblLicenseA.Caption = sCodigo
'Me.lblLicense.Caption = sCodigo

End Sub

Private Sub Form_Load()
    
    'Limpar Campos
    'lblTo.Caption = ""
    lblPlant.Caption = ""
    lblPartNumber.Caption = ""
    lblPeca.Caption = ""
    lblQtd.Caption = ""
    lblMaterial.Caption = ""
    lblReference.Caption = ""
'    lblLicense.Caption = ""
    lblLicenseA.Caption = ""
    lblLicenseB.Caption = ""
    lblContainerType.Caption = ""
    lblgrossWeight.Caption = ""
    lblRoute.Caption = ""
    lblCodMSB.Caption = ""
    lblEng.Caption = ""
    lblLot.Caption = ""
    lblParts.Caption = ""
    lblMfgDate.Caption = ""
    
    'Mostra form
    'Me.Width = 9810
    'Me.Height = 7215
    Me.Height = 6630
    Me.Width = 9780
    Me.Top = 0
    Me.Left = tWithOpcoes
    
    Dim v As Integer
    For v = 0 To (Forms.Count - 1)
        If Forms(v).Name = "frmGm" Then
            Me.Left = frmGm.Width
            Exit For
        End If
    Next
    Me.lbl_data.Caption = Format(Now(), "dd/mm/yyyy")
End Sub

'    If cRec.Fields("Cliente") <> "" Then
'        frmExibicao5.lblTo.Caption = cRec.Fields("Cliente")
'    End If
'    Rem incluido (marcos pedrosa) em 23-08-2007
'    If cRec.Fields("MOTIVO_ALTERACAO_OUTROS") <> "" Then
'        frmExibicao5.lblMotivo_alteracao_outros.Caption = Trim(cRec("MOTIVO_ALTERACAO_OUTROS").Value)
'    End If
'    If cRec.Fields("Ind_Suplementar") <> "" Then
'        frmExibicao5.lblPlant.Caption = cRec.Fields("Ind_Suplementar")
'    End If
'    Rem DEFINICAO ATE 23-08-2007 (MARCOS PEDROSA)
''''        If cRec.Fields("Cod_Util") <> "" Then
''''            frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & cRec.Fields("Cod_Util")
''''        End If
''''        If cRec.Fields("Desvio") <> "" Then
''''            frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & cRec.Fields("Desvio")
''''        End If
'    Rem INCLUIDO MARCOS PEDROSA EM 23-08-2007
'    If cRec.Fields("Cod_Embalagem_Pw") <> "" Then
'        frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & cRec.Fields("Cod_Embalagem_Pw")
'    End If
'    Rem INCLUIDO MARCOS PEDROSA EM 23-08-2007
'    If cRec.Fields("Desvio") <> "" Then
'        frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & cRec.Fields("Desvio")
'    End If
'
'    If cRec.Fields("Cod_Cliente") <> "" Then
'        frmExibicao5.lblPartNumber.Caption = cRec.Fields("Cod_Cliente")
'    End If
'    If cRec.Fields("Descr_Peca") <> "" Then
'        frmExibicao5.lblPeca.Caption = cRec.Fields("Descr_Peca")
'    End If
'    If cRec.Fields("Qtd_Caixa") <> 0 Then
'        frmExibicao5.lblQtd.Caption = cRec.Fields("Qtd_Caixa")
'    End If
'    If cRec.Fields("Modelo") <> "" Then
'        frmExibicao5.lblMaterial.Caption = cRec.Fields("Modelo")
'    End If
'    If cRec.Fields("Cod_Embalagem_Pw") <> "" Then
'        frmExibicao5.lblReference.Caption = cRec.Fields("Cod_Embalagem_Pw")
'    End If
'    If cRec.Fields("Pto_Entrega") <> "" Then
'        frmExibicao5.lblLicense.Caption = cRec.Fields("Pto_Entrega")
'        frmExibicao5.lblLicenseA.Caption = "*" & cRec.Fields("Pto_Entrega") & "*"
'        frmExibicao5.lblLicenseB.Caption = "*" & cRec.Fields("Pto_Entrega") & "*"
'    End If
'    Rem comentado em 23-08-2007 (marcos pedrosa)
''''        If cRec.Fields("Cod_Embalagem") <> "" Then
''''            frmExibicao5.lblContainerType.Caption = cRec.Fields("Cod_Embalagem")
''''        End If
'    Rem INCLUIDO EM 23-08-2007
'    If cRec.Fields("compl_peca1") <> "" Then
'        frmExibicao5.lblContainerType.Caption = cRec.Fields("compl_peca1")
'    End If
'    If cRec.Fields("Peso") <> 0 Then
'        frmExibicao5.lblgrossWeight.Caption = Format(cRec.Fields("Peso"), "0.00")
'    End If
'    Rem comentado em 23-08-2007 (marcos pedrosa)
''''        If cRec.Fields("Embarque_Controlado") <> "" Then
''''            frmExibicao5.lblRoute.Caption = cRec.Fields("Embarque_Controlado")
''''        End If
'    Rem INCLUIDO EM 23-08-2007
'    If cRec.Fields("compl_peca2") <> "" Then
'        frmExibicao5.lblRoute.Caption = cRec.Fields("compl_peca2")
'    End If
'    If cRec.Fields("Lote") <> "" Then
'        frmExibicao5.lblLot.Caption = cRec.Fields("Lote")
'    End If
'    Rem comentado em 23-08-2007 (marcos pedrosa)
''''        If cRec.Fields("Dum") <> "" Then
''''            frmExibicao5.lblEng.Caption = cRec.Fields("Dum")
''''        End If
'    Rem INCLUIDO EM 23-08-2007
'    If cRec.Fields("data_lote") <> "" Then
'        frmExibicao5.lblEng.Caption = Mid$(cRec.Fields("data_lote"), 1, 2) & _
'                                      Pega_Mes(Val(Mid$(cRec.Fields("data_lote"), 4, 2))) & _
'                                      Mid$(cRec.Fields("data_lote"), 7, 4)
'    End If
'    Rem INCLUIDO EM 23-08-2007
'    If cRec.Fields("envio_lote") = "1" Then
'        frmExibicao5.lblvalidade.Caption = "N"
'    Else
'        frmExibicao5.lblvalidade.Caption = ""
'    End If
'
'    If cRec.Fields("Data_expedicao") <> "" Then
'        frmExibicao5.lblMfgDate.Caption = Mid$(cRec.Fields("Data_expedicao"), 1, 2) & _
'                                          Pega_Mes(Val(Mid$(cRec.Fields("Data_expedicao"), 4, 2))) & _
'                                          Mid$(cRec.Fields("Data_expedicao"), 7, 4)
'    End If
'    If cRec.Fields("Cod_Peca") <> "" Then
'        frmExibicao5.lblCodMSB.Caption = cRec.Fields("Cod_Peca")
'    End If
''    If Not IsNull(cRec.Fields("EMBALAGEM")) Then
''        frmExibicao5.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
''    End If
'    frmExibicao5.lblEmbalagem2.Caption = nMatricula
'    If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
'        frmExibicao5.lblCodigoBarras.Caption = Format(cRec.Fields("ID_ETIQUETA"), "0000000000")
'        frmExibicao5.lblCodigoBarrasA.Caption = "*" & frmExibicao5.lblCodigoBarras.Caption & "*"
'        frmExibicao5.lblCodigoBarrasB.Caption = "*" & frmExibicao5.lblCodigoBarras.Caption & "*"
'     Else
'        frmExibicao5.lblCodigoBarras.Caption = ""
'        frmExibicao5.lblCodigoBarrasA.Caption = ""
'        frmExibicao5.lblCodigoBarrasB.Caption = ""
'     End If
'End If
'
'
Public Function CodeRefresh(codigo)
  'Construction du code barre / Build the barcode
  Dim i%, dummy$
  
'  If COliste.ListCount > 0 Then
  If Len(codigo) > 0 Then
    dummy$ = Chr$(207)
'    For i% = 0 To COliste.ListCount - 1
'      dummy$ = dummy$ & COliste.List(i%)
'    Next
    dummy$ = dummy$ & codigo
    If Right$(dummy$, 1) = Chr$(207) Then dummy$ = Left$(dummy$, Len(dummy$) - 1)
    dummy$ = ean128$(dummy$)
    sCodigo = dummy$
  Else
    sCodigo = ""
  End If
End Function

Public Function ean128$(chaine$)
  'Cette fonction est régie par la Licence Générale Publique Amoindrie GNU (GNU LGPL)
  'This function is governed by the GNU Lesser General Public License (GNU LGPL)
  'V 2.0.0
  'Paramètres : une chaine
  'Parameters : a string
  'Retour : * une chaine qui, affichée avec la police CODE128.TTF, donne le code barre
  '         * une chaine vide si paramètre fourni incorrect
  'Return : * a string which give the bar code when it is dispayed with CODE128.TTF font
  '         * an empty string if the supplied parameter is no good
  Dim i%, Checksum&, mini%, dummy%, tableB As Boolean
  ean128$ = ""
  If Len(chaine$) > 0 Then
  'Vérifier si caractères valides
  'Check for valid characters
    For i% = 1 To Len(chaine$)
      Select Case Asc(Mid$(chaine$, i%, 1))
      Case 32 To 126, 203, 207
      Case Else
        i% = 0
        Exit For
      End Select
    Next
    'Calculer la chaine de code en optimisant l'usage des tables B et C
    'Calculation of the code string with optimized use of tables B and C
    ean128$ = ""
    tableB = True
    If i% > 0 Then
      i% = 1 'i% devient l'index sur la chaine / i% become the string index
      Do While i% <= Len(chaine$)
        If tableB Then
          'Voir si intéressant de passer en table C / See if interesting to switch to table C
          'Oui pour 4 chiffres au début ou à la fin, sinon pour 6 chiffres / yes for 4 digits at start or end, else if 6 digits
          mini% = IIf(i% = 1 Or i% + 3 = Len(chaine$), 4, 6)
          GoSub TestNumOrFnc1
          If mini% < 0 Then 'Choix table C / Choice of table C
            If i% = 1 Then 'Débuter sur table C / Starting with table C
              ean128$ = Chr$(210)
            Else 'Commuter sur table C / Switch to table C
              ean128$ = ean128$ & Chr$(204)
            End If
            tableB = False
          Else
            If i% = 1 Then ean128$ = Chr$(209) 'Débuter sur table B / Starting with table B
          End If
        End If
        If Not tableB Then
          'On est sur la table C, essayer de traiter 2 chiffres ou Ê/ We are on table C, try to process 2 digits or Ê
          If Asc(Mid$(chaine$, i%, 2)) = 207 Then
            'On traite le Fnc1 (Ê) / We process the Fnc1 (Ê)
            ean128$ = ean128$ & Mid$(chaine$, i%, 1)
            i% = i% + 1
          Else
            mini% = 2
            GoSub TestNum
            If mini% < 0 Then 'OK pour 2 chiffres, les traiter / OK for 2 digits, process it
              dummy% = Val(Mid$(chaine$, i%, 2))
              dummy% = IIf(dummy% < 95, dummy% + 32, dummy% + 105)
              ean128$ = ean128$ & Chr$(dummy%)
              i% = i% + 2
            Else 'On n'a pas 2 chiffres, repasser en table B / We haven't 2 digits, switch to table B
              ean128$ = ean128$ & Chr$(205)
              tableB = True
            End If
          End If
        End If
        If tableB Then
          'Traiter 1 caractère en table B / Process 1 digit with table B
          ean128$ = ean128$ & Mid$(chaine$, i%, 1)
          i% = i% + 1
        End If
      Loop
      'Calcul de la clé de contrôle / Calculation of the checksum
      For i% = 1 To Len(ean128$)
        dummy% = Asc(Mid$(ean128$, i%, 1))
        dummy% = IIf(dummy% < 127, dummy% - 32, dummy% - 105)
        If i% = 1 Then Checksum& = dummy%
        Checksum& = (Checksum& + (i% - 1) * dummy%) Mod 103
      Next
      'Calcul du code ASCII de la clé / Calculation of the checksum ASCII code
      Checksum& = IIf(Checksum& < 95, Checksum& + 32, Checksum& + 105)
      'Ajout de la clé et du STOP / Add the checksum and the STOP
      ean128$ = ean128$ & Chr$(Checksum&) & Chr$(211)
    End If
  End If
  Exit Function
TestNum:
  'si les mini% caractères à partir de i% sont numériques, alors mini%=0
  'if the mini% characters from i% are numeric, then mini%=0
  mini% = mini% - 1
  If i% + mini% <= Len(chaine$) Then
    Do While mini% >= 0
      If Asc(Mid$(chaine$, i% + mini%, 1)) < 48 Or Asc(Mid$(chaine$, i% + mini%, 1)) > 57 Then Exit Do
      mini% = mini% - 1
    Loop
  End If
Return
TestNumOrFnc1:
  'si les mini% caractères à partir de i% sont numériques ou  FNC1, alors mini%=0
  'if the mini% characters from i% are numeric or Fnc1, then mini%=0
  mini% = mini% - 1
  If i% + mini% <= Len(chaine$) Then
    Do While mini% >= 0
      If (Asc(Mid$(chaine$, i% + mini%, 1)) < 48 Or Asc(Mid$(chaine$, i% + mini%, 1)) > 57) And Asc(Mid$(chaine$, i% + mini%, 1)) <> 207 Then Exit Do
      mini% = mini% - 1
    Loop
  End If
Return
End Function
