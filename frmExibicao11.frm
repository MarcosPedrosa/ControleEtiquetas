VERSION 5.00
Begin VB.Form frmExibicao11 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox DataToEncodeText 
      Height          =   390
      Left            =   8040
      MaxLength       =   640
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmExibicao11.frx":0000
      Top             =   6165
      Width           =   2250
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   7020
      TabIndex        =   0
      Top             =   5130
      Width           =   1305
   End
   Begin VB.Label lbl_QA 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7620
      TabIndex        =   66
      Top             =   120
      Width           =   75
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
      Left            =   7260
      TabIndex        =   65
      Top             =   7965
      Width           =   540
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
      Left            =   2850
      TabIndex        =   64
      Top             =   4080
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
      Left            =   10440
      TabIndex        =   63
      Top             =   2715
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
      Left            =   4680
      TabIndex        =   62
      Top             =   5955
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
      Left            =   10440
      TabIndex        =   61
      Top             =   3015
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
      Left            =   450
      TabIndex        =   60
      Top             =   9795
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
      Left            =   450
      TabIndex        =   59
      Top             =   9555
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
      Left            =   1440
      TabIndex        =   58
      Top             =   9165
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
      Left            =   7710
      TabIndex        =   57
      Top             =   5445
      Width           =   45
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   4245
      Left            =   150
      Top             =   90
      Width           =   8295
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
      Left            =   6090
      TabIndex        =   56
      Top             =   8445
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
      Left            =   6090
      TabIndex        =   55
      Top             =   8790
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
      Left            =   6090
      TabIndex        =   54
      Top             =   7890
      Width           =   1020
   End
   Begin VB.Label lbl_LPN_COD_BARRA 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5089060613000001"
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   26.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3285
      TabIndex        =   53
      Top             =   585
      Width           =   2160
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   8400
      X2              =   120
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label lblCodMSB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2G01740444"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   52
      Top             =   4680
      Width           =   1710
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
      Left            =   10425
      TabIndex        =   51
      Top             =   3435
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
      Left            =   10470
      TabIndex        =   50
      Top             =   2415
      Width           =   2775
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   8385
      X2              =   120
      Y1              =   2565
      Y2              =   2565
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
      Left            =   2490
      TabIndex        =   49
      Top             =   4710
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
      Left            =   9510
      TabIndex        =   48
      Top             =   4785
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
      Left            =   9480
      TabIndex        =   47
      Top             =   5085
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
      Left            =   9480
      TabIndex        =   46
      Top             =   4335
      Width           =   1065
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
      Left            =   8880
      TabIndex        =   45
      Top             =   4620
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
      Left            =   8880
      TabIndex        =   44
      Top             =   4875
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
      Left            =   5250
      TabIndex        =   43
      Top             =   6135
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
      Left            =   10155
      TabIndex        =   42
      Top             =   3315
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
      Left            =   8940
      TabIndex        =   41
      Top             =   2085
      Width           =   165
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   6060
      X2              =   6060
      Y1              =   570
      Y2              =   90
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
      Left            =   10995
      TabIndex        =   40
      Top             =   6990
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
      Left            =   10995
      TabIndex        =   39
      Top             =   6780
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
      Left            =   10995
      TabIndex        =   38
      Top             =   6585
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
      Left            =   10995
      TabIndex        =   37
      Top             =   6375
      Width           =   2235
   End
   Begin VB.Label lbl_LPN 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "LPN:"
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
      Left            =   150
      TabIndex        =   36
      Top             =   60
      Width           =   345
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
      Left            =   6150
      TabIndex        =   35
      Top             =   7185
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
      Left            =   8850
      TabIndex        =   34
      Top             =   1785
      Width           =   1365
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
      Left            =   8970
      TabIndex        =   33
      Top             =   1410
      Width           =   555
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   8385
      X2              =   120
      Y1              =   1875
      Y2              =   1875
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
      Left            =   3540
      TabIndex        =   32
      Top             =   6525
      Width           =   2325
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   8385
      X2              =   120
      Y1              =   1335
      Y2              =   1335
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
      Left            =   480
      TabIndex        =   31
      Top             =   6465
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
      Left            =   840
      TabIndex        =   30
      Top             =   8565
      Width           =   4080
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
      Left            =   930
      TabIndex        =   29
      Top             =   7020
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
      Left            =   10995
      TabIndex        =   28
      Top             =   6165
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   11070
      Picture         =   "frmExibicao11.frx":0004
      Stretch         =   -1  'True
      Top             =   1275
      Width           =   2085
   End
   Begin VB.Label lbl_LPN_COD 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "5089060613000001"
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
      Left            =   720
      TabIndex        =   27
      Top             =   165
      Width           =   2160
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   7200
      X2              =   7200
      Y1              =   570
      Y2              =   90
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
      Left            =   3510
      TabIndex        =   26
      Top             =   1125
      Width           =   1695
   End
   Begin VB.Label lbl_CODIGO 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO:"
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
      Left            =   150
      TabIndex        =   25
      Top             =   1335
      Width           =   645
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
      Left            =   210
      TabIndex        =   24
      Top             =   1515
      Width           =   2145
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   3600
      X2              =   3600
      Y1              =   1890
      Y2              =   1365
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   5790
      X2              =   5790
      Y1              =   1860
      Y2              =   1365
   End
   Begin VB.Label lbl_SUPPLIER 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER:"
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
      Left            =   3660
      TabIndex        =   23
      Top             =   1335
      Width           =   795
   End
   Begin VB.Label lbl_SUPPLIER_COD 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "5089"
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
      Left            =   4230
      TabIndex        =   22
      Top             =   1515
      Width           =   600
   End
   Begin VB.Label lbl_USER 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "USER:"
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
      Left            =   5910
      TabIndex        =   21
      Top             =   1335
      Width           =   450
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
      Left            =   6810
      TabIndex        =   20
      Top             =   1515
      Width           =   600
   End
   Begin VB.Label lbl_YAMAHA_BARRAS 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5089-0606130-00001"
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   24
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2925
      TabIndex        =   19
      Top             =   1875
      Width           =   2430
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
      Left            =   3300
      TabIndex        =   18
      Top             =   2355
      Width           =   1695
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4530
      X2              =   4530
      Y1              =   3570
      Y2              =   2580
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   8385
      X2              =   120
      Y1              =   3075
      Y2              =   3075
   End
   Begin VB.Label lbl_NOME 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NOME:"
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
      Left            =   150
      TabIndex        =   17
      Top             =   2565
      Width           =   480
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
      Left            =   150
      TabIndex        =   16
      Top             =   2775
      Width           =   4335
   End
   Begin VB.Label lbl_FORNECEDOR 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "FORNECEDOR:"
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
      Left            =   4590
      TabIndex        =   15
      Top             =   2565
      Width           =   1080
   End
   Begin VB.Label lbl_FORNECEDOR_NOME 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NISSIN BRAKE"
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
      Left            =   4680
      TabIndex        =   14
      Top             =   2775
      Width           =   1335
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   1620
      X2              =   1620
      Y1              =   3600
      Y2              =   3090
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   8385
      X2              =   120
      Y1              =   3615
      Y2              =   3615
   End
   Begin VB.Label lbl_QTDE 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "QTDE:"
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
      Left            =   150
      TabIndex        =   13
      Top             =   3105
      Width           =   465
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
      Left            =   630
      TabIndex        =   12
      Top             =   3255
      Width           =   450
   End
   Begin VB.Label lbl_NF 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NF/INVOICE:"
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
      Left            =   1680
      TabIndex        =   11
      Top             =   3105
      Width           =   900
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
      TabIndex        =   10
      Top             =   3255
      Width           =   1350
   End
   Begin VB.Label lbl_FIFO 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "FIFO:"
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
      Left            =   4590
      TabIndex        =   9
      Top             =   3105
      Width           =   390
   End
   Begin VB.Label lbl_MES 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "JUL"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   5310
      TabIndex        =   8
      Top             =   2955
      Width           =   1125
   End
   Begin VB.Label lbl_ANO 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "17"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   6750
      TabIndex        =   7
      Top             =   2940
      Width           =   690
   End
   Begin VB.Label lbl_QTDE_BARRAS 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   21.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1080
      TabIndex        =   6
      Top             =   3630
      Width           =   360
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
      Left            =   1080
      TabIndex        =   5
      Top             =   4110
      Width           =   345
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   2790
      X2              =   2790
      Y1              =   4320
      Y2              =   3630
   End
   Begin VB.Label lbl_MUSASHI 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "COD.MUSASHI:"
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
      Left            =   2820
      TabIndex        =   4
      Top             =   3600
      Width           =   1110
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
      Left            =   4530
      TabIndex        =   3
      Top             =   3660
      Width           =   2760
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
      Left            =   5580
      TabIndex        =   2
      Top             =   4110
      Width           =   885
   End
End
Attribute VB_Name = "frmExibicao11"
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

'Me.Height = 6630
'Me.Width = 9780
Me.Top = 0
Me.Left = frmOpcoes.Width + 100

'For v = 0 To (Forms.Count - 1)
'    If Forms(v).Name = "frmExibicao11" Then
'        Me.Left = frmExibicao11.Width - 1000
'        Exit For
'    End If
'Next

Me.lbl_data.Caption = Format(Now(), "dd/mm/yyyy")

End Sub


