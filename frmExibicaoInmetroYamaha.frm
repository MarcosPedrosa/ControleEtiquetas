VERSION 5.00
Begin VB.Form frmExibicaoInmetroYamaha 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão etiqueta Inmetro Yamaha"
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   FillColor       =   &H80000009&
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   418
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   289
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cbo_impressora 
      BackColor       =   &H00FFFF80&
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
      TabIndex        =   37
      Text            =   "Combo1"
      Top             =   5880
      Width           =   4155
   End
   Begin VB.CommandButton cmd_imprime 
      BackColor       =   &H00FFFF80&
      Caption         =   "&Impressão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2940
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Imprime esta Etiqueta"
      Top             =   5370
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   2580
      TabIndex        =   6
      Top             =   90
      Width           =   1605
      Begin VB.Image Image4 
         Height          =   585
         Left            =   810
         Picture         =   "frmExibicaoInmetroYamaha.frx":0000
         Stretch         =   -1  'True
         Top             =   300
         Width           =   645
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "OCP 0009"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   150
         TabIndex        =   11
         Top             =   750
         Width           =   645
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "I.Q.A."
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   300
         TabIndex        =   10
         Top             =   630
         Width           =   405
      End
      Begin VB.Label LBL_REGISTRO_INMETRO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "006297/2019"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   285
         TabIndex        =   9
         Top             =   1050
         Width           =   1140
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Registro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   480
         TabIndex        =   8
         Top             =   900
         Width           =   630
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Segurança"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   420
         TabIndex        =   7
         Top             =   120
         Width           =   780
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000007&
         BorderWidth     =   2
         Height          =   1125
         Left            =   30
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   1545
      End
      Begin VB.Image Image8 
         Height          =   405
         Left            =   240
         Picture         =   "frmExibicaoInmetroYamaha.frx":88FA
         Stretch         =   -1  'True
         Top             =   270
         Width           =   525
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   2580
      TabIndex        =   0
      Top             =   3090
      Width           =   1605
      Begin VB.Image Image3 
         Height          =   615
         Left            =   780
         Picture         =   "frmExibicaoInmetroYamaha.frx":1A451
         Stretch         =   -1  'True
         Top             =   240
         Width           =   675
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000007&
         BorderWidth     =   2
         Height          =   1215
         Left            =   30
         Shape           =   4  'Rounded Rectangle
         Top             =   30
         Width           =   1545
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Segurança"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   420
         TabIndex        =   5
         Top             =   30
         Width           =   780
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Registro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   480
         TabIndex        =   4
         Top             =   900
         Width           =   630
      End
      Begin VB.Label LBL_REGISTRO_INMETRO1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "006297/2019"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   285
         TabIndex        =   3
         Top             =   1050
         Width           =   1140
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "I.Q.A"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   6
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   165
         Left            =   270
         TabIndex        =   2
         Top             =   630
         Width           =   345
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "OCP 0009"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   150
         Left            =   150
         TabIndex        =   1
         Top             =   780
         Width           =   585
      End
      Begin VB.Image Image7 
         Height          =   435
         Left            =   180
         Picture         =   "frmExibicaoInmetroYamaha.frx":22D4B
         Stretch         =   -1  'True
         Top             =   210
         Width           =   525
      End
   End
   Begin VB.Label LBL_DESC_MOD8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "YAMAHA YBR350 FACTOR ED (2017 - 2018)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   120
      TabIndex        =   36
      Top             =   4800
      Width           =   3510
   End
   Begin VB.Label LBL_DESC_MOD7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "YAMAHA YBR350 FACTOR ED (2017 - 2018)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   90
      TabIndex        =   35
      Top             =   1800
      Width           =   3510
   End
   Begin VB.Label LBL_COD_PECA_CLIENTE1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2RP-F5439-10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   450
      TabIndex        =   34
      Top             =   3660
      Width           =   1740
   End
   Begin VB.Label LBL_COD_PECA_CLIENTE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2RP-F5439-10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   450
      TabIndex        =   33
      Top             =   630
      Width           =   1740
   End
   Begin VB.Label lblLicense2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MUSASHI DO BRASIL LTDA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   32
      Top             =   180
      Width           =   2220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CNPJ: 10.963.007/0001-62"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   31
      Top             =   420
      Width           =   2220
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4
      X2              =   174
      Y1              =   62
      Y2              =   62
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Modelos aplicaveis:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   30
      Top             =   1050
      Width           =   1635
   End
   Begin VB.Label LBL_DESC_PECA1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Coroa de transmissão: 39 dentes."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   90
      TabIndex        =   29
      Top             =   2010
      Width           =   2505
   End
   Begin VB.Label LBL_DESC_PECA2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Corrente correspondente: 428MX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   90
      TabIndex        =   28
      Top             =   2175
      Width           =   2535
   End
   Begin VB.Label LBL_DESC_PECA3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Fabricação: 10/2019."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   90
      TabIndex        =   27
      Top             =   2355
      Width           =   2265
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fabricado no Brasil"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1140
      TabIndex        =   26
      Top             =   2550
      Width           =   1590
   End
   Begin VB.Label LBL_DESC_MOD1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "YAMAHA YBR150 FACTOR E (2016-20218) "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   90
      TabIndex        =   25
      Top             =   1350
      Width           =   3465
   End
   Begin VB.Label LBL_DESC_MOD2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "YAMAHA YBR150 FACTOR ED (2016 - EM DIANTE)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   90
      TabIndex        =   24
      Top             =   1500
      Width           =   3960
   End
   Begin VB.Label LBL_DESC_MOD3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "YAMAHA YBR350 FACTOR ED (2017 - 2018)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   90
      TabIndex        =   23
      Top             =   1650
      Width           =   3510
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   2
      X2              =   278
      Y1              =   132
      Y2              =   132
   End
   Begin VB.Label LBL_DESC_MOD4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "YAMAHA YBR150 FACTOR E (2016-20218) "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   120
      TabIndex        =   22
      Top             =   4350
      Width           =   3465
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fabricado no Brasil"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1170
      TabIndex        =   21
      Top             =   5550
      Width           =   1590
   End
   Begin VB.Label LBL_DESC_PECA6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Fabricação: 10/2019."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   120
      TabIndex        =   20
      Top             =   5355
      Width           =   2265
   End
   Begin VB.Label LBL_DESC_PECA5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Corrente correspondente: 428MX"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   120
      TabIndex        =   19
      Top             =   5175
      Width           =   2535
   End
   Begin VB.Label LBL_DESC_PECA4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Coroa de transmissão: 39 dentes."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   120
      TabIndex        =   18
      Top             =   5010
      Width           =   2505
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CNPJ: 10.963.007/0001-62"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   17
      Top             =   3420
      Width           =   2220
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MUSASHI DO BRASIL LTDA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   16
      Top             =   3180
      Width           =   2220
   End
   Begin VB.Label LBL_DESC_MOD5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "YAMAHA YBR150 FACTOR ED (2016 - EM DIANTE)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   120
      TabIndex        =   15
      Top             =   4500
      Width           =   3960
   End
   Begin VB.Label LBL_DESC_MOD6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "YAMAHA YBR350 FACTOR ED (2017 - 2018)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   120
      TabIndex        =   14
      Top             =   4650
      Width           =   3510
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4
      X2              =   280
      Y1              =   332
      Y2              =   332
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   2
      X2              =   172
      Y1              =   266
      Y2              =   266
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Modelos aplicaveis:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   4110
      Width           =   1635
   End
End
Attribute VB_Name = "frmExibicaoInmetroYamaha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nQtde_Etiquetas As Integer

Private Sub cmd_imprime_Click()
Dim x As Printer
Dim nx As Integer
Dim bAchouImpressora As Boolean

nx = 0
bAchouImpressora = False

For Each x In Printers
   If x.DeviceName = Me.cbo_impressora.List(Me.cbo_impressora.ListIndex) Then
      Set Printer = x
      bAchouImpressora = True
      Exit For
   End If
Next

If Not bAchouImpressora Then
   MsgBox "Impressora - " & Me.cbo_impressora.List(Me.cbo_impressora.ListIndex) & " não Encontrada. Avise ao Responsável do TI. "
   Exit Sub
End If

Me.cmd_imprime.Visible = False
Me.cbo_impressora.Visible = False
If nQtde_Etiquetas > 1 Then nQtde_Etiquetas = nQtde_Etiquetas / 2
Printer.Copies = nQtde_Etiquetas
Printer.Orientation = 2

PrintForm
Printer.Orientation = 1: Printer.EndDoc
Me.cmd_imprime.Visible = True
Me.cbo_impressora.Visible = True
Unload Me
End Sub
Private Sub Form_Load()
Dim x As Printer
Dim nx As Integer

Me.Top = 0
Me.LBL_COD_PECA_CLIENTE.Caption = ""
Me.LBL_REGISTRO_INMETRO.Caption = ""

Me.LBL_DESC_MOD1.Caption = ""
Me.LBL_DESC_MOD2.Caption = ""
Me.LBL_DESC_MOD3.Caption = ""

Me.LBL_DESC_PECA1.Caption = ""
Me.LBL_DESC_PECA2.Caption = ""
Me.LBL_DESC_PECA3.Caption = ""

Me.LBL_COD_PECA_CLIENTE1.Caption = ""
Me.LBL_REGISTRO_INMETRO1.Caption = ""

Me.LBL_DESC_MOD4.Caption = ""
Me.LBL_DESC_MOD5.Caption = ""
Me.LBL_DESC_MOD6.Caption = ""

Me.LBL_DESC_MOD7.Caption = ""
Me.LBL_DESC_MOD8.Caption = ""

Me.LBL_DESC_PECA4.Caption = ""
Me.LBL_DESC_PECA5.Caption = ""
Me.LBL_DESC_PECA6.Caption = ""

For Each x In Printers
    If UCase(Mid$(x.DeviceName, 1, 8)) = "ETIQUETA" Then
       Me.cbo_impressora.AddItem x.DeviceName
    End If
Next

Me.cbo_impressora.ListIndex = 0
Rem verificar a impressora padrão para ser usada neste relatório
For nx = 0 To Me.cbo_impressora.ListCount - 1
    If Trim(UCase(sImpressoraInmetro)) = Trim(UCase(Me.cbo_impressora.List(nx))) Then
       Me.cbo_impressora.ListIndex = nx
    End If
Next

End Sub
