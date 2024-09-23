VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExibicaoInmetroYamaha 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressão etiqueta Inmetro Yamaha"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   FillColor       =   &H80000009&
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   424
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   514
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   4320
      TabIndex        =   40
      Top             =   3495
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   4890
      TabIndex        =   39
      Top             =   4380
      Width           =   915
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5010
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_impressaoD 
      BackColor       =   &H00FFFF80&
      Caption         =   "&ImpressãoD"
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
      Left            =   4350
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Imprime esta Etiqueta"
      Top             =   5370
      Width           =   1335
   End
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
Private Declare Function SHInvokePrinterCommand Lib "shell32.dll" Alias "SHInvokePrinterCommandA" _
(ByVal hwnd As Long, ByVal uAction As Long, _
ByVal lBuf1 As String, ByVal lBuf2 As String, ByVal fModal As Boolean) As Long
Public nQtde_Etiquetas As Integer


Private Declare Function PrinterProperties Lib "winspool.drv" _
  (ByVal hwnd As Long, ByVal hPrinter As Long) As Long

Private Declare Function OpenPrinter Lib "winspool.drv" _
  Alias "OpenPrinterA" (ByVal pPrinterName As String, _
  phPrinter As Long, pDefault As PRINTER_DEFAULTS) As Long

Private Declare Function ClosePrinter Lib "winspool.drv" _
  (ByVal hPrinter As Long) As Long

Private Type PRINTER_DEFAULTS
     pDatatype As Long ' String
     pDevMode As Long
     pDesiredAccess As Long
End Type

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const PRINTER_ACCESS_ADMINISTER = &H4
Private Const PRINTER_ACCESS_USE = &H8
Private Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
   PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)



Private Sub cmd_impressaoD_Click()

Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim nx As Double
Dim x As Printer
Dim rs As ADODB.Recordset
Dim sTexto As String
Dim nqtde As Double
Dim bAchouImpressora As Boolean

bAchouImpressora = False
nx = 0
For Each x In Printers
   If x.DeviceName = Me.cbo_impressora.List(Me.cbo_impressora.ListIndex) Then
      Set Printer = x
      bAchouImpressora = True
      Exit For
   End If
Next

If Not bAchouImpressora Then
   MsgBox "Impressora da Etiqueta não encontrada, chame o responsável. o nome da impressora será - 'ETIQUETA FABRICA'"
   Exit Sub
End If
nx = 0

'SHInvokePrinterCommand Me.hwnd, &H6, Me.cbo_impressora.List(Me.cbo_impressora.ListIndex), vbNull, False


On Error GoTo Erro
Set rs = New ADODB.Recordset

rs.Fields.Append "1_Num_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "2_Qtde_Emb", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "3_Classe_Func", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "4_Indicacao_Supl", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "5_Data_Fab_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "6_Cod_Emb", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "7_Vinculo", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "8_Lote_Sob_Desv", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "9_Qtde_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "10_Aplicacao", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "11_DUM", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "12_Embarque_Ctrl", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "13_Cod_Fornecedor", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "14_Num_Doc_Fis_BAM", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "15_Data", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "16_Pondo_Entrega", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "17_Denominacao", ADODB.DataTypeEnum.adChar, 80
rs.Fields.Append "18_Num_Desenho", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "19_Ctrl_Interno", ADODB.DataTypeEnum.adChar, 150
rs.Fields.Append "20_Ctrl_Oper_Log", ADODB.DataTypeEnum.adChar, 150
rs.Fields.Append "21_Codigo_Numero", ADODB.DataTypeEnum.adChar, 50
rs.Fields.Append "22_codigo_barras", ADODB.DataTypeEnum.adChar, 50

rs.Open

Me.MousePointer = vbHourglass

nx = 0
nqtde = 0

   
rs.AddNew

rs.Fields("1_Num_Lote").Value = "654"
rs.Fields("3_Classe_Func").Value = "000111"
rs.Fields("4_Indicacao_Supl").Value = "258"
rs.Fields("5_Data_Fab_Lote").Value = Format(Now(), "DD/MM/YYYY")
rs.Fields("6_Cod_Emb").Value = "258"
rs.Fields("7_Vinculo").Value = "258"
rs.Fields("8_Lote_Sob_Desv").Value = "258"
rs.Fields("9_Qtde_Lote").Value = "258"
rs.Fields("10_Aplicacao").Value = "258"
rs.Fields("11_DUM").Value = " "
rs.Fields("12_Embarque_Ctrl").Value = "258"
rs.Fields("13_Cod_Fornecedor").Value = "258"
rs.Fields("14_Num_Doc_Fis_BAM").Value = "258"
rs.Fields("15_Data").Value = "258"
rs.Fields("16_Pondo_Entrega").Value = "258"
rs.Fields("17_Denominacao").Value = "258"
rs.Fields("18_Num_Desenho").Value = "258"
rs.Fields("19_Ctrl_Interno").Value = "258"
rs.Fields("20_Ctrl_Oper_Log").Value = "258"
rs.Fields("21_Codigo_Numero").Value = "564656546"
rs.Fields("22_codigo_barras").Value = "*65468465*"

'   For nx = 1 To Grid1.Rows - 1
'       1: nqtde = nqtde + VBA.CDbl("258")
'       Grid1.Row = nx
'   Next

rs.Fields("2_Qtde_Emb").Value = "0001"
rs.Update



Me.MousePointer = vbHourglass

'If Me.Opt_Pallet.Value = True Then
'   Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_FIAT_P.rpt")
'Else
   Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_FIAT.rpt")
'End If

With CrystalReport1
     .Database.SetDataSource rs
End With
rs.Clone

CrystalReport1.PrintOutEx False
'CrystalReport1.SelectPrinter x.DeviceName
Me.MousePointer = vbDefault

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub

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
If nQtde_Etiquetas > 1 Then
   nQtde_Etiquetas = nQtde_Etiquetas / 2
Else
   nQtde_Etiquetas = 1
End If
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
'*********************************************************************
         Dim I As Integer

         ' List all available printers
         For I = 0 To Printers.Count - 1
            List1.AddItem Printers(I).DeviceName
            If Printers(I).DeviceName = Printer.DeviceName Then
               List1.Selected(I) = True  ' Select current default printer
            End If
         Next I
 
'  Dim cnt As Long
'
'   For cnt = 0 To 2
'
'      If cnt > 0 Then Load CommonDialog1(cnt)
'
'      With CommonDialog1(cnt)
'
'         .Move 200, 200 + (370 * cnt), 2000, 345
'         .Visible = True
'
'         Select Case cnt
'            Case 0: .Caption = "Connect Net Printer"
'            Case 1: .Caption = "Open Explorer"
'            Case 2: .Caption = "Exit"
'         End Select
'
'      End With
'
'   Next  'cnt





End Sub


Private Sub Command1_Click()
   Dim RetVal As Long, hPrinter As Long
   Dim PD As PRINTER_DEFAULTS

   PD.pDatatype = 0
   ' Note that you cannot request more rights than you have as a user
   PD.pDesiredAccess = STANDARD_RIGHTS_REQUIRED Or PRINTER_ACCESS_USE
   PD.pDevMode = 0
   RetVal = OpenPrinter(Printer.DeviceName, hPrinter, PD)
   If RetVal = 0 Then
       MsgBox "OpenPrinter Failed!"
   Else
       RetVal = PrinterProperties(Me.hwnd, hPrinter)
       RetVal = ClosePrinter(hPrinter)
   End If
End Sub

Private Sub List1_Click()
   Dim Prt As Printer

   ' Find and use the printer just selected in the ListBox
   For Each Prt In Printers
      If Prt.DeviceName = List1.Text Then
            Set Printer = Prt
         Exit For
      End If
   Next
End Sub

