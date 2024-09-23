VERSION 5.00
Begin VB.Form frmExibicao2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prévia de impressão - FIAT"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "frmExibicao2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   345
      Left            =   0
      TabIndex        =   52
      Top             =   7710
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblCodigoBarrasA 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0000186627"
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
      Left            =   360
      TabIndex        =   51
      Top             =   7650
      Width           =   5250
   End
   Begin VB.Label lblCodigoBarrasB 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0000186627"
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
      Left            =   360
      TabIndex        =   50
      Top             =   7890
      Width           =   5250
   End
   Begin VB.Label lblMadeInBrazil 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Made in Brazil"
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
      Left            =   3630
      TabIndex        =   49
      Top             =   7350
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cód. Musashi:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   570
      TabIndex        =   48
      Top             =   7380
      Width           =   885
   End
   Begin VB.Label lblCodigoBarras 
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
      Height          =   270
      Left            =   1620
      TabIndex        =   47
      Top             =   7350
      Width           =   1395
   End
   Begin VB.Label lblEmbalagem2 
      Alignment       =   2  'Center
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
      Left            =   5130
      TabIndex        =   46
      Top             =   7350
      Width           =   405
   End
   Begin VB.Label lblCod_Peca 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO MUSASHI"
      Height          =   195
      Left            =   3480
      TabIndex        =   45
      Top             =   630
      Width           =   1890
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   5460
      X2              =   630
      Y1              =   4830
      Y2              =   4830
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   3255
      X2              =   3255
      Y1              =   5405
      Y2              =   4185
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   5460
      X2              =   600
      Y1              =   5430
      Y2              =   5430
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   3240
      X2              =   3240
      Y1              =   6620
      Y2              =   6030
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   5460
      X2              =   600
      Y1              =   6630
      Y2              =   6615
   End
   Begin VB.Label lblDum2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "20/09/2001"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3510
      TabIndex        =   44
      Top             =   6180
      Width           =   1425
   End
   Begin VB.Label lblLoteSobDesvio2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NÃO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2400
      TabIndex        =   43
      Top             =   6210
      Width           =   585
   End
   Begin VB.Label lblEmbarqueControlado2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NÃO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1110
      TabIndex        =   42
      Top             =   6210
      Width           =   585
   End
   Begin VB.Label lblPontoEntrega2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "G.09-PWT ME"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3270
      TabIndex        =   41
      Top             =   5610
      Width           =   1755
   End
   Begin VB.Label lblTipoVeiculo2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   40
      Top             =   5670
      Width           =   135
   End
   Begin VB.Label lblIndicacaoSuplementar2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "00045"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   39
      Top             =   5040
      Width           =   1515
   End
   Begin VB.Label lblVinculo2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   38
      Top             =   4980
      Width           =   465
   End
   Begin VB.Label lblClasseFuncional2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1050
      TabIndex        =   37
      Top             =   5010
      Width           =   525
   End
   Begin VB.Label lblDum 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DUM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3960
      TabIndex        =   36
      Top             =   6030
      Width           =   330
   End
   Begin VB.Label lblLoteSobDesvio 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Lote Sob Desvio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   2160
      TabIndex        =   35
      Top             =   6030
      Width           =   1020
   End
   Begin VB.Label lblEmbarqueControlado 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Embarque Control."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   660
      TabIndex        =   34
      Top             =   6030
      Width           =   1170
   End
   Begin VB.Label lblPontoEntrega 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Ponto de Entrega"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3540
      TabIndex        =   33
      Top             =   5460
      Width           =   1080
   End
   Begin VB.Label lblTipoVeiculo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Veículo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1230
      TabIndex        =   32
      Top             =   5490
      Width           =   960
   End
   Begin VB.Label lblIndicacao 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Indica. Suplem."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3720
      TabIndex        =   31
      Top             =   4860
      Width           =   975
   End
   Begin VB.Label lblVinculo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Vínculo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   2400
      TabIndex        =   30
      Top             =   4860
      Width           =   450
   End
   Begin VB.Label lblClasseFuncional 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Classe Funcional"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   720
      TabIndex        =   29
      Top             =   4860
      Width           =   1035
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   2970
      X2              =   2970
      Y1              =   6020
      Y2              =   5430
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   2040
      X2              =   2040
      Y1              =   6620
      Y2              =   6030
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   2040
      X2              =   2040
      Y1              =   5410
      Y2              =   4185
   End
   Begin VB.Label lblQtdEmbalagem2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "00300"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3390
      TabIndex        =   28
      Top             =   4380
      Width           =   1875
   End
   Begin VB.Label lblQtdEmbalagem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quant. da Embalagem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3660
      TabIndex        =   27
      Top             =   4230
      Width           =   1410
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   375
      Left            =   780
      Top             =   6900
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   375
      Left            =   3300
      Top             =   6900
      Width           =   1935
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   5460
      X2              =   600
      Y1              =   2040
      Y2              =   2025
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5460
      X2              =   600
      Y1              =   960
      Y2              =   945
   End
   Begin VB.Label lblBAM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Doc. Fiscal (BAM)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3420
      TabIndex        =   25
      Top             =   1530
      Width           =   1335
   End
   Begin VB.Label lblDenominacao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Denominação do produto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   900
      TabIndex        =   24
      Top             =   1530
      Width           =   1545
   End
   Begin VB.Label lblDesenho2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "97648915612"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1290
      TabIndex        =   6
      Top             =   2010
      Width           =   4080
   End
   Begin VB.Label lblCodBarraCp2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "9764891561270000456"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   23
      Top             =   3090
      Width           =   4680
   End
   Begin VB.Label lblCodBarraCp1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "9764891561270000456"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   720
      TabIndex        =   22
      Top             =   2850
      Width           =   4680
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   7260
      Left            =   600
      Top             =   90
      Width           =   4890
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   5460
      X2              =   600
      Y1              =   6030
      Y2              =   6015
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   5490
      X2              =   600
      Y1              =   1530
      Y2              =   1530
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   5460
      X2              =   600
      Y1              =   3600
      Y2              =   3585
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   5460
      X2              =   600
      Y1              =   4170
      Y2              =   4185
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   2970
      X2              =   2970
      Y1              =   960
      Y2              =   2055
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2970
      X2              =   2970
      Y1              =   4190
      Y2              =   3600
   End
   Begin VB.Label lblDataExpedicao 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1470
      TabIndex        =   21
      Top             =   990
      Width           =   300
   End
   Begin VB.Label lblCodFornec 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Código Fornecedor"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3540
      TabIndex        =   20
      Top             =   990
      Width           =   1170
   End
   Begin VB.Label lblDataProducao 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Produção do Lote"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1020
      TabIndex        =   19
      Top             =   3615
      Width           =   1515
   End
   Begin VB.Label lblCodEmbalagem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código da Embalagem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3450
      TabIndex        =   18
      Top             =   3615
      Width           =   1410
   End
   Begin VB.Label lblQtdLote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quant. do Lote"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   2145
      TabIndex        =   17
      Top             =   4215
      Width           =   930
   End
   Begin VB.Label lblNumLote 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Lote"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   990
      TabIndex        =   16
      Top             =   4230
      Width           =   660
   End
   Begin VB.Label lblFIAT 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Fiat Chrysler Automoveis do Brasil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   14
      Top             =   735
      Width           =   3015
   End
   Begin VB.Label lblMusashi 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MUSASHI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3810
      TabIndex        =   13
      Top             =   270
      Width           =   1260
   End
   Begin VB.Label lblDesenho 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Desenho"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   720
      TabIndex        =   12
      Top             =   2070
      Width           =   540
   End
   Begin VB.Label lblAlmoxarifado 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Almoxarifado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1335
      TabIndex        =   11
      Top             =   6735
      Width           =   825
   End
   Begin VB.Label lblDataExpedicao2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "99/99/9999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   1125
      Width           =   1560
   End
   Begin VB.Label lblCodFornec2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "7894564845"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3210
      TabIndex        =   8
      Top             =   1125
      Width           =   1815
   End
   Begin VB.Label lblDenominacao2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Conexão p/ Freio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   1665
      Width           =   1815
   End
   Begin VB.Label lblCodBarra 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "9764891561270000456"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   720
      TabIndex        =   5
      Top             =   2580
      Width           =   4680
   End
   Begin VB.Label lblCodBarra2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "9764891561270000456"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1050
      TabIndex        =   4
      Top             =   3360
      Width           =   3990
   End
   Begin VB.Label lblDataProducao2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "21/06/2003"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   990
      TabIndex        =   3
      Top             =   3750
      Width           =   1560
   End
   Begin VB.Label lblCodEmbalagem2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "FP7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3750
      TabIndex        =   2
      Top             =   3750
      Width           =   510
   End
   Begin VB.Label lblQtdLote2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "3.000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2100
      TabIndex        =   1
      Top             =   4335
      Width           =   1110
   End
   Begin VB.Label lblNumLote2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "00045"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   630
      TabIndex        =   0
      Top             =   4335
      Width           =   1365
   End
   Begin VB.Image imgFiat 
      Height          =   330
      Left            =   6270
      Picture         =   "frmExibicao2.frx":030A
      Stretch         =   -1  'True
      Top             =   2490
      Width           =   945
   End
   Begin VB.Image imgGM 
      Height          =   315
      Left            =   6630
      Picture         =   "frmExibicao2.frx":110E
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   315
   End
   Begin VB.Label lblBam2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "133777"
      Height          =   195
      Left            =   3810
      TabIndex        =   26
      Top             =   1755
      Width           =   540
   End
   Begin VB.Label lblLinha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Linha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   4080
      TabIndex        =   10
      Top             =   6735
      Width           =   345
   End
   Begin VB.Label lblLocal 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Localização"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   2400
      TabIndex        =   15
      Top             =   6630
      Width           =   675
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   1830
      Picture         =   "frmExibicao2.frx":1291
      Top             =   120
      Width           =   690
   End
End
Attribute VB_Name = "frmExibicao2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nTamannhowidth As Integer

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
Printer.Orientation = 1
Me.PrintForm
Printer.Orientation = 2: Printer.EndDoc
Me.Command1.Visible = True

End Sub

Private Sub Form_Load()
    'Limpar campos
    'FIAT
    frmExibicao2.lblCod_Peca.Caption = " "
    frmExibicao2.lblDataExpedicao2.Caption = " "
    frmExibicao2.lblCodFornec2.Caption = " "
    frmExibicao2.lblDenominacao2.Caption = " "
    frmExibicao2.lblBam2.Caption = " "
    frmExibicao2.lblDesenho2.Caption = " "
    frmExibicao2.lblCodBarra.Caption = " "
    frmExibicao2.lblCodBarraCp1.Caption = " "
    frmExibicao2.lblCodBarraCp2.Caption = " "
    frmExibicao2.lblDataProducao2.Caption = " "
    frmExibicao2.lblCodEmbalagem2.Caption = " "
    frmExibicao2.lblNumLote2.Caption = " "
    frmExibicao2.lblQtdLote2.Caption = " "
    frmExibicao2.lblQtdEmbalagem2.Caption = " "
    frmExibicao2.lblClasseFuncional2.Caption = " "
    frmExibicao2.lblVinculo2.Caption = " "
    frmExibicao2.lblIndicacaoSuplementar2.Caption = " "
    frmExibicao2.lblPontoEntrega2.Caption = " "
    frmExibicao2.lblEmbarqueControlado2.Caption = " "
    frmExibicao2.lblLoteSobDesvio2.Caption = " "
    frmExibicao2.lblDum2.Caption = " "

    Me.Height = 9330
    Me.Width = 5850
    Me.Top = 0
    
    Dim v As Integer
    For v = 0 To (Forms.Count - 1)
'        If Forms(v).Name = "frmFiat" Then
'            Me.Left = frmFiat.Width
'            Exit Sub
'        End If
        Me.Left = nTamannhowidth
    Next
End Sub
