VERSION 5.00
Begin VB.Form frmExibicao2Ant 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prévia de impressão - FIAT"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5145
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   5145
   Begin VB.Image Image1 
      Height          =   690
      Left            =   5700
      Picture         =   "frmExibicao2Ant.frx":0000
      Top             =   3510
      Width           =   765
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
      Left            =   2220
      TabIndex        =   52
      Top             =   6735
      Width           =   675
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
      Left            =   3360
      TabIndex        =   51
      Top             =   6840
      Width           =   345
   End
   Begin VB.Label lblBam2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "133777"
      Height          =   195
      Left            =   3300
      TabIndex        =   50
      Top             =   1860
      Width           =   540
   End
   Begin VB.Image imgGM 
      Height          =   315
      Left            =   1920
      Picture         =   "frmExibicao2Ant.frx":20A2
      Stretch         =   -1  'True
      Top             =   435
      Width           =   315
   End
   Begin VB.Image imgFiat 
      Height          =   330
      Left            =   540
      Picture         =   "frmExibicao2Ant.frx":2225
      Stretch         =   -1  'True
      Top             =   375
      Width           =   945
   End
   Begin VB.Label lblNumLote2 
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
      Left            =   660
      TabIndex        =   49
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label lblQtdLote2 
      AutoSize        =   -1  'True
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
      Left            =   2220
      TabIndex        =   48
      Top             =   4440
      Width           =   750
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
      Left            =   3330
      TabIndex        =   47
      Top             =   3840
      Width           =   510
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
      Left            =   675
      TabIndex        =   46
      Top             =   3840
      Width           =   1560
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
      Left            =   870
      TabIndex        =   45
      Top             =   3465
      Width           =   3360
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
      Left            =   480
      TabIndex        =   44
      Top             =   2685
      Width           =   4170
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
      Left            =   540
      TabIndex        =   43
      Top             =   1770
      Width           =   1815
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
      Left            =   2700
      TabIndex        =   42
      Top             =   1140
      Width           =   1815
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
      Left            =   660
      TabIndex        =   41
      Top             =   1140
      Width           =   1560
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
      Left            =   1095
      TabIndex        =   40
      Top             =   6840
      Width           =   825
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
      Left            =   540
      TabIndex        =   39
      Top             =   2130
      Width           =   540
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
      Left            =   3015
      TabIndex        =   38
      Top             =   360
      Width           =   1260
   End
   Begin VB.Label lblPower 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "POWERTRAIN"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1500
      TabIndex        =   37
      Top             =   765
      Width           =   960
   End
   Begin VB.Label lblFIAT 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "FIAT - GM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   540
      TabIndex        =   36
      Top             =   690
      Width           =   915
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
      Left            =   540
      TabIndex        =   35
      Top             =   4335
      Width           =   660
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
      Left            =   1965
      TabIndex        =   34
      Top             =   4320
      Width           =   930
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
      Left            =   2700
      TabIndex        =   33
      Top             =   3720
      Width           =   1410
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
      Left            =   525
      TabIndex        =   32
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label lblCodFornec 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cód. Fornecedor"
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
      Left            =   2700
      TabIndex        =   31
      Top             =   975
      Width           =   1035
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
      Left            =   525
      TabIndex        =   30
      Top             =   975
      Width           =   300
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2580
      X2              =   2580
      Y1              =   4295
      Y2              =   3705
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   2580
      X2              =   2580
      Y1              =   375
      Y2              =   2135
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   4630
      X2              =   420
      Y1              =   4290
      Y2              =   4290
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   4630
      X2              =   420
      Y1              =   3690
      Y2              =   3690
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   4630
      X2              =   420
      Y1              =   1635
      Y2              =   1635
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   4630
      X2              =   420
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   7080
      Left            =   420
      Top             =   375
      Width           =   4260
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
      Left            =   480
      TabIndex        =   29
      Top             =   2955
      Width           =   4170
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
      Left            =   480
      TabIndex        =   28
      Top             =   3195
      Width           =   4170
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
      Left            =   900
      TabIndex        =   27
      Top             =   2115
      Width           =   3450
   End
   Begin VB.Label lblDenominacao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Denominação"
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
      Left            =   525
      TabIndex        =   26
      Top             =   1635
      Width           =   840
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
      Left            =   2940
      TabIndex        =   25
      Top             =   1635
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4630
      X2              =   420
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   4630
      X2              =   420
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   375
      Left            =   2580
      Top             =   7005
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   375
      Left            =   540
      Top             =   7005
      Width           =   1935
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
      Left            =   3180
      TabIndex        =   24
      Top             =   4335
      Width           =   1410
   End
   Begin VB.Label lblQtdEmbalagem2 
      AutoSize        =   -1  'True
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
      Left            =   3540
      TabIndex        =   23
      Top             =   4455
      Width           =   825
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   1860
      X2              =   1860
      Y1              =   5515
      Y2              =   4290
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   1860
      X2              =   1860
      Y1              =   6725
      Y2              =   6135
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   2580
      X2              =   2580
      Y1              =   6125
      Y2              =   5535
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
      Left            =   540
      TabIndex        =   22
      Top             =   4965
      Width           =   1035
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
      Left            =   2220
      TabIndex        =   21
      Top             =   4965
      Width           =   450
   End
   Begin VB.Label lblIndicacao 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Indicação Suplementar"
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
      Left            =   3105
      TabIndex        =   20
      Top             =   4965
      Width           =   1410
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
      Left            =   540
      TabIndex        =   19
      Top             =   5535
      Width           =   960
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
      Left            =   2700
      TabIndex        =   18
      Top             =   5535
      Width           =   1080
   End
   Begin VB.Label lblEmbarqueControlado 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Embarque Controlado"
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
      Left            =   480
      TabIndex        =   17
      Top             =   6135
      Width           =   1350
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
      Left            =   1980
      TabIndex        =   16
      Top             =   6135
      Width           =   1020
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
      Left            =   3780
      TabIndex        =   15
      Top             =   6135
      Width           =   330
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
      Left            =   990
      TabIndex        =   14
      Top             =   5055
      Width           =   195
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
      Left            =   2370
      TabIndex        =   13
      Top             =   5055
      Width           =   315
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
      Left            =   3420
      TabIndex        =   12
      Top             =   5175
      Width           =   855
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
      Left            =   1140
      TabIndex        =   11
      Top             =   5775
      Width           =   135
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
      Left            =   2730
      TabIndex        =   10
      Top             =   5655
      Width           =   1755
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
      Left            =   930
      TabIndex        =   9
      Top             =   6255
      Width           =   585
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
      Left            =   2235
      TabIndex        =   8
      Top             =   6255
      Width           =   585
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
      Left            =   3135
      TabIndex        =   7
      Top             =   6255
      Width           =   1425
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   4630
      X2              =   420
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   3060
      X2              =   3060
      Y1              =   6725
      Y2              =   6135
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   4630
      X2              =   420
      Y1              =   5535
      Y2              =   5535
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   3075
      X2              =   3075
      Y1              =   5510
      Y2              =   4290
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   4630
      X2              =   420
      Y1              =   4935
      Y2              =   4935
   End
   Begin VB.Label lblCod_Peca 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO MUSASHI"
      Height          =   195
      Left            =   2700
      TabIndex        =   6
      Top             =   735
      Width           =   1890
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
      Left            =   4140
      TabIndex        =   5
      Top             =   7455
      Width           =   45
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
      Left            =   1440
      TabIndex        =   4
      Top             =   7455
      Width           =   1395
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
      Left            =   540
      TabIndex        =   3
      Top             =   7455
      Width           =   885
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
      Left            =   3060
      TabIndex        =   2
      Top             =   7455
      Width           =   975
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
      Left            =   -90
      TabIndex        =   1
      Top             =   7995
      Width           =   5250
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
      Left            =   -90
      TabIndex        =   0
      Top             =   7755
      Width           =   5250
   End
End
Attribute VB_Name = "frmExibicao2Ant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Limpar campos
    'FIAT
    frmExibicao2.lblCod_Peca = ""
    frmExibicao2.lblDataExpedicao2 = ""
    frmExibicao2.lblCodFornec2 = ""
    frmExibicao2.lblDenominacao2 = ""
    frmExibicao2.lblBam2 = ""
    frmExibicao2.lblDesenho2 = ""
    frmExibicao2.lblCodBarra = ""
    frmExibicao2.lblCodBarraCp1 = ""
    frmExibicao2.lblCodBarraCp2 = ""
    frmExibicao2.lblDataProducao2 = ""
    frmExibicao2.lblCodEmbalagem2 = ""
    frmExibicao2.lblNumLote2 = ""
    frmExibicao2.lblQtdLote2 = ""
    frmExibicao2.lblQtdEmbalagem2 = ""
    frmExibicao2.lblClasseFuncional2 = ""
    frmExibicao2.lblVinculo2 = ""
    frmExibicao2.lblIndicacaoSuplementar2 = ""
    frmExibicao2.lblPontoEntrega2 = ""
    frmExibicao2.lblEmbarqueControlado2 = ""
    frmExibicao2.lblLoteSobDesvio2 = ""
    frmExibicao2.lblDum2 = ""
    
    Me.Height = 9330
    Me.Width = 5850
    Me.Top = 0
    
    Dim v As Integer
    For v = 0 To (Forms.Count - 1)
    If Forms(v).Name = "frmFiat" Then
        Me.Left = frmFiat.Width
        Exit Sub
    End If
    Me.Left = frmOpcoes.Width
    Next
End Sub

