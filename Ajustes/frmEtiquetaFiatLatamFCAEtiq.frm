VERSION 5.00
Begin VB.Form frmEtiquetaFiatLatamFCAEtiq 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Etiqueta Fiat Latam FCA"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   13470
   Begin VB.CommandButton cmd_imprime 
      Caption         =   "Impressão"
      Height          =   225
      Left            =   540
      TabIndex        =   67
      Top             =   5190
      Width           =   1215
   End
   Begin VB.Image IMG_CODE_128 
      Height          =   840
      Left            =   2400
      Stretch         =   -1  'True
      Top             =   4260
      Width           =   5880
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2160
      TabIndex        =   69
      Top             =   5850
      Width           =   6705
   End
   Begin VB.Label lbl_24_Codigo_Barra_C 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "*0005526798900130930000014322*"
      BeginProperty Font 
         Name            =   "Code 39"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2535
      TabIndex        =   68
      Top             =   7140
      Width           =   5865
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Desenho(Chrysler)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   2
      Left            =   2820
      TabIndex        =   3
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   8400
      X2              =   150
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line19 
      BorderWidth     =   2
      X1              =   3690
      X2              =   2790
      Y1              =   1050
      Y2              =   1050
   End
   Begin VB.Image IMG_32_QRCODE 
      Height          =   2310
      Left            =   180
      Stretch         =   -1  'True
      Top             =   630
      Width           =   2520
   End
   Begin VB.Label lbl_24_Codigo_Barra_B 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "*0005526798900130930000014322*"
      BeginProperty Font 
         Name            =   "Code 39"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2535
      TabIndex        =   66
      Top             =   6750
      Visible         =   0   'False
      Width           =   5865
   End
   Begin VB.Label lbl_24_Codigo_Barra_A 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "*0005526798900130930000014322*"
      BeginProperty Font 
         Name            =   "Code 39"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2535
      TabIndex        =   65
      Top             =   6960
      Visible         =   0   'False
      Width           =   5865
   End
   Begin VB.Label lbl_02_ODM 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "1234"
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
      Left            =   1680
      TabIndex        =   64
      Top             =   4890
      Width           =   375
   End
   Begin VB.Label lbl_14_Qtde_emb 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "3000"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   315
      TabIndex        =   61
      Top             =   4470
      Width           =   315
   End
   Begin VB.Label lbl_22_Num_Lote 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "00000045623"
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
      Left            =   6945
      TabIndex        =   60
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label lbl_15_Num_Sheda_Serial 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "543245"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   5550
      TabIndex        =   59
      Top             =   3960
      Width           =   465
   End
   Begin VB.Label lbl_16_Id_Inter_Fornecedor 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "0000041231"
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
      Left            =   3510
      TabIndex        =   58
      Top             =   3960
      Width           =   765
   End
   Begin VB.Label lbl_17_Embarque_Ctrl 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "20/05/2020"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1290
      TabIndex        =   56
      Top             =   4020
      Width           =   690
   End
   Begin VB.Label lbl_05_Data_Expedicao 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "20/05/2020"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   55
      Top             =   4020
      Width           =   690
   End
   Begin VB.Label lbl_30_Vinculo 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "NI"
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
      Left            =   2010
      TabIndex        =   48
      Top             =   3180
      Width           =   150
   End
   Begin VB.Label lbl_19_Classe_Func 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "1D"
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
      Left            =   1290
      TabIndex        =   47
      Top             =   3180
      Width           =   210
   End
   Begin VB.Label lbl_31_Restricoes 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "030F"
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
      Left            =   225
      TabIndex        =   46
      Top             =   3210
      Width           =   375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cód.Embal"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Index           =   12
      Left            =   6180
      TabIndex        =   14
      Top             =   3120
      Width           =   585
   End
   Begin VB.Label lbl_07_Quantidade 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "9000"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7230
      TabIndex        =   43
      Top             =   2970
      Width           =   720
   End
   Begin VB.Label lbl_27B_Num_Desenho 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "34708"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   42.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   4920
      TabIndex        =   40
      Top             =   1950
      Width           =   1965
   End
   Begin VB.Label lbl_27A_Num_Desenho 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "005182"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3360
      TabIndex        =   39
      Top             =   2250
      Width           =   1530
   End
   Begin VB.Label lbl_36_Incoterms 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "CIF"
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
      Left            =   3075
      TabIndex        =   38
      Top             =   1830
      Width           =   255
   End
   Begin VB.Label lbl_08_DOCA 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2985
      TabIndex        =   37
      Top             =   1170
      Width           =   465
   End
   Begin VB.Label lbl_28_Destino 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "FIAPE"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2910
      TabIndex        =   35
      Top             =   210
      Width           =   585
   End
   Begin VB.Label lbl_34_DUM 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "04/07/2020"
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
      Left            =   1950
      TabIndex        =   34
      Top             =   300
      Width           =   690
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   180
      Picture         =   "frmEtiquetaFiatLatamFCAEtiq.frx":0000
      Stretch         =   -1  'True
      Top             =   90
      Width           =   1065
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ODM"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   31
      Left            =   1590
      TabIndex        =   33
      Top             =   4740
      Width           =   300
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Data Validade"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   30
      Left            =   180
      TabIndex        =   32
      Top             =   4740
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Dados de Transporte"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   29
      Left            =   840
      TabIndex        =   31
      Top             =   4290
      Width           =   1185
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde.Emb"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   28
      Left            =   180
      TabIndex        =   30
      Top             =   4290
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Número do Lote"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   27
      Left            =   6660
      TabIndex        =   29
      Top             =   3780
      Width           =   930
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Núm. Scheda-Serial"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   26
      Left            =   5040
      TabIndex        =   28
      Top             =   3780
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "D.inter. Fornecedor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   25
      Left            =   3180
      TabIndex        =   27
      Top             =   3780
      Width           =   1110
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Data Prod."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   24
      Left            =   2310
      TabIndex        =   26
      Top             =   3840
      Width           =   585
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Embarq. Control"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   23
      Left            =   1110
      TabIndex        =   25
      Top             =   3840
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Data Exped."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   22
      Left            =   180
      TabIndex        =   24
      Top             =   3840
      Width           =   660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Indicação Suplementar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   21
      Left            =   6660
      TabIndex        =   23
      Top             =   3450
      Width           =   1275
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Lote sob desvio"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   20
      Left            =   5370
      TabIndex        =   22
      Top             =   3450
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Código Fornecedor"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   19
      Left            =   3180
      TabIndex        =   21
      Top             =   3450
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde. Lote"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   18
      Left            =   2310
      TabIndex        =   20
      Top             =   3480
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Nota Fiscal"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   17
      Left            =   1110
      TabIndex        =   19
      Top             =   3480
      Width           =   780
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Peso Bruto(Kg)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   16
      Left            =   180
      TabIndex        =   18
      Top             =   3480
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Vínculos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   15
      Left            =   1830
      TabIndex        =   17
      Top             =   3030
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Clas.Func."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   14
      Left            =   1110
      TabIndex        =   16
      Top             =   3030
      Width           =   555
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Restrições"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   13
      Left            =   180
      TabIndex        =   15
      Top             =   3030
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Forn.\Razão Social"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   11
      Left            =   4500
      TabIndex        =   13
      Top             =   3120
      Width           =   1020
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   10
      Left            =   7020
      TabIndex        =   12
      Top             =   2820
      Width           =   660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Control.Log/Qual."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   9
      Left            =   5700
      TabIndex        =   11
      Top             =   2820
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição do Produto"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   8
      Left            =   2820
      TabIndex        =   10
      Top             =   2790
      Width           =   1230
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Desenho(Part number)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   7
      Left            =   2820
      TabIndex        =   9
      Top             =   2070
      Width           =   1320
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Incoterms"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   6
      Left            =   2820
      TabIndex        =   8
      Top             =   1650
      Width           =   570
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Doca"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   5
      Left            =   2820
      TabIndex        =   7
      Top             =   1050
      Width           =   285
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pto.Entrega"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   4
      Left            =   2820
      TabIndex        =   6
      Top             =   540
      Width           =   645
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   3
      Left            =   2820
      TabIndex        =   5
      Top             =   60
      Width           =   450
   End
   Begin VB.Label lbl_29_Cod_Destino 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
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
      Left            =   1320
      TabIndex        =   2
      Top             =   210
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DUM"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   60
      Width           =   300
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Cód.Dest"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   0
      Left            =   1290
      TabIndex        =   0
      Top             =   60
      Width           =   525
   End
   Begin VB.Line Line30 
      BorderWidth     =   2
      X1              =   5010
      X2              =   5010
      Y1              =   4200
      Y2              =   3780
   End
   Begin VB.Line Line29 
      BorderWidth     =   2
      X1              =   5340
      X2              =   5340
      Y1              =   3750
      Y2              =   3480
   End
   Begin VB.Line Line28 
      BorderWidth     =   2
      X1              =   6630
      X2              =   6630
      Y1              =   4200
      Y2              =   3480
   End
   Begin VB.Line Line27 
      BorderWidth     =   2
      X1              =   4470
      X2              =   4470
      Y1              =   3450
      Y2              =   3150
   End
   Begin VB.Line Line26 
      BorderWidth     =   2
      X1              =   6150
      X2              =   6150
      Y1              =   3450
      Y2              =   3150
   End
   Begin VB.Line Line25 
      BorderWidth     =   2
      X1              =   5670
      X2              =   5670
      Y1              =   3150
      Y2              =   2820
   End
   Begin VB.Line Line24 
      BorderWidth     =   2
      X1              =   6990
      X2              =   6990
      Y1              =   3450
      Y2              =   2820
   End
   Begin VB.Line Line23 
      BorderWidth     =   2
      X1              =   6990
      X2              =   2790
      Y1              =   3150
      Y2              =   3150
   End
   Begin VB.Line Line22 
      BorderWidth     =   2
      X1              =   8400
      X2              =   2790
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Line Line21 
      BorderWidth     =   2
      X1              =   8400
      X2              =   2790
      Y1              =   2070
      Y2              =   2070
   End
   Begin VB.Line Line20 
      BorderWidth     =   2
      X1              =   3690
      X2              =   2760
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line18 
      BorderWidth     =   2
      X1              =   1560
      X2              =   1560
      Y1              =   5160
      Y2              =   4740
   End
   Begin VB.Line Line17 
      BorderWidth     =   2
      X1              =   810
      X2              =   810
      Y1              =   4740
      Y2              =   4290
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   1800
      X2              =   1800
      Y1              =   3450
      Y2              =   3030
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   1080
      X2              =   1080
      Y1              =   3030
      Y2              =   4290
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   2280
      X2              =   2280
      Y1              =   3480
      Y2              =   5160
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   2250
      X2              =   150
      Y1              =   4740
      Y2              =   4725
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   2250
      X2              =   150
      Y1              =   4290
      Y2              =   4290
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   2790
      X2              =   150
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   1890
      X2              =   1890
      Y1              =   540
      Y2              =   60
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000007&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   5085
      Left            =   120
      Top             =   90
      Width           =   8325
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   3150
      X2              =   3150
      Y1              =   3480
      Y2              =   4170
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   8430
      X2              =   2280
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   1260
      X2              =   1260
      Y1              =   540
      Y2              =   60
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   3150
      X2              =   150
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   2790
      X2              =   2790
      Y1              =   3480
      Y2              =   60
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   3690
      X2              =   150
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   8430
      X2              =   3150
      Y1              =   3780
      Y2              =   3780
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   3690
      X2              =   3690
      Y1              =   60
      Y2              =   2070
   End
   Begin VB.Label lbl_26_Descricao_Produto 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Forro de Teto"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2850
      TabIndex        =   41
      Top             =   2940
      Width           =   780
   End
   Begin VB.Label lbl_10_Control_Log_Qual 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "OK"
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
      Left            =   6060
      TabIndex        =   42
      Top             =   2940
      Width           =   210
   End
   Begin VB.Label lbl_23_Razao_Social 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "MUSASHI DO BRASIL LTDA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4560
      TabIndex        =   45
      Top             =   3270
      Width           =   1455
   End
   Begin VB.Label lbl_01_Peso_Bruto 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "10000"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   210
      TabIndex        =   49
      Top             =   3600
      Width           =   450
   End
   Begin VB.Label lbl_12_Num_Doc_Fiscal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "10000"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1380
      TabIndex        =   50
      Top             =   3600
      Width           =   585
   End
   Begin VB.Label lbl_21_Qtde_Lote 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "30000"
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
      Left            =   2460
      TabIndex        =   51
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label lbl_11_Cod_Fornecedor 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "0123456"
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
      Left            =   3720
      TabIndex        =   52
      Top             =   3570
      Width           =   525
   End
   Begin VB.Label lbl_13_Lote_Sob_Desv 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "123/4"
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
      Left            =   5640
      TabIndex        =   53
      Top             =   3570
      Width           =   345
   End
   Begin VB.Label lbl_18_Indicacao_Supl 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "PREPIL"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7170
      TabIndex        =   54
      Top             =   3570
      Width           =   510
   End
   Begin VB.Label lbl_03_Data_Producao 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "20/05/2020"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2355
      TabIndex        =   57
      Top             =   3960
      Width           =   705
   End
   Begin VB.Label lbl_20_Dados_Transporte 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Transporte"
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
      Left            =   870
      TabIndex        =   62
      Top             =   4440
      Width           =   1005
   End
   Begin VB.Label lbl_04_Data_validade 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "05/05/2020"
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
      Left            =   300
      TabIndex        =   63
      Top             =   4890
      Width           =   930
   End
   Begin VB.Label lbl_09_Ponto_Entrega 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "89"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3030
      TabIndex        =   36
      Top             =   690
      Width           =   315
   End
   Begin VB.Image IMG_32_PDF417 
      Height          =   1935
      Left            =   3900
      Stretch         =   -1  'True
      Top             =   210
      Width           =   4335
   End
   Begin VB.Label lbl_06_Cod_Emb 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "030F"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6345
      TabIndex        =   44
      Top             =   3270
      Width           =   375
   End
   Begin VB.Label lbl_25_Desenho_Chrysler 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "9658965896"
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
      Left            =   2880
      TabIndex        =   4
      Top             =   3240
      Width           =   900
   End
End
Attribute VB_Name = "frmEtiquetaFiatLatamFCAEtiq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sTextoQrCode As String
Public sTextoPDF417 As String

Private Function Limpar_campos()

Me.lbl_01_Peso_Bruto.Caption = ""
Me.lbl_02_ODM.Caption = ""
Me.lbl_03_Data_Producao.Caption = ""
Me.lbl_04_Data_validade.Caption = ""
Me.lbl_05_Data_Expedicao.Caption = ""
Me.lbl_06_Cod_Emb.Caption = ""
Me.lbl_07_Quantidade.Caption = ""
Me.lbl_08_DOCA.Caption = ""
Me.lbl_09_Ponto_Entrega.Caption = ""
Me.lbl_10_Control_Log_Qual.Caption = ""
Me.lbl_11_Cod_Fornecedor.Caption = ""
Me.lbl_12_Num_Doc_Fiscal.Caption = ""
Me.lbl_13_Lote_Sob_Desv.Caption = ""
Me.lbl_14_Qtde_emb.Caption = ""
Me.lbl_15_Num_Sheda_Serial.Caption = ""
Me.lbl_16_Id_Inter_Fornecedor.Caption = ""
Me.lbl_17_Embarque_Ctrl.Caption = ""
Me.lbl_18_Indicacao_Supl.Caption = ""
Me.lbl_19_Classe_Func.Caption = ""
Me.lbl_20_Dados_Transporte.Caption = ""
Me.lbl_21_Qtde_Lote.Caption = ""
Me.lbl_22_Num_Lote.Caption = ""
Me.lbl_23_Razao_Social.Caption = ""
Me.lbl_25_Desenho_Chrysler.Caption = ""
Me.lbl_26_Descricao_Produto.Caption = ""
'Me.lbl_27_Num_Desenho.Caption = ""
Me.lbl_28_Destino.Caption = ""
Me.lbl_29_Cod_Destino.Caption = ""
Me.lbl_30_Vinculo.Caption = ""
Me.lbl_31_Restricoes.Caption = ""
Me.lbl_34_DUM.Caption = ""
Me.lbl_36_Incoterms.Caption = ""

End Function

Private Sub cmd_imprime_Click()
    Dim Vezes As Integer
    Dim nSequencial As Integer ' sequencial para impressao do codigo de barras
    Dim nx As Double
    Dim x As Printer
               
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
       Exit Sub
    End If
    nx = 0

    Me.PrintForm

        Me.cmd_imprime.Visible = True
End Sub
Private Sub CarregarFoto()
'Dim Imagem As String
'Dim N As Integer
'
'
''Pega o caminho para a imagem com o Nome do TextBox'
'Imagem = App.Path & "\barcode\output.jpg" '    App.Path & "imagem do projeto" & txtNome.Text & ".jpg"
''Verifica se a imagem existe para evitar algum erro'
'N = 0
'INICIO:
'If N > 12 Then
'   If Dir$(Imagem) <> "" Then Kill App.Path & "\barcode\output.jpg"
'   Exit Sub
'End If
''Sleep 3000
'If Dir$(Imagem) = "" Then
''   Sleep 500
'   N = N + 1
'   GoTo INICIO
'Else
''Carrega a imagem caso exista'
'   Image1.Picture = LoadPicture(Imagem)
'End If
'Kill App.Path & "\barcode\output.jpg"
End Sub

Private Sub Form_Load()
Me.Top = 0
Call Limpar_campos

End Sub
