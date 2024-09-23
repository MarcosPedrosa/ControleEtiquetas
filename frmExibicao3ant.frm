VERSION 5.00
Begin VB.Form frmExibicao3ant 
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prévia de impressão - FORD - antiga"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   3840
      X2              =   3840
      Y1              =   2280
      Y2              =   1200
   End
   Begin VB.Label lblNumSerial 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "100000001"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1320
      TabIndex        =   43
      Top             =   3360
      Width           =   2760
   End
   Begin VB.Label lblNumPeca2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NÚMERO DA PEÇA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   165
      TabIndex        =   42
      Top             =   135
      Width           =   1560
   End
   Begin VB.Label lblDestino2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DESTINO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4560
      TabIndex        =   41
      Top             =   3360
      Width           =   795
   End
   Begin VB.Label lblLinhaUtil2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "LINHA UTILIZAÇÃO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6600
      TabIndex        =   40
      Top             =   2280
      Width           =   1620
   End
   Begin VB.Label lblNumSerial2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº SERIAL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   39
      Top             =   3360
      Width           =   885
   End
   Begin VB.Label lblLinhaUtil 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   38.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   6600
      TabIndex        =   38
      Top             =   2400
      Width           =   1680
   End
   Begin VB.Label lblPc 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2760
      TabIndex        =   37
      Top             =   1440
      Width           =   405
   End
   Begin VB.Label lblCodUtil2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CÓDIGO UTILIZAÇÃO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4560
      TabIndex        =   36
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblQtd2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTIDADE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   35
      Top             =   1200
      Width           =   1140
   End
   Begin VB.Label lblNumFornec2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº FORNECEDOR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   34
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label lblSufixo2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SUFIXO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3960
      TabIndex        =   33
      Top             =   1200
      Width           =   660
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   8400
      X2              =   120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   8400
      X2              =   120
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label lblCodUtil 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "5678"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   38.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   4590
      TabIndex        =   32
      Top             =   2400
      Width           =   1680
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   5400
      Left            =   120
      Top             =   120
      Width           =   8295
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   8400
      X2              =   120
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   6480
      X2              =   6480
      Y1              =   2280
      Y2              =   3360
   End
   Begin VB.Label lblQtdA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   600
      TabIndex        =   31
      Top             =   1680
      Width           =   630
   End
   Begin VB.Label lblQtdB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
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
      Left            =   600
      TabIndex        =   30
      Top             =   1920
      Width           =   630
   End
   Begin VB.Label lblSufixoA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A5TBA"
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
      Left            =   4800
      TabIndex        =   29
      Top             =   1680
      Width           =   1050
   End
   Begin VB.Label lblSufixoB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A5TBA"
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
      Left            =   4800
      TabIndex        =   28
      Top             =   1920
      Width           =   1050
   End
   Begin VB.Label lblNumFornecA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z999A"
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
      Left            =   960
      TabIndex        =   27
      Top             =   2760
      Width           =   1050
   End
   Begin VB.Label lblNumFornecB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z999A"
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
      Left            =   960
      TabIndex        =   26
      Top             =   3000
      Width           =   1050
   End
   Begin VB.Label lblNumFornec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z999A"
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
      Left            =   1920
      TabIndex        =   25
      Top             =   2280
      Width           =   870
   End
   Begin VB.Label lblNumSerialA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100000001"
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
      Left            =   360
      TabIndex        =   24
      Top             =   3960
      Width           =   1890
   End
   Begin VB.Label lblNumSerialB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100000001"
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
      Left            =   360
      TabIndex        =   23
      Top             =   3720
      Width           =   1890
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4440
      X2              =   4440
      Y1              =   2280
      Y2              =   5520
   End
   Begin VB.Label lblQtd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1560
      TabIndex        =   22
      Top             =   1200
      Width           =   720
   End
   Begin VB.Label lblMSB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MUSASHI - PRAÇA MOTOGEAR,111,CRUZ DE REBOUÇAS, IGARASSU, PE"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   120
      TabIndex        =   21
      Top             =   4320
      Width           =   4260
   End
   Begin VB.Label lblSufixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A5TBA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4680
      TabIndex        =   20
      Top             =   1200
      Width           =   1260
   End
   Begin VB.Label lblNumPeca 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E6VB 54233A33"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3165
      TabIndex        =   19
      Top             =   120
      Width           =   3075
   End
   Begin VB.Label lblNumPecaB 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "E6VB54233A33"
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
      Left            =   3450
      TabIndex        =   18
      Top             =   600
      Width           =   2520
   End
   Begin VB.Label lblNumPecaA 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E6VB54233A33"
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
      Left            =   3450
      TabIndex        =   17
      Top             =   840
      Width           =   2520
   End
   Begin VB.Label lblFordMotor 
      BackStyle       =   0  'Transparent
      Caption         =   "FORD MOTOR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   16
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblEndereco1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AV. CHARLES SCHNEIDER"
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
      Left            =   6360
      TabIndex        =   15
      Top             =   4125
      Width           =   1875
   End
   Begin VB.Label lblEndereco2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº 2222 - TAUBATÉ - SP"
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
      Left            =   6360
      TabIndex        =   14
      Top             =   4320
      Width           =   1740
   End
   Begin VB.Label lblDestino 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "WIXON"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   4515
      TabIndex        =   13
      Top             =   3555
      Width           =   1695
   End
   Begin VB.Label lblDestinoA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6360
      TabIndex        =   12
      Top             =   3480
      Width           =   165
   End
   Begin VB.Label lblDestinoB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6360
      TabIndex        =   11
      Top             =   3720
      Width           =   165
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lote MSB"
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Código MSB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   9
      Top             =   360
      Width           =   870
   End
   Begin VB.Label lblLote 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "LOTE MSB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   8
      Top             =   885
      Width           =   930
   End
   Begin VB.Label lblCod_Peca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO MSB "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   7
      Top             =   525
      Width           =   1245
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
      TabIndex        =   6
      Top             =   5040
      Width           =   45
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   8400
      X2              =   120
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label lblCodigoBarrasB 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   120
      TabIndex        =   5
      Top             =   5160
      Width           =   4320
   End
   Begin VB.Label lblCodigoBarrasA 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   120
      TabIndex        =   4
      Top             =   4950
      Width           =   4320
   End
   Begin VB.Label lblCodigoBarras 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   720
      TabIndex        =   3
      Top             =   4680
      Width           =   3000
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "COD MUSASHI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   4560
      Width           =   1245
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
      Left            =   7320
      TabIndex        =   1
      Top             =   5280
      Width           =   975
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
      Left            =   150
      TabIndex        =   0
      Top             =   5550
      Width           =   900
   End
End
Attribute VB_Name = "frmExibicao3ant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Limpar campos
    'FORD
    frmExibicao3ant.lblNumPeca.Caption = ""
    frmExibicao3ant.lblNumPecaA.Caption = ""
    frmExibicao3ant.lblNumPecaB.Caption = ""
    frmExibicao3ant.lblLote.Caption = ""
    frmExibicao3ant.lblQtd.Caption = ""
    frmExibicao3ant.lblQtdA.Caption = ""
    frmExibicao3ant.lblQtdB.Caption = ""
    frmExibicao3ant.lblSufixo.Caption = ""
    frmExibicao3ant.lblSufixoA.Caption = ""
    frmExibicao3ant.lblSufixoB.Caption = ""
    frmExibicao3ant.lblNumFornec.Caption = ""
    frmExibicao3ant.lblNumFornecA.Caption = ""
    frmExibicao3ant.lblNumFornecB.Caption = ""
    frmExibicao3ant.lblCodUtil.Caption = ""
    frmExibicao3ant.lblLinhaUtil.Caption = ""
    frmExibicao3ant.lblNumSerial.Caption = ""
    frmExibicao3ant.lblNumSerialA.Caption = ""
    frmExibicao3ant.lblNumSerialB.Caption = ""
    frmExibicao3ant.lblDestino.Caption = ""
    frmExibicao3ant.lblDestinoA.Caption = ""
    frmExibicao3ant.lblDestinoB.Caption = ""
    
    Me.Height = 6630
    Me.Width = 9780
    Me.Top = 0
    
    Dim v As Integer
    For v = 0 To (Forms.Count - 1)
        If Forms(v).Name = "frmFord" Then
            Me.Left = frmFord.Width
            Exit Sub
        End If
        Me.Left = frmOpcoes.Width
    Next
    Me.lbl_data.Caption = Format(Now(), "dd/mm/yyyy")
End Sub

