VERSION 5.00
Begin VB.Form frmExibicao3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prévia de impressão - FORD"
   ClientHeight    =   6105
   ClientLeft      =   1875
   ClientTop       =   2055
   ClientWidth     =   8745
   Icon            =   "frmExibicao3_BACKUP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblCodigoBarrasB 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "*0000000001*"
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
      Left            =   1800
      TabIndex        =   40
      Top             =   5640
      Width           =   5175
   End
   Begin VB.Label lblCodigoBarrasA 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "*0000000001*"
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
      Left            =   1800
      TabIndex        =   39
      Top             =   5400
      Width           =   5175
   End
   Begin VB.Label lblCodigoBarras 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0000000001"
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
      Left            =   1800
      TabIndex        =   38
      Top             =   5160
      Width           =   5175
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
      TabIndex        =   37
      Top             =   5280
      Width           =   45
   End
   Begin VB.Label lblCod_Peca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO MSB"
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
      Left            =   360
      TabIndex        =   36
      Top             =   840
      Width           =   1215
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
      Left            =   360
      TabIndex        =   35
      Top             =   1245
      Width           =   930
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
      Left            =   360
      TabIndex        =   34
      Top             =   660
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lote MSB"
      Height          =   195
      Left            =   360
      TabIndex        =   33
      Top             =   1065
      Width           =   705
   End
   Begin VB.Label lblDestinoB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Code 39"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   32
      Top             =   4320
      Width           =   240
   End
   Begin VB.Label lblDestinoA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Code 39"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   31
      Top             =   4080
      Width           =   240
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
      Left            =   4635
      TabIndex        =   30
      Top             =   4155
      Width           =   1695
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
      Left            =   6480
      TabIndex        =   29
      Top             =   4920
      Width           =   1740
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
      Left            =   6480
      TabIndex        =   28
      Top             =   4725
      Width           =   1875
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
      Left            =   4680
      TabIndex        =   27
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Label lblNumPecaA 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E6VB54233A33"
      BeginProperty Font 
         Name            =   "Code 39"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   26
      Top             =   1200
      Width           =   3600
   End
   Begin VB.Label lblNumPecaB 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "E6VB54233A33"
      BeginProperty Font 
         Name            =   "Code 39"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      TabIndex        =   25
      Top             =   960
      Width           =   3600
   End
   Begin VB.Label lblNumPeca 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E6VB 54233A33"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2880
      TabIndex        =   24
      Top             =   360
      Width           =   3885
   End
   Begin VB.Label lblSufixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A5TBA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4800
      TabIndex        =   23
      Top             =   1560
      Width           =   1620
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
      Left            =   240
      TabIndex        =   22
      Top             =   4920
      Width           =   4260
   End
   Begin VB.Label lblQtd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1680
      TabIndex        =   21
      Top             =   1560
      Width           =   900
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4560
      X2              =   4560
      Y1              =   2760
      Y2              =   5160
   End
   Begin VB.Label lblNumSerialB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100000001"
      BeginProperty Font 
         Name            =   "Code 39"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   20
      Top             =   4320
      Width           =   2700
   End
   Begin VB.Label lblNumSerialA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100000001"
      BeginProperty Font 
         Name            =   "Code 39"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   480
      TabIndex        =   19
      Top             =   4560
      Width           =   2700
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
      Left            =   2040
      TabIndex        =   18
      Top             =   2760
      Width           =   870
   End
   Begin VB.Label lblNumFornecB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z999A"
      BeginProperty Font 
         Name            =   "Code 39"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   17
      Top             =   3600
      Width           =   1500
   End
   Begin VB.Label lblNumFornecA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z999A"
      BeginProperty Font 
         Name            =   "Code 39"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   16
      Top             =   3360
      Width           =   1500
   End
   Begin VB.Label lblSufixoB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A5TBA"
      BeginProperty Font 
         Name            =   "Code 39"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      TabIndex        =   15
      Top             =   2400
      Width           =   1500
   End
   Begin VB.Label lblSufixoA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A5TBA"
      BeginProperty Font 
         Name            =   "Code 39"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      TabIndex        =   14
      Top             =   2160
      Width           =   1500
   End
   Begin VB.Label lblQtdB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Code 39"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      TabIndex        =   13
      Top             =   2400
      Width           =   900
   End
   Begin VB.Label lblQtdA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Code 39"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   720
      TabIndex        =   12
      Top             =   2160
      Width           =   900
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   6600
      X2              =   6600
      Y1              =   2760
      Y2              =   3960
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   8520
      X2              =   240
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   4800
      Left            =   240
      Top             =   360
      Width           =   8295
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
      Left            =   4710
      TabIndex        =   1
      Top             =   3000
      Width           =   1680
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   8520
      X2              =   240
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   8520
      X2              =   240
      Y1              =   1560
      Y2              =   1560
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
      Left            =   4080
      TabIndex        =   11
      Top             =   1560
      Width           =   660
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
      Left            =   360
      TabIndex        =   10
      Top             =   2760
      Width           =   1455
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
      Left            =   360
      TabIndex        =   9
      Top             =   1560
      Width           =   1140
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
      Left            =   4680
      TabIndex        =   8
      Top             =   2760
      Width           =   1815
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
      Left            =   2880
      TabIndex        =   7
      Top             =   1920
      Width           =   405
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
      Left            =   6720
      TabIndex        =   6
      Top             =   3000
      Width           =   1680
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
      Left            =   360
      TabIndex        =   5
      Top             =   3960
      Width           =   885
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
      Left            =   6720
      TabIndex        =   4
      Top             =   2760
      Width           =   1620
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
      Left            =   4680
      TabIndex        =   3
      Top             =   3960
      Width           =   795
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
      Left            =   285
      TabIndex        =   2
      Top             =   375
      Width           =   1560
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
      Left            =   1440
      TabIndex        =   0
      Top             =   3960
      Width           =   2760
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   3960
      X2              =   3960
      Y1              =   2760
      Y2              =   1560
   End
End
Attribute VB_Name = "frmExibicao3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Limpar campos
    'FORD
    frmExibicao3.lblNumPeca.Caption = ""
    frmExibicao3.lblNumPecaA.Caption = ""
    frmExibicao3.lblNumPecaB.Caption = ""
    frmExibicao3.lblLote.Caption = ""
    frmExibicao3.lblQtd.Caption = ""
    frmExibicao3.lblQtdA.Caption = ""
    frmExibicao3.lblQtdB.Caption = ""
    frmExibicao3.lblSufixo.Caption = ""
    frmExibicao3.lblSufixoA.Caption = ""
    frmExibicao3.lblSufixoB.Caption = ""
    frmExibicao3.lblNumFornec.Caption = ""
    frmExibicao3.lblNumFornecA.Caption = ""
    frmExibicao3.lblNumFornecB.Caption = ""
    frmExibicao3.lblCodUtil.Caption = ""
    frmExibicao3.lblLinhaUtil.Caption = ""
    frmExibicao3.lblNumSerial.Caption = ""
    frmExibicao3.lblNumSerialA.Caption = ""
    frmExibicao3.lblNumSerialB.Caption = ""
    frmExibicao3.lblDestino.Caption = ""
    frmExibicao3.lblDestinoA.Caption = ""
    frmExibicao3.lblDestinoB.Caption = ""
    
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
End Sub
