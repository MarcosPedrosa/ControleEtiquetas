VERSION 5.00
Object = "{B907CF17-F019-41BF-A9D4-8B1BEC2FCB54}#1.0#0"; "IDAutomationPDF417.dll"
Begin VB.Form frmExibicao3 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prévia de impressão - FORD"
   ClientHeight    =   4305
   ClientLeft      =   1875
   ClientTop       =   2055
   ClientWidth     =   6900
   Icon            =   "frmExibicao3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox DataToEncodeText 
      Height          =   390
      Left            =   4770
      MaxLength       =   640
      MultiLine       =   -1  'True
      TabIndex        =   50
      Text            =   "frmExibicao3.frx":030A
      Top             =   6090
      Width           =   2250
   End
   Begin VB.TextBox ImageHeight 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   300
      MaxLength       =   14
      TabIndex        =   41
      Text            =   "1310"
      Top             =   5370
      Width           =   780
   End
   Begin VB.TextBox ImageWidth 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   300
      MaxLength       =   14
      TabIndex        =   40
      Text            =   "2044"
      Top             =   5640
      Width           =   780
   End
   Begin VB.TextBox PDFErrorCorrectionLevel 
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   4800
      MaxLength       =   14
      TabIndex        =   39
      Text            =   "3"
      Top             =   5760
      Width           =   600
   End
   Begin VB.TextBox PDFColumns 
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   4800
      MaxLength       =   14
      TabIndex        =   38
      Text            =   "3"
      Top             =   5490
      Width           =   600
   End
   Begin VB.TextBox W2NRatio 
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   1800
      MaxLength       =   14
      TabIndex        =   37
      Text            =   "3.0"
      Top             =   5670
      Width           =   600
   End
   Begin VB.TextBox NarrowBarWidth 
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.0000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   1800
      MaxLength       =   14
      TabIndex        =   36
      Text            =   "0.03"
      Top             =   5400
      Width           =   600
   End
   Begin VB.TextBox TopMarginCM 
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   1800
      MaxLength       =   14
      TabIndex        =   35
      Text            =   "0.2"
      Top             =   5940
      Width           =   600
   End
   Begin VB.TextBox LeftMarginCM 
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   1800
      MaxLength       =   14
      TabIndex        =   34
      Text            =   "0.2"
      Top             =   6210
      Width           =   600
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   6870
      X2              =   4620
      Y1              =   4230
      Y2              =   4230
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   210
      X2              =   210
      Y1              =   3840
      Y2              =   3210
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   6870
      X2              =   6870
      Y1              =   4230
      Y2              =   3240
   End
   Begin VB.Label lblCodigoBarrasA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "*1234567890*"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   210
      TabIndex        =   56
      Top             =   3840
      Width           =   4290
   End
   Begin VB.Label lblCodigoBarrasB 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "*0005191340*"
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   29.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1740
      TabIndex        =   55
      Top             =   4440
      Width           =   4260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "LOT"
      Height          =   195
      Left            =   4080
      TabIndex        =   14
      Top             =   1260
      Width           =   315
   End
   Begin VB.Label lbl_Cliente 
      BackColor       =   &H8000000E&
      Caption         =   "VISTEON SALINE PLASTICS PLANT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1050
      TabIndex        =   4
      Top             =   150
      Width           =   3555
   End
   Begin VB.Label lbl_Cod_Fornecedor_Barras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "*56"
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   30
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   450
      TabIndex        =   53
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lbl_cod_cliente_Barras 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "*56"
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   30
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   900
      TabIndex        =   52
      Top             =   2310
      Width           =   4905
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "dddddddddd"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   195
      Left            =   5310
      TabIndex        =   51
      Top             =   240
      Width           =   1200
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4770
      X2              =   180
      Y1              =   810
      Y2              =   810
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Height"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1155
      TabIndex        =   49
      Top             =   5415
      Width           =   600
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1155
      TabIndex        =   48
      Top             =   5685
      Width           =   600
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Error Correction Level"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5475
      TabIndex        =   47
      Top             =   5805
      Width           =   1590
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PDF Columns"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5475
      TabIndex        =   46
      Top             =   5535
      Width           =   1005
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X to Y Ratio"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2430
      TabIndex        =   45
      Top             =   5715
      Width           =   960
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Top Margin"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2430
      TabIndex        =   44
      Top             =   5985
      Width           =   1005
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Left Margin"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2430
      TabIndex        =   43
      Top             =   6255
      Width           =   870
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Narrow Bar Width"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2430
      TabIndex        =   42
      Top             =   5445
      Width           =   1320
   End
   Begin VB.Label lbl_to 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "FORD TAUBATE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   4680
      TabIndex        =   32
      Top             =   3480
      Width           =   1050
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "CUST"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4770
      TabIndex        =   30
      Top             =   3690
      Width           =   450
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "MADE IN BRAZIL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3180
      TabIndex        =   29
      Top             =   3660
      Width           =   1020
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "DOC CODE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5760
      TabIndex        =   28
      Top             =   3270
      Width           =   945
   End
   Begin VB.Label lbl_id_etiqueta_barra 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "*56"
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   30
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Left            =   2040
      TabIndex        =   26
      Top             =   4950
      Width           =   2415
   End
   Begin VB.Label lbl_id_etiqueta 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "321654"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   25
      Top             =   3600
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "SERIAL NO (S)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   24
      Top             =   3630
      Width           =   1185
   End
   Begin VB.Label lbl_descr_peca 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "2527 CNSL FRT LO MED DK GRAPH"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   270
      TabIndex        =   23
      Top             =   3420
      Width           =   3915
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4620
      X2              =   4620
      Y1              =   4230
      Y2              =   3240
   End
   Begin VB.Label lbl_cod_peca 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "3S4X-A045A74"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2970
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "PART (P)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   270
      TabIndex        =   16
      Top             =   2340
      Width           =   480
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl_lote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "32903"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   4110
      TabIndex        =   15
      Top             =   1410
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "DATE"
      Height          =   195
      Left            =   2880
      TabIndex        =   11
      Top             =   1680
      Width           =   435
   End
   Begin VB.Label lbl_unidade 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "UN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2430
      TabIndex        =   9
      Top             =   1530
      Width           =   345
   End
   Begin VB.Label lbl_qtd_barras 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   30
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   270
      TabIndex        =   8
      Top             =   1350
      Width           =   1905
   End
   Begin VB.Label lbl_GROSS 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "GROSS WGT"
      Height          =   195
      Left            =   2880
      TabIndex        =   3
      Top             =   1260
      Width           =   1005
   End
   Begin VB.Label lbl_title 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "CONTAINER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2880
      TabIndex        =   2
      Top             =   810
      Width           =   975
   End
   Begin VB.Label lbl_qty 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "QTY (Q)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   270
      TabIndex        =   1
      Top             =   900
      Width           =   405
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl_supp 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "Supp (V)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   150
      Width           =   675
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   3600
      X2              =   3600
      Y1              =   2790
      Y2              =   3240
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   6870
      X2              =   210
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   3180
      Left            =   210
      Top             =   60
      Width           =   6675
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   6840
      X2              =   240
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   2850
      X2              =   2850
      Y1              =   1890
      Y2              =   840
   End
   Begin VB.Label lbl_peso 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      Caption         =   "440"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   1410
      Width           =   975
   End
   Begin VB.Label lbl_Data 
      BackColor       =   &H8000000E&
      Caption         =   "11NOV2003"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3390
      TabIndex        =   13
      Top             =   1620
      Width           =   1785
   End
   Begin VB.Label lbl_doc_code 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "R3"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   6000
      TabIndex        =   33
      Top             =   3540
      Width           =   795
   End
   Begin VB.Label lbl_qtd 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "999999"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   690
      TabIndex        =   7
      Top             =   750
      Width           =   1980
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "STR LOC 1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   2760
      Width           =   885
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "LINE FEED LOC 2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3630
      TabIndex        =   20
      Top             =   2760
      Width           =   1380
   End
   Begin VB.Label lbl_str_loc1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   1230
      TabIndex        =   21
      Top             =   2850
      Width           =   1995
   End
   Begin VB.Label lbl_line_feed_loc2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "6546464"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3690
      TabIndex        =   22
      Top             =   2910
      Width           =   1155
   End
   Begin VB.Label lbl_container 
      BackColor       =   &H8000000E&
      Caption         =   "SK32 C630L"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2880
      TabIndex        =   10
      Top             =   960
      Width           =   1665
   End
   Begin VB.Label lbl_Cod_Fornecedor 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "CFE0A"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2880
      TabIndex        =   5
      Top             =   330
      Width           =   1470
   End
   Begin VB.Label lbl_cust 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "FI05D"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4710
      TabIndex        =   31
      Top             =   3900
      Width           =   870
   End
   Begin VB.Label lbl_cod_cliente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "999999"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   300
      TabIndex        =   17
      Top             =   1890
      Width           =   6450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4680
      TabIndex        =   27
      Top             =   3240
      Width           =   240
   End
   Begin VB.Label lbl_cod_cliente_1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "3S4X-A045A74"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   270
      TabIndex        =   54
      Top             =   3240
      Width           =   1245
   End
   Begin PDF417LibCtl.PDF PDF1 
      Height          =   1185
      Left            =   4830
      TabIndex        =   6
      Top             =   240
      Width           =   1830
      _cx             =   3228
      _cy             =   2090
      BackColor       =   16777215
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Picture         =   "frmExibicao3.frx":030E
      DataToEncode    =   "IDAutomation.com Metafile Image Generator Example"
      Orientation     =   0
      XtoYRatio       =   3
      NarrowBarCM     =   0,03
      LeftMarginCM    =   0,2
      TopMarginCM     =   0,2
      Truncated       =   0
      PDFRows         =   0
      PDFColumns      =   5
      PDFErrorCorrectionLevel=   2
      PDFMode         =   0
      ApplyTilde      =   1
      FixedResolutionCM=   0
      MacroPDFEnable  =   0
      MacroPDFFileID  =   0
      MacroPDFSegmentIndex=   0
      MacroPDFLastSegment=   0
      WhiteBarIncrease=   0
   End
End
Attribute VB_Name = "frmExibicao3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sCodigo As String
Public nTamWidth As Integer

Private Sub DataToEncodeText_Change()
ImageHeight.Text = 1310
ImageWidth.Text = 2040
PDF1.Height = ImageHeight.Text
PDF1.Width = ImageWidth.Text
PDFColumns.Text = 3
PDF1.PDFColumns = PDFColumns.Text

Call CodeRefresh(lbl_Cod_Fornecedor.Caption)
lbl_Cod_Fornecedor_Barras.Caption = sCodigo

Call CodeRefresh(lbl_qtd.Caption)
lbl_qtd_barras.Caption = sCodigo

Call CodeRefresh(lbl_cod_cliente.Caption)
Me.lbl_cod_cliente_Barras.Caption = sCodigo

Call CodeRefresh(lbl_id_etiqueta.Caption)
Me.lblCodigoBarrasB.Caption = sCodigo

'lblCodigoBarrasB.Caption = "*" & lbl_id_etiqueta.Caption & "*"

lblCodigoBarrasA.Caption = "*" & lbl_id_etiqueta.Caption & "*"

Call CodeRefresh(Format(Val(lbl_id_etiqueta.Caption), "000000000"))
lbl_id_etiqueta_barra.Caption = sCodigo

End Sub

Private Sub Form_Activate()

'    PDF1.Height = ImageHeight.Text
'    PDF1.Width = ImageWidth.Text
'    PDF1.PDFColumns = PDFColumns.Text
'    PDF1.PDFErrorCorrectionLevel = PDFErrorCorrectionLevel.Text
'    PDF1.NarrowBarCM = NarrowBarWidth.Text
'    PDF1.TopMarginCM = TopMarginCM.Text
'    PDF1.LeftMarginCM = LeftMarginCM.Text

End Sub

Private Sub Form_Load()
    
    frmExibicao3.lbl_Cliente.Caption = ""
    frmExibicao3.lbl_Cod_Fornecedor.Caption = ""
    frmExibicao3.lbl_Cod_Fornecedor_Barras.Caption = ""
    frmExibicao3.lbl_qtd.Caption = ""
    frmExibicao3.lbl_peso.Caption = ""
        frmExibicao3.lbl_container.Caption = ""
    frmExibicao3.lbl_lote.Caption = ""
    frmExibicao3.lbl_Data.Caption = ""
    frmExibicao3.lbl_cod_cliente.Caption = ""
    frmExibicao3.lbl_cod_cliente_Barras.Caption = ""
    frmExibicao3.lbl_cod_peca.Caption = ""
    frmExibicao3.lbl_descr_peca.Caption = ""
    frmExibicao3.lbl_id_etiqueta.Caption = ""
    frmExibicao3.lbl_id_etiqueta_barra.Caption = ""
    frmExibicao3.DataToEncodeText.Text = ""
    
    frmExibicao3.lbl_to.Caption = ""
    frmExibicao3.lbl_cust.Caption = ""
    frmExibicao3.lbl_doc_code.Caption = ""
    
'    Me.Height = 6630
'    Me.Width = 9780
    Me.Top = 0
    
    Dim v As Integer
    For v = 0 To (Forms.Count - 1)
        If Forms(v).Name = "frmFord" Then
            Me.Left = frmFord.Width
            Exit Sub
        End If
        Me.Left = nTamWidth
    Next
    
   
End Sub

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

