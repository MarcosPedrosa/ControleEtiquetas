VERSION 5.00
Object = "{B907CF17-F019-41BF-A9D4-8B1BEC2FCB54}#1.0#0"; "IDAutomationPDF417.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmExibicao7GM 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pr�via de impress�o - GM"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6480
      TabIndex        =   83
      Top             =   5760
      Width           =   615
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "Code Datamatrix"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   8520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   82
      Text            =   "frmExibicao7GM.frx":0000
      Top             =   330
      Width           =   3165
   End
   Begin VB.TextBox DataToEncodeText2 
      Height          =   1095
      Left            =   10530
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   81
      Text            =   "frmExibicao7GM.frx":001A
      Top             =   3150
      Width           =   2805
   End
   Begin VB.ComboBox CodeType 
      Height          =   315
      ItemData        =   "frmExibicao7GM.frx":0034
      Left            =   10290
      List            =   "frmExibicao7GM.frx":005F
      TabIndex        =   80
      Text            =   "Combo1"
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CheckBox GS1 
      Caption         =   "GS1 Flag"
      Height          =   255
      Left            =   8550
      TabIndex        =   79
      Top             =   2700
      Width           =   1095
   End
   Begin VB.CheckBox Unicheck 
      Caption         =   "encode as Unicode"
      Height          =   255
      Left            =   8520
      TabIndex        =   78
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "AZTW"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   11460
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   5370
      Width           =   2475
   End
   Begin VB.TextBox cols 
      Height          =   285
      Left            =   9750
      TabIndex        =   76
      Text            =   "0"
      Top             =   2670
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   48
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   885
      Left            =   9180
      Locked          =   -1  'True
      TabIndex        =   75
      Top             =   3420
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   10710
      Locked          =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   74
      Text            =   "23234234234234234"
      Top             =   2670
      Width           =   2145
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
      Left            =   1380
      MaxLength       =   14
      TabIndex        =   50
      Text            =   "0.2"
      Top             =   8070
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
      Left            =   1380
      MaxLength       =   14
      TabIndex        =   49
      Text            =   "0.2"
      Top             =   7800
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
      Left            =   1380
      MaxLength       =   14
      TabIndex        =   48
      Text            =   "0.03"
      Top             =   7260
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
      Left            =   1380
      MaxLength       =   14
      TabIndex        =   47
      Text            =   "3,0"
      Top             =   7530
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
      Left            =   4380
      MaxLength       =   14
      TabIndex        =   46
      Text            =   "3"
      Top             =   7350
      Width           =   600
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
      Left            =   4380
      MaxLength       =   14
      TabIndex        =   45
      Text            =   "3"
      Top             =   7620
      Width           =   600
   End
   Begin VB.TextBox ImageWidth 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   960
      MaxLength       =   14
      TabIndex        =   44
      Text            =   "2044"
      Top             =   6840
      Width           =   780
   End
   Begin VB.TextBox ImageHeight 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   120
      MaxLength       =   14
      TabIndex        =   43
      Text            =   "1310"
      Top             =   6840
      Width           =   780
   End
   Begin VB.TextBox DataToEncodeText 
      Height          =   390
      Left            =   9090
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   42
      Text            =   "frmExibicao7GM.frx":00FD
      Top             =   5550
      Width           =   2250
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   47.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000012&
      Height          =   870
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   70
      Top             =   3030
      Width           =   5445
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12330
      Top             =   2070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   1845
      Left            =   5730
      Stretch         =   -1  'True
      Top             =   240
      Width           =   2445
   End
   Begin VB.Line Line12 
      X1              =   8520
      X2              =   12870
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Label lbl_id_etiqueta 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "6JUN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   66
      Top             =   3870
      Width           =   5445
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " DELIVERY NOTE or PUS OR INVOICE NUMBER:"
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
      Left            =   5280
      TabIndex        =   73
      Top             =   4200
      Width           =   2280
   End
   Begin VB.Label Label14 
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
      Left            =   12120
      TabIndex        =   72
      Top             =   1410
      Width           =   1200
   End
   Begin PDF417LibCtl.PDF PDF1 
      Height          =   1185
      Left            =   11610
      TabIndex        =   71
      Top             =   150
      Width           =   2250
      _cx             =   3969
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
      Picture         =   "frmExibicao7GM.frx":0101
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
   Begin VB.Label lblLicenseA 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   30
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   5310
      TabIndex        =   69
      Top             =   5730
      Width           =   1275
   End
   Begin VB.Label lblCodigoBarrasMusashiA 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "*1234567890*"
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
      Height          =   630
      Left            =   11880
      TabIndex        =   68
      Top             =   4320
      Width           =   930
   End
   Begin VB.Label lblCodFunc 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "14490"
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
      Left            =   1590
      TabIndex        =   67
      Top             =   5250
      Width           =   450
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
      Left            =   5670
      TabIndex        =   65
      Top             =   3915
      Width           =   465
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
      Left            =   5670
      TabIndex        =   64
      Top             =   3570
      Width           =   675
   End
   Begin VB.Label Label13 
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
      Left            =   5670
      TabIndex        =   63
      Top             =   3420
      Width           =   930
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
      Left            =   5670
      TabIndex        =   62
      Top             =   3780
      Width           =   825
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY:"
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
      Left            =   240
      TabIndex        =   61
      Top             =   1500
      Width           =   525
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
      Left            =   7470
      TabIndex        =   60
      Top             =   7920
      Width           =   2415
   End
   Begin VB.Label Label22 
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
      Left            =   7170
      TabIndex        =   59
      Top             =   7410
      Width           =   4260
   End
   Begin VB.Label Label21 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Narrow Bar Width"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2010
      TabIndex        =   58
      Top             =   7305
      Width           =   1320
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Left Margin"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2010
      TabIndex        =   57
      Top             =   8115
      Width           =   870
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Top Margin"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2010
      TabIndex        =   56
      Top             =   7845
      Width           =   1005
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X to Y Ratio"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2010
      TabIndex        =   55
      Top             =   7575
      Width           =   960
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PDF Columns"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5055
      TabIndex        =   54
      Top             =   7395
      Width           =   1005
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Error Correction Level"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5055
      TabIndex        =   53
      Top             =   7665
      Width           =   1590
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   735
      TabIndex        =   52
      Top             =   7545
      Width           =   600
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Height"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   735
      TabIndex        =   51
      Top             =   7275
      Width           =   600
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
      Left            =   255
      TabIndex        =   41
      Top             =   360
      Width           =   1995
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   5610
      X2              =   5610
      Y1              =   1500
      Y2              =   4170
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   8340
      X2              =   150
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   8370
      X2              =   150
      Y1              =   2850
      Y2              =   2850
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   40
      Top             =   195
      Width           =   315
   End
   Begin VB.Label lblLicense 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "UN 123456789 A2B4C6D8E"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4620
      TabIndex        =   39
      Top             =   6360
      Width           =   4965
   End
   Begin VB.Label lblEndereco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Av. Antonio Vicente "
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
      Left            =   255
      TabIndex        =   38
      Top             =   570
      Width           =   1770
   End
   Begin VB.Label lblIgarassu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Novelino 111 - IGARASSU"
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
      Left            =   255
      TabIndex        =   37
      Top             =   780
      Width           =   2280
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
      Left            =   255
      TabIndex        =   36
      Top             =   975
      Width           =   1110
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
      Left            =   255
      TabIndex        =   35
      Top             =   1185
      Width           =   1530
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   2610
      X2              =   2610
      Y1              =   1470
      Y2              =   120
   End
   Begin VB.Label lblTo2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TO:"
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
      Left            =   2655
      TabIndex        =   34
      Top             =   195
      Width           =   165
   End
   Begin VB.Label lblPlant2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PLANT/DOCK:"
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
      Left            =   2655
      TabIndex        =   33
      Top             =   1050
      Width           =   660
   End
   Begin VB.Label lblEng2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CROSS WEIGHT"
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
      Left            =   6960
      TabIndex        =   32
      Top             =   5550
      Width           =   765
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   8370
      X2              =   150
      Y1              =   4170
      Y2              =   4185
   End
   Begin VB.Label Label112 
      BackStyle       =   0  'Transparent
      Caption         =   "GENERAL MOTORS DO BRASIL"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   8760
      TabIndex        =   31
      Top             =   5850
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.Label lblPlant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "72479  A215"
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
      Left            =   3360
      TabIndex        =   30
      Top             =   1020
      Width           =   1860
   End
   Begin VB.Label lblCodMSB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2G01740"
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
      Left            =   3330
      TabIndex        =   29
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   3090
      X2              =   3090
      Y1              =   2220
      Y2              =   1530
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   5610
      X2              =   150
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C�d. MUSASHI"
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
      Left            =   240
      TabIndex        =   28
      Top             =   4230
      Width           =   690
   End
   Begin VB.Label lblShipmentDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "12/JAN/2006"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5640
      TabIndex        =   26
      Top             =   3000
      Width           =   1920
   End
   Begin VB.Label lblContainer2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "SHIPMENT DATE"
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
      Left            =   5640
      TabIndex        =   25
      Top             =   2880
      Width           =   780
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
      Left            =   7380
      TabIndex        =   24
      Top             =   5340
      Width           =   45
   End
   Begin VB.Label lblCodigoBarras 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Left            =   8640
      TabIndex        =   23
      Top             =   4500
      Width           =   3000
   End
   Begin VB.Label lblCodigoBarrasA 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   90
      TabIndex        =   22
      Top             =   6120
      Visible         =   0   'False
      Width           =   3600
   End
   Begin VB.Label lblCodigoBarrasB 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Code 128"
         Size            =   72
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   60
      TabIndex        =   21
      Top             =   5790
      Width           =   4680
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   5250
      X2              =   5250
      Y1              =   4200
      Y2              =   5220
   End
   Begin VB.Label lblCodigoProduto1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   1080
      TabIndex        =   20
      Top             =   4290
      Width           =   2400
   End
   Begin VB.Label lblQtde1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
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
      Left            =   270
      TabIndex        =   19
      Top             =   1740
      Width           =   360
   End
   Begin VB.Label lblComplPeca1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5700
      TabIndex        =   18
      Top             =   2400
      Width           =   2550
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER"
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
      Left            =   10920
      TabIndex        =   17
      Top             =   6810
      Width           =   1515
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LABEL"
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
      Left            =   10590
      TabIndex        =   16
      Top             =   7410
      Width           =   2175
   End
   Begin VB.Label lblQtdeContainers 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8610
      TabIndex        =   15
      Top             =   3990
      Width           =   300
   End
   Begin VB.Label lblPeso 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "120"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   9990
      TabIndex        =   14
      Top             =   4830
      Width           =   810
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000007&
      BorderWidth     =   2
      Height          =   5145
      Left            =   90
      Top             =   120
      Width           =   8295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PART"
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
      Left            =   240
      TabIndex        =   13
      Top             =   2250
      Width           =   255
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "NUMBER"
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
      Left            =   240
      TabIndex        =   12
      Top             =   2370
      Width           =   420
   End
   Begin VB.Label Label4 
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
      Left            =   3150
      TabIndex        =   11
      Top             =   1500
      Width           =   1350
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "REFERENCE"
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
      Left            =   5670
      TabIndex        =   10
      Top             =   2250
      Width           =   600
   End
   Begin VB.Label lblQtdeTot1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
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
      Left            =   11910
      TabIndex        =   9
      Top             =   5010
      Width           =   675
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   7920
      Shape           =   2  'Oval
      Top             =   5580
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Line Line3 
      Visible         =   0   'False
      X1              =   8160
      X2              =   7920
      Y1              =   5910
      Y2              =   5550
   End
   Begin VB.Line Line9 
      Visible         =   0   'False
      X1              =   8400
      X2              =   8160
      Y1              =   5550
      Y2              =   5910
   End
   Begin VB.Line Line11 
      Visible         =   0   'False
      X1              =   7920
      X2              =   8400
      Y1              =   5550
      Y2              =   5550
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "KG"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   7080
      TabIndex        =   8
      Top             =   6870
      Width           =   750
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTAINERS"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11130
      TabIndex        =   7
      Top             =   7890
      Width           =   1665
   End
   Begin VB.Label lblTo1 
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
      Height          =   255
      Left            =   2790
      TabIndex        =   6
      Top             =   510
      Width           =   2775
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
      Height          =   255
      Left            =   3660
      TabIndex        =   5
      Top             =   6750
      Width           =   2775
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
      Height          =   255
      Left            =   2790
      TabIndex        =   4
      Top             =   270
      Width           =   2385
   End
   Begin VB.Label lblMaterial 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "02W C32"
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
      Left            =   930
      TabIndex        =   3
      Top             =   2250
      Width           =   4515
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
      Left            =   11370
      TabIndex        =   2
      Top             =   4890
      Width           =   360
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
      Left            =   120
      TabIndex        =   1
      Top             =   5250
      Width           =   900
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "TAUBAT�"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2790
      TabIndex        =   0
      Top             =   750
      Width           =   2775
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   27
      Top             =   2880
      Width           =   915
   End
End
Attribute VB_Name = "frmExibicao7GM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sCodigo As String
Private CodeClair$, CodeBarre$
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function Bar2Ddmfx Lib "dmatdfu" (ByVal szIn As String, ByRef N As Long, ByRef ct As Long, ByRef fl As Long, ByRef col As Long, ByRef st As Long, ByRef sl As Long, ByRef szOut As String) As Long
Private Declare Function Bar2Ddmcx Lib "dmatdfu" (ByVal szIn As String, ByRef N As Long, ByRef ct As Long, ByRef fl As Long, ByRef col As Long, ByRef st As Long, ByRef sl As Long) As Long
Private Declare Function Bar2Ddmwx Lib "dmatdfu" (ByVal szIn As String, ByRef N As Long, ByRef ct As Long, ByRef fl As Long, ByRef col As Long, ByRef st As Long, ByRef sl As Long, ByVal szname As String) As Long
Private Declare Function Bar2Ddmd Lib "dmatdfu" (ByVal sz As String, ByRef so As String) As Long
Private Declare Function ErrorMessDM Lib "dmatdfu" (ByRef N As Long, ByRef ss As String) As Long
Private Declare Sub Sleep Lib "KERNEL32" _
        (ByVal dwMilliseconds As Long)

Private Sub CodeType_Click()
Text1_Change

End Sub

Private Sub Command1_Click()
Call CarregarFoto
End Sub

Private Sub DataToEncodeText_Change()

'Call CodeRefresh(lbl_Cod_Fornecedor.Caption)
'lbl_Cod_Fornecedor_Barras.Caption = sCodigo
'
'Call CodeRefresh(lbl_qtd.Caption)
'lbl_qtd_barras.Caption = sCodigo
'
'Call CodeRefresh(lbl_Cod_cliente.Caption)
'Me.lbl_cod_cliente_Barras.Caption = sCodigo
'

'  Dim CodeBarre$
'  CodeBarre$ = code128$(DataToEncodeText.Text)
'  Me.lblLicenseA.Caption = CodeBarre$
'  lblCodigoBarrasB.Caption = CodeBarre$
Call CodeRefresh(Replace(DataToEncodeText, " ", ""))
Me.lblLicenseA.Caption = sCodigo
Text3.Text = sCodigo
Text2.Text = sCodigo
lblCodigoBarrasMusashiA.Caption = sCodigo
'Me.lblLicenseB.Caption = sCodigo

'lblCodigoBarrasMusashiA.Caption = "*" & lblCodigoBarras.Caption & "*"

'Call CodeRefresh(Format(Val(lbl_id_etiqueta.Caption), "000000000"))
'lbl_id_etiqueta_barra.Caption = sCodigo
'lblLicenseA.Caption = sCodigo
'lblLicenseB.Caption = sCodigo

End Sub


Private Sub Form_Load()

    'Limpar Campos
    'lblTo.Caption = ""
    lblPlant.Caption = ""
    lblQtdeContainers.Caption = ""
    lblCodigoProduto1.Caption = ""
    lblQtde1.Caption = ""
    lblComplPeca1.Caption = ""
    
    lblQtdeTot1.Caption = ""
    
    lblLicense.Caption = ""
    lblLicenseA.Caption = ""
'    lblLicenseB.Caption = ""
    lblCodigoBarras.Caption = ""
    lblCodigoBarrasA.Caption = ""
    lblCodigoBarrasB.Caption = ""
    lblCodMSB.Caption = ""
    lblPeso.Caption = ""
    
    'Mostra form
'    Me.Height = 6630
    'Me.Width = 9780
    
    
'    Me.Width = 8565
    Me.Top = 0
    Me.Left = frmOpcoes.Width
    
'    Dim v As Integer
'    For v = 0 To (Forms.Count - 1)
'        If Forms(v).Name = "frmPaleteGm" Then
'            Me.Left = frmPaleteGm.Width
'            Exit For
'        End If
'    Next
    Me.lbl_data.Caption = Format(Now(), "dd/mm/yyyy")
    
'****************************************************************
    cols.Text = "0"
    CodeType.ListIndex = 0
    
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
  'Cette fonction est r�gie par la Licence G�n�rale Publique Amoindrie GNU (GNU LGPL)
  'This function is governed by the GNU Lesser General Public License (GNU LGPL)
  'V 2.0.0
  'Param�tres : une chaine
  'Parameters : a string
  'Retour : * une chaine qui, affich�e avec la police CODE128.TTF, donne le code barre
  '         * une chaine vide si param�tre fourni incorrect
  'Return : * a string which give the bar code when it is dispayed with CODE128.TTF font
  '         * an empty string if the supplied parameter is no good
  Dim i%, Checksum&, mini%, dummy%, tableB As Boolean
  ean128$ = ""
  If Len(chaine$) > 0 Then
  'V�rifier si caract�res valides
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
          'Voir si int�ressant de passer en table C / See if interesting to switch to table C
          'Oui pour 4 chiffres au d�but ou � la fin, sinon pour 6 chiffres / yes for 4 digits at start or end, else if 6 digits
          mini% = IIf(i% = 1 Or i% + 3 = Len(chaine$), 4, 6)
          GoSub TestNumOrFnc1
          If mini% < 0 Then 'Choix table C / Choice of table C
            If i% = 1 Then 'D�buter sur table C / Starting with table C
              ean128$ = Chr$(210)
            Else 'Commuter sur table C / Switch to table C
              ean128$ = ean128$ & Chr$(204)
            End If
            tableB = False
          Else
            If i% = 1 Then ean128$ = Chr$(209) 'D�buter sur table B / Starting with table B
          End If
        End If
        If Not tableB Then
          'On est sur la table C, essayer de traiter 2 chiffres ou �/ We are on table C, try to process 2 digits or �
          If Asc(Mid$(chaine$, i%, 2)) = 207 Then
            'On traite le Fnc1 (�) / We process the Fnc1 (�)
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
          'Traiter 1 caract�re en table B / Process 1 digit with table B
          ean128$ = ean128$ & Mid$(chaine$, i%, 1)
          i% = i% + 1
        End If
      Loop
      'Calcul de la cl� de contr�le / Calculation of the checksum
      For i% = 1 To Len(ean128$)
        dummy% = Asc(Mid$(ean128$, i%, 1))
        dummy% = IIf(dummy% < 127, dummy% - 32, dummy% - 105)
        If i% = 1 Then Checksum& = dummy%
        Checksum& = (Checksum& + (i% - 1) * dummy%) Mod 103
      Next
      'Calcul du code ASCII de la cl� / Calculation of the checksum ASCII code
      Checksum& = IIf(Checksum& < 95, Checksum& + 32, Checksum& + 105)
      'Ajout de la cl� et du STOP / Add the checksum and the STOP
      ean128$ = ean128$ & Chr$(Checksum&) & Chr$(211)
    End If
  End If
  Exit Function
TestNum:
  'si les mini% caract�res � partir de i% sont num�riques, alors mini%=0
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
  'si les mini% caract�res � partir de i% sont num�riques ou  FNC1, alors mini%=0
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

Private Sub DataToEncodeText2_Change()

ImageHeight.Text = 1310
ImageWidth.Text = 2040
PDF1.Height = ImageHeight.Text
PDF1.Width = ImageWidth.Text
PDFColumns.Text = 3
PDF1.PDFColumns = PDFColumns.Text
Me.Text5.Text = DataToEncodeText2.Text
End Sub


Private Sub lblShipmentDate_Change()
    Call CarregarFoto
End Sub

Private Sub lblShipmentDate_Click()
    Call CarregarFoto
End Sub



Private Sub Text1_Change()

Dim Start As Long
Dim security As Long
Dim i As Integer
'Dim fg As Integer
Dim ctype As Long
Dim icol As Long
Dim so As String
Dim sz As String
Dim ss As Long
Dim tt As Long
Dim fg As Long
Dim N As Long

ctype = CodeType.ListIndex
If (ctype < 0) Then ctype = 0
Start = 0
security = 0
so = String(8192, vbNullChar)
i = Val(cols.Text)
If (i > 0) Then security = -i
icol = i
If (Unicheck.Value > 0) Then
    fg = 129
Else
    fg = 1
End If
If GS1.Value <> 0 Then
    Start = 2
    fg = fg + 2
End If
so = String(8192, vbNullChar)
sz = Text1.Text
tt = 0
N = Bar2Ddmfx(ByVal sz, tt, ctype, fg, icol, Start, security, ByVal so)
If (N < 0) Then
    ss = String(60, vbNullChar)
    N = -N
    i = ErrorMessDM(N, ByVal ss)
'    Status.Text = StrConv(ss, vbFromUnicode)
    Text4.Text = ""
Else
'    Status.Text = "OK"
    Text4.Text = StrConv(so, vbFromUnicode)
End If

End Sub

Private Sub CarregarFoto()
Dim Imagem As String
Dim N As Integer


'Pega o caminho para a imagem com o Nome do TextBox'
Imagem = App.Path & "\barcode\output.jpg" '    App.Path & "imagem do projeto" & txtNome.Text & ".jpg"
'Verifica se a imagem existe para evitar algum erro'
N = 0
INICIO:
If N > 12 Then
   If Dir$(Imagem) <> "" Then Kill App.Path & "\barcode\output.jpg"
   Exit Sub
End If
'Sleep 3000
If Dir$(Imagem) = "" Then
   Sleep 500
   N = N + 1
   GoTo INICIO
Else
'Carrega a imagem caso exista'
   Image1.Picture = LoadPicture(Imagem)
End If
Kill App.Path & "\barcode\output.jpg"
End Sub
