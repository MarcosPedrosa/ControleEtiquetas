VERSION 5.00
Begin VB.Form frmEtiquetaHondaManaus 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Etiqueta honda Manaus"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Height          =   5685
      Left            =   -30
      TabIndex        =   5
      Top             =   -90
      Width           =   10815
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "02023656O00010000001501071"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   330
         TabIndex        =   61
         Top             =   780
         Width           =   5745
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   3450
         X2              =   3450
         Y1              =   3090
         Y2              =   3690
      End
      Begin VB.Label LBL_codigo_item_bar1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "28120KWB 9214 M1DA       000000001000PC "
         BeginProperty Font 
            Name            =   "CODE128"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   60
         Top             =   2010
         Width           =   7200
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   960
         X2              =   960
         Y1              =   3090
         Y2              =   3690
      End
      Begin VB.Line Line5 
         BorderWidth     =   2
         X1              =   8010
         X2              =   300
         Y1              =   3090
         Y2              =   3090
      End
      Begin VB.Shape Shp_borda 
         BorderWidth     =   2
         Height          =   4230
         Left            =   120
         Top             =   180
         Visible         =   0   'False
         Width           =   8235
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   8010
         X2              =   300
         Y1              =   3690
         Y2              =   3690
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   8100
         X2              =   300
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Label LBL_empresa 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "HDA"
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
         Left            =   300
         TabIndex        =   59
         Top             =   3270
         Width           =   555
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   2250
         X2              =   2250
         Y1              =   3690
         Y2              =   3090
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MUSASHI DA AMAZ�NIA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   300
         TabIndex        =   58
         Top             =   450
         Width           =   3360
      End
      Begin VB.Label LBL_pedido 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "02023692"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3780
         TabIndex        =   57
         Top             =   450
         Width           =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5070
         TabIndex        =   56
         Top             =   450
         Width           =   90
      End
      Begin VB.Label LBL_tipo_pedido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "O0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5130
         TabIndex        =   55
         Top             =   450
         Width           =   390
      End
      Begin VB.Label LBL_linha_pedido 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "001000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5550
         TabIndex        =   54
         Top             =   450
         Width           =   990
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5490
         TabIndex        =   53
         Top             =   450
         Width           =   90
      End
      Begin VB.Label LBL_nota_fiscal 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "000150111"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6570
         TabIndex        =   52
         Top             =   450
         Width           =   1485
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8025
         TabIndex        =   51
         Top             =   450
         Width           =   90
      End
      Begin VB.Label LBL_serie 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8070
         TabIndex        =   50
         Top             =   450
         Width           =   165
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "FORNECEDOR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   300
         TabIndex        =   49
         Top             =   300
         Width           =   540
      End
      Begin VB.Label LBL_NUMTIPLIN_PEDIDO 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "*02023656O00010000001501071 *"
         BeginProperty Font 
            Name            =   "CODE128"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   48
         Top             =   4680
         Width           =   6225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "NRO PEDIDO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   4050
         TabIndex        =   47
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "NOTA FISCAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   6750
         TabIndex        =   46
         Top             =   300
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   300
         TabIndex        =   45
         Top             =   1560
         Width           =   180
      End
      Begin VB.Label LBL_codigo_item 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "__14100KVS7410"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   300
         TabIndex        =   44
         Top             =   1650
         Width           =   2415
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "DESCRI��O DO ITEM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   300
         TabIndex        =   43
         Top             =   2670
         Width           =   765
      End
      Begin VB.Label LBL_descricao 
         BackStyle       =   0  'Transparent
         Caption         =   "COMANDO DE VALVULAS KVS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   300
         TabIndex        =   42
         Top             =   2790
         Width           =   3525
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTIDADE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   4620
         TabIndex        =   41
         Top             =   2670
         Width           =   510
      End
      Begin VB.Label LBL_quantidade 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "150,000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4095
         TabIndex        =   40
         Top             =   2760
         Width           =   1065
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "UN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   5580
         TabIndex        =   39
         Top             =   2670
         Width           =   120
      End
      Begin VB.Label LBL_unidade 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5445
         TabIndex        =   38
         Top             =   2760
         Width           =   405
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "QTD/VOL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   6000
         TabIndex        =   37
         Top             =   2670
         Width           =   360
      End
      Begin VB.Label LBL_volume_total 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "003 / 094"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5955
         TabIndex        =   36
         Top             =   2760
         Width           =   1230
      End
      Begin VB.Label LBLpsv1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PSV"
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
         Left            =   7275
         TabIndex        =   35
         Top             =   2790
         Width           =   705
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "EMPRESA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   300
         TabIndex        =   34
         Top             =   3120
         Width           =   345
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "DATA DE ENTREGA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   1050
         TabIndex        =   33
         Top             =   3120
         Width           =   690
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "HORA DE ENTREGA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   2340
         TabIndex        =   32
         Top             =   3150
         Width           =   720
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "SETOR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   3540
         TabIndex        =   31
         Top             =   3150
         Width           =   255
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "LOCAL DE ENTREGA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   5250
         TabIndex        =   30
         Top             =   3150
         Width           =   735
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "OUTROS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   6660
         TabIndex        =   29
         Top             =   3150
         Width           =   330
      End
      Begin VB.Line Line7 
         BorderWidth     =   2
         X1              =   5160
         X2              =   5160
         Y1              =   3090
         Y2              =   3690
      End
      Begin VB.Line Line9 
         BorderWidth     =   2
         X1              =   6540
         X2              =   6540
         Y1              =   3090
         Y2              =   3690
      End
      Begin VB.Label LBL_data 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "13/03/14"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1035
         TabIndex        =   28
         Top             =   3300
         Width           =   1140
      End
      Begin VB.Label LBL_hora_entrega 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "09:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2385
         TabIndex        =   27
         Top             =   3300
         Width           =   765
      End
      Begin VB.Label LBL_setor 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "M MOTOR 1"
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
         Left            =   3540
         TabIndex        =   26
         Top             =   3270
         Width           =   1575
      End
      Begin VB.Label LBL_local_entrega 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "LOCAL01"
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
         Left            =   5250
         TabIndex        =   25
         Top             =   3270
         Width           =   1275
      End
      Begin VB.Label LBL_outros 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0000069595"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6660
         TabIndex        =   24
         Top             =   3360
         Width           =   1320
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "NRO PEDIDO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   330
         TabIndex        =   23
         Top             =   3780
         Width           =   495
      End
      Begin VB.Label LBL_pedido1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "02023692"
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
         Left            =   870
         TabIndex        =   22
         Top             =   3720
         Width           =   840
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   2010
         TabIndex        =   21
         Top             =   3720
         Width           =   60
      End
      Begin VB.Label LBL_linha_pedido1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "001000"
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
         Left            =   2070
         TabIndex        =   20
         Top             =   3720
         Width           =   630
      End
      Begin VB.Label LBL_tipo_pedido2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "O0"
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
         Left            =   1770
         TabIndex        =   19
         Top             =   3720
         Width           =   225
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   1710
         TabIndex        =   18
         Top             =   3720
         Width           =   60
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "NOTA FISCAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   3690
         TabIndex        =   17
         Top             =   3780
         Width           =   480
      End
      Begin VB.Label LBL_serie1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "1"
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
         Left            =   5280
         TabIndex        =   16
         Top             =   3720
         Width           =   105
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   5190
         TabIndex        =   15
         Top             =   3720
         Width           =   60
      End
      Begin VB.Label LBL_nota_fiscal1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "000150111"
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
         Left            =   4230
         TabIndex        =   14
         Top             =   3720
         Width           =   945
      End
      Begin VB.Label LBLpsv2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PSV"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5580
         TabIndex        =   13
         Top             =   3810
         Width           =   2385
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   90
         Left            =   5400
         TabIndex        =   12
         Top             =   4230
         Width           =   180
      End
      Begin VB.Label LBL_codigo_item1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "__14100KVS7410"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   5640
         TabIndex        =   11
         Top             =   4200
         Width           =   870
      End
      Begin VB.Label LBL_NUMTIPLIN_PEDIDO1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "*02023656O00010000001501071 *"
         BeginProperty Font 
            Name            =   "CODE128"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   10
         Top             =   4830
         Width           =   6225
      End
      Begin VB.Label LBL_codigo_item_bar2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "28120KWB 9214 M1DA       000000001000PC "
         BeginProperty Font 
            Name            =   "CODE128"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   9
         Top             =   2280
         Width           =   7200
      End
      Begin VB.Label LBL_NUMTIPLIN_PEDIDO2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "*02023656O00010000001501071 *"
         BeginProperty Font 
            Name            =   "CODE128"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   8
         Top             =   5010
         Width           =   6225
      End
      Begin VB.Label LBL_NUMTIPLIN_PEDIDO3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "*02023656O00010000001501071 *"
         BeginProperty Font 
            Name            =   "CODE128"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   7
         Top             =   5130
         Width           =   6225
      End
      Begin VB.Label LBL_codigo_item_bar_t 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "28120KWB 2344234erererer23424234324"
         BeginProperty Font 
            Name            =   "CODE128"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   330
         TabIndex        =   6
         Top             =   3900
         Width           =   6300
      End
   End
   Begin VB.Frame frm_visible 
      Height          =   1695
      Left            =   -90
      TabIndex        =   0
      Top             =   5730
      Width           =   8415
      Begin VB.ComboBox cbo_impressora 
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
         Left            =   3510
         TabIndex        =   62
         Text            =   "Combo1"
         Top             =   300
         Width           =   4125
      End
      Begin VB.CommandButton CMD_IMP 
         Caption         =   "IMPRESS�O"
         Height          =   315
         Left            =   210
         TabIndex        =   3
         Top             =   300
         Width           =   1695
      End
      Begin VB.TextBox txt_Arq_Importacao 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   240
         MaxLength       =   100
         TabIndex        =   2
         Text            =   "C:\HONDA\ETIQUETAS.TXT"
         ToolTipText     =   "Arquivo que ser� analisado para importa��o dos dados"
         Top             =   930
         Width           =   6075
      End
      Begin VB.TextBox txtlidos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "0"
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Impressora:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2070
         TabIndex        =   63
         Top             =   330
         Width           =   1245
      End
      Begin VB.Label Label45 
         BackColor       =   &H80000009&
         Caption         =   "Lidos:"
         Height          =   255
         Left            =   6510
         TabIndex        =   4
         Top             =   990
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmEtiquetaHondaManaus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Flag_ativo As Boolean

Private Sub CMD_IMP_Click()

Dim nx As Double
Dim x As Printer
           
frm_visible.Visible = False
Shp_borda.Visible = False

nx = 0
For Each x In Printers
   If InStr(1, UCase(x.DeviceName), UCase(Me.cbo_impressora.List(Me.cbo_impressora.ListIndex))) > 0 Then
      Set Printer = x
      nx = 1
      Exit For
   End If
Next

If nx = 0 Then
   MsgBox "Impressora da Etiqueta n�o encontrada, chame o respons�vel. o nome da impressora ser� - 'ETIQUETA FABRICA'"
   Exit Sub
End If

Printer.Orientation = 2
frmEtiquetaHondaManaus.PrintForm
Printer.Orientation = 2: Printer.EndDoc
frm_visible.Visible = True
Shp_borda.Visible = True


End Sub

Private Sub Form_Activate()
If Flag_ativo = True Then
   Exit Sub
End If
Flag_ativo = True
Me.Top = 0
Me.Left = 0
Call Verificar_Rotina_Importacao
End
End Sub

Function Verificar_Rotina_Importacao()

Dim nx As Integer
Dim RESPOSTA As Integer
Dim CONTA As Double
Dim x As Double
Dim Y As Double
Dim Nada As String
Dim XP As Printer

nx = 0
For Each XP In Printers

   If InStr(1, UCase(XP.DeviceName), "ETIQUETA FABRICA") > 0 Then
      Set Printer = XP
      nx = 1
      Exit For
   End If

Next

If nx = 0 Then
   MsgBox "impressora da Etiqueta n�o encontrada, chame o respons�vel. o nome da impressora ser� - 'ETIQUETA FABRICA'"
   Exit Function
End If

Rem verificar exietencia do arquivo

Nada = Me.txt_Arq_Importacao.Text

If Dir$(Nada) = "" Then
   MsgBox "Arquivo de importa��o n�o encontrado, Procure o respons�vel! " & Nada, 16, "Programa Cancelado"
   Exit Function
End If

RESPOSTA = MsgBox("Ler dados DE IMPORTA��O DAS ETIQUETAS?", 20, "Sim/N�o?")
Close #11

Open Me.txt_Arq_Importacao.Text For Random Access Read Write As #11 Len = Len(Arquivo_AM)

If RESPOSTA = 6 Then

'      If Arquivo_AM.final_arq <> Chr$(13) + Chr$(10) Then
'         MsgBox "Arquivo lido , difere do tamanho do correspondente para sua importa��o."
'         Close #11
'         Exit Function
'      End If
      
      Y = LOF(11) / Len(Arquivo_AM)
      txtlidos.Text = Y
      x = 0
      
      For Y = 1 To LOF(11) / Len(Arquivo_AM)
      
        CONTA = CONTA + 1
        Get 11, Y, Arquivo_AM
        
        Me.LBL_pedido.Caption = Trim(Arquivo_AM.pedido)
        Me.LBL_tipo_pedido.Caption = Trim(Arquivo_AM.tipo_pedido)
        Me.LBL_linha_pedido.Caption = Trim(Arquivo_AM.linha_pedido)
        Me.LBL_pedido1.Caption = Trim(Arquivo_AM.pedido)
        Me.LBL_tipo_pedido2.Caption = Trim(Arquivo_AM.tipo_pedido)
        Me.LBL_linha_pedido1.Caption = Trim(Arquivo_AM.linha_pedido)
        Me.LBL_nota_fiscal.Caption = Trim(Arquivo_AM.nota_fiscal)
        Me.LBL_serie.Caption = Trim(Arquivo_AM.serie)
        Me.LBL_nota_fiscal1.Caption = Trim(Arquivo_AM.nota_fiscal)
        Me.LBL_serie1.Caption = Trim(Arquivo_AM.serie)
        
        Rem primeiro codigo de barras(Pedido)
        Me.LBL_NUMTIPLIN_PEDIDO.Caption = Format_Code128(Trim(Arquivo_AM.cod_bar_pedido))
        Me.LBL_NUMTIPLIN_PEDIDO1.Caption = Format_Code128(Trim(Arquivo_AM.cod_bar_pedido))
        Me.LBL_NUMTIPLIN_PEDIDO2.Caption = Format_Code128(Trim(Arquivo_AM.cod_bar_pedido))
        Me.LBL_NUMTIPLIN_PEDIDO3.Caption = Format_Code128(Trim(Arquivo_AM.cod_bar_pedido))
        
        Me.LBL_codigo_item.Caption = Trim(Arquivo_AM.codigo_item)
        Me.LBL_codigo_item1.Caption = Trim(Arquivo_AM.codigo_item)
        
        Rem segundo codigo de barras(item)
        Me.LBL_codigo_item_bar1.Caption = Format_Code128(Trim(Arquivo_AM.cod_bar_item))
        Me.LBL_codigo_item_bar2.Caption = Format_Code128(Trim(Arquivo_AM.cod_bar_item))
        
        Me.LBL_descricao.Caption = Trim(Arquivo_AM.descricao)
        Me.LBL_quantidade.Caption = Trim(Format(VBA.CDbl(Arquivo_AM.quantidade), "####0.000"))
        Me.LBL_unidade.Caption = Trim(Arquivo_AM.unidade)
        Me.LBL_volume_total.Caption = Trim(Arquivo_AM.volume) & "/" & Trim(Arquivo_AM.total_volume)
        Me.LBL_empresa.Caption = Trim(Arquivo_AM.empresa)
        Me.LBL_data.Caption = Trim(Replace(Arquivo_AM.data_entrega, ".", ""))
        Me.LBL_hora_entrega.Caption = Trim(Arquivo_AM.hora_entrega)
        Me.LBL_setor.Caption = Trim(Arquivo_AM.setor)
        Me.LBL_outros.Caption = Trim(Mid$(Trim(Arquivo_AM.outros), InStr(1, Trim(Arquivo_AM.outros), "N�"), Len(Trim(Arquivo_AM.outros))))
        Me.LBL_local_entrega.Caption = Trim(Arquivo_AM.local_entrega)
        Me.LBLpsv1.Caption = Trim(Arquivo_AM.psv)
        Me.LBLpsv2.Caption = Trim(Arquivo_AM.psv)
        
        Rem terceiro codigo de barras(pedido)
        LBL_codigo_item_bar_t.Caption = Format_Code128(Trim(Arquivo_AM.cod_bar_pedido))
        
        Printer.Orientation = 2
        frmEtiquetaHondaManaus.PrintForm
        Printer.Orientation = 2: Printer.EndDoc
      
      Next
'
      Close #11

      Kill Nada

End If

MsgBox "Final de impress�o, o programa ser� finalizado."

End Function


Private Sub Form_Load()

Dim Nada As String
Dim x As Printer
Dim nx As Integer


Rem verificar exietencia do arquivo

Nada = Me.txt_Arq_Importacao.Text

If Dir$(Nada) = "" Then
   MsgBox "Arquivo de importa��o n�o encontrado, Procure o respons�vel! " & Nada, 16, "Programa Cancelado"
   End
End If

nx = 0
For Each x In Printers
    If UCase(Mid$(x.DeviceName, 1, 8)) = "ETIQUETA" Then
       Me.cbo_impressora.AddItem x.DeviceName
    End If
Next

Rem verificar se ha impressoras cadastradas
If Me.cbo_impressora.ListCount = 0 Then
   MsgBox "Impressoras ETIQUETAS,n�o encontradas no sistema, Favor comunicar ao respons�vel para adiciona-las no sistema!"
   End
End If
Me.cbo_impressora.ListIndex = 0

Rem verificar a impressora padr�o para ser usada neste relat�rio
For nx = 0 To Me.cbo_impressora.ListCount - 1
    If Trim(UCase(sImpressoraFabrica)) = Trim(UCase(Me.cbo_impressora.List(nx))) Then
       Me.cbo_impressora.ListIndex = nx
    End If
Next

Me.LBL_pedido.Caption = ""
Me.LBL_tipo_pedido.Caption = ""
Me.LBL_linha_pedido.Caption = ""
Me.LBL_pedido1.Caption = ""
Me.LBL_tipo_pedido2.Caption = ""
Me.LBL_linha_pedido1.Caption = ""
Me.LBL_nota_fiscal.Caption = ""
Me.LBL_serie.Caption = ""
Me.LBL_nota_fiscal1.Caption = ""
Me.LBL_serie1.Caption = ""
Me.LBL_NUMTIPLIN_PEDIDO.Caption = ""

Me.LBL_codigo_item.Caption = ""
Me.LBL_codigo_item1.Caption = ""
Me.LBL_codigo_item_bar1.Caption = ""
Me.LBL_descricao.Caption = ""
Me.LBL_quantidade.Caption = ""
Me.LBL_unidade.Caption = ""
Me.LBL_volume_total.Caption = ""
Me.LBL_empresa.Caption = ""
Me.LBL_data.Caption = ""
Me.LBL_hora_entrega.Caption = ""
Me.LBL_setor.Caption = ""
Me.LBL_outros.Caption = ""
Me.LBL_local_entrega.Caption = ""
Me.LBLpsv1.Caption = ""
Me.LBLpsv2.Caption = ""

End Sub


Function Format_Code128(InString As String) As String
    Dim Sum As Integer, i As Integer
    Dim Checksum As Integer, Checkchar As Integer
    Dim MyString As String, CVal As Integer
    Dim Checkdigit As Integer
    '
    ' Initialize running total with value of
    ' Subset B start character
    '
    Sum = 104
    '
    ' Scan the string and add character value times position
    '
    For i = 1 To Len(InString)
        '
        ' Copy one character from InString position i to MyString
        '
        MyString = Mid$(InString, i, 1)
        '
        ' Get the numeric value of the character and subtract
        ' 32 to shift (the space character, ASCII value 32, has
        ' a numeric value of 0 as far as Code 128 is concerned)
        '
        CVal = Asc(MyString) - 32
        '
        ' Add the weighted value into the running sum
        '
        Sum = Sum + (CVal * i)
    Next i
    '
    ' Calculate the Modulo 103 checksum
    '
    Checksum = Sum Mod 103
    '
    ' Now convert this number to a character.  This conversion
    ' takes into account the particular mapping of the font
    ' being used (this example is for the font published by
    ' Azalea Software.
    '
    If Checksum = 0 Then
        Checkchar = 174
    ElseIf Checksum < 94 Then
        Checkchar = Checksum + 32
    Else
        Checkchar = Checksum + 71
    End If
    '
    ' Now format the final output string: start character,
    ' data, check character, and stop character
    '
'    MyString = Chr(162) + InString + Chr(Checkchar) + Chr(164)
    MyString = Chr(123) + InString + Chr(Checkchar) + Chr(125)
    Format_Code128 = MyString
End Function
