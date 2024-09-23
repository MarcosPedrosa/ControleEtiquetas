VERSION 5.00
Begin VB.Form frmExibicao9 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Prévia de Impressão - Palete GM"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PICT_CRISTAL 
      Height          =   525
      Left            =   150
      Picture         =   "frmExibicao9.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   33
      Top             =   300
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   345
      Left            =   7800
      TabIndex        =   32
      Top             =   2670
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   31
      Top             =   3210
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label LBL_MES 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "XXX"
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
      Left            =   2190
      TabIndex        =   34
      Top             =   5310
      Width           =   270
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
      Left            =   1200
      TabIndex        =   30
      Top             =   5280
      Width           =   900
   End
   Begin VB.Label lbl_id_etiqueta 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "              ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2100
      TabIndex        =   29
      Top             =   4530
      Width           =   675
   End
   Begin VB.Label lblCodFunc 
      Alignment       =   1  'Right Justify
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
      Left            =   6900
      TabIndex        =   28
      Top             =   5310
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MUSASHI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6180
      TabIndex        =   27
      Top             =   4860
      Width           =   1020
   End
   Begin VB.Label lblCodBar_Cod_cliente1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "*632503750014*"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2100
      TabIndex        =   26
      Top             =   1680
      Width           =   3570
   End
   Begin VB.Label lblCodBar_Desvio_Aviso_Mod1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*999999070312000004*"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2100
      TabIndex        =   25
      Top             =   4050
      Width           =   4800
   End
   Begin VB.Line Line4 
      X1              =   1380
      X2              =   7320
      Y1              =   4530
      Y2              =   4530
   End
   Begin VB.Line Line3 
      X1              =   1380
      X2              =   7320
      Y1              =   3735
      Y2              =   3735
   End
   Begin VB.Line Line2 
      X1              =   1380
      X2              =   7290
      Y1              =   2955
      Y2              =   2955
   End
   Begin VB.Line Line1 
      X1              =   1380
      X2              =   7350
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label lblCodBar_Desvio_Aviso_Mod 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*999999070312000004*"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2100
      TabIndex        =   24
      Top             =   4170
      Width           =   4800
   End
   Begin VB.Label lblCodBar_Cod_Peca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2W01210"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2100
      TabIndex        =   23
      Top             =   4770
      Width           =   2100
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Id:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1560
      TabIndex        =   22
      Top             =   3765
      Width           =   240
   End
   Begin VB.Label lbl_Cod_Peca 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "                X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   6210
      TabIndex        =   21
      Top             =   4650
      Width           =   825
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CM:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   20
      Top             =   4560
      Width           =   435
   End
   Begin VB.Label lblCodBar_lote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*0609015*"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2100
      TabIndex        =   19
      Top             =   3240
      Width           =   2700
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "PCs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6450
      TabIndex        =   18
      Top             =   2520
      Width           =   585
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Qty:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1560
      TabIndex        =   17
      Top             =   2370
      Width           =   510
   End
   Begin VB.Label lbl_data_expedicao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "12/03/2007"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   5910
      TabIndex        =   16
      Top             =   2040
      Width           =   930
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5190
      TabIndex        =   15
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label lbl_Fornecedor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MUSASHI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3510
      TabIndex        =   14
      Top             =   2070
      Width           =   825
   End
   Begin VB.Label lblCod_Fornecedor 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "F1250"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2610
      TabIndex        =   13
      Top             =   2070
      Width           =   585
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1560
      TabIndex        =   12
      Top             =   2040
      Width           =   1005
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Part:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1560
      TabIndex        =   11
      Top             =   750
      Width           =   570
   End
   Begin VB.Label LblVesao 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "V.2.3.4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6660
      TabIndex        =   10
      Top             =   390
      Width           =   765
   End
   Begin VB.Label lblCliente 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MWM INTERNACIONAL MOTORES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1140
      TabIndex        =   9
      Top             =   390
      Width           =   5640
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   4605
      Left            =   1050
      Top             =   660
      Width           =   6315
   End
   Begin VB.Label lblLote1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "LOT:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1560
      TabIndex        =   8
      Top             =   3000
      Width           =   540
   End
   Begin VB.Label lbl_Cod_cliente 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "90OB52L21I6G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2100
      TabIndex        =   7
      Top             =   720
      Width           =   3255
   End
   Begin VB.Label lbl_desc_peca 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TUBO ALIMENTADORwwwwwwwwwww"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2100
      TabIndex        =   6
      Top             =   1260
      Width           =   4830
   End
   Begin VB.Label lblqtd_caixa 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5880
      TabIndex        =   5
      Top             =   2370
      Width           =   540
   End
   Begin VB.Label lblLote 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "               ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2100
      TabIndex        =   4
      Top             =   3000
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   5235
      Left            =   900
      Top             =   330
      Width           =   6705
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
      Left            =   8910
      TabIndex        =   3
      Top             =   5760
      Width           =   45
   End
   Begin VB.Label lblCodBar_Cod_cliente 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "*632503750014*"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2100
      TabIndex        =   2
      Top             =   1530
      Width           =   3570
   End
   Begin VB.Label lblCodBarQtd_caixa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "*10*"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2100
      TabIndex        =   1
      Top             =   2460
      Width           =   1200
   End
   Begin VB.Label lblDesvio_Aviso_Mod 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "          ."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2100
      TabIndex        =   0
      Top             =   3765
      Width           =   495
   End
End
Attribute VB_Name = "frmExibicao9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Call Imprime_etiqueta_MWM_Cristal
End Sub

Private Sub Command2_Click()
Me.Command1.Visible = False
Me.Command2.Visible = False
Printer.Orientation = 2
Me.PrintForm
Printer.Orientation = 2: Printer.EndDoc
Me.Command1.Visible = True
Me.Command2.Visible = True
End Sub

Private Sub Form_Load()
    Dim v As Integer
    Me.Top = 0
    For v = 0 To (Forms.Count - 1)
        If Forms(v).Name = "frmExibicao9" Then
            Me.Left = frmExibicao9.Width
            Exit For
        End If
    Next
    lbl_data.Caption = Format(Now(), "dd/mm/yyyy")
End Sub
