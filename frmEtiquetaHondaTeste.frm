VERSION 5.00
Begin VB.Form frmEtiquetaHondaTeste 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "teste"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10365
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   10365
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   9840
      TabIndex        =   37
      Top             =   6420
      Width           =   465
   End
   Begin VB.CommandButton cmd_imprime 
      Caption         =   "imprime"
      Height          =   285
      Left            =   7740
      TabIndex        =   23
      Top             =   6720
      Width           =   1785
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame4"
      Height          =   6165
      Left            =   60
      TabIndex        =   18
      Top             =   0
      Width           =   10155
      Begin VB.PictureBox pic_codbar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4170
         ScaleHeight     =   375
         ScaleWidth      =   5955
         TabIndex        =   34
         Top             =   2190
         Width           =   5955
      End
      Begin VB.PictureBox pic_codbar1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4140
         ScaleHeight     =   375
         ScaleWidth      =   5955
         TabIndex        =   33
         Top             =   2820
         Width           =   5955
      End
      Begin VB.PictureBox pic_codbar2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3990
         ScaleHeight     =   375
         ScaleWidth      =   5955
         TabIndex        =   32
         Top             =   3570
         Width           =   5955
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000009&
         Height          =   1395
         Left            =   300
         TabIndex        =   29
         Top             =   390
         Width           =   9345
         Begin VB.Label LBL_NUMTIPLIN_PEDIDO12 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "{12}"
            BeginProperty Font 
               Name            =   "Code 128"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1440
            TabIndex        =   31
            Top             =   450
            Width           =   480
         End
         Begin VB.Label LBL_NUMTIPLIN_PEDIDO21 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "02023660O00010000001501101"
            BeginProperty Font 
               Name            =   "Code 128"
               Size            =   21.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   5700
            TabIndex        =   30
            Top             =   360
            Width           =   3120
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1725
         Left            =   510
         Picture         =   "frmEtiquetaHondaTeste.frx":0000
         ScaleHeight     =   1695
         ScaleWidth      =   4845
         TabIndex        =   22
         Top             =   180
         Visible         =   0   'False
         Width           =   4875
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1620
         Left            =   5460
         Picture         =   "frmEtiquetaHondaTeste.frx":18CD2
         ScaleHeight     =   1590
         ScaleWidth      =   4710
         TabIndex        =   21
         Top             =   270
         Visible         =   0   'False
         Width           =   4740
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000009&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1095
         Left            =   5970
         TabIndex        =   19
         Top             =   4650
         Width           =   3435
         Begin VB.Label Label4 
            Caption         =   "LBL_NUMTIPLIN_PEDIDO"
            Height          =   225
            Left            =   330
            TabIndex        =   20
            Top             =   210
            Width           =   2265
         End
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "{12}"
         BeginProperty Font 
            Name            =   "IDAutomationSC128XXL DEMO"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2715
         Left            =   3420
         TabIndex        =   36
         Top             =   3810
         Width           =   1140
      End
      Begin VB.Label LBL_NUMTIPLIN_PEDIDO23 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "{12}"
         BeginProperty Font 
            Name            =   "CODE128"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1050
         TabIndex        =   35
         Top             =   4410
         Width           =   1140
      End
      Begin VB.Label LBL_NUMTIPLIN_PEDIDO1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "{12}"
         BeginProperty Font 
            Name            =   "CODE128"
            Size            =   69.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1395
         Left            =   330
         TabIndex        =   28
         Top             =   3090
         Width           =   2820
      End
      Begin VB.Label LBL_NUMTIPLIN_PECA2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "28120KWB 9214 M1DA       000000100000PC"
         BeginProperty Font 
            Name            =   "CODE128"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   60
         TabIndex        =   27
         Top             =   840
         Width           =   8190
      End
      Begin VB.Label LBL_NUMTIPLIN_PECA 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "28120KWB 9214 M1DA       000000100000PC"
         BeginProperty Font 
            Name            =   "CODE128"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   60
         TabIndex        =   26
         Top             =   180
         Width           =   8190
      End
      Begin VB.Label LBL_NUMTIPLIN_PECA1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "28120KWB 9214 M1DA       000000100000PC"
         BeginProperty Font 
            Name            =   "CODE128"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   60
         TabIndex        =   25
         Top             =   540
         Width           =   8190
      End
      Begin VB.Label lbl_code39 
         AutoSize        =   -1  'True
         Caption         =   "1"
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
         Left            =   8490
         TabIndex        =   24
         Top             =   4080
         Width           =   180
      End
   End
   Begin VB.CommandButton cmd_fechar 
      Caption         =   "close"
      Height          =   645
      Left            =   7590
      TabIndex        =   14
      Top             =   4500
      Width           =   1515
   End
   Begin VB.CommandButton CMD_IMP 
      Caption         =   "calcula"
      Height          =   315
      Left            =   7710
      TabIndex        =   13
      Top             =   6300
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Frame1"
      Height          =   3945
      Left            =   180
      TabIndex        =   7
      Top             =   30
      Visible         =   0   'False
      Width           =   8955
      Begin VB.Label Label3 
         Caption         =   "LBL_codigo_item_bar_t"
         Height          =   225
         Left            =   180
         TabIndex        =   17
         Top             =   3090
         Width           =   1755
      End
      Begin VB.Label Label2 
         Caption         =   "LBL_codigo_item_bar2"
         Height          =   285
         Left            =   150
         TabIndex        =   16
         Top             =   1980
         Width           =   1725
      End
      Begin VB.Label Label1 
         Caption         =   "LBL_codigo_item_bar1"
         Height          =   195
         Left            =   210
         TabIndex        =   15
         Top             =   1260
         Width           =   1755
      End
      Begin VB.Label LBL_codigo_item_bar1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01200599O00010000000021521"
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
         Left            =   1440
         TabIndex        =   12
         Top             =   1560
         Width           =   5685
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   7980
         X2              =   270
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   7890
         X2              =   90
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Label LBL_codigo_item_bar2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01200599O00010000000021521"
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
         Left            =   1440
         TabIndex        =   11
         Top             =   2370
         Width           =   4680
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
         Left            =   210
         TabIndex        =   10
         Top             =   600
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
         Left            =   210
         TabIndex        =   9
         Top             =   750
         Width           =   6225
      End
      Begin VB.Label LBL_codigo_item_bar_t 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "01200599O00010000000021521"
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
         Left            =   840
         TabIndex        =   8
         Top             =   3450
         Width           =   4680
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1785
      Left            =   960
      Picture         =   "frmEtiquetaHondaTeste.frx":31F44
      ScaleHeight     =   1725
      ScaleWidth      =   6165
      TabIndex        =   6
      Top             =   3960
      Width           =   6225
   End
   Begin VB.Frame frm_visible 
      Height          =   1695
      Left            =   60
      TabIndex        =   0
      Top             =   5580
      Width           =   7515
      Begin VB.TextBox txtlidos 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Height          =   315
         Left            =   5940
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0"
         Top             =   960
         Width           =   555
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
         TabIndex        =   1
         Text            =   "C:\HONDA\ETIQUETAS.TXT"
         ToolTipText     =   "Arquivo que será analisado para importação dos dados"
         Top             =   930
         Width           =   4905
      End
      Begin VB.Label Label45 
         BackColor       =   &H80000009&
         Caption         =   "Lidos:"
         Height          =   255
         Left            =   5370
         TabIndex        =   3
         Top             =   990
         Width           =   555
      End
   End
   Begin VB.Label LBL_cod_bar_pedido 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "02023656O00010000001501071 "
      BeginProperty Font 
         Name            =   "CODE128"
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
      Top             =   5130
      Width           =   3690
   End
   Begin VB.Label LBL_cod_bar_pedido1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "02023656O00010000001501071 "
      BeginProperty Font 
         Name            =   "CODE128"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3660
      TabIndex        =   4
      Top             =   5280
      Width           =   3690
   End
End
Attribute VB_Name = "frmEtiquetaHondaTeste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Flag_ativo As Boolean
Dim I As Integer
Dim J As Integer
Dim F As Integer
Dim DataToPrint As String
Dim OnlyCorrectData As String
Dim PrintableString As String
Dim Encoding As String
Dim WeightedTotal As Long
Dim WeightValue As Integer
Dim CurrentValue As Long
Dim CheckDigitValue As Integer
Dim Factor As Integer
Dim CheckDigit As Integer
Dim CurrentEncoding As String
Dim NewLine As String
Dim msg As String
Dim CurrentChar As String
Dim CurrentCharNum As Integer
Dim C128_StartA As String
Dim C128_StartB As String
Dim C128_StartC As String
Dim C128_Stop As String
Dim C128Start As String
Dim C128CheckDigit As String
Dim StartCode As String
Dim StopCode As String
Dim Fnc1 As String
Dim LeadingDigit As Integer
Dim EAN2AddOn As String
Dim EAN5AddOn As String
Dim EANAddOnToPrint As String
Dim HumanReadableText As String
Dim StringLength As Integer
Dim CorrectFNC As Integer
Dim CID As Integer
Dim FID As Integer
Dim NCID As Integer


Private Sub cmd_fechar_Click()
Unload Me
End Sub

Private Sub CMD_IMP_Click()


'        Me.LBL_NUMTIPLIN_PECA.Caption = Format_Code128("28120KWB 9214 M1DA       000000100000PC")
'        Me.LBL_NUMTIPLIN_PECA1.Caption = Format_Code128("28120KWB 9214 M1DA       000000100000PC")
'        Me.LBL_NUMTIPLIN_PECA2.Caption = Format_Code128("28120KWB 9214 M1DA       000000100000PC")
'        LBL_NUMTIPLIN_PEDIDO1.Caption = Format_Code128("02023660O00010000001501101")
'        LBL_NUMTIPLIN_PEDIDO12.Caption = Format_Code128("02023660O00010000001501101")
        LBL_NUMTIPLIN_PEDIDO1.Caption = Format_Code128("02023660O00010000001501101")
        LBL_NUMTIPLIN_PEDIDO12.Caption = Format_Code128("02023660O00010000001501101")
        LBL_NUMTIPLIN_PEDIDO21.Caption = Format_Code128("02023660O00010000001501101")
        LBL_NUMTIPLIN_PEDIDO23.Caption = Format_Code128("02023660O00010000001501101")
 
        Call preenche_pic



''        Me.LBL_NUMTIPLIN_PEDIDO1.Caption = Chr(123) & "12" & Chr(93) & Chr(126)
'        Me.LBL_NUMTIPLIN_PEDIDO2.Caption = Chr(123) & "01200599O00010000000021521" & Chr(93) & Chr(126)
'        Me.LBL_NUMTIPLIN_PEDIDO3.Caption = Chr(123) & "01200599O00010000000021521" & Chr(93) & Chr(126)
''        Me.LBL_codigo_item.Caption = Trim(Arquivo_AM.codigo_item)
''        Me.LBL_codigo_item1.Caption = Trim(Arquivo_AM.codigo_item)
'        Me.LBL_codigo_item_bar1.Caption = Chr(123) & "01200599O00010000000021521" & Chr(93) & Chr(126)
'        Me.LBL_codigo_item_bar2.Caption = Chr(123) & "01200599O00010000000021521" & Chr(93) & Chr(126)
'        lbl_code39.Caption = Azalea_Code_128_A("1")
''       LBL_NUMTIPLIN_PEDIDO1.Caption = "{11~  "



End Sub

Private Sub cmd_imprime_Click()
Dim RESPOSTA As Integer
Dim CONTA As Double
Dim X As Double
Dim Y As Double
Dim Nada As String
Dim XP As Printer
Dim svar As String * 1
Dim nx As Integer

nx = 0
For Each XP In Printers

   If InStr(1, UCase(XP.DeviceName), "ETIQUETA FABRICA") > 0 Then
      Set Printer = XP
      nx = 1
      Exit For
   End If

Next

Printer.Orientation = 2
Me.PrintForm
Printer.EndDoc

End Sub

Private Sub Command1_Click()
LBL_NUMTIPLIN_PEDIDO23.Caption = Code128(1)
'"01200599O00010000000021521"
End Sub

Private Sub Form_Activate()
If Flag_ativo = True Then
   Exit Sub
End If
Flag_ativo = True
Me.Top = 0
Me.Left = 0
'Call Verificar_Rotina_Importacao
'End
End Sub

Function Verificar_Rotina_Importacao()

Dim nx As Integer
Dim RESPOSTA As Integer
Dim CONTA As Double
Dim X As Double
Dim Y As Double
Dim Nada As String
Dim XP As Printer
Dim svar As String * 1
Dim svar1 As String * 1

'nx = 0
'For Each XP In Printers
'
'   If InStr(1, UCase(XP.DeviceName), "ETIQUETA FABRICA") > 0 Then
'      Set Printer = XP
'      nx = 1
'      Exit For
'   End If
'
'Next
'
'If nx = 0 Then
'   MsgBox "impressora da Etiqueta não encontrada, chame o responsável. o nome da impressora será - 'ETIQUETA FABRICA'"
'   Exit Function
'End If
'
'Rem verificar exietencia do arquivo
'
'Nada = Me.txt_Arq_Importacao.Text
'
'If Dir$(Nada) = "" Then
'   MsgBox "Arquivo de importação não encontrado, Procure o responsável! " & Nada, 16, "Programa Cancelado"
'   Exit Function
'End If
'
'RESPOSTA = MsgBox("Ler dados DE IMPORTAÇÃO DAS ETIQUETAS?", 20, "Sim/Não?")
'Close #11
'
'Open Me.txt_Arq_Importacao.Text For Random Access Read Write As #11 Len = Len(Arquivo_AM)
'
'If RESPOSTA = 6 Then
'        If CONTA = 1 Then svar = ""
'        If CONTA = 2 Then svar = "*"
'        If CONTA = 3 Then svar = "!"
'        svar = "{"
'        svar1 = "?"
'        Me.LBL_pedido.Caption = Trim(Arquivo_AM.pedido)
'        Me.LBL_tipo_pedido.Caption = Trim(Arquivo_AM.tipo_pedido)
'        Me.LBL_linha_pedido.Caption = Trim(Arquivo_AM.linha_pedido)
'        Me.LBL_pedido1.Caption = Trim(Arquivo_AM.pedido)
'        Me.LBL_tipo_pedido2.Caption = Trim(Arquivo_AM.tipo_pedido)
'        Me.LBL_linha_pedido1.Caption = Trim(Arquivo_AM.linha_pedido)
'        Me.LBL_nota_fiscal.Caption = Trim(Arquivo_AM.nota_fiscal)
'        Me.LBL_serie.Caption = Trim(Arquivo_AM.serie)
'        Me.LBL_nota_fiscal1.Caption = Trim(Arquivo_AM.nota_fiscal)
'        Me.LBL_serie1.Caption = Trim(Arquivo_AM.serie)
        Me.LBL_NUMTIPLIN_PEDIDO.Caption = Chr(123) & "01200599O00010000000021521" & Chr(93) & Chr(126)
        Me.LBL_NUMTIPLIN_PEDIDO1.Caption = Chr(123) & "01200599O00010000000021521" & Chr(93) & Chr(126)
        Me.LBL_NUMTIPLIN_PEDIDO2.Caption = Chr(123) & "01200599O00010000000021521" & Chr(93) & Chr(126)
        Me.LBL_NUMTIPLIN_PEDIDO3.Caption = Chr(123) & "01200599O00010000000021521" & Chr(93) & Chr(126)
'        Me.LBL_codigo_item.Caption = Trim(Arquivo_AM.codigo_item)
'        Me.LBL_codigo_item1.Caption = Trim(Arquivo_AM.codigo_item)
        Me.LBL_codigo_item_bar1.Caption = Chr(123) & "01200599O00010000000021521" & Chr(93) & Chr(126)
        Me.LBL_codigo_item_bar2.Caption = Chr(123) & "01200599O00010000000021521" & Chr(93) & Chr(126)
'        Me.LBL_descricao.Caption = Trim(Arquivo_AM.descricao)
'        Me.LBL_quantidade.Caption = Trim(Format(VBA.CDbl(Arquivo_AM.quantidade), "####0.000"))
'        Me.LBL_unidade.Caption = Trim(Arquivo_AM.unidade)
'        Me.LBL_volume_total.Caption = Trim(Arquivo_AM.volume) & "/" & Trim(Arquivo_AM.total_volume)
'        Me.LBL_empresa.Caption = Trim(Arquivo_AM.empresa)
'        Me.LBL_data.Caption = Trim(Replace(Arquivo_AM.data_entrega, ".", ""))
'        Me.LBL_hora_entrega.Caption = Trim(Arquivo_AM.hora_entrega)
'        Me.LBL_setor.Caption = Trim(Arquivo_AM.setor)
'        Me.LBL_outros.Caption = Trim(Mid$(Trim(Arquivo_AM.outros), InStr(1, Trim(Arquivo_AM.outros), "Nº"), Len(Trim(Arquivo_AM.outros))))
'        Me.LBL_cod_bar_pedido.Caption = svar & Arquivo_AM.cod_bar_pedido & svar1
'        Me.LBL_cod_bar_pedido1.Caption = svar & Arquivo_AM.cod_bar_pedido & svar1
'        Me.LBL_local_entrega.Caption = Trim(Arquivo_AM.local_entrega)
'        Me.LBLpsv1.Caption = Trim(Arquivo_AM.psv)
'        Me.LBLpsv2.Caption = Trim(Arquivo_AM.psv)
        LBL_codigo_item_bar_t.Caption = Format_Code128(Arquivo_AM.cod_bar_item)
'        Printer.Orientation = 2
'        frmEtiquetaHondaManaus.PrintForm
'        Printer.EndDoc
'
'      Next
''
'      Close #11
'
'      Kill Nada
'
'End If
'
'MsgBox "Final de impressão, o programa será finalizado."

End Function


Private Sub Form_Load()
'
'Dim Nada As String
'
'Rem verificar exietencia do arquivo
'
'Nada = Me.txt_Arq_Importacao.Text
'
'If Dir$(Nada) = "" Then
'   MsgBox "Arquivo de importação não encontrado, Procure o responsável! " & Nada, 16, "Programa Cancelado"
'   End
'End If

End Sub


Function Format_Code128(InString As String) As String
    Dim Sum As Double
    Dim I As Integer
    Dim Checksum As Double
    Dim Checkchar As Double
    Dim MyString As String
    Dim CVal As Double
    Dim CheckDigit As Double
    '
    ' Initialize running total with value of
    ' Subset B start character
    '
    Sum = 104
    '
    ' Scan the string and add character value times position
    '
    For I = 1 To Len(InString)
        '
        ' Copy one character from InString position i to MyString
        '
        MyString = Mid$(InString, I, 1)
        '
        ' Get the numeric value of the character and subtract
        ' 32 to shift (the space character, ASCII value 32, has
        ' a numeric value of 0 as far as Code 128 is concerned)
        '
        CVal = Asc(MyString) - 32
        '
        ' Add the weighted value into the running sum
        '
        Sum = Sum + (CVal * I)
    Next I
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
Function Azalea_Code_128_A(ByVal yourData As String) As String
' C128Tools 29may12 jwhiting
' Copyright 2012 Azalea Software, Inc. All rights reserved. www.azalea.com

' Creating a Code 128 code set A barcode using C128Tools.
' Your input, yourData, is a string to be encoded as a Code 128 code set A symbol.
' yourData must be the Code 128 code set A character set. Input error checking is your responsibility.

  Dim temp As String                 ' a temporary placeholder
  Dim chunk As String                ' loop chunk
  Dim I As Integer                   ' our loop counter
  Dim checkDigitSubtotal As Integer  ' a check digit throwaway
  Dim e As Integer                   ' a placeholder variable
    
  ' seed the variables
  temp = Chr(181)                   ' code set A start glyph
  checkDigitSubtotal = 103                ' code set A start checkdigit value

  ' map the input string into the C128Tools character set
  For I = 1 To Len(yourData) Step 1
    chunk = Mid(yourData, I, 1)
    Select Case Asc(chunk)
      ' map from above ASCII 182 placeholders to the font's character assignments
      Case Is > 95
        temp = temp & Chr(Asc(chunk) - 66)
      Case Is = 32 ' move the space character
        temp = temp & Chr(206)
      Case Else
        temp = temp & chunk
    End Select
  Next I

  ' Calculate the Code 128 mod 103 check digit
  For I = 1 To Len(yourData)
    e = Asc(Mid(yourData, I, 1)) - 32
    If e <> 142 Then
      checkDigitSubtotal = checkDigitSubtotal + (e * I)
    End If
  Next I
  checkDigitSubtotal = checkDigitSubtotal Mod 103

  ' Put together the final output string
  ' code set A start bar, the encoded string, check digit, stop bar
  Select Case checkDigitSubtotal
    Case 0
      Azalea_Code_128_A = temp & Chr(206) & Chr(196)
    Case 1 To 93
      Azalea_Code_128_A = temp & Chr(checkDigitSubtotal + 32) & Chr(196)
    Case Is > 93
      Azalea_Code_128_A = temp & Chr(checkDigitSubtotal + 103) & Chr(196)
  End Select

End Function

Function Azalea_Code_128_B(ByVal yourData As String) As String
' C128Tools 29may12 jwhiting
' Copyright 2012 Azalea Software, Inc. All rights reserved. www.azalea.com

' Creating a Code 128 code set B barcode using C128Tools.
' Your input, yourData, is a string to be encoded as a Code 128 code set B symbol.
' yourData must be the Code 128 code set B character set. Input error checking is your responsibility.

  Dim temp As String                 ' a temporary placeholder
  Dim chunk As String                ' loop chunk
  Dim I As Integer                   ' our loop counter
  Dim checkDigitSubtotal As Integer  ' a check digit throwaway
  Dim e As Integer                   ' a placeholder variable
    
  ' seed the variables
  temp = Chr(182)                   ' code set B start glyph
  checkDigitSubtotal = 104          ' code set B start checkdigit value

  ' map the input string into the C128Tools character set
  For I = 1 To Len(yourData) Step 1
    chunk = Mid(yourData, I, 1)
    Select Case Asc(chunk)
      ' map from above ASCII 200 to the actual character assignments
      Case Is > 200
        temp = temp & Chr(Asc(chunk) - 35)
        ' The space character is at ASCII 194 because TrueType
        ' doesn't allow a glyph in the ASCII 32 slot
      Case Is = 32
        temp = temp & Chr(206)
      Case Else
        temp = temp & chunk
    End Select
  Next I

  ' Calculate the Code 128 mod 103 check digit
  For I = 1 To Len(yourData)
    e = Asc(Mid(yourData, I, 1)) - 32
    If e <> 142 Then
      checkDigitSubtotal = checkDigitSubtotal + (e * I)
    End If
  Next I
  checkDigitSubtotal = checkDigitSubtotal Mod 103

  ' Put together the final output string
  ' code set A start bar, the encoded string, check digit, stop bar
  Select Case checkDigitSubtotal
    Case 0
      Azalea_Code_128_B = temp & Chr(206) & Chr(196)
    Case 1 To 93
      Azalea_Code_128_B = temp & Chr(checkDigitSubtotal + 32) & Chr(196)
    Case Is > 93
      Azalea_Code_128_B = temp & Chr(checkDigitSubtotal + 103) & Chr(196)
  End Select

End Function

Function Azalea_Code_128_C(ByVal yourData As String) As String
' C128Tools 29may12 jwhiting
' Copyright 2012 Azalea Software, Inc. All rights reserved. www.azalea.com

' Creating a Code 128 code set C barcode using C128Tools.
' Your input, yourData, is a string to be encoded as a Code 128 code set C symbol.
' yourData must be the Code 128 code set C character set. Input error checking is your responsibility.

  Dim temp As String                 ' a temporary placeholder
  Dim checkDigitSubtotal As Integer  ' a check digit throwaway
  Dim I As Integer                   ' our loop counter
  Dim fontString As String           ' a temporary placeholder
  Dim chunk As String                ' loop chunk
   
  ' seed the variables
  fontString = Chr(183)                   ' code set C start glyph
  checkDigitSubtotal = 105                ' code set C start checkdigit value
  temp = yourData                         ' 2 character chunks

  ' pad odd length input with a leading zero goesHere PRN
  
  ' Calculate the Code 128 mod 103 check digit
  For I = 1 To Len(yourData) / 2
    chunk = Left$(temp, 2)
    checkDigitSubtotal = checkDigitSubtotal + Val(chunk) * I
    Select Case Val(chunk)
      Case 0
        fontString = fontString + Chr(206)
      Case 1 To 93
        fontString = fontString & Chr(Val(chunk) + 32)
      Case Is > 93
        fontString = fontString & Chr(Val(chunk) + 103)
    End Select
    temp = Right$(temp, Len(temp) - 2)
  Next I
  checkDigitSubtotal = checkDigitSubtotal Mod 103

  ' Put together the final output string
  ' code set C start bar, the encoded string, check digit, stop bar
  Select Case checkDigitSubtotal
    Case 0
      Azalea_Code_128_C = fontString & Chr(206) & Chr(196)
    Case 1 To 93
      Azalea_Code_128_C = fontString & Chr(checkDigitSubtotal + 32) & Chr(196)
    Case Is > 93
      Azalea_Code_128_C = fontString & Chr(checkDigitSubtotal + 103) & Chr(196)
  End Select

End Function

Function Azalea_Code_39(ByVal Code39 As String) As String
' C39Tools 24mar09 jwhiting
' Copyright 2009 Azalea Software, Inc. All rights reserved. www.azalea.com

' Creating a Code 39 barcode in Excel
' Your input, Code39, is a string to be encoded as a Code 39 symbol.
' yourData must be the Code 39 character set. Input error checking is your responsibility.
' The standard Code 39 character set is: A-Z (upppercase), 0-9, $ % + - . / and the space character.

  Dim I As Integer                   ' our loop counter
  Dim chunk As String                ' loop chunk
  Dim temp As String                 ' a temporary placeholder
  
  ' TrueType doesn't support glyphs in the space slot (ASCII 32)
  ' We've moved the space character to the underscore ( _ ).
  ' Therefore "APPLE PIE" is formatted as  *APPLE_PIE*
  ' Here's the search and replace, underscore for space:
  For I = 1 To Len(Code39)
    chunk$ = Mid$(Code39, I, 1)
    If chunk = " " Then
      temp = temp + "_"
    Else
      temp = temp + chunk
    End If
  Next I

  ' Add the start and stop bars, the asterisk, before and after the input string.
  Azalea_Code_39 = "*" + temp + "*"

  ' Excel: B1=Azalea_Code_39(A1)
  ' Or put another way, yourContainer.text=Azalea_Code_39(yourInputString)
  
End Function

Private Function preenche_pic()
  pic_codbar.Cls
  pic_codbar.FontName = "CODE128"   ' nome da fonte usada
  pic_codbar.FontSize = 12                ' tamanho da fonte usada
  pic_codbar.CurrentX = 10.3               ' posiciona
  pic_codbar.CurrentY = 3.3
  pic_codbar.Print Format_Code128("02023660O00010000001501101")             ' imprime no picturebox  codigo gerado
  
  pic_codbar1.Cls
  pic_codbar1.FontName = "CODE128"   ' nome da fonte usada
  pic_codbar1.FontSize = 12                ' tamanho da fonte usada
  pic_codbar1.CurrentX = 10.3               ' posiciona
  pic_codbar1.CurrentY = 3.3
  pic_codbar1.Print Format_Code128("02023660O00010000001501101")             ' imprime no picturebox  codigo gerado
  
  pic_codbar2.Cls
  pic_codbar2.FontName = "CODE128"   ' nome da fonte usada
  pic_codbar2.FontSize = 12                ' tamanho da fonte usada
  pic_codbar2.CurrentX = 10.3               ' posiciona
  pic_codbar2.CurrentY = 3.3
  pic_codbar2.Print Format_Code128("02023660O00010000001501101")             ' imprime no picturebox  codigo gerado





End Function

'*********************************************************************
'*  IDAutomation Barcode Font Formulas for Crystal Reports 4.01
'*  Copyright, IDAutomation.com, Inc. 2000-2004. All rights reserved.
'*
'*  You MUST use the fully functional Code 128 (dated 12/2000 or later)
'*  font for this code to create and print a proper barcode
'*
'*  To create UCC/EAN128 barcodes, use the appropriate
'*  ASCII 0202 and AIs included as documented at:
'*  http://www.idautomation.com/code128faq.html#EAN128andUCC128
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid License from IDAutomation.com, Inc.
'*  for the associated font and the copyright notices are not
'*  removed from the source code.
'*
'*  Formula: Code128 - formats output to IDAutomationC128 fonts
'*  Tutorial: http://www.BizFonts.com/crystal/
'*********************************************************************
Function elabora_code_128()

Dim DataToEncode As String
'Change the next line to connect to your data source; for example:
'DataToEncode = ({Table.Field})
DataToEncode = "Ê8100712345Ê2112345678"

    Dim PrintableString As String
    Dim DataToFormat As String
    Dim WeightedTotal As Double
    Dim CurrentValue As Double
    Dim CheckDigitValue As Double
    Dim C128CheckDigit As String
    Dim StringLength As Double
    Dim I As Double
    Dim CurrentCharNum As Double
    Dim CurrentEncoding As String
    Dim C128Start As String
    Dim CorrectFNC As Double
    Dim CurrentChar As String
    Dim Formula As String
    
    CorrectFNC = 0
    PrintableString = ""
    DataToFormat = DataToEncode
    DataToEncode = ""
    'Here we select character set A, B or C for the START character
    StringLength = Len(DataToFormat)
    CurrentCharNum = Asc(Mid(DataToFormat, 1, 1))
    If CurrentCharNum < 32 Then C128Start = Chr(203)
    If CurrentCharNum > 31 And CurrentCharNum < 127 Then C128Start = Chr(204)
    If ((StringLength > 4) And IsNumeric(Mid(DataToFormat, 1, 4))) Then C128Start = Chr(205)
    '202 & 212-215 is for the FNC1, with this Start C is mandatory
    If CurrentCharNum = 202 Then C128Start = Chr(205)
    If CurrentCharNum = 212 Then C128Start = Chr(205)
    If CurrentCharNum = 213 Then C128Start = Chr(205)
    If CurrentCharNum = 214 Then C128Start = Chr(205)
    If CurrentCharNum = 215 Then C128Start = Chr(205)
    If C128Start = Chr(203) Then CurrentEncoding = "A"
    If C128Start = Chr(204) Then CurrentEncoding = "B"
    If C128Start = Chr(205) Then CurrentEncoding = "C"
    For I = 1 To StringLength
        'check for FNC1 in any set which is ASCII 202 and ASCII 212-215
        CurrentCharNum = Asc(Mid(DataToFormat, I, 1))
        If ((CurrentCharNum = 202) Or (CurrentCharNum = 212) Or (CurrentCharNum = 213) Or (CurrentCharNum = 214) Or (CurrentCharNum = 215)) Then
            DataToEncode = DataToEncode & Chr(202)
        'check for switching to character set C
        ElseIf ((I < StringLength - 2) And (IsNumeric(Mid(DataToFormat, I, 1))) And (IsNumeric(Mid(DataToFormat, I + 1, 1))) And (IsNumeric(Mid(DataToFormat, I, 4)))) Or ((I < StringLength) And (IsNumeric(Mid(DataToFormat, I, 1))) And (IsNumeric(Mid(DataToFormat, I + 1, 1))) And (CurrentEncoding = "C")) Then
        'switch to set C if not already in it
            If CurrentEncoding <> "C" Then DataToEncode = DataToEncode & Chr(199)
            CurrentEncoding = "C"
            CurrentChar = Mid(DataToFormat, I, 2)
            CurrentValue = Val(CurrentChar)
        'set the CurrentValue to the number of String CurrentChar
            If (CurrentValue < 95 And CurrentValue > 0) Then DataToEncode = DataToEncode & Chr(CurrentValue + 32)
            If CurrentValue > 94 Then DataToEncode = DataToEncode & Chr(CurrentValue + 100)
            If CurrentValue = 0 Then DataToEncode = DataToEncode & Chr(194)
            I = I + 1
        'check for switching to character set A
        ElseIf (I <= StringLength) And ((Asc(Mid(DataToFormat, I, 1)) < 31) Or ((CurrentEncoding = "A") And (Asc(Mid(DataToFormat, I, 1)) > 32 And (Asc(Mid(DataToFormat, I, 1))) < 96))) Then
        'switch to set A if not already in it
            If CurrentEncoding <> "A" Then DataToEncode = DataToEncode & Chr(201)
            CurrentEncoding = "A"
        'Get the ASCII value of the next character
            CurrentCharNum = Asc(Mid(DataToFormat, I, 1))
            If CurrentCharNum = 32 Then
                DataToEncode = DataToEncode & Chr(194)
            ElseIf CurrentCharNum < 32 Then
                DataToEncode = DataToEncode & Chr(CurrentCharNum + 96)
            ElseIf CurrentCharNum > 32 Then
                DataToEncode = DataToEncode & Chr(CurrentCharNum)
            End If
        'check for switching to character set B
        ElseIf (I <= StringLength) And (((Asc(Mid(DataToFormat, I, 1))) > 31) And ((Asc(Mid(DataToFormat, I, 1)))) < 127) Then
        'switch to set B if not already in it
            If CurrentEncoding <> "B" Then DataToEncode = DataToEncode & Chr(200)
            CurrentEncoding = "B"
        'Get the ASCII value of the next character
            CurrentCharNum = Asc(Mid(DataToFormat, I, 1))
            If CurrentCharNum = 32 Then
                DataToEncode = DataToEncode & Chr(194)
            Else
                DataToEncode = DataToEncode & Chr(CurrentCharNum)
            End If
        End If
    Next I
    
 
    '<<<< Calculate Modulo 103 Check Digit >>>>
    WeightedTotal = Asc(C128Start) - 100
    StringLength = Len(DataToEncode)
    For I = 1 To StringLength
        CurrentCharNum = Asc(Mid(DataToEncode, I, 1))
        If CurrentCharNum < 135 Then CurrentValue = CurrentCharNum - 32
        If CurrentCharNum > 134 Then CurrentValue = CurrentCharNum - 100
        If CurrentCharNum = 194 Then CurrentValue = 0
        CurrentValue = CurrentValue * I
        WeightedTotal = WeightedTotal + CurrentValue
        If CurrentCharNum = 32 Then CurrentCharNum = 194
        PrintableString = PrintableString & Chr(CurrentCharNum)
    Next I
    CheckDigitValue = (WeightedTotal Mod 103)
    If CheckDigitValue < 95 And CheckDigitValue > 0 Then C128CheckDigit = Chr(CheckDigitValue + 32)
    If CheckDigitValue > 94 Then C128CheckDigit = Chr(CheckDigitValue + 100)
    If CheckDigitValue = 0 Then C128CheckDigit = Chr(194)
    DataToEncode = ""
    Formula = C128Start & PrintableString & C128CheckDigit & Chr(206) & " "

    Me.Label5.Caption = Formula

End Function

Public Function Code128(DataToFormat As String, Optional ReturnType As Integer = 0, Optional ApplyTilde As Boolean = False) As String
    'This method has been updated to support ReturnTypes 6 to 9,
    'which increases its complexity. IDAutomation suggests using
    'the prior version (http://www.idautomation.com/fonts/tools/barcodeapp/module1.txt)
    'of this code when performing conversions or modifications.
    'ReturnTypes are explained at http://www.idautomation.com/barcode/return-type.html
    '
    'The next 12 lines were added to support ReturnTypes 6-9
    CID = 0 'Character ID
    NCID = 0 'Numbers Character ID (for set C)
    FID = 0 'Function ID used for start, stop and check characters
    If ReturnType = 6 Or ReturnType = 7 Then CID = 11000
    If ReturnType = 8 Then CID = 11300
    If ReturnType = 9 Then CID = 11500
    If ReturnType = 6 Or ReturnType = 9 Then FID = 11500
    If ReturnType = 7 Or ReturnType = 8 Then FID = 11300
    If ReturnType = 6 Or ReturnType = 7 Then NCID = 12000
    If ReturnType = 8 Then NCID = 11300
    If ReturnType = 9 Then NCID = 11500
    Dim SetString As String
     
    Dim DataToEncode As String
    Dim SetAry As String
    CorrectFNC = 0
    PrintableString = ""
    DataToEncode = ""
    SetString = ""
    'ProcessTilde was modified to support ReturnTypes 6-9 and
    'support using ( and ) to define AIs for GS1-128
    If ApplyTilde Then DataToFormat = ProcessTilde(DataToFormat)
    
If ReturnType = 0 Or ReturnType = 2 Or ReturnType > 5 Then
    'Select the character set A, B or C for the START character
    'The next 15 lines were modified to support ReturnTypes 6-9
    'by the addition of the FID. The SetAry records the character set
    'so that the correct character can be displayed for HR text.
    SetAry = Array("0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0")
    CurrentChar = Left(DataToFormat, 1)
    CurrentCharNum = AscW(CurrentChar)
    StringLength = Len(DataToFormat)
    C128Start = ChrW(204 + FID)
    If CurrentCharNum < 32 Then C128Start = ChrW(203 + FID)
    If CurrentCharNum > 31 And CurrentCharNum < 127 Then C128Start = ChrW(204 + FID)
    If ((StringLength > 3) And IsNumeric(Mid(DataToFormat, 1, 4))) Then C128Start = ChrW(205 + FID)
    '202 & 212-215 is for the FNC1, with this Start C is mandatory
    If CurrentCharNum = 197 Then C128Start = ChrW(204 + FID)
    If CurrentCharNum > 201 Then C128Start = ChrW(205 + FID)
    If C128Start = ChrW(203 + FID) Then CurrentEncoding = "A"
    If C128Start = ChrW(204 + FID) Then CurrentEncoding = "B"
    If C128Start = ChrW(205 + FID) Then CurrentEncoding = "C"
    'The next line was added to support ReturnTypes 6-9
    J = 0
    For I = 1 To StringLength
        'The next line was added to support ReturnTypes 6-9
        J = J + 1
        'Check for FNC1 in any set which is ASCII 202 and ASCII 212-217
        CurrentCharNum = AscW(Mid(DataToFormat, I, 1))
        
        If CurrentCharNum > 201 Then
            'The next 5 lines were added to support ReturnTypes 6-9
            'Switch to set C if not already in it BEFORE the AI
            'Change SetAry(J) to B so the number 99 does not show up for the set switch
            If CurrentEncoding <> "C" Then
                SetAry(J) = "B"
                DataToEncode = DataToEncode & ChrW(199)
                J = J + 1
                CurrentEncoding = "C"
            End If
            DataToEncode = DataToEncode & ChrW(202)
            'The next 14 lines were added to support ReturnTypes 6-9
            'Change SetAry(J) to B so the number 99 does not show up for the AI
            SetAry(J) = "B"
            If CurrentCharNum > 211 Then
                'Change SetAry to reflect the location of proper HR text
                'E indicates the ) is at the end of a number pair instead of the Middle
                If CurrentCharNum = 212 Then SetAry(J + 1) = "E"
                If CurrentCharNum = 213 Then SetAry(J + 2) = "M"
                If CurrentCharNum = 214 Then SetAry(J + 2) = "E"
                If CurrentCharNum = 215 Then SetAry(J + 3) = "M"
                If CurrentCharNum = 216 Then SetAry(J + 3) = "E"
                If CurrentCharNum = 217 Then SetAry(J + 4) = "M"
            End If
        'Check for switching
        ElseIf CurrentCharNum = 195 Then
            If CurrentEncoding = "C" Then
                DataToEncode = DataToEncode & ChrW(200)
                'The next line was added to support ReturnTypes 6-9
                If SetAry(J) = "0" Then SetAry(J) = "B"
                CurrentEncoding = "B"
            End If
            DataToEncode = DataToEncode & ChrW(195)
            'The next line was added to support ReturnTypes 6-9
            If SetAry(J) = "0" Then SetAry(J) = "B"
        ElseIf CurrentCharNum = 196 Then
            If CurrentEncoding = "C" Then
                DataToEncode = DataToEncode & ChrW(200)
                'The next line was added to support ReturnTypes 6-9
                If SetAry(J) = "0" Then SetAry(J) = "B"
                CurrentEncoding = "B"
            End If
            DataToEncode = DataToEncode & ChrW(196)
            'The next line was added to support ReturnTypes 6-9
            If SetAry(J) = "0" Then SetAry(J) = "B"
        ElseIf CurrentCharNum = 197 Then
            If CurrentEncoding = "C" Then
                DataToEncode = DataToEncode & ChrW(200)
                'The next line was added to support ReturnTypes 6-9
                If SetAry(J) = "0" Then SetAry(J) = "B"
                CurrentEncoding = "B"
            End If
            DataToEncode = DataToEncode & ChrW(197)
            'The next line was added to support ReturnTypes 6-9
            If SetAry(J) = "0" Then SetAry(J) = "B"
        ElseIf CurrentCharNum = 198 Then
            If CurrentEncoding = "C" Then
                DataToEncode = DataToEncode & ChrW(200)
                'The next line was added to support ReturnTypes 6-9
                If SetAry(J) = "0" Then SetAry(J) = "B"
                CurrentEncoding = "B"
            End If
            DataToEncode = DataToEncode & ChrW(198)
            'The next line was added to support ReturnTypes 6-9
            If SetAry(J) = "0" Then SetAry(J) = "B"
        ElseIf CurrentCharNum = 200 Then
            If CurrentEncoding = "C" Then
                DataToEncode = DataToEncode & ChrW(200)
                'The next line was added to support ReturnTypes 6-9
                If SetAry(J) = "0" Then SetAry(J) = "B"
                CurrentEncoding = "B"
            End If
            DataToEncode = DataToEncode & ChrW(200)
            'The next line was added to support ReturnTypes 6-9
            If SetAry(J) = "0" Then SetAry(J) = "B"
        ElseIf ((I < StringLength - 2) And (IsNumeric(Mid(DataToFormat, I, 1))) And (IsNumeric(Mid(DataToFormat, I + 1, 1))) And (IsNumeric(Mid(DataToFormat, I, 4)))) Or ((I < StringLength) And (IsNumeric(Mid(DataToFormat, I, 1))) And (IsNumeric(Mid(DataToFormat, I + 1, 1))) And (CurrentEncoding = "C")) Then
            'check to see if there is an odd number of digits to encode,
            'if so stay in current set for 1 number and then switch to save space
            'This IF statement was modified to support ReturnTypes 6-9; changed counter variable name from J to F
            If CurrentEncoding <> "C" Then
                F = I
                Factor = 3
                Do While F <= StringLength And IsNumeric(Mid(DataToFormat, F, 1))
                    Factor = 4 - Factor
                    F = F + 1
                Loop
                If Factor = 1 Then
                    'if so stay in current set for 1 character to save space
                    DataToEncode = DataToEncode & ChrW(CurrentCharNum)
                    'The next line was added to support ReturnTypes 6-9
                    If SetAry(J) = "0" Then SetAry(J) = CurrentEncoding
                    I = I + 1
                    J = J + 1
                End If
            End If
            'Switch to set C if not already in it
            If CurrentEncoding <> "C" Then DataToEncode = DataToEncode & ChrW(199)
            'The next 2 lines of code were added to support ReturnTypes 6-9
            'Sets the encoding in SetAry to the previous mode to keep switch characters from showing up in HR text.
            If CurrentEncoding <> "C" Then SetAry(J) = CurrentEncoding
            If CurrentEncoding <> "C" Then J = J + 1
            CurrentEncoding = "C"
            CurrentChar = (Mid(DataToFormat, I, 2))
            CurrentValue = Val(CurrentChar)
            'Set the CurrentValue to the number of String CurrentChar
            If (CurrentValue < 95 And CurrentValue > 0) Then DataToEncode = DataToEncode & ChrW(CurrentValue + 32)
            If CurrentValue > 94 Then DataToEncode = DataToEncode & ChrW(CurrentValue + 100)
            If CurrentValue = 0 Then DataToEncode = DataToEncode & ChrW(194)
            'The next line was added to support ReturnTypes 6-9
            If SetAry(J) = "0" Then SetAry(J) = CurrentEncoding
            I = I + 1
        'Check for switching to character set A
        ElseIf (I <= StringLength) And ((AscW(Mid(DataToFormat, I, 1)) < 31) Or ((CurrentEncoding = "A") And (AscW(Mid(DataToFormat, I, 1)) > 32 And (AscW(Mid(DataToFormat, I, 1))) < 96))) Then
        'Switch to set A if not already in it
            If CurrentEncoding <> "A" Then DataToEncode = DataToEncode & ChrW(201)
            'The next 2 lines were added to support ReturnTypes 6-9
            If CurrentEncoding <> "A" Then SetAry(J) = "A"
            If CurrentEncoding <> "A" Then J = J + 1
            CurrentEncoding = "A"
            'Get the ASCII value of the next character
            CurrentCharNum = AscW(Mid(DataToFormat, I, 1))
            If CurrentCharNum = 32 Then
                DataToEncode = DataToEncode & ChrW(194)
            ElseIf CurrentCharNum < 32 Then
                DataToEncode = DataToEncode & ChrW(CurrentCharNum + 96)
            ElseIf CurrentCharNum > 32 Then
                DataToEncode = DataToEncode & ChrW(CurrentCharNum)
            End If
            'The next line was added to support ReturnTypes 6-9
            If SetAry(J) = "0" Then SetAry(J) = CurrentEncoding
        'Check for switching to character set B
        ElseIf (I <= StringLength) And ((AscW(Mid(DataToFormat, I, 1))) > 31 And (AscW(Mid(DataToFormat, I, 1)))) < 127 Then
        'Switch to set B if not already in it
            If CurrentEncoding <> "B" Then DataToEncode = DataToEncode & ChrW(200)
           'The next 2 lines were added to support ReturnTypes 6-9
            If CurrentEncoding <> "B" Then SetAry(J) = "B"
            If CurrentEncoding <> "B" Then J = J + 1
            'J = J + 1
            CurrentEncoding = "B"
        'Get the ASCII value of the next character
            CurrentCharNum = AscW(Mid(DataToFormat, I, 1))
            If CurrentCharNum = 32 Then
                DataToEncode = DataToEncode & ChrW(194)
            Else
                DataToEncode = DataToEncode & ChrW(CurrentCharNum)
            End If
            'The next line was added to support ReturnTypes 6-9
            If SetAry(J) = "0" Then SetAry(J) = CurrentEncoding
        End If
    Next I
End If

'FORMAT TEXT FOR AIs
If ReturnType = 1 Then
    'ReturnType 1 = format the data for human readable text only
    HumanReadableText = ""
    StringLength = Len(DataToFormat)
    For I = 1 To StringLength
        CorrectFNC = 0
        'Get ASCII value of each character
        CurrentCharNum = AscW(Mid(DataToFormat, I, 1))
        'Check for FNC1
        If ((I < StringLength - 2) And ((CurrentCharNum = 202) Or ((CurrentCharNum > 211) And (CurrentCharNum < 219)))) Then
        '2005.12 BDA updated the next if/else to eliminate errors from text after the AI
        'It appears that there is an AI
        'Get the value of the next 2 digits to try to determine the length of the AI, if text is used here
        'Set this value to 81 for a 4 digit AI
        If IsNumeric(Mid(DataToFormat, I + 1, 1)) And IsNumeric(Mid(DataToFormat, I + 2, 1)) Then
            CurrentChar = Mid(DataToFormat, I + 1, 2)
            CurrentCharNum = Val(CurrentChar)
        Else
            CurrentCharNum = 81
        End If
        'Is 2 digit AI by entering ASCII 212?
            If ((CorrectFNC = 0) And (AscW(Mid(DataToFormat, I, 1)) = 212)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 2)) & ") "
                I = I + 2
                CorrectFNC = 1
        'Is 3 digit AI by entering ASCII 213?
            ElseIf ((I < StringLength - 3) And (CorrectFNC = 0) And (AscW(Mid(DataToFormat, I, 1)) = 213)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 3)) & ") "
                I = I + 3
                CorrectFNC = 1
        'Is 4 digit AI by entering ASCII 214?
            ElseIf ((I < StringLength - 4) And (CorrectFNC = 0) And (AscW(Mid(DataToFormat, I, 1)) = 214)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 4)) & ") "
                I = I + 4
                CorrectFNC = 1
        'Is 5 digit AI by entering ASCII 215?
            ElseIf ((I < StringLength - 5) And (CorrectFNC = 0) And (AscW(Mid(DataToFormat, I, 1)) = 215)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 5)) & ") "
                I = I + 5
                CorrectFNC = 1
        'Is 6 digit AI by entering ASCII 216?
            ElseIf ((I < StringLength - 6) And (CorrectFNC = 0) And (AscW(Mid(DataToFormat, I, 1)) = 216)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 6)) & ") "
                I = I + 6
                CorrectFNC = 1
        'Is 7 digit AI by entering ASCII 217?
            ElseIf ((I < StringLength - 7) And (CorrectFNC = 0) And (AscW(Mid(DataToFormat, I, 1)) = 217)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 7)) & ") "
                I = I + 7
                CorrectFNC = 1
        'Is 8 digit AI by entering ASCII 218?
            ElseIf ((I < StringLength - 8) And (CorrectFNC = 0) And (AscW(Mid(DataToFormat, I, 1)) = 218)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 8)) & ") "
                I = I + 8
                CorrectFNC = 1
        'Is 4 digit AI by detection?
            ElseIf ((I < StringLength - 4) And (CorrectFNC = 0) And ((CurrentCharNum <= 81 And CurrentCharNum >= 80) Or (CurrentCharNum <= 34 And CurrentCharNum >= 31))) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 4)) & ") "
                I = I + 4
                CorrectFNC = 1
        'Is 3 digit AI by detection?
            ElseIf ((I < StringLength - 3) And (CorrectFNC = 0) And ((CurrentCharNum <= 49 And CurrentCharNum >= 40) Or (CurrentCharNum <= 25 And CurrentCharNum >= 23))) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 3)) & ") "
                I = I + 3
                CorrectFNC = 1
        'Is 2 digit AI by detection?
            ElseIf ((CurrentCharNum <= 30 And (CorrectFNC = 0) And CurrentCharNum >= 0) Or (CurrentCharNum <= 99 And CurrentCharNum >= 90)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 2)) & ") "
                I = I + 2
                CorrectFNC = 1
        'If no AI was detected, set default to 4 digit AI:
            ElseIf ((I < StringLength - 4) And (CorrectFNC = 0)) Then
                HumanReadableText = HumanReadableText & " (" & (Mid(DataToFormat, I + 1, 4)) & ") "
                I = I + 4
                CorrectFNC = 1
            End If
        ElseIf (AscW(Mid(DataToFormat, I, 1)) < 32) Then
            HumanReadableText = HumanReadableText & " "
        ElseIf ((AscW(Mid(DataToFormat, I, 1)) > 31) And (AscW(Mid(DataToFormat, I, 1)) < 128)) Then
            HumanReadableText = HumanReadableText & Mid(DataToFormat, I, 1)
        End If
    Next I
End If

'The next line was modified to support ReturnTypes 3-5
If ReturnType > 2 And ReturnType < 6 Then
    'ReturnType 3, 4 or 5 = format the data for human readable text only
    'inserting a space for every 3, 4 or 5 characters
    HumanReadableText = ""
    StringLength = Len(DataToFormat)
    J = 0
    For I = 1 To StringLength
        CurrentCharNum = AscW(Mid(DataToFormat, I, 1))
        If CurrentCharNum > 31 And CurrentCharNum < 128 Then
            HumanReadableText = HumanReadableText & Mid(DataToFormat, I, 1)
            J = J + 1
        End If
        If (J Mod ReturnType) = 0 Then HumanReadableText = HumanReadableText & " "
    Next I
End If

If ReturnType = 0 Or ReturnType = 2 Or ReturnType > 5 Then
    DataToFormat = ""
    'The next line was modified to support ReturnTypes 6-9
    WeightedTotal = AscW(C128Start) - (FID + 100)
    StringLength = Len(DataToEncode)
    For I = 1 To StringLength
        CurrentCharNum = AscW(Mid(DataToEncode, I, 1))
        If CurrentCharNum < 135 Then CurrentValue = CurrentCharNum - 32
        If CurrentCharNum > 134 Then CurrentValue = CurrentCharNum - 100
        If CurrentCharNum = 194 Then CurrentValue = 0
        CurrentValue = CurrentValue * I
        WeightedTotal = WeightedTotal + CurrentValue
        If CurrentCharNum = 32 Then CurrentCharNum = 194
        'The next 10 lines were modified/added to support ReturnTypes 6-9
        'so that set C characters show the correct HR number pairs
        If ReturnType > 5 And SetAry(I) = "C" Then
            PrintableString = PrintableString & ChrW(CurrentCharNum + NCID)
        ElseIf (ReturnType = 6 Or ReturnType = 7) And SetAry(I) = "E" Then
            PrintableString = PrintableString & ChrW(CurrentCharNum + 10500)
        ElseIf (ReturnType = 6 Or ReturnType = 7) And SetAry(I) = "M" Then
            PrintableString = PrintableString & ChrW(CurrentCharNum + 10700)
        Else
            PrintableString = PrintableString & ChrW(CurrentCharNum + CID)
        End If
    Next I
    CheckDigitValue = (WeightedTotal Mod 103)
    If CheckDigitValue < 95 And CheckDigitValue > 0 Then C128CheckDigit = ChrW(CheckDigitValue + 32 + FID)
    If CheckDigitValue > 94 Then C128CheckDigit = ChrW(CheckDigitValue + 100 + FID)
    If CheckDigitValue = 0 Then C128CheckDigit = ChrW(194 + FID)
End If
    
    DataToEncode = ""
    'ReturnType 0 returns data formatted to the barcode font
    If ReturnType = 0 Or ReturnType > 2 Then Code128 = C128Start & PrintableString & C128CheckDigit & ChrW(206 + FID)
    'ReturnType 1 returns data formatted for human readable text
    If ReturnType = 1 Then Code128 = HumanReadableText
    'ReturnType 2 returns the check digit for the data supplied
    If ReturnType = 2 Then Code128 = C128CheckDigit
End Function

Public Function ProcessTilde(StringToProcess As String) As String
        ProcessTilde = ""
        Dim OutString As String
        Dim CharsAdded As Integer
        StringLength = Len(StringToProcess)
        For I = 1 To StringLength
            If (I < StringLength - 2) And Mid(StringToProcess, I, 2) = "~m" And IsNumeric(Mid(StringToProcess, I + 2, 2)) Then
                Dim StringToCheck As String
                WeightValue = Val(Mid(StringToProcess, I + 2, 2))
                'Dim CharsAdded As Integer
                For J = I To 1 Step -1
                    If IsNumeric(Mid(OutString, J, 1)) Then
                        StringToCheck = StringToCheck & Mid(OutString, J, 1)
                        CharsAdded = CharsAdded + 1
                    End If
                    'when the number of digits added to StringToCheck equals the weight value exit the for loop
                    If CharsAdded = WeightValue Then
                        Exit For
                    End If
                Next J
                CheckDigitValue = MOD10(StrReverse(StringToCheck))
                OutString = OutString & ChrW(CheckDigitValue + 48)
                I = I + 3
            ElseIf (I < StringLength - 2) And Mid(StringToProcess, I, 1) = "~" And IsNumeric(Mid(StringToProcess, I + 1, 3)) Then
                CurrentCharNum = Val(Mid(StringToProcess, I + 1, 3))
                OutString = OutString & ChrW(CurrentCharNum)
                I = I + 3
            'This ElseIf was modified to support using () to add in AIs
            ElseIf (I < StringLength - 4) And Mid(StringToProcess, I, 1) = "(" And (Mid(StringToProcess, I + 2, 1) = ")" Or Mid(StringToProcess, I + 3, 1) = ")" Or Mid(StringToProcess, I + 4, 1) = ")" Or (Mid(StringToProcess, I + 5, 1) = ")" And (I < StringLength - 5)) Or (Mid(StringToProcess, I + 6, 1) = ")" And (I < StringLength - 6)) Or (Mid(StringToProcess, I + 7, 1) = ")" And (I < StringLength - 4)) Or (Mid(StringToProcess, I + 8, 1) = ")" And (I < StringLength - 4))) Then
                'Assign ASCII 212-217 depending on how many digits between ()
                If Mid(StringToProcess, I + 3, 1) = ")" Then OutString = OutString & ChrW(212)
                If Mid(StringToProcess, I + 4, 1) = ")" Then OutString = OutString & ChrW(213)
                If Mid(StringToProcess, I + 5, 1) = ")" Then OutString = OutString & ChrW(214)
                If Mid(StringToProcess, I + 6, 1) = ")" Then OutString = OutString & ChrW(215)
                If Mid(StringToProcess, I + 7, 1) = ")" Then OutString = OutString & ChrW(216)
                If Mid(StringToProcess, I + 8, 1) = ")" Then OutString = OutString & ChrW(217)
            'This ElseIf was modified to exclude ")" from being encoded
            'ElseIf (I < StringLength - 2) And Mid(StringToProcess, I, 1) = ")" Then
            '6/2/2010, TB, Fix bug #392, removed - 2 from the if statement
            ElseIf (I < StringLength) And Mid(StringToProcess, I, 1) = ")" Then
            'Skip this character by breaking out of the else if
            Else
               OutString = OutString & Mid(StringToProcess, I, 1)
            End If
        Next I
        ProcessTilde = OutString
        StringToProcess = ""
End Function


 

