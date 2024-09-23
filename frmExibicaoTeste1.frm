VERSION 5.00
Begin VB.Form frmExibicaoTeste1 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Teste1"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   8190
   Begin VB.CommandButton CMD_IMPRIME 
      Caption         =   "IMPRIME"
      Height          =   375
      Left            =   6930
      TabIndex        =   19
      Top             =   5550
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Label2"
      Height          =   3435
      Left            =   240
      TabIndex        =   18
      Top             =   750
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      Caption         =   "KMS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   17
      Top             =   420
      Width           =   675
   End
   Begin VB.Image LogoMsb 
      Height          =   735
      Left            =   1800
      Picture         =   "frmExibicaoTeste1.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   4275
   End
   Begin VB.Label lblQtd1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "QTD.:"
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
      Left            =   1050
      TabIndex        =   16
      Top             =   3300
      Width           =   1155
   End
   Begin VB.Label lblPeso1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PESO:"
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
      Left            =   4530
      TabIndex        =   15
      Top             =   3300
      Width           =   1185
   End
   Begin VB.Label lblLote1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "LOTE:"
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
      Left            =   1050
      TabIndex        =   14
      Top             =   3780
      Width           =   1170
   End
   Begin VB.Label lblCliente 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "MUSASHI DA AMAZONIA LTDA "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   810
      TabIndex        =   13
      Top             =   960
      Width           =   6480
   End
   Begin VB.Label lblPeca 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "2H26241"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2670
      TabIndex        =   12
      Top             =   1860
      Width           =   2595
   End
   Begin VB.Label lblDescricao 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ENGRENAGEM SECUNDARIA C-4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1380
      TabIndex        =   11
      Top             =   2460
      Width           =   5400
   End
   Begin VB.Label lblCodCliente2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "23481-KRM-8410"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2370
      TabIndex        =   10
      Top             =   2910
      Width           =   3825
   End
   Begin VB.Label lblQtd2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2370
      TabIndex        =   9
      Top             =   3270
      Width           =   300
   End
   Begin VB.Label lblPeso2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0,18"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   5730
      TabIndex        =   8
      Top             =   3270
      Width           =   1050
   End
   Begin VB.Label lblLote2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "11322472"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   2400
      TabIndex        =   7
      Top             =   3750
      Width           =   2400
   End
   Begin VB.Label lblCodCliente1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COD..:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1050
      TabIndex        =   6
      Top             =   2910
      Width           =   1185
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   5235
      Left            =   60
      Top             =   120
      Width           =   7935
   End
   Begin VB.Label lblEmbalagem2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0525"
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
      TabIndex        =   5
      Top             =   4980
      Width           =   360
   End
   Begin VB.Label lblCodigoBarrasB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0003079152"
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
      Left            =   1710
      TabIndex        =   4
      Top             =   4890
      Width           =   5175
   End
   Begin VB.Label lblCodigoBarras 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0003079152"
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
      Left            =   1710
      TabIndex        =   3
      Top             =   4410
      Width           =   5175
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
      Left            =   210
      TabIndex        =   2
      Top             =   5040
      Width           =   975
   End
   Begin VB.Label lblCodigoBarrasA 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0003079152"
      BeginProperty Font 
         Name            =   "3 of 9 Barcode"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1710
      TabIndex        =   1
      Top             =   4740
      Width           =   5175
   End
   Begin VB.Label lbl_data 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "11/08/2011"
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
      Left            =   240
      TabIndex        =   0
      Top             =   4710
      Width           =   840
   End
End
Attribute VB_Name = "frmExibicaoTeste1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMD_IMPRIME_Click()
Dim X As Printer
Dim nx As Integer

nx = 0
For Each X In Printers
   If InStr(1, UCase(X.DeviceName), "ETIQUETA FABRICA") > 0 Then
      Set Printer = X
      nx = 1
      Exit For
   End If
Next

If nx = 0 Then
   MsgBox "impressora da Etiqueta não encontrada, chame o responsável. o nome da impressora será - 'ETIQUETA FABRICA'"
   Exit Sub
End If
nx = 0
CMD_IMPRIME.Visible = False
Printer.Orientation = 2
Me.PrintForm
Printer.EndDoc
CMD_IMPRIME.Visible = True
End Sub
