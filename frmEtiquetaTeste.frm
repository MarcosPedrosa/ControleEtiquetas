VERSION 5.00
Begin VB.Form frmEtiquetaTeste 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prévia de Impressão - Etiqueta Teste"
   ClientHeight    =   3795
   ClientLeft      =   8490
   ClientTop       =   1590
   ClientWidth     =   6060
   Icon            =   "frmEtiquetaTeste.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   6060
   Begin VB.Label lblFuncionario 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "IGOR DE ALMEIDA RODRIGUES"
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   1755
      Width           =   2445
   End
   Begin VB.Label lblCodigoBarra 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "2468"
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
      Left            =   2520
      TabIndex        =   6
      Top             =   3285
      Width           =   1020
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   4380
      X2              =   4395
      Y1              =   2205
      Y2              =   2670
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   1455
      X2              =   4395
      Y1              =   2655
      Y2              =   2670
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   1493
      X2              =   1508
      Y1              =   2670
      Y2              =   2025
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   503
      X2              =   1508
      Y1              =   2655
      Y2              =   2670
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   510
      X2              =   4365
      Y1              =   2025
      Y2              =   2025
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   3825
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   510
      X2              =   4365
      Y1              =   1395
      Y2              =   1395
   End
   Begin VB.Image Image1 
      Height          =   345
      Left            =   1088
      Picture         =   "frmEtiquetaTeste.frx":0442
      Stretch         =   -1  'True
      Top             =   810
      Width           =   2625
   End
   Begin VB.Label lblNomeSetor 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Departamento de Informática"
      Height          =   195
      Left            =   1590
      TabIndex        =   5
      Top             =   2385
      Width           =   2055
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Setor"
      Height          =   195
      Left            =   1583
      TabIndex        =   4
      Top             =   2115
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Foto"
      Height          =   195
      Left            =   4598
      TabIndex        =   3
      Top             =   2295
      Width           =   645
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "2468"
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   2385
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Código"
      Height          =   195
      Left            =   593
      TabIndex        =   1
      Top             =   2115
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Funcionário"
      Height          =   195
      Left            =   593
      TabIndex        =   0
      Top             =   1485
      Width           =   825
   End
   Begin VB.Shape shpFoto 
      BackColor       =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   1455
      Left            =   4380
      Top             =   765
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   2445
      Left            =   503
      Top             =   675
      Width           =   5100
   End
End
Attribute VB_Name = "frmEtiquetaTeste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'Limpar Campos
'EtiquetaTeste
frmEtiquetaTeste.lblFuncionario = ""
frmEtiquetaTeste.lblCodigo = ""
frmEtiquetaTeste.lblNomeSetor = ""
frmEtiquetaTeste.lblCodigoBarra = ""

Me.Top = 0
Me.Width = 6180
Me.Height = 4200

Dim v As Integer
For v = 0 To (Forms.Count - 1)
If Forms(v).Name = "frmEtiquetaTesteOpcoes" Then
    Me.Left = frmEtiquetaTesteOpcoes.Width
    Exit Sub
End If
Me.Left = frmOpcoes.Width
Next

End Sub
