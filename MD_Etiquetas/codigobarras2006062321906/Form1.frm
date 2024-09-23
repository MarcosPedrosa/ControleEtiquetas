VERSION 5.00
Object = "{BFAB5AD1-BEA8-4DA8-8177-AA5ED87B533A}#1.0#0"; "CTKBCDC.OCX"
Begin VB.Form Form1 
   Caption         =   "Cria Codigo Barras"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   840
      MaxLength       =   10
      TabIndex        =   6
      Top             =   2760
      Width           =   1875
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   840
      MaxLength       =   8
      TabIndex        =   4
      Top             =   1830
      Width           =   1875
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   1140
      Width           =   1875
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   840
      MaxLength       =   13
      TabIndex        =   0
      Top             =   180
      Width           =   1875
   End
   Begin ctk_BarCode.EAN8 EAN81 
      Height          =   795
      Left            =   2910
      Top             =   1620
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1402
   End
   Begin ctk_BarCode.EAN13 EAN131 
      Height          =   915
      Left            =   2910
      Top             =   30
      Width           =   3090
      _ExtentX        =   5450
      _ExtentY        =   1614
   End
   Begin ctk_BarCode.Code39 Code391 
      Height          =   405
      Left            =   2910
      Top             =   1050
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   714
   End
   Begin ctk_BarCode.I2x5 I2x51 
      Height          =   585
      Left            =   2910
      Top             =   2550
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   1032
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "I2x51"
      Height          =   195
      Left            =   150
      TabIndex        =   7
      Top             =   2850
      Width           =   390
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "EAN81"
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   1920
      Width           =   510
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Code391"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   1230
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "EAN131"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Top             =   270
      Width           =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()
   EAN131.Code = Text1
End Sub

Private Sub Text2_Change()
   Code391.Code = Text2
End Sub

Private Sub Text3_Change()
   EAN81.Code = Text3
End Sub

Private Sub Text4_Change()
   I2x51.Code = Text4
End Sub
