VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   10665
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Gerar"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   180
      MaxLength       =   12
      TabIndex        =   1
      Top             =   660
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   315
      Left            =   1140
      TabIndex        =   5
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "String"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   2700
      Width           =   795
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Cod impresso"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   1020
      TabIndex        =   0
      Top             =   1080
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1.Caption = EAN13(Text1.Text)
Label4.Caption = EAN13(Text1.Text)
End Sub

