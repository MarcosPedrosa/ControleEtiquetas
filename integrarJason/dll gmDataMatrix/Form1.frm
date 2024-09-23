VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   LinkTopic       =   "0"
   ScaleHeight     =   6840
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2295
      Left            =   2070
      Negotiate       =   -1  'True
      ScaleHeight     =   2235
      ScaleWidth      =   2805
      TabIndex        =   1
      Top             =   3150
      Width           =   2865
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   585
      Left            =   390
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1755
      Left            =   3810
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   810
      Width           =   2925
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Simage As String

Simage = "C:\Documents and Settings\IEUser\My Documents\Sistemas\mussashi\Desenvolvimento\Fontes\MD_Etiquetas\output.jpg"

Me.Image1 = LoadPicture(Simage)

'Shell App.Path & "DataMatrix.exe"

End Sub

Private Sub Form_Load()

End Sub
w
