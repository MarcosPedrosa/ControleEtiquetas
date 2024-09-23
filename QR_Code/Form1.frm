VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QR Code"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdCodificTxt 
      Caption         =   "Codificar Texto"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton CmdDecodificar 
      Caption         =   "Decodificar Imagen"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   3990
      Width           =   1815
   End
   Begin VB.CommandButton CmdWebCam 
      Caption         =   "WebCam"
      Height          =   375
      Left            =   8040
      TabIndex        =   2
      Top             =   3960
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   3060
      Left            =   6000
      ScaleHeight     =   200
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   1
      Top             =   120
      Width           =   3060
   End
   Begin VB.TextBox Text1 
      Height          =   2655
      Left            =   120
      MaxLength       =   700
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'----------------------------------
'Autor:  Leandro Ascierto
'Web:    leandroascierto.com
'Date:   09/09/2011
'----------------------------------
Dim cQrCode As ClsQrCode

Private Sub CmdCodificTxt_Click()
    Picture1.Picture = cQrCode.GetPictureQrCode(Text1.Text, Picture1.ScaleWidth, Picture1.ScaleHeight)
    If Picture1.Picture Is Nothing Then MsgBox "Error!"
    Picture1.Picture = cQrCode.GetPictureQrCode(Text1.Text, 200, 200, "UTF-8", "L", vbBlack, vbWhite, 3)
End Sub

Private Sub CmdDecodificar_Click()
    Dim strDecode As String
    If cQrCode.DecodeFromPicture(Picture1.Picture, strDecode) Then
        MsgBox strDecode
    Else
        MsgBox "Error!"
    End If
End Sub



Private Sub CmdWebCam_Click()
    FrmWebCam.Show , Me
End Sub

Private Sub Form_Load()
    Set cQrCode = New ClsQrCode
    CmdCodificTxt_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cQrCode = Nothing
End Sub

