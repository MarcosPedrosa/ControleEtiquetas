VERSION 5.00
Begin VB.Form frmEtiquetaDigitaQtde 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Digitação Quantidade no Pallet"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   4035
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_qtde 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2010
      TabIndex        =   1
      Top             =   330
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Quantidade :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmEtiquetaDigitaQtde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável para MDIapp
Public Flag_ativo As Boolean 'Conterá true se o form ja foi ativado

Private Sub txt_qtde_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   If Not VBA.IsNumeric(Me.txt_qtde.Text) Then
      MsgBox "O valor digitado não é numérico.Favor, digite a quantidade com valores numéricos."
      Me.txt_qtde.Text = ""
      Me.txt_qtde.SetFocus
      Exit Sub
   End If
   Me.Hide
End If

End Sub
