VERSION 5.00
Begin VB.MDIForm MDIFiat 
   BackColor       =   &H8000000C&
   Caption         =   "Emissão de Etiquetas - FIAT"
   ClientHeight    =   4470
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5700
   Icon            =   "MDIFiat.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuEtiqueta 
      Caption         =   "&Etiqueta"
      Begin VB.Menu MnuEntrada 
         Caption         =   "Entrar dados"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrevia 
         Caption         =   "Prévia de impressão"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImprimir 
         Caption         =   "Imprimir"
      End
   End
   Begin VB.Menu mnuSair 
      Caption         =   "Sai&r"
   End
End
Attribute VB_Name = "MDIFiat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    frmInformacoes.Show
End Sub

Private Sub mnuSair_Click()
    End
End Sub


