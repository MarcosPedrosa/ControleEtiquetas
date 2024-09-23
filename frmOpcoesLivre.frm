VERSION 5.00
Begin VB.Form frmOpcoesLivre 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impressao Etiquetas"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimi&r"
      Height          =   735
      Left            =   3540
      Picture         =   "frmOpcoesLivre.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Prepare a impressora clique no Botão ao lado para Impressão da Etiqueta."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   210
      TabIndex        =   1
      Top             =   210
      Width           =   3105
   End
End
Attribute VB_Name = "frmOpcoesLivre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Peca As String
Public Lote As String

Public FlagGeral As Boolean
Public FlagPecas As Boolean
Public FlagLote As Boolean
Dim gNumeroArquivo As Integer
Dim ultimoTipoEtiqueta As String

Private Sub cmdImprimir_Click()
Rem aqui marcos botao imprimir

    Dim Vezes As Integer
    Dim nx As Double
    Dim X As Printer
               
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
    
    If Forms(v).Name = "frmExibicao41" Then
       If objApplication.filial = adMusashiDaAmazonia Then
          Printer.Orientation = 1
       Else
          Printer.Orientation = 2
       End If
    
       frmAvulsoPadraoPonteiro.PrintForm
       Printer.EndDoc
       Exit For
    End If
                
        
    If Dir("etiq.txt") = "etiq.txt" Then
        Close gNumeroArquivo
        Kill "etiq.txt"
    End If
End Sub

Private Sub Form_Activate()
Dim Y As Integer

Close 11
Open "C:\Sistemas\mussashi\Desenvolvimento\Fontes\MD_Etiquetas\etiqshowa.txt" For Random Access Read Write As #11 Len = Len(Arq_Mov_Showa)

For Y = 1 To LOF(11) / Len(Arq_Mov_Showa)
    Get 11, Y, Arq_Mov_Showa
  
  If Arq_Mov_Showa.final <> Chr$(13) + Chr$(10) Then
     MsgBox "Arquivo lido , difere do tamanho do correspondente para sua importação."
     Close #11
     Exit Function
  End If
  
Next

Kill "C:\Sistemas\mussashi\Desenvolvimento\Fontes\MD_Etiquetas\etiqshowa.txt"
Open "C:\Sistemas\mussashi\Desenvolvimento\Fontes\MD_Etiquetas\etiqshowa.txt" For Random Access Read Write As #11 Len = Len(Arq_Mov_Showa)

End Sub

Private Sub Form_Load()
    
    'Habilita o flag geral, optGeral está selecionado
    FlagGeral = True
    
    Me.Left = 0
    Me.Top = 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Cancel = 0 Then
   Call Fechar_Form_Etiqueta
End If
End Sub
