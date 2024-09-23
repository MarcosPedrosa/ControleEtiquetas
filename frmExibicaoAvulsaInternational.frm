VERSION 5.00
Begin VB.Form frmExibicaoAvulsaInternational 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Previa solicitacao Avulsa Internation"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   8745
   Begin VB.Frame frmDigitacao 
      Caption         =   "Digite o Número do Pallet"
      Height          =   2055
      Left            =   150
      TabIndex        =   26
      Top             =   30
      Width           =   8325
      Begin VB.ComboBox cbo_impressora 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1500
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   1260
         Width           =   4125
      End
      Begin VB.CommandButton btoConfirma 
         Height          =   435
         Left            =   7050
         Picture         =   "frmExibicaoAvulsaInternational.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Confirmar sequencial"
         Top             =   420
         Width           =   495
      End
      Begin VB.CommandButton cmd_limpar 
         Height          =   435
         Left            =   7560
         Picture         =   "frmExibicaoAvulsaInternational.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Fecha a solicitação"
         Top             =   420
         Width           =   495
      End
      Begin VB.CommandButton cmdfechar 
         Height          =   735
         Left            =   7110
         Picture         =   "frmExibicaoAvulsaInternational.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Fecha tela de desmenbramento"
         Top             =   1020
         Width           =   975
      End
      Begin VB.CommandButton cmd_Impressao 
         Caption         =   "&Imprime"
         Enabled         =   0   'False
         Height          =   735
         Left            =   6060
         Picture         =   "frmExibicaoAvulsaInternational.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir"
         Top             =   1020
         Width           =   975
      End
      Begin VB.TextBox txtsequencial 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1500
         MaxLength       =   20
         TabIndex        =   0
         Text            =   "0000000000"
         Top             =   390
         Width           =   6585
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Impressora:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   60
         TabIndex        =   33
         Top             =   1290
         Width           =   1245
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Número :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   60
         TabIndex        =   27
         Top             =   510
         Width           =   945
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   435
      Left            =   6180
      TabIndex        =   32
      Top             =   5070
      Width           =   2325
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   405
      Left            =   210
      TabIndex        =   30
      Top             =   900
      Width           =   8265
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   405
      Left            =   210
      TabIndex        =   11
      Top             =   90
      Width           =   8265
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   435
      Left            =   210
      TabIndex        =   16
      Top             =   5040
      Width           =   2325
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   8460
      X2              =   180
      Y1              =   5070
      Y2              =   5070
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "INTERNATIONAL MOTORES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   330
      Left            =   2580
      TabIndex        =   31
      Top             =   5130
      Width           =   3645
   End
   Begin VB.Label lblNumPeca 
      BackStyle       =   0  'Transparent
      Caption         =   "INTERNATIONAL MOTORES - CANOAS/RS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   510
      Left            =   450
      TabIndex        =   29
      Top             =   480
      Width           =   7755
   End
   Begin VB.Label lblnota2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "123456 - 123456 - 123456 - 123456 - 123456 - 123456"
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
      Left            =   660
      TabIndex        =   28
      Top             =   3480
      Width           =   7380
   End
   Begin VB.Label lblqtdecaixa 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      TabIndex        =   25
      Top             =   4470
      Width           =   855
   End
   Begin VB.Label lblnota1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "123456 - 123456 - 123456 - 123456 - 123456 - 123456"
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
      Left            =   660
      TabIndex        =   24
      Top             =   2700
      Width           =   7380
   End
   Begin VB.Label lblPeso 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "1234,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2250
      TabIndex        =   23
      Top             =   4500
      Width           =   1380
   End
   Begin VB.Label lblano2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7890
      TabIndex        =   22
      Top             =   4530
      Width           =   225
   End
   Begin VB.Label lblano1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7170
      TabIndex        =   21
      Top             =   4530
      Width           =   225
   End
   Begin VB.Label lblmes2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6480
      TabIndex        =   20
      Top             =   4530
      Width           =   225
   End
   Begin VB.Label lblMes1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5790
      TabIndex        =   19
      Top             =   4530
      Width           =   225
   End
   Begin VB.Label lbldia2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5085
      TabIndex        =   18
      Top             =   4530
      Width           =   225
   End
   Begin VB.Label lbldia1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000E&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4380
      TabIndex        =   17
      Top             =   4530
      Width           =   225
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   7650
      X2              =   7650
      Y1              =   4470
      Y2              =   5025
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   6945
      X2              =   6945
      Y1              =   4470
      Y2              =   5025
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   6255
      X2              =   6255
      Y1              =   4470
      Y2              =   5025
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   5550
      X2              =   5550
      Y1              =   4470
      Y2              =   5025
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº da Nota Fiscal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   300
      TabIndex        =   15
      Top             =   2340
      Width           =   1575
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4860
      X2              =   4860
      Y1              =   4470
      Y2              =   5025
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4140
      TabIndex        =   14
      Top             =   4170
      Width           =   435
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Peso Bruto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2220
      TabIndex        =   13
      Top             =   4170
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde de Caixa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   300
      TabIndex        =   12
      Top             =   4170
      Width           =   1275
   End
   Begin VB.Label lblNumFornec2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Fornecedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   300
      TabIndex        =   10
      Top             =   1380
      Width           =   1935
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   8490
      X2              =   210
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Shape shp_pallet 
      BorderWidth     =   2
      Height          =   5400
      Left            =   210
      Top             =   90
      Width           =   8295
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   8490
      X2              =   210
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   2100
      X2              =   2100
      Y1              =   4110
      Y2              =   5040
   End
   Begin VB.Label lblNumFornec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MUSASHI DO BRASIL LTDA."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   420
      TabIndex        =   9
      Top             =   1650
      Width           =   5250
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4080
      X2              =   4080
      Y1              =   4110
      Y2              =   5040
   End
   Begin VB.Label lblEmbalagem2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   8040
      TabIndex        =   8
      Top             =   5010
      Width           =   45
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   8490
      X2              =   210
      Y1              =   4080
      Y2              =   4080
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
      Left            =   7440
      TabIndex        =   7
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label lbl_data 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "dd/mm/yyyy"
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
      Left            =   270
      TabIndex        =   6
      Top             =   5520
      Width           =   900
   End
End
Attribute VB_Name = "frmExibicaoAvulsaInternational"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cRecP As ADODB.Recordset 'conterá os dados do registro corrente da etiqueta principal
Public bAtivo As Boolean
Public bTelaImp As Boolean
Public bJafoi As Boolean

Private Sub cmd_Impressao_Click()
Dim v As Integer
Dim sData As String
Dim x As Printer
Dim nx As Integer
Dim nlblqtdecaixa As Integer
Dim nlblPeso As Double
Dim nNumImp As Integer

nx = 0
For Each x In Printers
    If UCase(Mid$(x.DeviceName, 1, 8)) = "ETIQUETA" Then
       Me.cbo_impressora.AddItem x.DeviceName
    End If
Next

Rem verificar se ha impressoras cadastradas
If Me.cbo_impressora.ListCount = 0 Then
   MsgBox "Impressoras ETIQUETAS,não encontradas no sistema, Favor comunicar ao responsável para adiciona-las no sistema!"
   End
End If
Me.cbo_impressora.ListIndex = 0

Rem verificar a impressora padrão para ser usada neste relatório
For nx = 0 To Me.cbo_impressora.ListCount - 1
    If Trim(UCase(sImpressoraFabrica)) = Trim(UCase(Me.cbo_impressora.List(nx))) Then
       Me.cbo_impressora.ListIndex = nx
    End If
Next
nx = 0

Rem************************************************
Rem AJUSTAR A TELA PARA AUMENTAR O TAMANHO E RETIRAR O FRAME
Rem************************************************

frmExibicaoAvulsaInternational.BackColor = &H8000000E
frmExibicaoAvulsaInternational.Height = 6060
Me.frmDigitacao.Visible = False
Rem************************************************
Rem************************************************


Rem************************************************
Rem preencher as datas da etiqueta
Rem************************************************
Me.lbldia1.Caption = Mid$(Format(Now(), "DD"), 1, 1)
Me.lbldia2.Caption = Mid$(Format(Now(), "DD"), 2, 1)
Me.lblMes1.Caption = Mid$(Format(Now(), "MM"), 1, 1)
Me.lblmes2.Caption = Mid$(Format(Now(), "MM"), 2, 1)
Me.lblano1.Caption = Mid$(Format(Now(), "YYYY"), 3, 1)
Me.lblano2.Caption = Mid$(Format(Now(), "YYYY"), 4, 1)
Rem************************************************
Rem************************************************

Rem************************************************
Rem calcular o peso total e as quantidades de caixas
Rem************************************************
nlblqtdecaixa = 0
nlblPeso = 0
cRecP.MoveFirst
While Not cRecP.EOF
      nlblqtdecaixa = nlblqtdecaixa + cRecP!QTDE
      nlblPeso = nlblPeso + cRecP!Peso
      cRecP.MoveNext
Wend
lblqtdecaixa.Caption = Format(nlblqtdecaixa, "0")
lblPeso.Caption = Format(nlblPeso, "0.00")
LBL_data.Caption = Trim(Me.txtsequencial.Text)
Rem************************************************
Rem************************************************

Rem************************************************
Rem imprimir as etiquetas ou etiquetas
Rem************************************************
cRecP.MoveFirst
nNumImp = 1

While Not cRecP.EOF
      lblnota1.Caption = ""
      lblnota2.Caption = ""
      For nx = 1 To 12
          If cRecP.EOF Then Exit For
          
          If nx < 7 Then
             If nx = 1 Then
                lblnota1.Caption = Format(cRecP!XBLNR, "000000")
             Else
                lblnota1.Caption = lblnota1.Caption & " - " & Format(cRecP!XBLNR, "000000")
             End If
          Else
             If nx = 7 Then
                lblnota2.Caption = Format(cRecP!XBLNR, "000000")
             Else
                lblnota2.Caption = lblnota2.Caption & " - " & Format(cRecP!XBLNR, "000000")
             End If
          End If
          cRecP.MoveNext
      Next
      
      If nNumImp > 1 Then
         lblqtdecaixa.Caption = ""
         lblPeso.Caption = ""
      End If
      
      Printer.Orientation = 2
      Me.PrintForm
      Printer.Orientation = 2: Printer.EndDoc
      nNumImp = nNumImp + 1
Wend

frmExibicaoAvulsaInternational.BackColor = &H8000000F
frmExibicaoAvulsaInternational.Height = 2250
Me.frmDigitacao.Visible = True


End Sub

Private Sub cmd_limpar_Click()
Me.txtsequencial.Text = ""
Me.txtsequencial.SetFocus
End Sub

Private Sub Form_Activate()
If bAtivo Then Exit Sub
bAtivo = True
bJafoi = False
Me.txtsequencial.Text = ""
Me.txtsequencial.SetFocus
End Sub

Private Sub Form_Load()
Dim x As Printer
Dim nx As Integer

For Each x In Printers
    If UCase(Mid$(x.DeviceName, 1, 8)) = "ETIQUETA" Then
       Me.cbo_impressora.AddItem x.DeviceName
    End If
Next

Rem verificar se ha impressoras cadastradas
If Me.cbo_impressora.ListCount = 0 Then
   MsgBox "Impressoras ETIQUETAS,não encontradas no sistema, Favor comunicar ao responsável para adiciona-las no sistema!"
   End
End If
Me.cbo_impressora.ListIndex = 0

Rem verificar a impressora padrão para ser usada neste relatório
For nx = 0 To Me.cbo_impressora.ListCount - 1
    If Trim(UCase(sImpressoraFabrica)) = Trim(UCase(Me.cbo_impressora.List(nx))) Then
       Me.cbo_impressora.ListIndex = nx
    End If
Next

Me.Left = 0
Me.Top = 0
bTelaImp = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
bAtivo = False
'Fecha o form q estiver aberto
'Call Fechar_Form_Etiqueta
End Sub
Private Sub btoConfirma_Click()

Dim nx As Integer
Dim cFields As Collection

On Error GoTo Erro

Me.MousePointer = vbHourglass

Set cFields = New Collection

If Len(Trim(Me.txtsequencial.Text)) = 0 Then
   MsgBox "Digite o número de um pallet."
   Me.txtsequencial.SetFocus
   Me.MousePointer = vbDefault
   Exit Sub
End If

cFields.Add Me.txtsequencial.Text

Set cRecP = New ADODB.Recordset

Set cRecP = CCTempneMov_Etiq.Mov_Pallet_Consulta(sBancoMusashi, _
                                                 cFields)

If cRecP.RecordCount = 0 Then
   MsgBox "Não existe Pallet com este número, Redigite."
   Me.txtsequencial.Text = ""
   Me.txtsequencial.SetFocus
   Me.MousePointer = vbDefault
   Set cRecP = Nothing
   Exit Sub
End If

Me.cmd_Impressao.Enabled = True
Me.MousePointer = vbDefault

Exit Sub

Erro:

MsgBox Err.Description
Me.MousePointer = vbDefault
Set cRecP = Nothing

End Sub

Private Sub cmdfechar_Click()
Unload Me
End Sub

Private Sub txtsequencial_GotFocus()
btoConfirma.Default = True
End Sub
