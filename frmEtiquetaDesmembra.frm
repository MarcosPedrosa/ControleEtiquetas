VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmEtiquetaDesmembra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desmembrar uma etiqueta"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5700
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
      Left            =   120
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   4170
      Width           =   3315
   End
   Begin VB.Frame frm_usuario 
      BackColor       =   &H00FF8080&
      Caption         =   "Login do usu�rio"
      Height          =   1665
      Left            =   660
      TabIndex        =   23
      Top             =   5100
      Width           =   3945
      Begin VB.TextBox txtPassword 
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   855
         PasswordChar    =   "*"
         TabIndex        =   25
         Top             =   615
         Width           =   2325
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   390
         Left            =   2025
         TabIndex        =   27
         Top             =   1125
         Width           =   1140
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   390
         Left            =   855
         TabIndex        =   26
         Top             =   1125
         Width           =   1140
      End
      Begin VB.TextBox txtUserName 
         Height          =   345
         Left            =   855
         TabIndex        =   24
         Top             =   240
         Width           =   2325
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Senha:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   29
         Top             =   645
         Width           =   510
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Usu�rio:"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   28
         Top             =   255
         Width           =   585
      End
   End
   Begin VB.CommandButton cmd_Impressao 
      Enabled         =   0   'False
      Height          =   735
      Left            =   3600
      Picture         =   "frmEtiquetaDesmembra.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprimir"
      Top             =   3870
      Width           =   975
   End
   Begin VB.Frame frm_desmembra 
      Caption         =   "Definir quantidade de caixas a serem geradas"
      Height          =   2055
      Left            =   60
      TabIndex        =   13
      Top             =   1710
      Width           =   5565
      Begin VB.CommandButton cmd_proximo 
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         Height          =   585
         Left            =   4770
         MaskColor       =   &H8000000F&
         Picture         =   "frmEtiquetaDesmembra.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Pr�ximo"
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton cmd_anterior 
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         Height          =   585
         Left            =   4050
         Picture         =   "frmEtiquetaDesmembra.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Anterior"
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton cmd_visualizar 
         Enabled         =   0   'False
         Height          =   585
         Left            =   4440
         Picture         =   "frmEtiquetaDesmembra.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Gerar etiquetas desmembradas de acordo as quantidades pedidas"
         Top             =   390
         Width           =   615
      End
      Begin VB.ComboBox cbo_qtd_Caixa 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   270
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Defina quantas caixa vai dividir as quantidades"
         Top             =   630
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid Grd_qtde 
         Height          =   1695
         Left            =   1560
         TabIndex        =   5
         ToolTipText     =   "Digite as quantidades em cada caixa, caso seja para desmenbrar digite 'S' ou brancos"
         Top             =   300
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   2990
         _Version        =   393216
         Cols            =   3
         ForeColorFixed  =   16711680
         BackColorSel    =   16711935
         AllowBigSelection=   0   'False
         Enabled         =   0   'False
         TextStyle       =   2
         TextStyleFixed  =   2
         FocusRect       =   2
         HighLight       =   2
         GridLines       =   2
         FormatString    =   $"frmEtiquetaDesmembra.frx":0FD0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lbl_totalqtde 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   390
         TabIndex        =   22
         Top             =   1620
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   420
         TabIndex        =   21
         Top             =   1380
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Qtde.Caixas :"
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
         Left            =   150
         TabIndex        =   14
         Top             =   300
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdfechar 
      Height          =   735
      Left            =   4650
      Picture         =   "frmEtiquetaDesmembra.frx":0FD9
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Fecha tela de desmenbramento"
      Top             =   3870
      Width           =   975
   End
   Begin VB.Frame frm_etiqueta 
      Caption         =   "Filtro do sequencial ou Escanei a etiqueta"
      Height          =   1485
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   5595
      Begin VB.CommandButton cmd_limpar 
         Height          =   435
         Left            =   4860
         Picture         =   "frmEtiquetaDesmembra.frx":141B
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Fecha a solicita��o"
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton btoConfirma 
         Height          =   435
         Left            =   4350
         Picture         =   "frmEtiquetaDesmembra.frx":185D
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Confirmar sequencial"
         Top             =   360
         Width           =   495
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
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "0000000000"
         Top             =   330
         Width           =   3825
      End
      Begin VB.Label lbl_qtd 
         AutoSize        =   -1  'True
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3750
         TabIndex        =   20
         Top             =   1200
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Qtde.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3120
         TabIndex        =   19
         Top             =   1200
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lbl_peca 
         AutoSize        =   -1  'True
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1560
         TabIndex        =   18
         Top             =   1200
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label lbl_produto 
         AutoSize        =   -1  'True
         Caption         =   "Pe�a.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   1200
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label lbl_descricao 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   1560
         TabIndex        =   16
         Top             =   930
         Visible         =   0   'False
         Width           =   3915
      End
      Begin VB.Label label33 
         AutoSize        =   -1  'True
         Caption         =   "Descri��o.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   930
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sequencial :"
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
         Left            =   150
         TabIndex        =   12
         Top             =   450
         Width           =   1305
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Impressora:."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1050
      TabIndex        =   31
      Top             =   3900
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label4 
      Caption         =   "aten��o a tela de login do usuario esta abaixo, aumente o form"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   210
      TabIndex        =   30
      Top             =   4560
      Visible         =   0   'False
      Width           =   3075
   End
End
Attribute VB_Name = "frmEtiquetaDesmembra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public cRecP As ADODB.Recordset 'conter� os dados do registro corrente da etiqueta principal
Public cRec As ADODB.Recordset 'conter� os dados do registro corrente
Public bAtivo As Boolean
Public bTelaImp As Boolean
Public bJafoi As Boolean
Public nLogin As Integer ' conter� o codigo do usuario, quando confirmar a senha
Public nTipo As Integer ' conter� o tipo do usuario, quando confirmar a senha
Public nMatricula As Double ' conter� a matricula do usuario, quando confirmar a senha
Public ultimoTipoEtiqueta As String
Public sEtiqueta As String
Public oTela As frmExibicao12
Public bJaModificou As Boolean

Private Sub cbo_qtd_Caixa_Click()
If bAtivo Then Call carrega_Grd_qtde
End Sub

Private Sub cmd_Anterior_Click()

If cRec Is Nothing Then Exit Sub

If cRec.BOF Then
   cRec.MoveFirst
   Me.cmd_anterior.BackColor = &HFF& 'vermelho
   Me.cmd_proximo.BackColor = &HFF00& ' verde
Else
   cRec.MovePrevious
   If cRec.BOF Then
      Me.cmd_anterior.BackColor = &HFF& 'vermelho
      Me.cmd_proximo.BackColor = &HFF00& ' verde
      cRec.MoveFirst
   Else
      Me.cmd_proximo.BackColor = &HFF00& ' verde
   End If
End If
Call Visualizar_Etiqueta

End Sub

Private Sub cmd_Impressao_Click()

Dim v As Integer
Dim sData As String
Dim x As Printer
Dim nx As Integer

nx = 0
For Each x In Printers
   If InStr(1, UCase(x.DeviceName), UCase(Me.cbo_impressora.List(Me.cbo_impressora.ListIndex))) > 0 Then
      Set Printer = x
      nx = 1
      Exit For
   End If
Next

If nx = 0 Then
   MsgBox "Impressora da Etiqueta n�o encontrada, chame o respons�vel. o nome da impressora ser� - 'ETIQUETA FABRICA'"
   Exit Sub
End If
nx = 0

cRec.MoveFirst

While Not cRec.EOF

      Call Visualizar_Etiqueta
      
      For v = 0 To (Forms.Count - 1)
      
'          If Forms(v).Name = "frmExibicao1" Or Forms(v).Name = "frmExibicao4" Then
'                If objApplication.filial = adMusashiDaAmazonia Then
'                   Printer.Orientation = 1
'                Else
'                   Printer.Orientation = 2
'                End If
'                frmAvulsoPadraoPonteiro.PrintForm
'                Printer.Orientation = 2 : Printer.EndDoc
'              Exit For
'          End If
          
          If Forms(v).Name = "frmExibicao1" Or Forms(v).Name = "frmExibicao4" Or Forms(v).Name = "frmExibicao4MDA" Then
             If objApplication.filial = adMusashiDaAmazonia Then
                Printer.Orientation = 1
             Else
                Printer.Orientation = 2
             End If
             If Forms(v).Name = "frmExibicao4MDA" Then
                frmExibicao4MDA.PrintForm
             Else
                frmAvulsoPadraoPonteiro.PrintForm
             End If
             Printer.Orientation = 2: Printer.EndDoc
             Exit For
          End If
         If Forms(v).Name = "frmExibicao2" Then
             
             If ultimoTipoEtiqueta = "F" Then
                Call Imprime_Etiqueta_Fiat_FTP
             ElseIf ultimoTipoEtiqueta = "Z" Then
                Printer.Orientation = 1
                frmExibicao2.nTamannhowidth = Me.Width
                frmExibicao2.PrintForm
                Call Imprime_Etiqueta_Fiat_FTP
                Call Imprime_Etiqueta_Fiat_FTP
             Else
                Printer.Orientation = 1
                frmExibicao2.nTamannhowidth = Me.Width
                frmExibicao2.PrintForm
                oTela.Visible = True
                oTela.Show
                Printer.Orientation = 2
                oTela.PrintForm
                oTela.PrintForm
                Unload oTela: Set oTela = Nothing
                Printer.Orientation = 2: Printer.EndDoc
             End If
              
'              Printer.Orientation = 1
'              frmExibicao2.PrintForm
'              Printer.Orientation = 2 : Printer.EndDoc
              Exit For
          End If
          If Forms(v).Name = "frmExibicao3" Then
              Printer.Orientation = 2
              frmExibicao3.PrintForm
              Printer.Orientation = 2: Printer.EndDoc
              Exit For
          End If
          If Forms(v).Name = "frmExibicao5" Then
              Printer.Orientation = 2
              frmExibicao5.PrintForm
              Printer.Orientation = 2: Printer.EndDoc
              Exit For
          End If
          If Forms(v).Name = "frmExibicao6" Then
              Printer.Orientation = 2
              frmExibicao6.PrintForm
              Printer.Orientation = 2: Printer.EndDoc
              Exit For
          End If
          If Forms(v).Name = "frmExibicao7UmProduto" Or Forms(v).Name = "frmExibicao7VariosProdutos" Then
              Printer.Orientation = 2
              frmExibicao7Ref.PrintForm
              Printer.Orientation = 2: Printer.EndDoc
              Exit For
          End If
      Rem incluido para impressao do novo formulario da mvm international motores
          If Forms(v).Name = "frmExibicao9" Then
              If frmExibicao9.PICT_CRISTAL.Visible = True Then
                 Call Imprime_etiqueta_MWM_Cristal
              Else
                 Printer.Orientation = 2
                 frmExibicao9.PrintForm
                 Printer.Orientation = 2: Printer.EndDoc
              End If
              Exit For
          End If
          If Forms(v).Name = "frmExibicao10" Then
              Printer.Orientation = 2
              frmExibicao10.PrintForm
              Printer.Orientation = 2: Printer.EndDoc
              Exit For
          End If
          If Forms(v).Name = "frmExibicao11" Then
              Printer.Orientation = 2
              frmExibicao11.PrintForm
              Printer.Orientation = 2: Printer.EndDoc
              Exit For
          End If
      Next
      cRec.MoveNext
      
Wend
        
MsgBox "Impress�o conclu�da com sucesso! a tela da etiqueta ser� encerrada!", vbOKOnly + vbInformation, "Tarefa Conclu�da"

Call Fechar_Form_Etiqueta
cmd_limpar_Click

End Sub

Private Sub cmd_limpar_Click()

Call Limpar_Grd_qtde
Call Fechar_Form_Etiqueta

Me.txtsequencial.Text = ""
Me.LBL_descricao.Caption = ""
Me.lbl_peca.Caption = ""
Me.lbl_qtd.Caption = ""

Me.LBL_descricao.Visible = False
Me.lbl_peca.Visible = False
Me.lbl_qtd.Visible = False
Me.label33.Visible = False
Me.lbl_produto.Visible = False
Me.Label5.Visible = False
Me.cbo_qtd_Caixa.Enabled = False
Me.Grd_qtde.Enabled = False

Me.cmd_anterior.Enabled = False
Me.cmd_proximo.Enabled = False
Me.cmd_anterior.BackColor = &HFF& 'vermelho
Me.cmd_proximo.BackColor = &HFF& 'vermelho
Me.cmd_Impressao.Enabled = False
Me.cmd_Visualizar.Enabled = False

Me.txtsequencial.Enabled = True
Me.cbo_qtd_Caixa.ListIndex = 0
Me.frm_etiqueta.Enabled = True
Me.txtsequencial.Enabled = True
Me.txtsequencial.Locked = False
Me.txtsequencial.SetFocus

End Sub

Private Sub cmd_Proximo_Click()

If cRec Is Nothing Then Exit Sub

If cRec.EOF Then
      Me.cmd_anterior.BackColor = &HFF00& ' verde
      Me.cmd_proximo.BackColor = &HFF& 'vermelho
Else
   cRec.MoveNext
   If cRec.EOF Then
      Me.cmd_anterior.BackColor = &HFF00& ' verde
      Me.cmd_proximo.BackColor = &HFF& 'vermelho
      cRec.MovePrevious
   Else
      Me.cmd_anterior.BackColor = &HFF00& ' verde
   End If
End If
Call Visualizar_Etiqueta

End Sub

Private Sub cmd_visualizar_Click()

Dim nx As Integer
Dim cFields As Collection
Dim cFieldsDes As Collection
Dim sdata_aux As String

On Error GoTo Erro

If MsgBox("Voc� desmembrar� esta etiqueta ap�s confirma��o, Deseja continuar?", vbQuestion + vbYesNo, "ATEN��O !!!") = vbNo Then Exit Sub

Me.MousePointer = vbHourglass

Me.txtsequencial.Text = Format(Me.txtsequencial.Text, "0000000000")

Set cFields = New Collection
Set cFieldsDes = New Collection

Me.Grd_qtde.col = 1
Me.Grd_qtde.Row = 1

cFields.Add Me.lbl_qtd.Caption ' poe a quantidade total para critica do desmembramento, e as que devem ser deletadas
cFields.Add nMatricula ' Atualizar o campo da matricula da etiqueta desmembrada, saber quem desmembrou na impressao da etiq.

'''Rem estes campos � para ficar com o mesmo numero de itens para a proxima camada
'''cFields.Add ""
'''cFields.Add ""

For nx = 1 To Me.Grd_qtde.Rows - 1
    Me.Grd_qtde.Row = nx
    Me.Grd_qtde.col = 1
    cFields.Add Me.Grd_qtde.Text
    Me.Grd_qtde.col = 2
    cFieldsDes.Add Trim(Me.Grd_qtde.Text)
Next

Set cRec = New ADODB.Recordset

Set cRec = CCTempneMov_Etiq.Mov_Etiq_Desmembrar(sBancoMusashi, _
                                                Me.txtsequencial.Text, _
                                                str$(nLogin), _
                                                cFields, _
                                                cFieldsDes)

cRec.MoveFirst
Call Visualizar_Etiqueta

Me.cmd_Visualizar.Enabled = False
Me.Grd_qtde.Enabled = False
Me.cbo_qtd_Caixa.Enabled = False
Me.cmd_Impressao.Enabled = True
Me.cmd_anterior.BackColor = &HFF00& ' verde
Me.cmd_proximo.BackColor = &HFF00& 'verde
Me.cmd_anterior.Enabled = True
Me.cmd_proximo.Enabled = True

Me.MousePointer = vbDefault

Exit Sub

Erro:

MsgBox Err.Description
Me.MousePointer = vbDefault
If Err.Number = 50001 Then
   Me.Grd_qtde.SetFocus
End If
If Err.Number = 50002 Then
   cmd_limpar_Click
   Me.txtsequencial.SetFocus
End If

End Sub

Private Sub cmd_visualizar_GotFocus()
Dim nQtdeAux As Integer
Dim nx As Integer

Me.Grd_qtde.Row = 1

nQtdeAux = 0
Me.Grd_qtde.col = 1
For nx = 1 To Me.Grd_qtde.Rows - 1
   Me.Grd_qtde.Row = nx
   nQtdeAux = nQtdeAux + CInt(Me.Grd_qtde.Text)
Next

Me.lbl_totalqtde.Caption = nQtdeAux
Me.Label3.Visible = True
Me.lbl_totalqtde.Visible = True

Me.Grd_qtde.Row = 1

End Sub

Private Sub cmd_visualizar_LostFocus()
Me.Label3.Visible = False
Me.lbl_totalqtde.Visible = False
End Sub

Private Sub cmdCancel_Click()
Me.frm_usuario.Top = 4950
Me.frm_usuario.Left = 900
Me.frm_desmembra.Enabled = True
Me.frm_etiqueta.Enabled = True

cmd_limpar_Click
End Sub

Private Sub cmdOk_Click()
Dim rRecLogin As ADODB.Recordset

On Error GoTo Erro

Me.MousePointer = vbHourglass

Set rRecLogin = New ADODB.Recordset

Set rRecLogin = CCTempneUsuario.USUARIO_Confirmar_Login(sBancoMusashi, _
                                                        Me.txtUserName.Text, _
                                                        Me.txtPassword.Text)

nLogin = rRecLogin!codigo
nTipo = rRecLogin!Tipo
nMatricula = rRecLogin!matricula

Me.MousePointer = vbDefault
Me.frm_usuario.Top = 4950
Me.frm_usuario.Left = 900
Me.frm_desmembra.Enabled = True
Me.frm_etiqueta.Enabled = True
Call carrega_Grd_qtde
Me.cbo_qtd_Caixa.SetFocus

Exit Sub

Erro:

MsgBox Err.Description
Me.MousePointer = vbDefault
Me.txtPassword.SetFocus

End Sub

Private Sub Form_Activate()

'If bTelaImp Then
'   cmd_limpar_Click
'   bTelaImp = False
'End If
Me.cbo_qtd_Caixa.ListIndex = 0

If bAtivo Then Exit Sub
Me.cmd_limpar.Default = False
Me.cmdCancel.Default = False
bAtivo = True
bJafoi = False
Call Limpar_Grd_qtde
Me.txtsequencial.Text = ""
Me.txtsequencial.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' para funcionar , tem que mudar o keyPreviwe=true
If KeyCode = 13 Then
   If Me.ActiveControl.TabIndex > 0 Then
      If Me.ActiveControl.TabIndex = 5 Then
         Me.cmd_Visualizar.Enabled = True ' Habilitar o bot�o de gerar etiqueta
      End If
'      SendKeys "{TAB}"
   End If
ElseIf KeyCode = 27 Then
   If Me.ActiveControl.TabIndex < 1 Then
      If 6 = MsgBox("Deseja realmente sair deste m�dulo?", 32 + 4) Then
         Unload Me
      End If
   Else
       SendKeys "+{TAB}" ' retornar campo
   End If
End If

End Sub

Private Sub Form_Load()
Dim x As Printer
Dim nx As Integer

nx = 0
For Each x In Printers
    If UCase(Mid$(x.DeviceName, 1, 8)) = "ETIQUETA" Then
       Me.cbo_impressora.AddItem x.DeviceName
    End If
Next

Rem verificar se ha impressoras cadastradas
If Me.cbo_impressora.ListCount = 0 Then
   MsgBox "Impressoras ETIQUETAS,n�o encontradas no sistema, Favor comunicar ao respons�vel para adiciona-las no sistema!"
   End
End If
Me.cbo_impressora.ListIndex = 0

Rem verificar a impressora padr�o para ser usada neste relat�rio
For nx = 0 To Me.cbo_impressora.ListCount - 1
    If Trim(UCase(sImpressoraFabrica)) = Trim(UCase(Me.cbo_impressora.List(nx))) Then
       Me.cbo_impressora.ListIndex = nx
    End If
Next
nx = 0
Me.Left = 0
Me.Top = 0
bTelaImp = False

Me.cbo_qtd_Caixa.Clear
For nx = 2 To 20
    Me.cbo_qtd_Caixa.AddItem nx
Next

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
bAtivo = False
'Fecha o form q estiver aberto
Call Fechar_Form_Etiqueta
End Sub
Private Sub btoConfirma_Click()

Dim nx As Integer

On Error GoTo Erro

Me.MousePointer = vbHourglass

Set cRecP = New ADODB.Recordset

Set cRecP = CCTempneMov_Etiq.Mov_Etiq_Consultar_Filtro(sBancoMusashi, _
                                                       Format(Me.txtsequencial.Text, "0000000000"), _
                                                       "", _
                                                       "", _
                                                       "", _
                                                       "")

Call Limpar_Grd_qtde

Me.LBL_descricao.Caption = cRecP!Descr_Peca
Me.lbl_peca.Caption = cRecP!Cod_Peca
Me.lbl_qtd.Caption = cRecP!Qtd_Caixa

Me.LBL_descricao.Visible = True
Me.lbl_peca.Visible = True
Me.lbl_qtd.Visible = True
Me.label33.Visible = True
Me.lbl_produto.Visible = True
Me.Label5.Visible = True
Me.cbo_qtd_Caixa.Enabled = True
Me.Grd_qtde.Enabled = True
Me.txtsequencial.Enabled = False

If MsgBox("Para desmembrar uma etquieta, precisar� ser usu�rio cadastrado, Deseja continuar?", vbQuestion + vbYesNo, "ATEN��O !!!") = vbNo Then
   Me.frm_usuario.Top = 4950
   Me.frm_usuario.Left = 900
   Me.MousePointer = vbDefault
   Call cmd_limpar_Click
   
   Exit Sub
End If
bJaModificou = True
Me.txtsequencial.Text = Format(Me.txtsequencial.Text, "0000000000")
bJaModificou = False
Me.txtUserName.Text = ""
Me.txtPassword.Text = ""
Me.frm_desmembra.Enabled = False
Me.frm_etiqueta.Enabled = False
Me.frm_usuario.Top = 2000
Me.frm_usuario.Left = 1000
Me.txtUserName.SetFocus

Me.MousePointer = vbDefault

Exit Sub

Erro:

MsgBox Err.Description
Me.MousePointer = vbDefault

End Sub

Private Sub cmdfechar_Click()
Unload Me
End Sub
Private Function Fechar_Form_Etiqueta()

Dim v As Integer

For v = 0 To (Forms.Count - 1)
    If Forms(v).Name = "frmExibicao1" Or Forms(v).Name = "frmExibicao4" Then
        Unload frmAvulsoPadraoPonteiro
        Exit For
    End If
    If Forms(v).Name = "frmExibicao4MDA" Then
        Unload frmExibicao4MDA
        Exit For
    End If
    If Forms(v).Name = "frmExibicao2" Then
        Unload frmExibicao2
        Exit For
    End If
    If Forms(v).Name = "frmExibicao3" Then
        Unload frmExibicao3
        Exit For
    End If
    If Forms(v).Name = "frmExibicao5" Then
        Unload frmExibicao5
        Exit For
    End If
    If Forms(v).Name = "frmExibicao7UmProduto" Then
        Unload frmExibicao7UmProduto
        Exit For
    End If
    If Forms(v).Name = "frmExibicao7VariosProdutos" Then
        Unload frmExibicao7VariosProdutos
        Exit For
    End If
    If Forms(v).Name = "frmExibicao9" Then
        Unload frmExibicao9
        Exit For
    End If
    If Forms(v).Name = "frmExibicao10" Then
        Unload frmExibicao10
        Exit For
    End If
    If Forms(v).Name = "frmExibicao11" Then
        Unload frmExibicao11
        Exit For
    End If
Next
Unload frmOpcoes

End Function

Private Sub Grd_qtde_GotFocus()
Me.Grd_qtde.col = 1
Me.Grd_qtde.Row = 1
Me.cmd_Visualizar.Enabled = False
End Sub

Private Sub Grd_qtde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
   Me.cbo_qtd_Caixa.SetFocus
End If
End Sub

Private Sub Grd_qtde_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
   Case vbKeyReturn, vbKeyTab
   'move para a proxima celula.
   
   With Grd_qtde
   
     If .col + 1 <= .cols - 1 Then
        .col = .col + 1
     Else
        If .Row + 1 <= .Rows - 1 Then
            .Row = .Row + 1
            .col = 0
        Else
            .Row = 1
            .col = 0
        End If
     End If
   End With
   
   Case vbKeyBack
   
      With Grd_qtde
      'remove o ultimo caractere
         If Len(.Text) Then
            .Text = Left(.Text, Len(.Text) - 1)
         End If
      End With
   
   Case Is < 32
   
   Case Else
   
       If KeyAscii > 47 And KeyAscii < 58 Then
          With Grd_qtde
             If .col = 1 Then
                .Text = .Text & Chr(KeyAscii)
             End If
          End With
       ElseIf KeyAscii = 115 Or KeyAscii = 83 Then
          With Grd_qtde
             If .col = 2 Then
                If Len(Trim(.Text)) = 0 Then .Text = .Text & UCase(Chr(KeyAscii))
             End If
          End With
       End If
End Select

End Sub

Private Sub txtsequencial_Change()

If bJaModificou Then Exit Sub

If Len(Trim(txtsequencial.Text)) = 10 Then
   btoConfirma_Click
End If

End Sub

Private Sub txtsequencial_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'   Me.btoConfirma.SetFocus
'End If
'If KeyAscii = 27 Then
'   If MsgBox("Deseja sair deste m�dulo?", vbQuestion + vbYesNo, "ATEN��O !!!") = vbNo Then
'      Me.txtsequencial.Text = ""
'      Me.txtsequencial.SetFocus
'   Else
'      Unload Me
'   End If
'End If

End Sub

Private Sub Limpar_Grd_qtde()
Dim nx As Double
Dim nLinhas As Double
Dim nLinhas1 As Double

Grd_qtde.Clear
nLinhas = Grd_qtde.Rows
'Grd_qtde.Col = 1
'If Grd_qtde.Rows = 2 Then
'   If Grd_qtde.Text = "" Then Exit Sub
'End If

If Grd_qtde.Rows > 2 Then
   For nx = Grd_qtde.Rows To nLinhas1 - 2 Step -1
       If nx > 2 Then Grd_qtde.RemoveItem (nx)
   Next
End If

Grd_qtde.Row = 0

If nTipo = 1 Then
   Me.Grd_qtde.Font = 14
   Grd_qtde.col = 0: Grd_qtde.ColWidth(0) = 600:  Grd_qtde.Text = "SEQ"
   Grd_qtde.col = 1: Grd_qtde.ColWidth(1) = 800: Grd_qtde.Text = "QTDE": Me.Grd_qtde.TextStyle = flexTextInsetLight
   Grd_qtde.col = 1: Grd_qtde.BackColor = &H80FFFF
   Grd_qtde.col = 2: Grd_qtde.ColWidth(2) = 600:  Grd_qtde.Text = "Des": Me.Grd_qtde.TextStyle = flexTextInsetLight
   Grd_qtde.col = 2: Grd_qtde.BackColor = &H80FFFF
   Grd_qtde.ColAlignment(2) = flexAlignCenterCenter
Else
   Me.Grd_qtde.Font = 16
   Grd_qtde.col = 0: Grd_qtde.ColWidth(0) = 800:  Grd_qtde.Text = "SEQ"
   Grd_qtde.col = 1: Grd_qtde.ColWidth(1) = 1000: Grd_qtde.Text = "QTDE": Me.Grd_qtde.TextStyle = flexTextInsetLight
   Grd_qtde.col = 2: Grd_qtde.ColWidth(2) = 0
   Grd_qtde.col = 1: Grd_qtde.BackColor = &H80FFFF
End If
Grd_qtde.Row = 0

Grd_qtde.HighLight = False
Grd_qtde.ColAlignment(0) = flexAlignCenterCenter
Grd_qtde.ColAlignment(1) = flexAlignCenterCenter

End Sub

Private Sub carrega_Grd_qtde()
Dim nx As Double

Call Limpar_Grd_qtde

'If Me.cbo_qtd_Caixa.ListIndex = 0 Then Exit Sub

Grd_qtde.Row = 1
bJafoi = True

For nx = 1 To Me.cbo_qtd_Caixa.List(Me.cbo_qtd_Caixa.ListIndex)
    Grd_qtde.col = 0: Grd_qtde.Text = nx
    Grd_qtde.col = 1: Grd_qtde.Text = 0
    Grd_qtde.col = 2: Grd_qtde.Text = "S"
    
    If nx <> Me.cbo_qtd_Caixa.List(Me.cbo_qtd_Caixa.ListIndex) Then
       Grd_qtde.Rows = Grd_qtde.Rows + 1
       Grd_qtde.Row = Grd_qtde.Row + 1
    End If
Next

bJafoi = False
End Sub

Private Sub Visualizar_Etiqueta()
Rem  *********************************************************************************************************
Rem  *********************************************************************************************************
Rem  *********************************************************************************************************
Rem  *********************************************************************************************************
Rem  *********************************************************************************************************
    
Dim qtdeConteiners As Integer
Dim executaUnloadForm As Boolean
Dim sSeqId As String * 18 ' usado para formatar o campo ID, da MVM
Dim sSequencial As String
Dim nQuantidade As Integer
Dim sMes As String
Dim sdata_aux As String

    
Rem leitura no banco para emiss�o da etiqueta

On Error GoTo Erro

executaUnloadForm = False

'Fecha o form q estiver aberto
Call Fechar_Form_Etiqueta

ultimoTipoEtiqueta = cRec.Fields("Tipo").Value
nQuantidade = cRec.Fields("Qtd_Etiq")

'---------------------------------------------------------------------------------------------------
'Atualiza o form frmExibicao de acordo com o tipo
'Se tipo 1 opcao padr�o etiqueta pequena
If cRec.Fields("Tipo") = "1" Then
    frmAvulsoPadraoPonteiro.Show
    frmAvulsoPadraoPonteiro.Left = Me.Width
    If cRec.Fields("Cliente") <> "" Then
        frmAvulsoPadraoPonteiro.lblCliente.Caption = cRec.Fields("Cliente")
    End If
    
    If cRec.Fields("Cod_Cliente") <> "" Then
        frmAvulsoPadraoPonteiro.lblCodCliente2.Caption = cRec.Fields("Cod_Cliente")
    End If
    If cRec.Fields("Descr_Peca") <> "" Then
        frmAvulsoPadraoPonteiro.lblDescricao.Caption = cRec.Fields("Descr_Peca")
    End If
    If cRec.Fields("Lote") <> "" Then
        frmAvulsoPadraoPonteiro.lblLote2.Caption = cRec.Fields("Lote")
    End If
    If cRec.Fields("Peso") <> "" Then
        frmAvulsoPadraoPonteiro.lblPeso2.Caption = Format(cRec.Fields("Peso"), "0.00")
    End If
    If cRec.Fields("Qtd_Caixa") <> "" Then
        frmAvulsoPadraoPonteiro.lblQtd2.Caption = cRec.Fields("Qtd_Caixa")
    End If
    If cRec.Fields("Cod_Peca") <> "" Then
        frmAvulsoPadraoPonteiro.lblPeca.Caption = cRec.Fields("Cod_Peca")
    End If
    If Not IsNull(cRec.Fields("EMBALAGEM")) Then
        frmAvulsoPadraoPonteiro.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
    End If
    
    frmAvulsoPadraoPonteiro.lbl_data.Caption = Format(cRec.Fields("data_etiq"), "dd/mm/yyyy")
    frmAvulsoPadraoPonteiro.LBL_MES.Caption = Pega_Mes(Mid$(Format(cRec.Fields("data_etiq"), "dd/mm/yyyy"), 4, 2))
    
    If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
        frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption = Trim(cRec.Fields("ID_ETIQUETA"))
        frmAvulsoPadraoPonteiro.lblCodigoBarrasA.Caption = "*" & Trim(frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption) & "*"
        frmAvulsoPadraoPonteiro.lblCodigoBarrasB.Caption = "*" & Trim(frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption) & "*"
     Else
        frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption = ""
        frmAvulsoPadraoPonteiro.lblCodigoBarrasA.Caption = ""
        frmAvulsoPadraoPonteiro.lblCodigoBarrasB.Caption = ""
     End If
ElseIf (cRec.Fields("Tipo") = "4") Then
       frmAvulsoPadraoPonteiro.lblMsgProduto.Visible = False
       If (Val((cRec.Fields("id_cliente")) = 1) Or (Val(cRec.Fields("id_cliente")) = 2) Or (Val(cRec.Fields("id_cliente")) = 3)) Then
          frmAvulsoPadraoPonteiro.lblMsgProduto.Visible = True
       End If
End If

If cRec.Fields("Tipo") = "2" Or _
   cRec.Fields("Tipo") = "F" Or _
   cRec.Fields("Tipo") = "Z" Then
    frmExibicao2.nTamannhowidth = Me.Width
    frmExibicao2.Show
    'Mostra FIAT
    frmExibicao2.lblCod_Peca.Caption = cRec.Fields("Cod_Peca")
    If cRec.Fields("Data_Expedicao") <> "" Then
        frmExibicao2.lblDataExpedicao2.Caption = cRec.Fields("Data_Expedicao")
    End If
    If cRec.Fields("Cod_Fornecedor") <> "" Then
        frmExibicao2.lblCodFornec2.Caption = Format(cRec.Fields("Cod_Fornecedor"), "000000")
    End If
    If cRec.Fields("Descr_Peca") <> "" Then
        frmExibicao2.lblDenominacao2.Caption = cRec.Fields("Descr_Peca")
    End If
    If cRec.Fields("Num_Doc_Fiscal") <> "" Then
        frmExibicao2.lblBam2.Caption = cRec.Fields("Num_Doc_Fiscal")
    End If
    If cRec.Fields("Cod_Cliente") <> "" Then
        frmExibicao2.lblDesenho2.Caption = Format(cRec.Fields("Cod_Cliente"), "00000000000")
    End If
    
    frmExibicao2.lblCodBarra.Caption = Format(cRec.Fields("Cod_Cliente"), "00000000000") & Format(cRec.Fields("Qtd_Caixa"), "00000") _
                                  & cRec.Fields("Cod_Embalagem_pw") & Format(cRec.Fields("Cod_Fornecedor"), "000000")
    frmExibicao2.lblCodBarra2.Caption = frmExibicao2.lblCodBarra.Caption
    'Adicionar o * de inicio e fim
    frmExibicao2.lblCodBarra.Caption = "*" & frmExibicao2.lblCodBarra.Caption & "*"
    frmExibicao2.lblCodBarraCp1.Caption = frmExibicao2.lblCodBarra.Caption
    frmExibicao2.lblCodBarraCp2.Caption = frmExibicao2.lblCodBarra.Caption
    
    If cRec.Fields("Data_Lote") <> "" Then
        frmExibicao2.lblDataProducao2.Caption = cRec.Fields("Data_Lote")
    End If
    If cRec.Fields("Cod_Embalagem") <> "" Then
        frmExibicao2.lblCodEmbalagem2.Caption = cRec.Fields("Cod_Embalagem")
    End If
    If cRec.Fields("Lote") <> "" Then
        frmExibicao2.lblNumLote2.Caption = cRec.Fields("Lote")
    End If
    If cRec.Fields("Qtd_Lote") <> "" Then
        frmExibicao2.lblQtdLote2.Caption = cRec.Fields("Qtd_Lote")
    End If
    If cRec.Fields("Qtd_Caixa") <> "" Then
        frmExibicao2.lblQtdEmbalagem2.Caption = Format(cRec.Fields("Qtd_Caixa"), "00000")
    End If
    If cRec.Fields("Classe_Funcional") <> "" Then
        frmExibicao2.lblClasseFuncional2.Caption = cRec.Fields("Classe_Funcional")
    End If
    If cRec.Fields("Vinculo") <> "" Then
        frmExibicao2.lblVinculo2.Caption = cRec.Fields("Vinculo")
    End If
    If cRec.Fields("Ind_Suplementar") <> "" Then
        frmExibicao2.lblIndicacaoSuplementar2.Caption = cRec.Fields("Ind_Suplementar")
    End If
    If cRec.Fields("Embarque_Controlado") <> "" Then
        frmExibicao2.lblEmbarqueControlado2.Caption = cRec.Fields("Embarque_Controlado")
    End If
    If cRec.Fields("Desvio") <> "" Then
        frmExibicao2.lblLoteSobDesvio2.Caption = cRec.Fields("Desvio")
    End If
    If cRec.Fields("DUM") <> "" Then
        frmExibicao2.lblDum2.Caption = cRec.Fields("DUM")
    End If
    If Not IsNull(cRec.Fields("EMBALAGEM")) Then
        frmExibicao2.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
    Else
        frmExibicao2.lblEmbalagem2.Caption = ""
    End If
Rem acrescentado o pto de entrega - 09-09-2016
    If cRec.Fields("Pto_Entrega") <> "" Then
        frmExibicao2.lblPontoEntrega2.Caption = Trim(cRec.Fields("Pto_Entrega"))
    End If
    
    If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
        frmExibicao2.lblCodigoBarras.Caption = ""
        frmExibicao2.lblCodigoBarras.Caption = Trim(cRec.Fields("ID_ETIQUETA"))
        frmExibicao2.lblCodigoBarrasA.Caption = "*" & Trim(frmExibicao2.lblCodigoBarras.Caption) & "*"
        frmExibicao2.lblCodigoBarrasB.Caption = "*" & Trim(frmExibicao2.lblCodigoBarras.Caption) & "*"
'            frmExibicao2.lblCodigoBarrasC.Caption = "*" & frmExibicao2.lblCodigoBarras.Caption & "*"
'            frmExibicao2.lblCodigoBarrasD.Caption = "*" & frmExibicao2.lblCodigoBarras.Caption & "*"
     Else
        frmExibicao2.lblCodigoBarras.Caption = ""
        frmExibicao2.lblCodigoBarrasA.Caption = ""
        frmExibicao2.lblCodigoBarrasB.Caption = ""
     End If
     
     If cRec.Fields("Tipo") = "2" Then ' aqui ser� preenchido os dados para a outra etiqueta tipo expotacao da fiat italiana

        Set oTela = New frmExibicao12
        oTela.slbl_01 = IIf(IsNull(cRec.Fields("Cliente")), " ", Trim(cRec.Fields("Cliente")))
        oTela.slbl_02 = IIf(IsNull(cRec.Fields("Pto_Entrega")), " ", Trim(cRec.Fields("Pto_Entrega")))
        oTela.slbl_03 = " "
        oTela.slbl_04 = "MUSASHI DO BRASIL LTDA"
        oTela.slbl_07 = " "
        oTela.slbl_08 = " " ' FALTA VER COM MAURO
        oTela.slbl_09 = IIf(IsNull(cRec.Fields("Qtd_Caixa")), " ", Trim(cRec.Fields("Qtd_Caixa")))
        oTela.slbl_10 = IIf(IsNull(cRec.Fields("Descr_Peca")), " ", Trim(cRec.Fields("Descr_Peca")))
        oTela.slbl_11 = IIf(IsNull(cRec.Fields("Cod_Fornecedor")), " ", Trim(cRec.Fields("Cod_Fornecedor")))
        oTela.slbl_12 = IIf(IsNull(cRec.Fields("Cod_Embalagem")), " ", Trim(cRec.Fields("Cod_Embalagem")))
        oTela.slbl_13 = Format(Now(), "DD/MM/YYYY")
        oTela.slbl_14 = " "
        oTela.slbl_15 = " "
        oTela.slbl_16 = IIf(IsNull(cRec.Fields("Lote")), " ", Trim(cRec.Fields("Lote")))
        oTela.ldl_usuario.Caption = IIf(IsNull(cRec.Fields("EMBALAGEM")), " ", cRec.Fields("EMBALAGEM"))
        oTela.lbl_Sequencial.Caption = Trim(cRec.Fields("ID_ETIQUETA"))
        oTela.lbl_barras1.Caption = "*" & oTela.lbl_Sequencial.Caption & "*"
        oTela.lbl_barras2.Caption = "*" & oTela.lbl_Sequencial.Caption & "*"
        oTela.Visible = False
        
     End If

     If cRec.Fields("Tipo") = "Z" Then ' aqui ser� preenchido os dados para a outra etiqueta tipo expotacao da fiat italiana

        Set oTela = New frmExibicao12
        oTela.slbl_01 = IIf(IsNull(cRec.Fields("Cliente")), " ", Trim(cRec.Fields("Cliente")))
        oTela.slbl_02 = IIf(IsNull(cRec.Fields("Pto_Entrega")), " ", Trim(cRec.Fields("Pto_Entrega")))
        oTela.slbl_03 = " "
        oTela.slbl_04 = "MUSASHI DO BRASIL LTDA"
        oTela.slbl_07 = " "
        oTela.slbl_08 = " " ' FALTA VER COM MAURO
        oTela.slbl_09 = IIf(IsNull(cRec.Fields("Qtd_Caixa")), " ", Trim(cRec.Fields("Qtd_Caixa")))
        oTela.slbl_10 = IIf(IsNull(cRec.Fields("Descr_Peca")), " ", Trim(cRec.Fields("Descr_Peca")))
        oTela.slbl_11 = IIf(IsNull(cRec.Fields("Cod_Fornecedor")), " ", Trim(cRec.Fields("Cod_Fornecedor")))
        oTela.slbl_12 = IIf(IsNull(cRec.Fields("Cod_Embalagem")), " ", Trim(cRec.Fields("Cod_Embalagem")))
        oTela.slbl_13 = Format(Now(), "DD/MM/YYYY")
        oTela.slbl_14 = " "
        oTela.slbl_15 = " "
        oTela.slbl_16 = IIf(IsNull(cRec.Fields("Lote")), " ", Trim(cRec.Fields("Lote")))
        oTela.ldl_usuario.Caption = IIf(IsNull(cRec.Fields("EMBALAGEM")), " ", cRec.Fields("EMBALAGEM"))
        oTela.lbl_Sequencial.Caption = Trim(cRec.Fields("ID_ETIQUETA"))
        oTela.lbl_barras1.Caption = "*" & oTela.lbl_Sequencial.Caption & "*"
        oTela.lbl_barras2.Caption = "*" & oTela.lbl_Sequencial.Caption & "*"
        oTela.Visible = False
     End If

End If


'---------------------------------------------------------------------------------------------------
'Se tipo 3 opcao FORD
If cRec.Fields("Tipo") = "3" Then
        sdata_aux = Mid$(Trim(cRec.Fields("data_etiq")), 1, 2) & _
                    Pega_Mes(Val(Mid$(Trim(cRec.Fields("data_etiq")), 4, 2))) & _
                    Mid$(Trim(cRec.Fields("data_etiq")), 7, 4)
        frmExibicao3.lbl_Cliente.Caption = "MUSASHI DO BRASIL LTDA" 'Trim(crec.Fields("Cliente"))
        frmExibicao3.lbl_Cod_Fornecedor.Caption = "CFEOA"  'Left(crec.Fields("Cod_Fornecedor"), 16)
        frmExibicao3.lbl_Cod_Fornecedor_Barras.Caption = "CFEOA" 'Trim(crec.Fields("Cod_Fornecedor"))
        frmExibicao3.lbl_qtd.Caption = Left(Trim(cRec.Fields("Qtd_Caixa")), 16)
        frmExibicao3.lbl_qtd_barras.Caption = Left(Trim(cRec.Fields("Qtd_Caixa")), 16)
        frmExibicao3.lbl_peso.Caption = Left(Trim(cRec.Fields("Peso")), 16)
        frmExibicao3.lbl_container.Caption = "KLT 4314 CFEOA" 'Left(Trim(crec.Fields("Lote"), 16))
        frmExibicao3.lbl_lote.Caption = Trim(cRec.Fields("Lote"))
        frmExibicao3.lbl_data.Caption = sdata_aux
        frmExibicao3.lbl_Cod_cliente.Caption = Left(Trim(cRec.Fields("Cod_Cliente")), 16)
        frmExibicao3.lbl_cod_cliente_1.Caption = Left(Trim(cRec.Fields("Cod_Cliente")), 16)
        frmExibicao3.lbl_cod_cliente_Barras.Caption = Left(Trim(cRec.Fields("Cod_Cliente")), 16)
        frmExibicao3.lbl_Cod_Peca.Caption = Trim(cRec.Fields("Cod_Peca"))
        frmExibicao3.lbl_line_feed_loc2.Caption = Trim(cRec.Fields("Cod_Peca")) & " " & Trim(cRec.Fields("Lote"))

        frmExibicao3.lbl_descr_peca.Caption = Trim(cRec.Fields("Descr_Peca"))
        frmExibicao3.lbl_id_etiqueta.Caption = Left(Trim(cRec.Fields("id_etiqueta")), 16)
        frmExibicao3.lbl_id_etiqueta_barra.Caption = Left(Trim(cRec.Fields("id_etiqueta")), 16)
        sdata_aux = Mid$(Trim(cRec.Fields("data_etiq")), 9, 2) & _
                    Mid$(Trim(cRec.Fields("data_etiq")), 4, 2) & _
                    Mid$(Trim(cRec.Fields("data_etiq")), 1, 2)
        frmExibicao3.DataToEncodeText.Text = Format(cRec.Fields("id_etiqueta"), "00000") & " (P)" & _
                                             Format(Trim(cRec.Fields("Qtd_Caixa")), "0") & " (Q)" & _
                                             Trim(cRec.Fields("Cod_Fornecedor")) & " (V)" & _
                                             sdata_aux & " (D)" & _
                                             Format(Trim(cRec.Fields("Serial")), "000") & " (S)"
        
        frmExibicao3.lbl_to.Caption = "FORD TAUBATE"
        frmExibicao3.lbl_cust.Caption = "FI05D"
        frmExibicao3.lbl_doc_code.Caption = "R3"
    
        frmExibicao3.PDF1.DataToEncode = frmExibicao3.DataToEncodeText.Text
    
End If

'---------------------------------------------------------------------------------------------------
'Se tipo 4 opcao padr�o etiqueta grande
If cRec.Fields("Tipo") = "4" Or cRec.Fields("Tipo") = "8" Then
    'frmExibicao4.Show
    frmAvulsoPadraoPonteiro.Show
    frmAvulsoPadraoPonteiro.Left = Me.Width
    If cRec.Fields("Cliente") <> "" Then
        'frmExibicao4.lblCliente.Caption = cRec.Fields("Cliente")
        frmAvulsoPadraoPonteiro.lblCliente.Caption = cRec.Fields("Cliente")
    End If
    If cRec.Fields("Cod_Cliente") <> "" Then
        'frmExibicao4.lblCodCliente2.Caption = cRec.Fields("Cod_Cliente")
        frmAvulsoPadraoPonteiro.lblCodCliente2.Caption = cRec.Fields("Cod_Cliente")
    End If
    If cRec.Fields("Descr_Peca") <> "" Then
        'frmExibicao4.lblDescricao.Caption = cRec.Fields("Descr_Peca")
        frmAvulsoPadraoPonteiro.lblDescricao.Caption = cRec.Fields("Descr_Peca")
    End If
    If cRec.Fields("Lote") <> "" Then
        'frmExibicao4.lblLote2.Caption = cRec.Fields("Lote")
        frmAvulsoPadraoPonteiro.lblLote2.Caption = cRec.Fields("Lote")
    End If
    If cRec.Fields("Peso") <> "" Then
        'frmExibicao4.lblPeso2.Caption = cRec.Fields("Peso")
        frmAvulsoPadraoPonteiro.lblPeso2.Caption = Format(cRec.Fields("Peso"), "0.00")
    End If
    If cRec.Fields("Qtd_Caixa") <> "" Then
        'frmExibicao4.lblQtd2.Caption = cRec.Fields("Qtd_Caixa")
        frmAvulsoPadraoPonteiro.lblQtd2.Caption = cRec.Fields("Qtd_Caixa")
    End If
    If cRec.Fields("Cod_Peca") <> "" Then
        'frmExibicao4.lblPeca.Caption = cRec.Fields("Cod_Peca")
        frmAvulsoPadraoPonteiro.lblPeca.Caption = cRec.Fields("Cod_Peca")
    End If
    If Not IsNull(cRec.Fields("EMBALAGEM")) Then
        frmAvulsoPadraoPonteiro.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
     End If
     If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
       frmAvulsoPadraoPonteiro.lbl_Seq_Milhar.Caption = Mid$(Trim(Format(cRec.Fields("ID_ETIQUETA"), "0000000000")), _
                                                        Len(Trim(Format(cRec.Fields("ID_ETIQUETA"), "0000000000"))) - 5, _
                                                        Len(Trim(Format(cRec.Fields("ID_ETIQUETA"), "0000000000"))))
     
        frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption = Format(cRec.Fields("ID_ETIQUETA"), "0000000000")
        frmAvulsoPadraoPonteiro.lblCodigoBarrasA.Caption = "*" & frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption & "*"
        frmAvulsoPadraoPonteiro.lblCodigoBarrasB.Caption = "*" & frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption & "*"
     Else
        frmAvulsoPadraoPonteiro.lblCodigoBarras.Caption = ""
        frmAvulsoPadraoPonteiro.lblCodigoBarrasA.Caption = ""
        frmAvulsoPadraoPonteiro.lblCodigoBarrasB.Caption = ""
        frmAvulsoPadraoPonteiro.lbl_Seq_Milhar.Caption = ""
     End If
     
     frmAvulsoPadraoPonteiro.lbl_data.Caption = Format(cRec.Fields("data_etiq"), "dd/mm/yyyy")
     frmAvulsoPadraoPonteiro.LBL_MES.Caption = Pega_Mes(Mid$(Format(cRec.Fields("data_etiq"), "dd/mm/yyyy"), 4, 2))

     If (cRec.Fields("Tipo") = "4") Then
        frmAvulsoPadraoPonteiro.lblMsgProduto.Visible = False
        If (Val((cRec.Fields("id_cliente")) = 1) Or (Val(cRec.Fields("id_cliente")) = 2) Or (Val(cRec.Fields("id_cliente")) = 3)) Then
           frmAvulsoPadraoPonteiro.lblMsgProduto.Visible = True
        End If
     End If
     
     If (cRec.Fields("Tipo") = "8") Then
        frmAvulsoPadraoPonteiro.lblCodCliente1.Caption = "CODE:"
        frmAvulsoPadraoPonteiro.lblQtd1.Caption = "QTY.:"
        frmAvulsoPadraoPonteiro.lblPeso1.Caption = "WEIGHT:"
        frmAvulsoPadraoPonteiro.lblLote1.Caption = "LOT.:"
     End If
          
End If

'---------------------------------------------------------------------------------------------------
'Se tipo 4 opcao padr�o etiqueta NOVA MDA 26/04/2019
If cRec.Fields("Tipo") = "M" Then
    frmExibicao4MDA.Show
    frmExibicao4MDA.Left = Me.Width
    
    If cRec.Fields("Cliente") <> "" Then
        frmExibicao4MDA.lblCliente.Caption = cRec.Fields("Cliente")
    End If
    If cRec.Fields("Cod_Cliente") <> "" Then
        frmExibicao4MDA.lblCodCliente2.Caption = cRec.Fields("Cod_Cliente")
        frmExibicao4MDA.LBL_CODIGO_NOVO.Caption = "*" & Trim(cRec.Fields("Cod_Cliente")) & "*"
    End If
    If cRec.Fields("Descr_Peca") <> "" Then
        frmExibicao4MDA.lblDescricao.Caption = cRec.Fields("Descr_Peca")
    End If
    If cRec.Fields("Lote") <> "" Then
        frmExibicao4MDA.lblLote2.Caption = cRec.Fields("Lote")
        frmExibicao4MDA.LBL_LOTE_NOVO.Caption = "*" & Trim(cRec.Fields("Lote")) & "*"
    End If
    If cRec.Fields("Peso") <> "" Then
        frmExibicao4MDA.lblPeso2.Caption = Format(cRec.Fields("Peso"), "0.00")
    End If
    If cRec.Fields("Qtd_Caixa") <> "" Then
        frmExibicao4MDA.lblQtd2.Caption = cRec.Fields("Qtd_Caixa")
        frmExibicao4MDA.LBL_QTDE_NOVO.Caption = "*" & cRec.Fields("Qtd_Caixa") & "*"
    End If
    If cRec.Fields("Cod_Peca") <> "" Then
        frmExibicao4MDA.lblPeca.Caption = cRec.Fields("Cod_Peca")
    End If
    If Not IsNull(cRec.Fields("EMBALAGEM")) Then
        frmExibicao4MDA.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
     End If
     If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
        frmExibicao4MDA.lblCodigoBarras.Caption = Format(cRec.Fields("ID_ETIQUETA"), "0000000000")
        frmExibicao4MDA.lblCodigoBarrasA.Caption = "*" & frmExibicao4MDA.lblCodigoBarras.Caption & "*"
        frmExibicao4MDA.lblCodigoBarrasB.Caption = "*" & frmExibicao4MDA.lblCodigoBarras.Caption & "*"
     Else
        frmExibicao4MDA.lblCodigoBarras.Caption = ""
        frmExibicao4MDA.lblCodigoBarrasA.Caption = ""
        frmExibicao4MDA.lblCodigoBarrasB.Caption = ""
     End If
     
     frmExibicao4MDA.lbl_data.Caption = Format(cRec.Fields("data_etiq"), "dd/mm/yyyy")
     frmExibicao4MDA.LBL_MES.Caption = Pega_Mes(Mid$(Format(cRec.Fields("data_etiq"), "dd/mm/yyyy"), 4, 2))

End If
'---------------------------------------------------------------------------------------------------
'Se tipo 5 opcao GM etiqueta grande
'MODIFICADO EM 23-08-2007 (MARCOS PEDROSA) ACRESCENTAR CAMPOS LAY-OUT NOVO
If cRec.Fields("Tipo") = "5" Then
    frmExibicao5.tWithOpcoes = Me.Width
    frmExibicao5.Show
    If cRec.Fields("Cliente") <> "" Then
        frmExibicao5.lblTo.Caption = cRec.Fields("Cliente")
    End If
    Rem incluido (marcos pedrosa) em 23-08-2007
    If cRec.Fields("MOTIVO_ALTERACAO_OUTROS") <> "" Then
        frmExibicao5.lblMotivo_alteracao_outros.Caption = Trim(cRec("MOTIVO_ALTERACAO_OUTROS").Value)
    End If
    If cRec.Fields("Ind_Suplementar") <> "" Then
        frmExibicao5.lblPlant.Caption = cRec.Fields("Ind_Suplementar")
    End If
    Rem DEFINICAO ATE 23-08-2007 (MARCOS PEDROSA)
'''        If cRec.Fields("Cod_Util") <> "" Then
'''            frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & cRec.Fields("Cod_Util")
'''        End If
'''        If cRec.Fields("Desvio") <> "" Then
'''            frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & cRec.Fields("Desvio")
'''        End If
    Rem INCLUIDO MARCOS PEDROSA EM 23-08-2007
    If cRec.Fields("Cod_Embalagem_Pw") <> "" Then
        frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & cRec.Fields("Cod_Embalagem_Pw")
    End If
    Rem INCLUIDO MARCOS PEDROSA EM 23-08-2007
    If cRec.Fields("Desvio") <> "" Then
        frmExibicao5.lblPlant.Caption = frmExibicao5.lblPlant.Caption & "-" & cRec.Fields("Desvio")
    End If

    If cRec.Fields("Cod_Cliente") <> "" Then
        frmExibicao5.lblPartNumber.Caption = cRec.Fields("Cod_Cliente")
    End If
    If cRec.Fields("Descr_Peca") <> "" Then
        frmExibicao5.lblPeca.Caption = cRec.Fields("Descr_Peca")
    End If
    If cRec.Fields("Qtd_Caixa") <> 0 Then
        frmExibicao5.lblQtd.Caption = cRec.Fields("Qtd_Caixa")
    End If
    If cRec.Fields("Modelo") <> "" Then
        frmExibicao5.lblMaterial.Caption = cRec.Fields("Modelo")
    End If
    If cRec.Fields("Cod_Embalagem_Pw") <> "" Then
        frmExibicao5.lblReference.Caption = cRec.Fields("Cod_Embalagem_Pw")
    End If
'    If cRec.Fields("Pto_Entrega") <> "" Then
'        frmExibicao5.lblLicense.Caption = cRec.Fields("Pto_Entrega")
'        frmExibicao5.lblLicenseA.Caption = "*" & cRec.Fields("Pto_Entrega") & "*"
'        frmExibicao5.lblLicenseB.Caption = "*" & cRec.Fields("Pto_Entrega") & "*"
'    End If
    Rem comentado em 23-08-2007 (marcos pedrosa)
'''        If cRec.Fields("Cod_Embalagem") <> "" Then
'''            frmExibicao5.lblContainerType.Caption = cRec.Fields("Cod_Embalagem")
'''        End If
    Rem INCLUIDO EM 23-08-2007
    If cRec.Fields("compl_peca1") <> "" Then
        frmExibicao5.lblContainerType.Caption = cRec.Fields("compl_peca1")
    End If
    If cRec.Fields("Peso") <> 0 Then
        frmExibicao5.lblgrossWeight.Caption = Format(cRec.Fields("Peso"), "0.00")
    End If
    Rem comentado em 23-08-2007 (marcos pedrosa)
'''        If cRec.Fields("Embarque_Controlado") <> "" Then
'''            frmExibicao5.lblRoute.Caption = cRec.Fields("Embarque_Controlado")
'''        End If
    Rem INCLUIDO EM 23-08-2007
    If cRec.Fields("compl_peca2") <> "" Then
        frmExibicao5.lblRoute.Caption = cRec.Fields("compl_peca2")
    End If
    If cRec.Fields("Lote") <> "" Then
        frmExibicao5.lblLot.Caption = cRec.Fields("Lote")
    End If
    Rem comentado em 23-08-2007 (marcos pedrosa)
'''        If cRec.Fields("Dum") <> "" Then
'''            frmExibicao5.lblEng.Caption = cRec.Fields("Dum")
'''        End If
    Rem INCLUIDO EM 23-08-2007
    If cRec.Fields("data_lote") <> "" Then
        frmExibicao5.lblEng.Caption = Mid$(cRec.Fields("data_lote"), 1, 2) & _
                                      Pega_Mes(Val(Mid$(cRec.Fields("data_lote"), 4, 2))) & _
                                      Mid$(cRec.Fields("data_lote"), 7, 4)
    End If
    Rem INCLUIDO EM 23-08-2007
    If cRec.Fields("envio_lote") = "1" Then
        frmExibicao5.lblvalidade.Caption = "N"
    Else
        frmExibicao5.lblvalidade.Caption = ""
    End If
    
    If cRec.Fields("Data_expedicao") <> "" Then
        frmExibicao5.lblMfgDate.Caption = Mid$(cRec.Fields("Data_expedicao"), 1, 2) & _
                                          Pega_Mes(Val(Mid$(cRec.Fields("Data_expedicao"), 4, 2))) & _
                                          Mid$(cRec.Fields("Data_expedicao"), 7, 4)
    End If
    If cRec.Fields("Cod_Peca") <> "" Then
        frmExibicao5.lblCodMSB.Caption = cRec.Fields("Cod_Peca")
    End If
    If Not IsNull(cRec.Fields("EMBALAGEM")) Then
        frmExibicao5.lblEmbalagem2.Caption = cRec.Fields("EMBALAGEM")
    End If
    If Not IsNull(cRec.Fields("ID_ETIQUETA")) Then
        frmExibicao5.lblCodigoBarras.Caption = Format(cRec.Fields("ID_ETIQUETA"), "0000000000")
        frmExibicao5.lblCodigoBarrasA.Caption = "*" & frmExibicao5.lblCodigoBarras.Caption & "*"
        frmExibicao5.lblCodigoBarrasB.Caption = "*" & frmExibicao5.lblCodigoBarras.Caption & "*"
     Else
        frmExibicao5.lblCodigoBarras.Caption = ""
        frmExibicao5.lblCodigoBarrasA.Caption = ""
        frmExibicao5.lblCodigoBarrasB.Caption = ""
     End If
End If

'-----------------------------------------------------------------------------------------------------
'Se tipo 6 opcao Cliente MVM INTERNACIONAL MOTORES
If cRec.Fields("Tipo") = "6" Or cRec.Fields("Tipo") = "9" Then
    frmExibicao9.Show
    If cRec.Fields("Tipo") = "9" Then
       frmExibicao9.PICT_CRISTAL.Visible = True
    Else
       frmExibicao9.PICT_CRISTAL.Visible = False
    End If
    
    frmExibicao9.Left = Me.Width
    If cRec.Fields("Cliente") <> "" Then
        frmExibicao9.lblCliente.Caption = cRec.Fields("Cliente")
    End If
    If cRec.Fields("Cod_Cliente") <> "" Then
        frmExibicao9.lblCodBar_Cod_cliente.Caption = "*" & Trim(cRec.Fields("Cod_Cliente")) & "*"
        frmExibicao9.lblCodBar_Cod_cliente1.Caption = "*" & Trim(cRec.Fields("Cod_Cliente")) & "*"
        frmExibicao9.lbl_Cod_cliente.Caption = Trim(cRec.Fields("Cod_Cliente"))
    End If
    If cRec.Fields("Descr_Peca") <> "" Then
        frmExibicao9.lbl_desc_peca.Caption = cRec.Fields("Descr_Peca")
    End If
    If cRec.Fields("cod_fornecedor") <> "" Then
        frmExibicao9.lblCod_Fornecedor.Caption = cRec.Fields("cod_fornecedor")
    End If

    frmExibicao9.lbl_Fornecedor.Caption = "MUSASHI"
    If cRec.Fields("data_expedicao") <> "" Then
        frmExibicao9.lbl_data_expedicao.Caption = cRec.Fields("data_expedicao")
    End If
    If cRec.Fields("qtd_caixa") <> "" Then
        frmExibicao9.lblCodBarQtd_caixa.Caption = "*" & cRec.Fields("qtd_caixa") & "*"
        frmExibicao9.lblqtd_caixa.Caption = cRec.Fields("qtd_caixa")
    End If
    If cRec.Fields("Lote") <> "" Then
        frmExibicao9.lblCodBar_lote.Caption = "*" & cRec.Fields("Lote") & "*"
        frmExibicao9.lblLote.Caption = cRec.Fields("Lote")
    End If
    If cRec.Fields("Cod_Peca") <> "" Then
        frmExibicao9.lbl_id_etiqueta.Caption = Format(cRec.Fields("ID_ETIQUETA"), "0000000000")
        frmExibicao9.lblCodBar_Cod_Peca.Caption = "*" & cRec.Fields("id_etiqueta") & "*"
        frmExibicao9.lbl_Cod_Peca.Caption = cRec.Fields("cod_peca")
    End If
    
    If Not IsNull(cRec.Fields("EMBALAGEM")) Then
        frmExibicao9.lblCodFunc.Caption = cRec.Fields("EMBALAGEM")
    Else
        frmExibicao9.lblCodFunc.Caption = ""
    End If
    
    Rem preparar o "id" da etiqueta com sua formatacao conforme documento fornecedor(6)+dtEtiqueta(6-aammdd)+seq. do dia
    sSeqId = " "
    
    If Not IsNull(cRec.Fields("desvio_aviso_mod")) Then
'            frmExibicao9.lblDesvio_Aviso_Mod.Caption = cRec.Fields("desvio_aviso_mod")
        sSeqId = Mid$(cRec.Fields("desvio_aviso_mod"), 1, 6) ' fornecedor
        sSeqId = Trim(sSeqId) & Mid$(cRec.Fields("desvio_aviso_mod"), 13, 2) ' (aa)mmdd
        sSeqId = Trim(sSeqId) & Mid$(cRec.Fields("desvio_aviso_mod"), 9, 2)  ' aa(mm)dd
        sSeqId = Trim(sSeqId) & Mid$(cRec.Fields("desvio_aviso_mod"), 7, 2) ' aamm(dd)
        sSeqId = Trim(sSeqId) & Format(Mid$(cRec.Fields("desvio_aviso_mod"), 15, 11), "000000") ' sequencial
        
        frmExibicao9.lblCodBar_Desvio_Aviso_Mod.Caption = "*" & sSeqId & "*"
        frmExibicao9.lblCodBar_Desvio_Aviso_Mod1.Caption = "*" & sSeqId & "*"
        frmExibicao9.lblDesvio_Aviso_Mod.Caption = sSeqId
        
    End If
     
End If

'-----------------------------------------------------------------------------------------------------
'Se tipo 7 opcao Palete GM
If cRec.Fields("Tipo") = "7" Then
    If (IsNull(cRec.Fields("CODIGO_PRODUTO2").Value)) Then
        Set frmExibicao7Ref = frmExibicao7UmProduto
    ElseIf (Trim(cRec.Fields("CODIGO_PRODUTO2").Value) = "") Then
        Set frmExibicao7Ref = frmExibicao7UmProduto
    Else
        Set frmExibicao7Ref = frmExibicao7VariosProdutos
    End If
    
    
    frmExibicao7Ref.Show
    
    'Cliente
    If Not IsNull(cRec.Fields("Cliente").Value) Then
        frmExibicao7Ref.lblTo.Caption = cRec.Fields("Cliente").Value
    Else
        frmExibicao7Ref.lblTo.Caption = ""
    End If
    If Not IsNull(cRec.Fields("IND_SUPLEMentar").Value) Then
        frmExibicao7Ref.lblPlant.Caption = cRec.Fields("IND_SUPLEMentar").Value
    Else
        frmExibicao7Ref.lblPlant.Caption = ""
    End If
    Rem INCLUIDO MARCOS PEDROSA EM 28-08-2007
    If cRec.Fields("Cod_Embalagem_Pw") <> "" Then
        frmExibicao7Ref.lblPlant.Caption = frmExibicao7Ref.lblPlant.Caption & "-" & cRec.Fields("Cod_Embalagem_Pw")
    End If
    Rem INCLUIDO MARCOS PEDROSA EM 28-08-2007
    If cRec.Fields("Embarque_Controlado") <> "" Then
        frmExibicao7Ref.lblPlant.Caption = frmExibicao7Ref.lblPlant.Caption & "-" & cRec.Fields("Embarque_Controlado")
    End If
    Rem INCLUIDO MARCOS PEDROSA EM 28-08-2007
    If cRec.Fields("Modelo") <> "" Then
        frmExibicao7Ref.lblMaterial.Caption = cRec.Fields("Modelo")
    End If
    
    Rem INCLUIDO EM 28-08-2007
    If cRec.Fields("envio_lote") = "1" Then
        frmExibicao7Ref.lblvalidade.Caption = "N"
    Else
        frmExibicao7Ref.lblvalidade.Caption = ""
    End If
    
    'Pode ser gravado nulo no caso da etiqueta N: 7 (Palete GM)
    If Not IsNull(cRec.Fields("Pto_Entrega").Value) Then
        frmExibicao7Ref.lblLicense.Caption = cRec.Fields("Pto_Entrega").Value
    Else
        frmExibicao7Ref.lblLicense.Caption = ""
    End If
    frmExibicao7Ref.lblLicenseA.Caption = "*" & cRec.Fields("Pto_Entrega").Value & "*"
    frmExibicao7Ref.lblLicenseB.Caption = "*" & cRec.Fields("Pto_Entrega").Value & "*"
    frmExibicao7Ref.lblShipmentDate.Caption = UCase(Format(Date, "DD/MMM/YYYY"))
    frmExibicao7Ref.lblPeso.Caption = Format(cRec.Fields("Peso"), "0.00")
    
    frmExibicao7Ref.lblCodigoProduto1.Caption = cRec.Fields("CODIGO_PRODUTO1").Value
    If (IsNumeric(cRec.Fields("QTDE_CAIXA1").Value) And IsNumeric(cRec.Fields("PECAS_CAIXA1").Value)) Then
        frmExibicao7Ref.lblQtde1.Caption = cRec.Fields("QTDE_CAIXA1").Value & " X " & CStr(CInt(cRec.Fields("PECAS_CAIXA1").Value)) & " PC"
    Else
        frmExibicao7Ref.lblQtde1.Caption = ""
    End If
    If Not IsNull(cRec.Fields("COMPL_PECA1").Value) Then
        frmExibicao7Ref.lblComplPeca1.Caption = cRec.Fields("COMPL_PECA1").Value
    Else
        frmExibicao7Ref.lblComplPeca1.Caption = ""
    End If
    
    
    If frmExibicao7Ref.Name = "frmExibicao7UmProduto" Then
        If (IsNumeric(cRec.Fields("QTDE_CAIXA1").Value) And IsNumeric(cRec.Fields("PECAS_CAIXA1").Value)) Then
            If CInt(cRec.Fields("QTDE_CAIXA1").Value) <> 0 And CInt(cRec.Fields("PECAS_CAIXA1").Value) <> 0 Then
                frmExibicao7UmProduto.lblQtdeTot1.Caption = cRec.Fields("QTDE_CAIXA1").Value * cRec.Fields("PECAS_CAIXA1").Value
            End If
        End If
    ElseIf frmExibicao7Ref.Name = "frmExibicao7VariosProdutos" Then
        frmExibicao7Ref.lblCodigoProduto2.Caption = cRec.Fields("CODIGO_PRODUTO2").Value
        If (IsNumeric(cRec.Fields("QTDE_CAIXA2").Value) And IsNumeric(cRec.Fields("PECAS_CAIXA2").Value)) Then
            If CInt(cRec.Fields("QTDE_CAIXA2").Value) <> 0 And CInt(cRec.Fields("PECAS_CAIXA2").Value) <> 0 Then
                frmExibicao7Ref.lblQtde2.Caption = cRec.Fields("QTDE_CAIXA2").Value & " X " & CStr(CInt(cRec.Fields("PECAS_CAIXA2").Value)) & " PC"
            End If
        Else
            frmExibicao7Ref.lblQtde2.Caption = ""
        End If
        If Not IsNull(cRec.Fields("COMPL_PECA2").Value) Then
            frmExibicao7Ref.lblComplPeca2.Caption = cRec.Fields("COMPL_PECA2").Value
        Else
            frmExibicao7Ref.lblComplPeca2.Caption = ""
        End If
        
        If Not IsNull(cRec.Fields("CODIGO_PRODUTO3").Value) Then
            frmExibicao7Ref.lblCodigoProduto3.Caption = cRec.Fields("CODIGO_PRODUTO3").Value
        Else
            frmExibicao7Ref.lblCodigoProduto3.Caption = ""
        End If
        If (IsNumeric(cRec.Fields("QTDE_CAIXA3").Value) And IsNumeric(cRec.Fields("PECAS_CAIXA3").Value)) Then
            If CInt(cRec.Fields("QTDE_CAIXA3").Value) <> 0 And CInt(cRec.Fields("PECAS_CAIXA3").Value) <> 0 Then
                frmExibicao7Ref.lblQtde3.Caption = cRec.Fields("QTDE_CAIXA3").Value & " X " & CStr(CInt(cRec.Fields("PECAS_CAIXA3").Value)) & " PC"
            End If
        Else
            frmExibicao7Ref.lblQtde3.Caption = ""
        End If
        If Not IsNull(cRec.Fields("COMPL_PECA3").Value) Then
            frmExibicao7Ref.lblComplPeca3.Caption = cRec.Fields("COMPL_PECA3").Value
        Else
            frmExibicao7Ref.lblComplPeca3.Caption = ""
        End If
        
        If Not IsNull(cRec.Fields("CODIGO_PRODUTO4").Value) Then
            frmExibicao7Ref.lblCodigoProduto4.Caption = cRec.Fields("CODIGO_PRODUTO4").Value
        Else
            frmExibicao7Ref.lblCodigoProduto4.Caption = ""
        End If
        If (IsNumeric(cRec.Fields("QTDE_CAIXA4").Value) And IsNumeric(cRec.Fields("PECAS_CAIXA4").Value)) Then
            If CInt(cRec.Fields("QTDE_CAIXA4").Value) <> 0 And CInt(cRec.Fields("PECAS_CAIXA4").Value) <> 0 Then
                frmExibicao7Ref.lblQtde4.Caption = cRec.Fields("QTDE_CAIXA4").Value & " X " & CStr(CInt(cRec.Fields("PECAS_CAIXA4").Value)) & " PC"
            End If
        Else
            frmExibicao7Ref.lblQtde4.Caption = ""
        End If
        If Not IsNull(cRec.Fields("COMPL_PECA4").Value) Then
            frmExibicao7Ref.lblComplPeca4.Caption = cRec.Fields("COMPL_PECA4").Value
        Else
            frmExibicao7Ref.lblComplPeca4.Caption = ""
        End If
    End If
    
    qtdeConteiners = 0
    If (IsNumeric(cRec.Fields("QTDE_CAIXA1").Value)) Then
        qtdeConteiners = qtdeConteiners + CInt(cRec.Fields("QTDE_CAIXA1").Value)
    End If
    If (IsNumeric(cRec.Fields("QTDE_CAIXA2").Value)) Then
        qtdeConteiners = qtdeConteiners + CInt(cRec.Fields("QTDE_CAIXA2").Value)
    End If
    If (IsNumeric(cRec.Fields("QTDE_CAIXA3").Value)) Then
        qtdeConteiners = qtdeConteiners + CInt(cRec.Fields("QTDE_CAIXA3").Value)
    End If
    If (IsNumeric(cRec.Fields("QTDE_CAIXA4").Value)) Then
        qtdeConteiners = qtdeConteiners + CInt(cRec.Fields("QTDE_CAIXA4").Value)
    End If
    
    
    frmExibicao7Ref.lblQtdeContainers.Caption = CStr(qtdeConteiners)
    
End If


'-------------------------------------------------------------------
'Se tipo A - Identifica��o de altera��es do produto e processo
If cRec.Fields("Tipo") = "A" Then
    
    frmExibicao6.Show
    
    With frmExibicao6
        .lblDesenho.Caption = Trim(cRec("COD_CLIENTE").Value & "")
        .lblDesvio.Caption = Trim(cRec("DESVIO_AVISO_MOD").Value & "")
        .lblData1.Caption = Format(Date, "dd/mm/yyyy")
        .lblData2.Caption = .lblData1.Caption
        .lblNotaFiscal1.Caption = Trim(cRec("NUM_DOC_FISCAL").Value & "")
        .lblNotaFiscal2.Caption = .lblNotaFiscal1.Caption
        
        .lblOptDefinitiva.Visible = False
        .lblOptProvisoria.Visible = False
        .lblOptLoteUnico.Visible = False
        
        Select Case Trim(cRec("TIPO_ALTERACAO").Value)
        Case "1"
            .lblOptDefinitiva.Visible = True
        Case "2"
            .lblOptProvisoria.Visible = True
        Case "3"
            .lblOptLoteUnico.Visible = True
        End Select
        
        .lblOptDesvio.Visible = False
        .lblOptMaterialSelecionado.Visible = False
        .lblOptOutros.Visible = False
        .lblOptReparoRetrabaho.Visible = False
        .lblOptProdutoNovo.Visible = False
        Select Case Trim(cRec("MOTIVO_ALTERACAO").Value)
        Case "1"
            frmExibicao6.lblOptDesvio.Visible = True
        Case "2"
            frmExibicao6.lblOptMaterialSelecionado.Visible = True
        Case "3"
            frmExibicao6.lblOptOutros.Visible = True
        Case "4"
            frmExibicao6.lblOptReparoRetrabaho.Visible = True
        Case "5"
            frmExibicao6.lblOptProdutoNovo.Visible = True
        End Select
        
        .lblMotivoAlteracaoOutros.Caption = Trim(cRec("MOTIVO_ALTERACAO_OUTROS").Value & "")
        
        .lblOptPrimeiroEnvio.Visible = False
        .lblOptLoteIntermediario.Visible = False
        .lblOptUltimoLote.Visible = False
        Select Case Trim(cRec("ENVIO_LOTE").Value)
        Case "1"
            .lblOptPrimeiroEnvio.Visible = True
        Case "2"
            .lblOptLoteIntermediario.Visible = True
        Case "3"
            .lblOptUltimoLote.Visible = True
        End Select
        
        .lblNumAm.Caption = Trim(cRec("NUM_AM").Value & "")
        
    
    End With
    
    nQuantidade = cRec("QTD_ETIQ").Value
    
End If

Rem aqui marcos
    
'Se tipo y - Identifica��o de altera��es do produto e processo
If cRec("Tipo") = "Y" Then
    If objApplication.filial = adMusashiDaAmazonia Then
       frmExibicao11.Show
       frmExibicao11.Left = Me.Width
       With frmExibicao11
'        .lblDesenho.Caption = Trim(cRec("COD_CLIENTE").Value & "")
           
           .lbl_LPN_COD.Caption = cRec("Cod_Fornecedor") & Format(Now(), "DDMMYY") & Format(cRec("Sequencia_Dia"), "000000")
            Call CodeRefresh(.lbl_LPN_COD.Caption)
           .lbl_LPN_COD_BARRA.Caption = sCodigo128
           .lbl_LPN_COD_B.Caption = .lbl_LPN_COD.Caption
           .lbl_CODIGO_NUM.Caption = Trim(cRec("Cod_Cliente"))
           .lbl_SUPPLIER_COD.Caption = IIf(IsNull(cRec("Cod_Fornecedor")), " ", cRec("Cod_Fornecedor"))
           .lbl_USER_COD.Caption = "9219"
           .lbl_YAMAHA_COD_BARRAS.Caption = cRec("Cod_Cliente") & "-" & _
                                            .lbl_SUPPLIER_COD.Caption & "-" & _
                                            "9219"
            Call CodeRefresh(.lbl_YAMAHA_COD_BARRAS.Caption)
           .lbl_YAMAHA_BARRAS.Caption = sCodigo128
           
           .lbl_NOME_DESCRICAO.Caption = cRec("descr_peca")
'            If objApplication.filial = adMusashiDaAmazonia Then
              .lbl_FORNECEDOR_NOME.Caption = "MUSASHI DO BRASIL LTDA"
'            Else
'               .lbl_FORNECEDOR_NOME.Caption = ""
'            End If
           .lbl_QTDE_NUM.Caption = cRec("Qtd_Caixa")
           .lbl_NF_NUM.Caption = ""
           .LBL_MES.Caption = Pega_Mes(Mid$(Format(cRec("data_etiq"), "dd/mm/yyyy"), 4, 2))
           .lbl_ANO.Caption = Format(cRec("data_etiq"), "YY")
           .lbl_QTDE_BARRAS.Caption = cRec("Qtd_Caixa")
           .lbl_QTDE_NUM1.Caption = cRec("Qtd_Caixa")
           .lbl_COD_MUSASHI_BARRAS.Caption = "*" & Trim(cRec("ID_ETIQUETA")) & "*"
           .lbl_COD_MUSASHI_NUM.Caption = Trim(cRec("ID_ETIQUETA"))
       End With
    Else
       frmExibicao10.Show
       With frmExibicao10
           .lbl_LPN_COD.Caption = cRec("Cod_Fornecedor") & Format(Now(), "DDMMYY") & Format(cRec("Sequencia_Dia"), "000000")
            Call CodeRefresh(.lbl_LPN_COD.Caption)
           .lbl_LPN_COD_BARRA.Caption = sCodigo128
           .lbl_LPN_COD_B.Caption = .lbl_LPN_COD.Caption
           .lbl_CODIGO_NUM.Caption = Trim(cRec("Cod_Cliente"))
           .lbl_SUPPLIER_COD.Caption = IIf(IsNull(cRec("Cod_Fornecedor")), " ", cRec("Cod_Fornecedor"))
           .lbl_USER_COD.Caption = "9219"
           .lbl_YAMAHA_COD_BARRAS.Caption = cRec("Cod_Cliente") & "-" & _
                                            .lbl_SUPPLIER_COD.Caption & "-" & _
                                            "9219"
            Call CodeRefresh(.lbl_YAMAHA_COD_BARRAS.Caption)
           .lbl_YAMAHA_BARRAS.Caption = sCodigo128
           
           .lbl_NOME_DESCRICAO.Caption = cRec("descr_peca")
'            If objApplication.filial = adMusashiDaAmazonia Then
              .lbl_FORNECEDOR_NOME.Caption = "MUSASHI DO BRASIL LTDA"
'            Else
'               .lbl_FORNECEDOR_NOME.Caption = ""
'            End If
           .lbl_QTDE_NUM.Caption = cRec("Qtd_Caixa")
           .lbl_NF_NUM.Caption = ""
           .LBL_MES.Caption = Pega_Mes(Mid$(Format(cRec("data_etiq"), "dd/mm/yyyy"), 4, 2))
           .lbl_ANO.Caption = Format(cRec("data_etiq"), "YY")
           .lbl_QTDE_BARRAS.Caption = cRec("Qtd_Caixa")
           .lblCodMSB.Caption = cRec.Fields("Cod_Peca")
           Rem aqui marcos feitoi aumento da ultima letra do codigo da peca 31/01/2024
           .lblCodMSB_Letra = Mid$(Trim(cRec.Fields("Cod_Peca")), Len(Trim(cRec.Fields("Cod_Peca"))), 1)
           .lblLot.Caption = cRec.Fields("Lote")
           .lbl_QTDE_NUM1.Caption = cRec("Qtd_Caixa")
           .lbl_COD_MUSASHI_BARRAS.Caption = "*" & Trim(cRec("ID_ETIQUETA")) & "*"
           .lbl_COD_MUSASHI_NUM.Caption = Trim(cRec("ID_ETIQUETA"))
       End With
    End If
    
End If

bTelaImp = False
Me.cmd_Impressao.Enabled = True
Me.SetFocus
    
Exit Sub

Erro:

MsgBox Err.Description
Me.MousePointer = vbDefault
Me.cmd_Visualizar.SetFocus

    '-------------------------------------------------------------------
End Sub

Private Sub Imprime_Etiqueta_Fiat_FTP()

Dim CrystalReport1 As New CRAXDRT.Report
Dim Application As New CRAXDRT.Application
Dim nx As Double
Dim x As Printer
Dim rs As ADODB.Recordset
Dim cRec_Aux As ADODB.Recordset
Dim sData As String
Dim oTela As frmEscRelCristalReport

On Error GoTo Erro

Me.MousePointer = vbHourglass

Set cRec_Aux = New ADODB.Recordset

Set cRec_Aux = CCTempneMov_Etiq.Mov_Etiq_Consultar_FIAT_FTP(sBancoMusashi, _
                                                        cRec.Fields("id_etiqueta"))

If cRec_Aux.RecordCount = 0 Then
   MsgBox "Etiqueta n�o encontrada, anote a etiqueta e procure o respons�vel, etiq: " & dteEtiquetas.rsEtiquetas.Fields("id_etiqueta")
   Me.MousePointer = vbDefault
   Exit Sub
End If

nx = 0

On Error GoTo Erro
Set rs = New ADODB.Recordset

rs.Fields.Append "1_Num_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "2_Qtde_Emb", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "3_Classe_Func", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "4_Indicacao_Supl", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "5_Data_Fab_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "6_Cod_Emb", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "7_Vinculo", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "8_Lote_Sob_Desv", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "9_Qtde_Lote", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "10_Aplicacao", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "11_DUM", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "12_Embarque_Ctrl", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "13_Cod_Fornecedor", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "14_Num_Doc_Fis_BAM", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "15_Data", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "16_Pondo_Entrega", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "17_Denominacao", ADODB.DataTypeEnum.adChar, 80
rs.Fields.Append "18_Num_Desenho", ADODB.DataTypeEnum.adChar, 20
rs.Fields.Append "19_Ctrl_Interno", ADODB.DataTypeEnum.adChar, 150
rs.Fields.Append "20_Ctrl_Oper_Log", ADODB.DataTypeEnum.adChar, 150
rs.Fields.Append "21_Codigo_Numero", ADODB.DataTypeEnum.adChar, 50
rs.Fields.Append "22_codigo_barras", ADODB.DataTypeEnum.adChar, 50

rs.Open
nx = 0

cRec_Aux.MoveFirst
rs.AddNew
rs.Fields("1_Num_Lote").Value = IIf(IsNull(cRec_Aux.Fields("Num_Lote")), " ", Trim(cRec_Aux.Fields("Num_Lote")))
rs.Fields("2_Qtde_Emb").Value = IIf(IsNull(cRec_Aux.Fields("Qtde_Emb")), "0", Format(cRec_Aux.Fields("Qtde_Emb"), "00"))
rs.Fields("3_Classe_Func").Value = IIf(IsNull(cRec_Aux.Fields("Classe_Func")), " ", cRec_Aux.Fields("Classe_Func"))
rs.Fields("4_Indicacao_Supl").Value = IIf(IsNull(cRec_Aux.Fields("Indicacao_Supl")), " ", cRec_Aux.Fields("Indicacao_Supl"))
If Format(cRec_Aux.Fields("DUM"), "DD/MM/YYYY") = "01/01/1900" Then
   rs.Fields("5_Data_Fab_Lote").Value = " "
Else
   rs.Fields("5_Data_Fab_Lote").Value = IIf(IsNull(cRec_Aux.Fields("Data_Fab_Lote")), " ", Format(cRec_Aux.Fields("Data_Fab_Lote"), "DD/MM/YYYY"))
End If
rs.Fields("6_Cod_Emb").Value = IIf(IsNull(cRec_Aux.Fields("Cod_Emb")), " ", cRec_Aux.Fields("Cod_Emb"))
rs.Fields("7_Vinculo").Value = IIf(IsNull(cRec_Aux.Fields("Vinculo")), " ", cRec_Aux.Fields("Vinculo"))
rs.Fields("8_Lote_Sob_Desv").Value = IIf(IsNull(cRec_Aux.Fields("embalagem")), " ", cRec_Aux.Fields("embalagem"))
'rs.Fields("8_Lote_Sob_Desv").Value = IIf(IsNull(cRec_Aux.Fields("Lote_Sob_Desv")), " ", cRec_Aux.Fields("Lote_Sob_Desv"))
rs.Fields("9_Qtde_Lote").Value = IIf(IsNull(cRec_Aux.Fields("Qtde_Lote")), "0", cRec_Aux.Fields("Qtde_Lote"))
rs.Fields("10_Aplicacao").Value = IIf(IsNull(cRec_Aux.Fields("Aplicacao")), " ", Mid$(cRec_Aux.Fields("Aplicacao"), 1, 14))
If Format(cRec_Aux.Fields("DUM"), "DD/MM/YYYY") = "01/01/1900" Then
   rs.Fields("11_DUM").Value = " "
Else
   rs.Fields("11_DUM").Value = IIf(IsNull(cRec_Aux.Fields("DUM")), " ", Format(cRec_Aux.Fields("DUM"), "DD/MM/YYYY"))
End If
rs.Fields("12_Embarque_Ctrl").Value = IIf(IsNull(cRec_Aux.Fields("Embarque_Ctrl")), " ", cRec_Aux.Fields("Embarque_Ctrl"))
rs.Fields("13_Cod_Fornecedor").Value = "13093"
rs.Fields("14_Num_Doc_Fis_BAM").Value = IIf(IsNull(cRec_Aux.Fields("Num_Doc_Fis_BAM")), " ", cRec_Aux.Fields("Num_Doc_Fis_BAM"))
rs.Fields("15_Data").Value = IIf(IsNull(cRec_Aux.Fields("Data")), " ", cRec_Aux.Fields("Data"))
rs.Fields("16_Pondo_Entrega").Value = IIf(IsNull(cRec_Aux.Fields("Ponto_Entrega")), " ", Trim(cRec_Aux.Fields("Ponto_Entrega")))
rs.Fields("17_Denominacao").Value = IIf(IsNull(cRec_Aux.Fields("Denominacao")), " ", cRec_Aux.Fields("Denominacao"))
rs.Fields("18_Num_Desenho").Value = IIf(IsNull(cRec_Aux.Fields("Num_Desenho")), " ", Mid$(cRec_Aux.Fields("Num_Desenho"), 1, 20))
rs.Fields("19_Ctrl_Interno").Value = IIf(IsNull(cRec_Aux.Fields("Ctrl_Interno")), " ", cRec_Aux.Fields("Ctrl_Interno"))
rs.Fields("20_Ctrl_Oper_Log").Value = IIf(IsNull(cRec_Aux.Fields("Ctrl_Oper_Log")), " ", cRec_Aux.Fields("Ctrl_Oper_Log"))
rs.Fields("21_Codigo_Numero").Value = IIf(IsNull(cRec_Aux.Fields("Num_Desenho")), " ", cRec_Aux.Fields("Num_Desenho")) & Format(rs.Fields("2_Qtde_Emb").Value, "00000") & cRec_Aux.Fields("Cod_Emb") & "013093"
rs.Fields("22_codigo_barras").Value = IIf(IsNull(cRec_Aux.Fields("Num_Desenho")), " ", "*" & Mid$(cRec_Aux.Fields("Num_Desenho"), 1, 11) & Format(rs.Fields("2_Qtde_Emb").Value, "00000") & cRec_Aux.Fields("Cod_Emb") & "013093" & "*")
rs.Update

Set CrystalReport1 = Application.OpenReport(App.Path & "\rpt_Etiquetas_RelEtiqueta_FIAT.rpt")
CrystalReport1.Database.SetDataSource rs

rs.Clone

Rem mostrar a tela
'Set oTela = New frmEscRelCristalReport
'oTela.CRViewer1.ReportSource = CrystalReport1
'oTela.CRViewer1.ViewReport
'oTela.Show 0
Rem *************************

Rem nao mostrar a tela
CrystalReport1.PrintOutEx False
Rem *************************

Me.MousePointer = vbDefault

Exit Sub

Erro:
MsgBox Err.Description, , Me.Caption
Me.MousePointer = vbDefault

End Sub


