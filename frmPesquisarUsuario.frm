VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPesquisarUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pesquisar usu�rio"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   3735
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   8325
      Begin MSFlexGridLib.MSFlexGrid Grid1 
         Height          =   3375
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   7905
         _ExtentX        =   13944
         _ExtentY        =   5953
         _Version        =   393216
         Cols            =   4
         AllowBigSelection=   0   'False
         TextStyle       =   3
         TextStyleFixed  =   2
         HighLight       =   2
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdfechar 
      Caption         =   "&Fechar"
      Height          =   330
      Left            =   7440
      TabIndex        =   5
      Top             =   3900
      Width           =   1005
   End
   Begin VB.CommandButton cmdSelecionar 
      Caption         =   "&Selecionar"
      Height          =   330
      Left            =   6390
      TabIndex        =   4
      Top             =   3900
      Width           =   1005
   End
   Begin VB.TextBox txtlidos 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H8000000F&
      Height          =   315
      Left            =   1350
      MaxLength       =   6
      TabIndex        =   3
      Top             =   4290
      Width           =   1005
   End
   Begin VB.TextBox txtNome 
      Height          =   315
      Left            =   810
      TabIndex        =   2
      Top             =   3930
      Width           =   5265
   End
   Begin VB.OptionButton Opt_nome 
      Caption         =   "Nome"
      Height          =   255
      Left            =   4410
      TabIndex        =   1
      Top             =   4320
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.OptionButton Opt_secao 
      Caption         =   "Se��o"
      Height          =   255
      Left            =   5310
      TabIndex        =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Total registros : "
      Height          =   285
      Left            =   150
      TabIndex        =   10
      Top             =   4320
      Width           =   1185
   End
   Begin VB.Label Label2 
      Caption         =   "Nome :"
      Height          =   285
      Left            =   150
      TabIndex        =   9
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Classifica��o:"
      Height          =   195
      Left            =   3270
      TabIndex        =   8
      Top             =   4350
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmPesquisarUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ccodigo_pesquisa As String 'Codigo escolhido pelo usu�rio.
Public cnome As String 'Nome do escolhido pelo usu�rio.
Public rs As ADODB.Recordset
Public nTeclou_Enter As Integer

Private Sub cmdfechar_Click()
ccodigo_pesquisa = ""
cnome = ""
Me.Hide
End Sub

Private Sub cmdSelecionar_Click()
Dim nCod As Double

Me.Grid1.Col = 0
nCod = Me.Grid1.Text
ccodigo_pesquisa = nCod
Me.Grid1.Col = 1
cnome = Me.Grid1.Text
Me.Hide
End Sub

Private Sub Form_Activate()

Call Limpar_Grid
Call Carrega_januspesquisa

End Sub

Function Carrega_januspesquisa()

Dim nx As Double
Dim nLinhas As Double
Dim sClass As String

On Error GoTo Erro

Me.Grid1.Visible = False
Me.MousePointer = vbHourglass
Set rs = New ADODB.Recordset

If Me.Opt_nome.Value = True Then
   sClass = "0"
Else
   sClass = "1"
End If

Set rs = CCTempneUsuario.USUARIO_Consultar(sBancoMusashi)

Call Limpar_Grid
Call Carregar_Grid

txtNome.SetFocus
nTeclou_Enter = 0

Me.MousePointer = vbDefault
Me.Grid1.Visible = True

Exit Function

Erro:

Set rs = Nothing
Me.MousePointer = vbDefault
MsgBox Err.Description

ccodigo_pesquisa = ""
cnome = ""
Me.Hide
End Function

Private Sub Limpar_Grid()
Dim nx As Double
Dim nLinhas As Double
Dim nLinhas1 As Double

Grid1.Clear
nLinhas = Grid1.Rows

If Grid1.Rows > 2 Then
   For nx = Grid1.Rows To nLinhas1 - 2 Step -1
       If nx > 2 Then Grid1.RemoveItem (nx)
   Next
End If

Grid1.Row = 0
Grid1.Col = 0: Grid1.ColWidth(0) = 700:  Grid1.Text = "COD"
Grid1.Col = 1:  Grid1.ColWidth(1) = 4500: Grid1.Text = "NOME"
Grid1.Col = 2: Grid1.ColWidth(2) = 1500: Grid1.Text = "LOGIN"
Grid1.Col = 3: Grid1.ColWidth(3) = 2: Grid1.Text = ""
Grid1.Col = 3: Grid1.BackColor = &H80FFFF

Grid1.Row = 0

Grid1.HighLight = False

End Sub



Private Sub Grid1_DblClick()
Dim nCod As Double

Me.Grid1.Col = 0
nCod = Me.Grid1.Text
ccodigo_pesquisa = nCod
Me.Grid1.Col = 1
cnome = Me.Grid1.Text
Me.Hide
End Sub

Private Sub Opt_Nome_Click()
Call Carrega_januspesquisa
End Sub

Private Sub Opt_secao_Click()
Call Carrega_januspesquisa
End Sub

Private Sub txtNome_Change()
Dim nx As Integer
Dim sPesquisa As String
Dim sCritica As String

On Error GoTo ERROR

If Len(Trim(txtNome.Text)) = 0 Then Exit Sub

sPesquisa = "NOME LIKE "
sPesquisa = sPesquisa & "'%" & UCase(Trim(txtNome.Text)) & "*'"
sCritica = sPesquisa
rs.Filter = sCritica

If rs.RecordCount = 0 Then
   sPesquisa = "NOME > "
   sPesquisa = sPesquisa & Chr$(39) & UCase(Trim(txtNome.Text)) & Chr$(39)
   sPesquisa = sPesquisa & " or NOME="
   sPesquisa = sPesquisa & Chr$(39) & UCase(Trim(txtNome.Text)) & Chr$(39)
   sCritica = sPesquisa
   rs.Filter = sCritica
End If

Call Carregar_Grid

Exit Sub

ERROR:

MsgBox "Nada encontrado com essa digita��o no campo da pesquisa"
Me.MousePointer = vbDefault
End Sub

Public Function Carregar_Grid()
   Dim nx As Double
   Dim nLinhas As String
   
   Me.Grid1.Visible = False
   Me.MousePointer = vbHourglass

   Call Limpar_Grid
   
   Grid1.Row = 1
   rs.MoveFirst
   
   For nx = 1 To rs.RecordCount
       nLinhas = rs.Fields("codigo")
       Grid1.Col = 0: Grid1.Text = nLinhas
       Grid1.Col = 1: Grid1.Text = rs.Fields("nome")
       Grid1.Col = 2: Grid1.Text = rs.Fields("login")
       Grid1.Col = 3: Grid1.Text = rs.Fields("senha")
       rs.MoveNext
       If Not rs.EOF Then
          Grid1.Rows = Grid1.Rows + 1
          Grid1.Row = Grid1.Row + 1
       End If
   Next
   
Me.Grid1.Visible = True
Me.txtlidos.Text = rs.RecordCount
Me.MousePointer = vbDefault
   
Exit Function

ERROR:

'MsgBox "Algun caracter digitado indevido, Redigite!"

End Function
