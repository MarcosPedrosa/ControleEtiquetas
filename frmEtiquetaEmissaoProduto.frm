VERSION 5.00
Begin VB.Form frmEtiquetaEmissaoProduto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de etiquetas dos produtos"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   10965
   Begin VB.CommandButton cmd_Impressao 
      Enabled         =   0   'False
      Height          =   735
      Left            =   8880
      Picture         =   "frmEtiquetaEmissaoProduto.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Imprimir"
      Top             =   5670
      Width           =   975
   End
   Begin VB.CommandButton cmdfechar 
      Height          =   735
      Left            =   9930
      Picture         =   "frmEtiquetaEmissaoProduto.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Fecha a tela de etiqueta"
      Top             =   5670
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Etiquetas"
      Height          =   5505
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   10815
      Begin VB.CommandButton cmd_somar_todos 
         Caption         =   "Incluir &Todos"
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Top             =   2520
         Width           =   1245
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Classificação pela descrição"
         Height          =   285
         Left            =   4740
         TabIndex        =   10
         Top             =   2550
         Width           =   2475
      End
      Begin VB.OptionButton opt_cl_codigo 
         Caption         =   "Classificação por Codigo"
         Height          =   285
         Left            =   2310
         TabIndex        =   9
         Top             =   2550
         Value           =   -1  'True
         Width           =   2235
      End
      Begin VB.ListBox lst_Posto 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   2220
         Left            =   150
         TabIndex        =   6
         Top             =   240
         Width           =   10515
      End
      Begin VB.ListBox lst_postoBanco 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   2460
         Left            =   150
         TabIndex        =   5
         Top             =   2880
         Width           =   10515
      End
      Begin VB.CommandButton cmd_excluir_todos 
         Caption         =   "Excluir T&odos"
         Height          =   315
         Left            =   9390
         TabIndex        =   3
         Top             =   2520
         Width           =   1245
      End
      Begin VB.CommandButton cmd_somar 
         Caption         =   "&Somar"
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton cmd_excluir 
         Caption         =   "&Excluir"
         Height          =   315
         Left            =   10020
         TabIndex        =   4
         Top             =   2520
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmEtiquetaEmissaoProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável para MDIapp
Public Flag_ativo As Boolean 'Conterá true se o form ja foi ativado
Public cRec As ADODB.Recordset 'conterá os dados do registro corrente
Public cData_alteracao As String 'Data de alteracao vinda do registro a ser alterado ou excluido
Public Confirma_Mudanca As Boolean 'Servirá para confirmar as mudanças de alteracoes dos campos na tela
Public sManutencao As Boolean ' Servirá para confirmar alteração dos campos
Public cTipo_Movimentacao As Integer  'Se for 1=Inclusão;2=Alteração




Private Sub cmd_excluir_Click()
If Me.lst_postoBanco.ListCount = 0 Then
   Me.cmd_Impressao.Enabled = False
   Exit Sub
End If
Me.lst_postoBanco.RemoveItem (Me.lst_postoBanco.ListIndex)
Me.lst_postoBanco.Refresh
Me.cmd_Impressao.Enabled = True
End Sub

Private Sub cmd_excluir_todos_Click()
Me.lst_postoBanco.Clear
Me.lst_postoBanco.Refresh
Me.cmd_Impressao.Enabled = False
End Sub

Private Sub cmd_Impressao_Click()
Dim oTela As Object
Dim nx As Double

Dim x As Printer
               
For Each x In Printers
   If InStr(1, x.DeviceName, "Etiqueta Produto") > 0 Then
      Set Printer = x
      Exit For
   End If
Next

If x.DeviceName <> "Etiqueta Produto" Then
   MsgBox "impressora da Etiqueta não encontrada, chame o responsável. o nome da impressora será - 'Etiqueta Produto'"
   Exit Sub
End If

Printer.Orientation = 1
'Printer.Height = 1000
'Printer.Width = 500

For nx = 0 To Me.lst_postoBanco.ListCount - 1
    Printer.FontSize = 12
    Printer.FontName = "Times New Roman"
    Printer.Print Space(15) & Mid$(Me.lst_postoBanco.List(nx), 1, 9)
    Printer.FontSize = 28
    Printer.FontName = "3 of 9 Barcode"
    Printer.Print "*" & Mid$(Me.lst_postoBanco.List(nx), 1, 9) & "*"
'    Printer.FontSize = 12
'    Printer.FontName = "Times New Roman"
'    Printer.Print "1 "
'    Printer.Print "2 "
'    Printer.Print "3 "
'    Printer.Print "4 "
'    Printer.Print "5 "
    
    Printer.NewPage
'    printer.Height = 100
Next

Printer.Orientation = 2: Printer.EndDoc

End Sub

Private Sub cmd_somar_Click()
Call Somar_Posto
End Sub

Private Sub cmd_somar_todos_Click()
Call Somar_Todos_Posto
End Sub

Private Sub cmdfechar_Click()
Unload Me
End Sub



Private Sub Form_Activate()
If Flag_ativo = True Then
   Exit Sub
End If
Me.Top = 0
Me.Left = 0
Flag_ativo = True
Confirma_Mudanca = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
' para funcionar , tem que mudar o keyPreviwe=true
If KeyCode = 13 Then
   If Me.ActiveControl.TabIndex > 3 Then
      SendKeys "{TAB}"
   End If
ElseIf KeyCode = 27 Then
   Unload Me
End If

End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
Call Carrega_Etiquetas
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'If cmd_Impressao.Enabled = True Then
'    If 7 = MsgBox("Possiveis alterações foram realizadas sem salvar,Deseja abandonar?", 32 + 4) Then
'        'respondeu não
'        Cancel = True
'    End If
'End If

End Sub
Public Sub Carrega_Etiquetas()
Dim cRec_Etiq As ADODB.Recordset
Dim y As Integer
Dim sNomePosto As String * 45
Dim sCodigoPosto As String * 9
Dim sArea As String * 20

On Error GoTo Erro
Me.MousePointer = vbHourglass
Set cRec_Etiq = New ADODB.Recordset

lst_Posto.Clear

If Me.opt_cl_codigo.Value = True Then
   Set cRec_Etiq = CCTempPecaAvulso.Peca_Avulso_Consultar(sBancoAccess, "0") 'por codigo
Else
   Set cRec_Etiq = CCTempPecaAvulso.Peca_Avulso_Consultar(sBancoAccess, "1") 'pode descricao
End If
If cRec_Etiq.RecordCount > 0 Then
'    Set cCodPosto = New Collection
    For y = 1 To cRec_Etiq.RecordCount
        sNomePosto = cRec_Etiq!Descr_Peca
        sCodigoPosto = cRec_Etiq!Cod_Peca
        sArea = cRec_Etiq!Descr_Peca
        lst_Posto.AddItem sCodigoPosto & " - " & sNomePosto
        If y = cRec_Etiq.RecordCount Then
        Else
           cRec_Etiq.MoveNext
        End If
    Next
Else
    MsgBox "Não existem Produtos avulsos cadastrados, por favor cadastre-os!!"
End If


Set cRec_Etiq = Nothing
Me.MousePointer = vbDefault

Exit Sub

Erro:
   Set cRec_Etiq = Nothing
   MsgBox Err.Description, , Me.Caption
   Me.MousePointer = vbDefault

End Sub

Private Sub lst_Posto_Click()
Call Somar_Posto
End Sub
Private Function Somar_Posto()
Dim nx As Integer


'ver se ja existe o mesmo selecionado
For nx = 1 To Me.lst_postoBanco.ListCount
    If Mid$(Me.lst_Posto.List(Me.lst_Posto.ListIndex), 1, 9) = Mid$(Me.lst_postoBanco.List(nx - 1), 1, 9) Then
       MsgBox "Item ja selecionado. Marque outro posto!"
       nx = Me.lst_postoBanco.ListCount
       Exit Function
    End If
Next

Me.lst_postoBanco.AddItem Me.lst_Posto.List(Me.lst_Posto.ListIndex)
Me.cmd_Impressao.Enabled = True
End Function
Private Function Somar_Todos_Posto()
Dim nx As Integer

Me.lst_postoBanco.Clear

'ver se ja existe o mesmo selecionado
For nx = 1 To Me.lst_Posto.ListCount
    Me.lst_postoBanco.AddItem Me.lst_Posto.List(nx - 1)
Next


End Function

Private Sub lst_postoBanco_Click()
Me.lst_postoBanco.RemoveItem (Me.lst_postoBanco.ListIndex)
Me.lst_postoBanco.Refresh
If Me.lst_postoBanco.ListCount = 0 Then
   Me.cmd_Impressao.Enabled = False
End If

End Sub
Private Function CARREGA_GRUPO_POSTO()
Dim nx As Integer
Dim sNomePosto As String * 45
Dim sCodigoPosto As String * 9

'ver se ja existe o mesmo selecionado
Me.lst_postoBanco.Clear
For nx = 1 To Me.cRec.RecordCount
    sNomePosto = Me.cRec.Fields(3)
    sCodigoPosto = Me.cRec.Fields(2)
    Me.lst_postoBanco.AddItem sNomePosto & " - " & sCodigoPosto
    Me.cRec.MoveNext
Next
Set cRec = Nothing
Me.cmd_Impressao.Enabled = True

End Function

Private Sub opt_cl_codigo_Click()
Call Carrega_Etiquetas
End Sub

Private Sub Option1_Click()
Call Carrega_Etiquetas
End Sub
