VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1905
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125.536
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_banco 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1560
      Width           =   2745
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   720
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   720
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   1890
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   720
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   510
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Usuário:"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   585
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Senha:"
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   510
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Usuario As String
Public LoginSucceeded As Boolean


Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
    End
End Sub
Private Sub cmdOk_Click()
     Me.MousePointer = vbHourglass
    'check for correct password
    
    dteValeTransporte.rsUsuariosVale.Close
    dteValeTransporte.rsUsuariosVale.Source = "Select * from UsuariosVale where login = '" & txtUserName.Text & "'"
    dteValeTransporte.rsUsuariosVale.Open
    Me.MousePointer = vbDefault
    
    If dteValeTransporte.rsUsuariosVale.RecordCount = 0 Then
        MsgBox "Usuário não cadastrado!", vbCritical + vbOKOnly, "Atenção!!!"
        txtUserName.SetFocus
        Exit Sub
    Else
        dteValeTransporte.rsUsuariosVale.MoveFirst
        If Trim(txtPassword.Text) = dteValeTransporte.rsUsuariosVale.Fields("Senha") Then
            LoginSucceeded = True
            sUsuario = dteValeTransporte.rsUsuariosVale.Fields("Login")
            Usuario = dteValeTransporte.rsUsuariosVale.Fields("Login")
            mdiValeTransporte.Show
            Me.Hide
        Else
            MsgBox "Senha incorreta!!", vbCritical + vbOKOnly, "Atenção!!!"
            txtPassword.SetFocus
            SendKeys "{Home}+{End}"
            Exit Sub
        End If
    End If
End Sub
Private Sub Form_Load()

Rem carregar variaveis do banco para abertura
 

Rem *************************  A T E N Ç Ã O *****************************************
Rem *************************  A T E N Ç Ã O *****************************************
Rem **********************************************************************************
Rem

' Me.txt_banco.Text = "Base teste"
' sBancoMusashi = "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sysdba;Initial Catalog=BkpMusashi;Data Source=msb-2"
' sBancoRM = "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=BkpRM;Data Source=msb-2"
 
 Me.txt_banco.Text = "Produção"
 sBancoMusashi = "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sysdba;Initial Catalog=Musashi;Data Source=msb-2"
 sBancoRM = "Provider=SQLOLEDB.1;Password=sa;Persist Security Info=True;User ID=sa;Initial Catalog=CorporeRM;Data Source=msb-2"

Rem quando mudar o nome do banco, alterar tambem no form, dteValeTransporte,
Rem na coneccção "convaletransporte" a propriedade "Connection Sourse" ,
Rem mudar o nome do banco tambem
Rem
Rem **********************************************************************************
Rem *************************  A T E N Ç Ã O *****************************************
Rem *************************  A T E N Ç Ã O *****************************************



    dteValeTransporte.rsUsuariosVale.Open
    dteValeTransporte.rsUsuarioNivel.Open
    Me.cmdOK.Default = False
End Sub

Private Sub txtPassword_GotFocus()
Me.cmdOK.Default = True
End Sub

Private Sub txtPassword_LostFocus()
Me.cmdOK.Default = False
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then Me.txtPassword.SetFocus
End Sub


