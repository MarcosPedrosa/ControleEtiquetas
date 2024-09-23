VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Acceso via WS"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   10755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Crear JSON"
      Height          =   360
      Left            =   2085
      TabIndex        =   5
      Top             =   4050
      Width           =   1065
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   2130
      TabIndex        =   4
      Top             =   2835
      Width           =   990
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1575
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1890
      Width           =   2130
   End
   Begin VB.TextBox txtUsuario 
      Height          =   345
      Left            =   1575
      TabIndex        =   0
      Top             =   1080
      Width           =   2130
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Index           =   1
      Left            =   1545
      TabIndex        =   2
      Top             =   1650
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   795
      Width           =   540
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Ejemplo para  consumir ws por GET
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Autor del proyecto: Yvan Acosta (YAcosta)
' Este proyecto no hubiera sido posible sin la colaboracion
' de mis amigos Leandro Ascierto y Alberto Miñano
' Visite: leandroascierto.com
Dim CadenaJSON       As String


Private Sub btnAceptar_Click()
Dim p             As Object
Dim Texto         As String
Dim sInputJson    As String
Dim cab           As Integer

Set httpURL = New WinHttp.WinHttpRequest

usua = Trim(txtUsuario)
Pass = Trim(txtPassword)

Cadena = "http://tupagina.com/prueba/login.php?USUARIO=" & usua & "&PASSWORD=" & Pass  'Aloja tu archivo php en tu hosting y
                                                                                       'cambia esta direccion
httpURL.Open "GET", Cadena
httpURL.send
Texto = httpURL.responseText
If Texto = "[]" Then
   MsgBox ("No se obtuvo resultados")
   Exit Sub
End If

sInputJson = "{items:" & Texto & "}"

Set p = JSON.parse(sInputJson)

NOMBRE = p.Item("items").Item(1).Item("NOMBRE")

MsgBox ("Bienvenido " & NOMBRE)

End Sub

Private Sub Command1_Click()
Call CreandoJson("JUAN PEREZ", "JPEREZ")

Debug.Print CadenaJSON
End Sub

Public Sub CreandoJson(NOMBRE As String, USUARIO As String)
'Creando string
CadenaJSON = ""
Dim sBuffer As String

AddParamJSON CadenaJSON, "NOMBRE", NOMBRE
AddParamJSON CadenaJSON, "USUARIO", USUARIO, True
'enviando al ws
End Sub
'***************************************************************************************************


Private Sub cmdChangeDirectory_Click()
   ' Cambia al directorio txtRemotePath.
   Inet1.Execute txtURL.Text, "CD " & _
   txtRemotePath.Text
End Sub
  
Private Sub cmdDELETE_Click()
   ' Elimina el directorio indicado en txtRemotePath.
   Inet1.Execute txtURL.Text, "DELETE " & _
   txtRemotePath.Text
End Sub
  
Private Sub cmdDIR_Click()
   Inet1.Execute txtURL.Text, "DIR BuscaEsto.txt"
End Sub
  
Private Sub cmdGET_Click()
   Inet1.Execute txtURL.Text, _
   "GET TomaEsto.txt C:\MisDocumentos\TengoEsto.txt"
End Sub
  
Private Sub cmdSEND_Click()
   Inet1.Execute txtURL.Text, _
   "SEND C:\MisDocumentos\Enviar.txt DocsEnviados\Enviado.txt"
End Sub
  
Private Sub Inet1_StateChanged(ByVal State As Integer)
   ' Obtiene la respuesta del servidor con el método
   ' GetChunk cuando State = 12.
  
   Dim vtData As Variant ' Variable de datos.
   Select Case State
   ' ... Otros casos no mostrados.
   Case icError ' 11
      ' En caso de error, devuelve ResponseCode
      ' y ResponseInfo.
      vtData = Inet1.ResponseCode & ":" & _
      Inet1.ResponseInfo
   Case icResponseCompleted ' 12
      Dim vtData As Variant
      Dim strData As String
      Dim bDone As Boolean: bDone = False
  
      ' Obtiene el primer bloque.
      vtData = Inet1.GetChunk(1024, icString)
      DoEvents
  
      Do While Not bDone
         strData = strData & vtData
         ' Obtiene el siguiente bloque.
         vtData = Inet1.GetChunk(1024, icString)
         DoEvents
  
         If Len(vtData) = 0 Then
            bDone = True
         End If
      Loop
      txtData.Text = strData
   End Select
     
End Sub

Private Sub Form_Load()
Dim xhr, method, url, contents, formatcontent, doc

Set xhr = CreateObject("MSXML2.XMLHTTP")

method = "GET" 'Escolhe o método HTTP de envio
url = "https://ws.printwayy.com/api/Printer?api_token=1F61D333-CCA5-423A-A764-F8577119A9FE&company_token=&serialNumbers=AK18054352&initialDate=&endDate=" 'url da API
contents = "" 'conteudo
formatcontent = "application/xml" 'Se a API usar outro formato basta alterar aqui

xhr.Open method, url, False

'Necessário pra sua API retornar XML ao invés de JSON
xhr.setRequestHeader "Accept", "application/xml"

If method = "POST" Or method = "PUT" Then
    xhr.setRequestHeader "Content-Type", formatcontent
    xhr.setRequestHeader "Content-Length", Len(contents)
    xhr.send contents
Else
    xhr.send
End If

If xhr.Status < 200 Or xhr.Status >= 300 Then
    'Algo falhou, as vezes pode haver uma descrição em `xhr.responseText` ou pode retornar vazio, o `xhr.status` indica o tipo de erro
    MsgBox "Erro HTTP:" & xhr.Status & " - Detalhes: " & xhr.responseText
Else
    'Faz o parse da String para XML
    Set doc = CreateObject("MSXML2.DOMDocument")
    doc.loadXML (xhr.responseText)

    'Seleciona com XPATH
    Set nodes = doc.selectNodes("//IPAddress")

    MsgBox "Elementos encontrados para IPAddress: " & nodes.Length

    For Each node In nodes
        MsgBox "Endereço IP: " & node.Text
    Next
End If
End Sub
