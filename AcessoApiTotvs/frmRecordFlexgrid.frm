VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{8C566473-B65B-4850-8FAC-1CB63545425C}#2.0#0"; "PrjHttpService.ocx"
Begin VB.Form frmRecordFlexgrid 
   Caption         =   "Some JSON to RS to FlexGrid"
   ClientHeight    =   3510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7455
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   3510
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin PrjHttpService.HttpService HttpService 
      Left            =   6390
      Top             =   1110
      _ExtentX        =   609
      _ExtentY        =   556
   End
   Begin VB.CommandButton cmd_teste 
      Caption         =   "Lêr Registro"
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   2280
      Width           =   1065
   End
   Begin VB.TextBox txtOutput 
      Height          =   1035
      Left            =   210
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2160
      Width           =   5595
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1785
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   3149
      _Version        =   393216
      BackColorBkg    =   -2147483636
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmRecordFlexgrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' CodePage constant for UTF-8
Private Const CP_UTF8 = 65001
Private pvReadFile As String
''' Maps a character string to a UTF-16 (wide character) string
Private Declare Function MultiByteToWideChar Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cchMultiByte As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long _
    ) As Long


''' WinApi function that maps a UTF-16 (wide character) string to a new character string
Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long
    
' CodePage constant for UTF-8

Private Sub cmd_teste_Click()

HttpService.Host = "localhost"
HttpService.Port = 8080
HttpService.Path = "/poc"
HttpService.QueryStringParameter("cFields") = "2022-4-15"
'HttpService.QueryStringParameter("var2") = "pass"

txtOutput.Text = DecodeURI(HttpService.Get_)

End Sub

Private Sub Form_Load()
    With New ADODB.Stream
        .Type = adTypeText
        .Charset = "utf-8"
        .Open
        .LoadFromFile App.Path & "\json.txt"
        Me.txtOutput.Text = .ReadText(adReadAll)
        Set MSHFlexGrid1.DataSource = Transform.JsonToRecordset(Me.txtOutput.Text)

        .Close
    End With
    With MSHFlexGrid1
        .ColWidth(0) = 300
        .ColWidth(1) = 600
        .ColWidth(2) = 1200
        .ColWidth(3) = 1800
        .ColWidth(4) = 1800
    End With
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        MSHFlexGrid1.Move 0, 0, ScaleWidth, ScaleHeight
    End If
End Sub





'------------------------------------------------------------------
' NAME:         DecodeURI (PUBLIC)
' DESCRIPTION:  Decodes a UTF8 encoded string
' CALLED BY:    HandleNavigate
' PARAMETERS:
'  EncodedURL (I,REQ) - the UTF-8 encoded string to decode
' RETURNS:      the the decoded UTF-8 string
'------------------------------------------------------------------
Private Function DecodeURI(ByVal EncodedURI As String) As String
Dim bANSI() As Byte
Dim bUTF8() As Byte
Dim lIndex As Long
Dim lUTFIndex As Long

If Len(EncodedURI) = 0 Then
    Exit Function
End If

EncodedURI = Replace$(EncodedURI, "+", " ")         ' In case encoding isn't used.
bANSI = StrConv(EncodedURI, vbFromUnicode)          ' Convert from unicode text to ANSI values
ReDim bUTF8(UBound(bANSI))                          ' Declare dynamic array, get length
For lIndex = 0 To UBound(bANSI)                     ' from 0 to length of ANSI
    If bANSI(lIndex) = &H25 Then                    ' If we have ASCII 37, %, then
        bUTF8(lUTFIndex) = Val("&H" & Mid$(EncodedURI, lIndex + 2, 2)) ' convert hex to ANSI
        lIndex = lIndex + 2                         ' this character was encoded into two bytes
    Else
        bUTF8(lUTFIndex) = bANSI(lIndex)            ' otherwise don't need to do anything special
    End If
    lUTFIndex = lUTFIndex + 1                       ' advance utf index
Next
DecodeURI = FromUTF8(bUTF8, lUTFIndex)              ' convert to string
End Function
'------------------------------------------------------------------
' NAME:         FromUTF8 (Private)
' DESCRIPTION:  Use the system call MultiByteToWideChar to
'               get chars using more than one byte and return
'               return the whole string
' CALLED BY:    DecodeURI
' PARAMETERS:
'  UTF8 (I,REQ)   - the ID of the element to return
'  Length (I,REQ) - length of the string
' RETURNS:      the full raw data of this field
'------------------------------------------------------------------
Private Function FromUTF8(ByRef UTF8() As Byte, ByVal Length As Long) As String
    Dim lDataLength As Long

    lDataLength = MultiByteToWideChar(CP_UTF8, 0, VarPtr(UTF8(0)), Length, 0, 0)  ' Get the length of the data.
    FromUTF8 = String$(lDataLength, 0)                                         ' Create array big enough
    MultiByteToWideChar CP_UTF8, 0, VarPtr(UTF8(0)), _
                        Length, StrPtr(FromUTF8), lDataLength                  '
End Function

