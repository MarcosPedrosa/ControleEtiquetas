VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   8415
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_Carrega_RecordSet 
      Caption         =   "Carga Recordset"
      Height          =   405
      Left            =   3210
      TabIndex        =   3
      Top             =   330
      Width           =   1005
   End
   Begin VB.CommandButton converter 
      Caption         =   "converter"
      Height          =   675
      Left            =   6720
      TabIndex        =   2
      Top             =   1080
      Width           =   945
   End
   Begin AcessoCloud.HttpService HttpServiceAcesso 
      Left            =   90
      Top             =   420
      _extentx        =   926
      _extenty        =   714
   End
   Begin VB.CommandButton cmd_teste 
      Caption         =   "Lêr Registro"
      Height          =   495
      Left            =   5310
      TabIndex        =   1
      Top             =   2910
      Width           =   1065
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   900
      TabIndex        =   0
      Top             =   900
      Width           =   5595
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1875
      Left            =   1590
      TabIndex        =   4
      Top             =   3540
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   3307
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Private Const CP_UTF8 = 65001

Private Sub cmd_Carrega_RecordSet_Click()
With New ADODB.Stream
    .Type = adTypeText
    .Charset = "utf-8"
    .Open
    .LoadFromFile App.Path & "\json.txt"
    Set MSHFlexGrid1.DataSource = Transform.JsonToRecordset(.ReadText(adReadAll))
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
Private Sub cmd_teste_Click()

HttpServiceAcesso.Host = "localhost"
HttpServiceAcesso.Port = 8080
HttpServiceAcesso.Path = "/poc"
HttpServiceAcesso.QueryStringParameter("cFields") = "2022-4-15"
'HttpService.QueryStringParameter("var2") = "pass"

txtOutput.Text = DecodeURI(HttpServiceAcesso.Get_)

End Sub


''' Return length of byte array or zero if uninitialized
Private Function BytesLength(abBytes() As Byte) As Long
    ' Trap error if array is uninitialized
    On Error Resume Next
    BytesLength = UBound(abBytes) - LBound(abBytes) + 1
End Function

''' Return VBA "Unicode" string from byte array encoded in UTF-8
Public Function Utf8BytesToString(abUtf8Array() As Byte) As String
    Dim nBytes As Long
    Dim nChars As Long
    Dim strOut As String
    Utf8BytesToString = ""
    ' Catch uninitialized input array
    nBytes = BytesLength(abUtf8Array)
    If nBytes <= 0 Then Exit Function
    ' Get number of characters in output string
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, 0&, 0&)
    ' Dimension output buffer to receive string
    strOut = String(nChars, 0)
    nChars = MultiByteToWideChar(CP_UTF8, 0&, VarPtr(abUtf8Array(0)), nBytes, StrPtr(strOut), nChars)
    Utf8BytesToString = Left$(strOut, nChars)
End Function
Public Function Utf8BytesFromString(strInput As String) As Byte()
    Dim nBytes As Long
    Dim abBuffer() As Byte
    ' Catch empty or null input string
    Utf8BytesFromString = vbNullString
    If Len(strInput) < 1 Then Exit Function
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, 0&, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
End Function

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


Public Function UTF8ENCODE(ByVal sStr As String) As String
Dim L As Integer
Dim lChar As String
Dim sUtf8 As String

    For L = 1 To Len(sStr)

        lChar = AscW(Mid(sStr, L, 1))

        If lChar < 128 Then
            sUtf8 = sUtf8 + Mid(sStr, L, 1)
        ElseIf ((lChar > 127) And (lChar < 2048)) Then

            sUtf8 = sUtf8 + Chr(((lChar \ 64) Or 192))
            sUtf8 = sUtf8 + Chr(((lChar And 63) Or 128))

        Else

            sUtf8 = sUtf8 + Chr(((lChar \ 144) Or 234))
            sUtf8 = sUtf8 + Chr((((lChar \ 64) And 63) Or 128))
            sUtf8 = sUtf8 + Chr(((lChar And 63) Or 128))

        End If
    Next L

    UTF8ENCODE = sUtf8

End Function


Private Sub converter_Click()
Dim FSO As New FileSystemObject
Dim JsonTS As TextStream
Dim JsonText As String
Dim Parsed As Dictionary
Dim vVariant As Variant
Dim elemento
Dim itemArray

'' Read .json file
'Set JsonTS = FSO.OpenTextFile("example.json", ForReading)
'JsonText = JsonTS.ReadAll
'JsonTS.Close
'
'' Parse json to Dictionary
'' "values" is parsed as Collection
'' each item in "values" is parsed as Dictionary

Set Parsed = New Dictionary


With Parsed
     .CompareMode = TextCompare
End With

Set Parsed = JsonConverter.ParseJson(txtOutput.Text)

vVariant = convertJsonToVariantArray(txtOutput.Text)


For Each elemento In Parsed.Keys
    MsgBox "Item - " & elemento
Next

MsgBox vVariant(1)(1)

' Prepare and write values to sheet
Dim Values As Object
'ReDim Values(Parsed("NOME").Count, 3)

Dim Value As Dictionary
Dim I As Long

'i = 0
'For Each Value In Parsed("values")
'  Values(i, 0) = Value("a")
'  Values(i, 1) = Value("b")
'  Values(i, 2) = Value("c")
'  i = i + 1
'Next Value

'Sheets("example").Range(Cells(1, 1), Cells(Parsed("values").Count, 3)) = Values
End Sub





Private Function convertJsonToVariantArray(ByVal JsonString As String) As Variant()
Dim cleanedUpArray() As Variant
Dim brokenUpRows() As String

'Remove the first and last square bracket in the string
JsonString = Right$(JsonString, Len(JsonString) - 2)
JsonString = Left$(JsonString, Len(JsonString) - 2)
'Break up the string in an array
brokenUpRows = Split(JsonString, "], [")

Dim counter As Integer
counter = 0
Dim counter2 As Integer
Dim brokenUpCols As Variant

ReDim linkArray(UBound(brokenUpRows)) As String

For counter = 0 To UBound(brokenUpRows)
    brokenUpCols = Split(brokenUpRows(counter), ",")
    If counter = 0 Then
       ReDim cleanedUpArray(UBound(brokenUpRows), UBound(brokenUpCols)) As Variant
    End If
    For counter2 = 0 To UBound(brokenUpCols)
        cleanedUpArray(counter, counter2) = brokenUpCols(counter2)
    Next
Next
convertJsonToVariantArray = cleanedUpArray
End Function




