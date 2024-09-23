VERSION 5.00
Object = "{B907CF17-F019-41BF-A9D4-8B1BEC2FCB54}#1.0#0"; "IDAutomationPDF417.dll"
Begin VB.Form frmTelaFormacaoPDF417 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tela auxilio formação da etiqueta PDF417"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   10530
   Begin VB.Frame Frame3 
      Caption         =   "Dados para a etiqueta pdf417"
      Height          =   3105
      Left            =   270
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.TextBox LeftMarginCM 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1380
         MaxLength       =   14
         TabIndex        =   12
         Text            =   "0.2"
         Top             =   2520
         Width           =   600
      End
      Begin VB.TextBox TopMarginCM 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1380
         MaxLength       =   14
         TabIndex        =   11
         Text            =   "0.2"
         Top             =   2250
         Width           =   600
      End
      Begin VB.TextBox NarrowBarWidth 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1380
         MaxLength       =   14
         TabIndex        =   10
         Text            =   "0.03"
         Top             =   1710
         Width           =   600
      End
      Begin VB.TextBox W2NRatio 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1380
         MaxLength       =   14
         TabIndex        =   9
         Text            =   "3.0"
         Top             =   1980
         Width           =   600
      End
      Begin VB.TextBox PDFColumns 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   6150
         MaxLength       =   14
         TabIndex        =   8
         Text            =   "3"
         Top             =   1620
         Width           =   600
      End
      Begin VB.TextBox PDFErrorCorrectionLevel 
         BackColor       =   &H00FFFFFF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   6150
         MaxLength       =   14
         TabIndex        =   7
         Text            =   "3"
         Top             =   1890
         Width           =   600
      End
      Begin VB.TextBox DataToEncodeText 
         Height          =   690
         Left            =   60
         MaxLength       =   640
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   270
         Width           =   4650
      End
      Begin VB.TextBox ImageWidth 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   90
         MaxLength       =   14
         TabIndex        =   2
         Text            =   "2044"
         Top             =   1380
         Width           =   780
      End
      Begin VB.TextBox ImageHeight 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   90
         MaxLength       =   14
         TabIndex        =   1
         Text            =   "1310"
         Top             =   1110
         Width           =   780
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Narrow Bar Width"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2010
         TabIndex        =   18
         Top             =   1755
         Width           =   1320
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Left Margin"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2010
         TabIndex        =   17
         Top             =   2565
         Width           =   870
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Top Margin"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2010
         TabIndex        =   16
         Top             =   2295
         Width           =   1005
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "X to Y Ratio"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2010
         TabIndex        =   15
         Top             =   2025
         Width           =   960
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PDF Columns"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6825
         TabIndex        =   14
         Top             =   1665
         Width           =   1005
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Error Correction Level"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6825
         TabIndex        =   13
         Top             =   1935
         Width           =   1590
      End
      Begin PDF417LibCtl.PDF PDF1 
         Height          =   1170
         Left            =   4980
         TabIndex        =   6
         Top             =   240
         Width           =   2610
         _cx             =   4604
         _cy             =   2064
         BackColor       =   16777215
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         Picture         =   "frmTelaFormacaoPDF417.frx":0000
         DataToEncode    =   "IDAutomation.com Metafile Image Generator Example"
         Orientation     =   0
         XtoYRatio       =   3
         NarrowBarCM     =   0,03
         LeftMarginCM    =   0,2
         TopMarginCM     =   0,2
         Truncated       =   0
         PDFRows         =   0
         PDFColumns      =   5
         PDFErrorCorrectionLevel=   2
         PDFMode         =   0
         ApplyTilde      =   1
         FixedResolutionCM=   0
         MacroPDFEnable  =   0
         MacroPDFFileID  =   0
         MacroPDFSegmentIndex=   0
         MacroPDFLastSegment=   0
         WhiteBarIncrease=   0
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   945
         TabIndex        =   5
         Top             =   1425
         Width           =   600
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   945
         TabIndex        =   4
         Top             =   1155
         Width           =   600
      End
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   3360
      Picture         =   "frmTelaFormacaoPDF417.frx":3752
      Stretch         =   -1  'True
      Top             =   3450
      Width           =   4305
   End
End
Attribute VB_Name = "frmTelaFormacaoPDF417"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sNomeArquivo As String

'Variável para MDIapp
'Public cRec As ADODB.Recordset 'conterá os dados do registro corrente
Private Sub Form_Load()

'This is how we set the data to encode in the barcode
'    PDF1.DataToEncode = DataToEncodeText.Text
'    If AutoSize Then AutoSizeButton_Click
'    Picture1.Font = "39251" 'PODE SER ARIAL, COURIER, ETC...
'    Picture1.FontSize = 10
'    Picture1.Print PDF1.Picture
'    ImageHeight.Text = 40
'    ImageWidth.Text = 40
'    If AutoSize Then
'       ImageHeight.Text = 960
'    Else
'       ImageHeight.Text = PDF1.GetYPixels * 15
'       ImageWidth.Text = PDF1.GetXPixels * 16

PDF1.Height = ImageHeight.Text
PDF1.Width = ImageWidth.Text
PDF1.PDFColumns = PDFColumns.Text
PDF1.PDFErrorCorrectionLevel = PDFErrorCorrectionLevel.Text
PDF1.NarrowBarCM = NarrowBarWidth.Text
PDF1.TopMarginCM = TopMarginCM.Text
PDF1.LeftMarginCM = LeftMarginCM.Text

'    End If
    sNomeArquivo = App.Path & "\ImagemsPDF417\" & sNomeArquivo & ".wmf"
    'PDF1.SaveBarCode SaveAs.Text
    PDF1.SaveEnhWMF sNomeArquivo

End Sub

'Function CopyFileToField()
'    Dim ChunkSize As Long
'    Dim filenum As Integer
'    Dim Buffer()  As Byte
'    Dim BytesNeeded As Long
'    Dim Buffers As Long
'    Dim Remainder As Long
'    Dim FileName As String
'    FileName = App.Path & "\\FotoTemp.jpg"
'    Dim I As Long
'    If Len(FileName) = 0 Then
'        Exit Function
'    End If
'    If Dir(FileName) = "" Then
'        Err.Raise vbObjectError, , "File not found: """ & FileName & """"
'    End If
'    ChunkSize = 65536
'    filenum = FreeFile
'    Open FileName For Binary As #filenum
'    BytesNeeded = LOF(filenum)
'    Buffers = BytesNeeded '\\ ChunkSize
'    Remainder = BytesNeeded Mod ChunkSize
'    For I = 0 To Buffers - 1
'        ReDim Buffer(ChunkSize)
'        Get #filenum, , Buffer
'        rsclientes.Fields("foto").AppendChunk (Buffer)
'    Next
'    ReDim Buffer(Remainder)
'    Get #filenum, , Buffer
'    rsclientes.Fields("foto").AppendChunk (Buffer)
'    Close #filenum
'    Kill FileName
'End Function
'
'
'Function CopyFieldToFile(strFileName As String, Controle As PictureBox) As String
'    Dim filenum As Integer
'    Dim Buffer() As Byte
'    Dim BytesNeeded As Long
'    Dim Buffers As Long
'    Dim Remainder As Long
'    Dim Offset As Long
'    Dim r As Integer
'    Dim I As Long
'    Dim ChunkSize As Long
'
'    ChunkSize = 65536
'    BytesNeeded = rsclientes.Fields("foto").FieldSize
'    If BytesNeeded > 0 Then
'        Buffers = BytesNeeded '\\ ChunkSize
'        Remainder = BytesNeeded Mod ChunkSize
'        If Dir(strFileName) <> "" Then
'            Kill strFileName
'        End If
'        filenum = FreeFile
'        Open strFileName For Binary As #filenum
'        For I = 0 To Buffers - 1
'           ReDim Buffer(ChunkSize)
'           Buffer = rsclientes.Fields("foto").GetChunk(Offset, ChunkSize)
'           Put #filenum, , Buffer()
'           Offset = Offset + ChunkSize
'        Next
'        ReDim Buffer(Remainder)
'        Buffer = rsclientes.Fields("foto").GetChunk(Offset, Remainder)
'        Put #filenum, , Buffer()
'        Close #filenum
'    End If
'    CopyFieldToFile = strFileName
'    Controle.Picture = LoadPicture(strFileName)
'    Kill strFileName
'End Function

Private Sub ImageHeight_Change()
    If Not IsNumeric(ImageHeight.Text) Then ImageHeight.Text = "1000"
    PDF1.Height = ImageHeight.Text
End Sub

Private Sub ImageWidth_Change()
    If Not IsNumeric(ImageWidth.Text) Then ImageWidth.Text = "3000"
    PDF1.Width = ImageWidth.Text
End Sub
