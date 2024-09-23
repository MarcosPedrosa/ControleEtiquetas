VERSION 5.00
Object = "{D2F3FE5B-BCF5-4274-9EE3-8E402EAD4E12}#1.0#0"; "TBarCode6.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Open Report"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1095
      ScaleWidth      =   6495
      TabIndex        =   0
      Top             =   2280
      Width           =   6495
   End
   Begin TBARCODE6LibCtl.TBarCode6 tbc 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   6495
      _cx             =   11456
      _cy             =   1931
      BackColor       =   16777215
      BackStyle       =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      Text            =   "Adjust Properties"
      TextAlignment   =   0
      BarCode         =   20
      CDMethod        =   1
      CountCheckDigits=   0
      EscapeSequences =   0   'False
      Format          =   ""
      GuardWidth      =   0
      ModuleWidth     =   ""
      Orientation     =   0
      PrintDataText   =   -1  'True
      PrintTextAbove  =   0   'False
      Ratio           =   ""
      RatioHint       =   "1B:2B:3B:4B:1S:2S:3S:4S"
      RatioDefault    =   "1:2:3:4:1:2:3:4"
      TextColor       =   0
      LastError       =   ""
      LastErrorNo     =   0
      MustFit         =   0   'False
      TextDistance    =   -1
      NotchHeight     =   -1
      PDF417_Rows     =   -1
      PDF417_Columns  =   -1
      PDF417_ECLevel  =   -1
      PDF417_RowHeight=   -1
      MAXI_Mode       =   4
      MAXI_AppendIndex=   -1
      MAXI_AppendCount=   -1
      MAXI_Undercut   =   -1
      MAXI_Preamble   =   0   'False
      MAXI_PostalCode =   ""
      MAXI_CountryCode=   ""
      MAXI_ServiceClass=   ""
      MAXI_Date       =   "96"
      CountModules    =   222
      DrawStatus      =   0
      SuppressErrorMsg=   0   'False
      CountRows       =   1
      DM_Size         =   0
      DM_Rectangular  =   0   'False
      DM_Format       =   0
      DM_AppendIndex  =   -1
      DM_AppendCount  =   -1
      DM_AppendFileID =   -1
      PDF417_SegmentIndex=   -1
      PDF417_FileID   =   ""
      PDF417_LastSegment=   0   'False
      PDF417_FileName =   ""
      PDF417_SegmentCount=   -1
      PDF417_TimeStamp=   -1
      PDF417_Sender   =   ""
      PDF417_Addressee=   ""
      PDF417_FileSize =   -1
      PDF417_CheckSum =   -1
      PDF417_RatioRowCol=   ""
      MicroPDF_Mode   =   0
      MicroPDF_Version=   0
      QR_Version      =   0
      QR_Format       =   0
      QR_FmtAppIndicator=   ""
      QR_ECLevel      =   1
      QR_Mask         =   -1
      QR_AppendIndex  =   -1
      QR_AppendCount  =   -1
      QR_AppendParity =   -1
      QR_CompactKanji =   0   'False
      CBF_Rows        =   -1
      CBF_Columns     =   -1
      CBF_RowHeight   =   -1
      CBF_RowSeparatorHeight=   -1
      CBF_Format      =   0
      InterpretInputAs=   0
      OptResolution   =   0   'False
      DisplayText     =   ""
      BarWidthReduction=   -1
      Quality         =   0
      CompositeComponent=   0
      RSS_SegmPerRow  =   -1
      TrimSpaces      =   0
      BarSimmDefauls  =   0
      QuietZone       =   0
   End
   Begin VB.Label Label3 
      Caption         =   "(c) TEC-IT Datenverarbeitung GmbH (many thanks to Mr. Andreas Hochstöger, Fa. Kastner)"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   "TBarCode OCX and the PictureBox are used as a ""template"" to insert bar codes in CR."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Function BarcodeGenerate(Id As String, Width As Long, Height As Long) As PictureBox
Dim nSizing, cm

    tbc.Text = Id
    cm = tbc.CountModules
    Picture1.Cls
    Picture1.ScaleMode = vbPixels
    Form1.ScaleMode = vbPixels
    
    Width = ScaleX(Width, vbTwips, vbPixels)
    Height = ScaleY(Height, vbTwips, vbPixels)
    nSizing = Int(Width / cm)
    If nSizing < 1 Then nSizing = 1
    
    Picture1.Width = cm * nSizing
    Picture1.Height = Height
    DoEvents
    tbc.BCDraw Picture1.hDC, 0, 0, Picture1.Width, Picture1.Height
    Set BarcodeGenerate = Picture1
End Function

Private Sub Command1_Click()
    Form2.Show
End Sub

Private Sub tbc_Click()
    tbc.PropertyDialog
End Sub
