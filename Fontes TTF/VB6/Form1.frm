VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check4 
      Alignment       =   1  'Right Justify
      Caption         =   "CompactPDF"
      Height          =   255
      Left            =   480
      TabIndex        =   31
      Top             =   2340
      Width           =   1335
   End
   Begin VB.ComboBox Combo7 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1620
      List            =   "Form1.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   1860
      Width           =   1575
   End
   Begin VB.ComboBox Combo6 
      Height          =   315
      ItemData        =   "Form1.frx":0047
      Left            =   1620
      List            =   "Form1.frx":0069
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":008E
      Left            =   1620
      List            =   "Form1.frx":00B0
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   1020
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "SaveToImageFile"
      Height          =   375
      Left            =   4260
      TabIndex        =   23
      Top             =   2940
      Width           =   1515
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Print1"
      Height          =   375
      Left            =   4260
      TabIndex        =   22
      Top             =   4620
      Width           =   1515
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Print"
      Height          =   375
      Left            =   4260
      TabIndex        =   21
      Top             =   4080
      Width           =   1515
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Draw to the Form"
      Height          =   375
      Left            =   4260
      TabIndex        =   20
      Top             =   3540
      Width           =   1515
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CopyToClipboard"
      Height          =   375
      Left            =   4260
      TabIndex        =   19
      Top             =   5160
      Width           =   1515
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ShowProperties"
      Height          =   375
      Left            =   6720
      TabIndex        =   18
      Top             =   5160
      Width           =   1515
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   375
      Left            =   6720
      TabIndex        =   17
      Top             =   4620
      Width           =   1515
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   5880
      TabIndex        =   16
      Text            =   "c:\PDF417.gif"
      Top             =   2940
      Width           =   2475
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "Form1.frx":00D5
      Left            =   1620
      List            =   "Form1.frx":00E2
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   4860
      Width           =   1575
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      ItemData        =   "Form1.frx":00FB
      Left            =   1620
      List            =   "Form1.frx":0108
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4500
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "QuietZone"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   255
      Left            =   1620
      TabIndex        =   8
      Top             =   3840
      Width           =   435
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   255
      Left            =   1620
      TabIndex        =   7
      Top             =   3480
      Width           =   435
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Transparent"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Form1.frx":0121
      Left            =   1620
      List            =   "Form1.frx":0131
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2700
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Form1.frx":0146
      Left            =   1620
      List            =   "Form1.frx":0156
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1620
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   180
      Width           =   2535
   End
   Begin VB.PictureBox PDF417Ctrl1 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   2295
      Left            =   4320
      ScaleHeight     =   2235
      ScaleWidth      =   3915
      TabIndex        =   24
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "DataRows"
      Height          =   195
      Left            =   180
      TabIndex        =   29
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "DataColumns"
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   1500
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "ErrorCorrectionLevel"
      Height          =   195
      Left            =   -120
      TabIndex        =   25
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "AlignV"
      Height          =   195
      Left            =   180
      TabIndex        =   15
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "AlidnH"
      Height          =   195
      Left            =   180
      TabIndex        =   14
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "ForeColor"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   3900
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "BackColor"
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   3540
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Orientation"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "CompactionMode"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   660
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "DataToEncode"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    PDF417Ctrl1.Transparent = Check1.Value
End Sub

Private Sub Check2_Click()
    PDF417Ctrl1.QuietZone = Check2.Value
End Sub

Private Sub Check4_Click()
    PDF417Ctrl1.CompactPDF = Check4.Value
End Sub

Private Sub Combo1_Click()
    PDF417Ctrl1.ErrorCorrectionLevel = Combo1.ListIndex
End Sub

Private Sub Combo2_Click()
    PDF417Ctrl1.CompactionMode = Combo2.ListIndex
End Sub

Private Sub Combo3_Click()
    PDF417Ctrl1.Orientation = Combo3.ListIndex * 90
End Sub

Private Sub Combo4_Click()
    PDF417Ctrl1.AlignH = Combo4.ListIndex
End Sub

Private Sub Combo5_Click()
    PDF417Ctrl1.AlignV = Combo5.ListIndex
End Sub

Private Sub Combo6_Click()
    PDF417Ctrl1.DataColumns = Combo6.ListIndex
End Sub

Private Sub Combo7_Click()
    PDF417Ctrl1.DataRows = Combo7.ListIndex
End Sub

Private Sub Command1_Click()
    PDF417Ctrl1.BackColor = &HFF0000
End Sub

Private Sub Command2_Click()
    PDF417Ctrl1.ForeColor = &HFF00FF
End Sub

Private Sub Command3_Click()
    PDF417Ctrl1.AboutBox
End Sub

Private Sub Command4_Click()
    PDF417Ctrl1.ShowProperties
End Sub

Private Sub Command5_Click()
    Call PDF417Ctrl1.CopyToClipboard(260, 90)
End Sub

Private Sub Command6_Click()
    Call PDF417Ctrl1.DrawPDF417ToSize(0, 0, 200, 200, dmPixels, Form1.hDC)
End Sub

Private Sub Command7_Click()
    Printer.CurrentX = 2048
    Printer.Print "BarcodeTools.com, VB Example"

    'PDF417-ActiveX will use the Visual Basic Printer object
    Call PDF417Ctrl1.SetPrinterHDC(Printer.hDC)

    'print the barcodes
    Call PDF417Ctrl1.DrawPDF417ToSize(10, 10, 30, 30, dmMM)
    Call PDF417Ctrl1.DrawPDF417ToSize(70, 10, 30, 30, dmMM)

    Printer.EndDoc
End Sub

Private Sub Command8_Click()
    'Create a new instance of the PDF417-ActiveX
    Dim oPDF417 As Object
    Set oPDF417 = CreateObject("PDF417ActiveX.PDF417Ctrl.1")

    'set the PDF417-ActiveX properties
    oPDF417.DataToEncode = "Hello world"

    'open a current printer
    Call oPDF417.BeginPrint("")
    
    'print the PDF417
    Call oPDF417.DrawPDF417ToSize(10, 10, 30, 30, dmMM)

    Call oPDF417.EndPrint

    Set oPDF417 = Nothing
End Sub

Private Sub Command9_Click()
    Call PDF417Ctrl1.SaveToImageFile(270, 90, Text2)
End Sub

Private Sub Form_Load()
    Text1 = PDF417Ctrl1.DataToEncode
    Combo2.ListIndex = PDF417Ctrl1.CompactionMode
    Combo1.ListIndex = PDF417Ctrl1.ErrorCorrectionLevel
    Combo3.ListIndex = PDF417Ctrl1.Orientation / 90
    Combo6.ListIndex = PDF417Ctrl1.DataColumns
    Combo7.ListIndex = PDF417Ctrl1.DataRows
    If (PDF417Ctrl1.Transparent) Then Check1.Value = 1 Else Check1.Value = 0
    If (PDF417Ctrl1.QuietZone) Then Check2.Value = 1 Else Check2.Value = 0
    If (PDF417Ctrl1.CompactPDF) Then Check4.Value = 1 Else Check4.Value = 0
    Combo4.ListIndex = PDF417Ctrl1.AlignH
    Combo5.ListIndex = PDF417Ctrl1.AlignV
End Sub

Private Sub Text1_Change()
    PDF417Ctrl1.DataToEncode = Text1
End Sub
