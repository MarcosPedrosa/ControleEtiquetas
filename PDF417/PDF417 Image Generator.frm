VERSION 5.00
Object = "{B907CF17-F019-41BF-A9D4-8B1BEC2FCB54}#1.0#0"; "IDAutomationPDF417.dll"
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "IDAutomation.com Metafile Image Generator Example"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "PDF417 Image Generator.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   465
      Left            =   7950
      TabIndex        =   38
      Top             =   4170
      Width           =   705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   285
      Left            =   8010
      TabIndex        =   37
      Top             =   2970
      Width           =   585
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "MW6 PDF417R3"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   5490
      ScaleHeight     =   1215
      ScaleWidth      =   3645
      TabIndex        =   36
      Top             =   4680
      Width           =   3705
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Sample"
      Height          =   315
      Left            =   4800
      TabIndex        =   35
      Top             =   240
      Width           =   1500
   End
   Begin VB.TextBox SaveAs 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   6480
      MaxLength       =   40
      TabIndex        =   4
      Text            =   "SavedBarCode.wmf"
      Top             =   660
      Width           =   2745
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4440
      TabIndex        =   14
      Text            =   "0"
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Symbology Properties (Measurements in CM)"
      Height          =   1455
      Left            =   2040
      TabIndex        =   22
      Top             =   960
      Width           =   7215
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
         Left            =   4860
         MaxLength       =   14
         TabIndex        =   17
         Text            =   "3"
         Top             =   450
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
         Left            =   4860
         MaxLength       =   14
         TabIndex        =   16
         Text            =   "4"
         Top             =   180
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
         Left            =   90
         MaxLength       =   14
         TabIndex        =   10
         Text            =   "3.0"
         Top             =   540
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
         Left            =   90
         MaxLength       =   14
         TabIndex        =   9
         Text            =   "0.03"
         Top             =   270
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
         Left            =   90
         MaxLength       =   14
         TabIndex        =   11
         Text            =   "0.2"
         Top             =   810
         Width           =   600
      End
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
         Left            =   90
         MaxLength       =   14
         TabIndex        =   12
         Text            =   "0.2"
         Top             =   1080
         Width           =   600
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2400
         TabIndex        =   13
         Text            =   "Binary"
         Top             =   240
         Width           =   990
      End
      Begin VB.CheckBox Truncate 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Truncate"
         Height          =   240
         Left            =   2400
         TabIndex        =   15
         Top             =   960
         Width           =   1590
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Error Correction Level"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5535
         TabIndex        =   33
         Top             =   495
         Width           =   1590
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PDF Columns"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5535
         TabIndex        =   32
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "X to Y Ratio"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   720
         TabIndex        =   29
         Top             =   585
         Width           =   960
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Orientation"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3480
         TabIndex        =   28
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "PDF Mode"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3465
         TabIndex        =   26
         Top             =   300
         Width           =   885
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Top Margin"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   720
         TabIndex        =   25
         Top             =   855
         Width           =   1005
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Left Margin"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   720
         TabIndex        =   24
         Top             =   1125
         Width           =   870
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Narrow Bar Width"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   720
         TabIndex        =   23
         Top             =   315
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Image Size in Twips"
      Height          =   1455
      Left            =   120
      TabIndex        =   19
      Top             =   960
      Width           =   1815
      Begin VB.TextBox ImageWidth 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   120
         MaxLength       =   14
         TabIndex        =   8
         Text            =   "3940"
         Top             =   1110
         Width           =   780
      End
      Begin VB.TextBox ImageHeight 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   120
         MaxLength       =   14
         TabIndex        =   7
         Text            =   "1710"
         Top             =   840
         Width           =   780
      End
      Begin VB.CommandButton AutoSizeButton 
         Caption         =   "Auto Size Now"
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Top             =   480
         Width           =   1500
      End
      Begin VB.CheckBox AutoSize 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Auto Size Image"
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Width"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   975
         TabIndex        =   21
         Top             =   1155
         Width           =   600
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Height"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   975
         TabIndex        =   20
         Top             =   885
         Width           =   600
      End
   End
   Begin VB.CommandButton CopyToClipboard 
      Caption         =   "Copy To Clipboard"
      Height          =   315
      Left            =   4800
      TabIndex        =   1
      Top             =   600
      Width           =   1500
   End
   Begin VB.CommandButton SaveToFile 
      Caption         =   "Save To File"
      Height          =   255
      Left            =   7950
      TabIndex        =   3
      Top             =   390
      Width           =   1260
   End
   Begin VB.TextBox DataToEncodeText 
      Height          =   690
      Left            =   120
      MaxLength       =   640
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "PDF417 Image Generator.frx":030A
      Top             =   210
      Width           =   4650
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   4410
      Picture         =   "PDF417 Image Generator.frx":033E
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   4305
   End
   Begin VB.Label Label15 
      BackColor       =   &H8000000E&
      Caption         =   "Label15"
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   780
      TabIndex        =   39
      Top             =   2760
      Width           =   1965
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "For help with this application, please refer to the manual provided with the ActiveX Control"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2460
      TabIndex        =   34
      Top             =   2460
      Width           =   6816
   End
   Begin PDF417LibCtl.PDF PDF1 
      DragMode        =   1  'Automatic
      Height          =   2040
      Left            =   120
      TabIndex        =   31
      Top             =   3270
      Width           =   4170
      _cx             =   7355
      _cy             =   3598
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
      Picture         =   "PDF417 Image Generator.frx":486E
      DataToEncode    =   ""
      Orientation     =   0
      XtoYRatio       =   3
      NarrowBarCM     =   0,03
      LeftMarginCM    =   0,2
      TopMarginCM     =   0,2
      Truncated       =   0
      PDFRows         =   0
      PDFColumns      =   6
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
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "File name and path:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6480
      TabIndex        =   30
      Top             =   480
      Width           =   2640
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Code 128 Set"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8160
      TabIndex        =   27
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Barcode Image Preview:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   90
      TabIndex        =   18
      Top             =   2430
      Width           =   5550
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Data To Encode in Barcode:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3345
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AutoSizeButton_Click()
    ImageHeight.Text = 40
    ImageWidth.Text = 40
    If AutoSize Then
       ImageHeight.Text = 960
    Else
       ImageHeight.Text = PDF1.GetYPixels * 15
       ImageWidth.Text = PDF1.GetXPixels * 16
    End If
End Sub

Private Sub BarColor_Change()
    'If Not IsNumeric(BarHeight.Text) Then BarHeight.Text = "1"
    PDF1.ForeColor = BarColor.Text
    'If AutoSize Then AutoSizeButton_Click
End Sub

Private Sub BarHeight_Change()
    If Not IsNumeric(BarHeight.Text) Then BarHeight.Text = "1"
    PDF1.BarHeight = BarHeight.Text
    AutoSizeButton_Click
End Sub

Private Sub cmdPrint_Click()
    'Print text directly to the printer, start at position x=2048
    Printer.CurrentX = 200
    Printer.Print "IDAutomation.com, Inc. Print Sample"
    Printer.CurrentX = 200
    Printer.Print "Printing the barcode at X=200, Y=1024"
    'Print the barcode directly to the printer, start at position x=200,y=1024
    Printer.PaintPicture PDF1.Picture, 200, 1024
    'Printer.Print "This above barcode is encoding the data:"
    'Printer.Print DataToEncodeText.Text
    'Eject the page...
    Printer.EndDoc
End Sub

Private Sub Combo1_Click()
    If Combo1.ListIndex = 0 Then PDF1.PDFMode = 0
    If Combo1.ListIndex = 1 Then PDF1.PDFMode = 1
    If AutoSize Then AutoSizeButton_Click
End Sub


Private Sub Combo3_Click()
    If Combo3.ListIndex = 0 Then PDF1.Orientation = 0
    If Combo3.ListIndex = 1 Then PDF1.Orientation = 90
    If Combo3.ListIndex = 2 Then PDF1.Orientation = 180
    If Combo3.ListIndex = 3 Then PDF1.Orientation = 270
    If AutoSize Then AutoSizeButton_Click
End Sub

Private Sub Command1_Click()
'' Set Me.Image = Me.PDF1.Picture
' With Image1
'   .Width = Picture1.Width
'   .Height = Picture1.Height
'   .Stretch = True
'   .Picture = Me.PDF1.Picture
'   Picture1.Picture = .Picture
' End With
'
'' With Picture1
''   .Width = PDF1.Width
''   .Height = PDF1.Height
'''   .Stretch = True
''   .Picture = Me.PDF1.Picture
'' End With
ImageHeight.Text = 1310
ImageWidth.Text = 2040
PDF1.Height = ImageHeight.Text
PDF1.Width = ImageWidth.Text
PDFColumns.Text = 3
PDF1.PDFColumns = PDFColumns.Text
End Sub

Private Sub Command2_Click()
Dim esquerda_picture As Double
Dim superior_picture As Double
Dim diferenca_tamanho_hor As Double
Dim diferenca_tamanho_ver As Double

'For i = 1 To 10
  esquerda_picture = Picture1.Left
  superior_picture = Picture1.Top
  diferenca_tamanho_hor = Picture1.Width - 100
  diferenca_tamanho_ver = Picture1.Height - 500
'  Me.Image1.Left = esquerda_picture + (diferenca_tamanho_hor / 2)
'  Image1.Top = superior_picture + (diferenca_tamanho_ver / 2)
'Next

End Sub

Private Sub CopyToClipboard_Click()
    'First, clear the clipboard
    Clipboard.Clear
    'Copy the metafile image into the clipboard
    Clipboard.SetData PDF1.Picture, vbCFMetafile
End Sub

Private Sub DataToEncodeText_Change()
    'This is how we set the data to encode in the barcode
    PDF1.DataToEncode = DataToEncodeText.Text
    If AutoSize Then AutoSizeButton_Click
'    Picture1.Font = "39251" 'PODE SER ARIAL, COURIER, ETC...
    Picture1.FontSize = 10
    Picture1.Print PDF1.Picture
End Sub





Private Sub ImageHeight_Change()
    If Not IsNumeric(ImageHeight.Text) Then ImageHeight.Text = "1000"
    PDF1.Height = ImageHeight.Text
End Sub

Private Sub ImageWidth_Change()
    If Not IsNumeric(ImageWidth.Text) Then ImageWidth.Text = "3000"
    PDF1.Width = ImageWidth.Text
End Sub

Private Sub LeftMarginCM_Change()
    If Val(LeftMarginCM.Text) > 0 And Val(LeftMarginCM.Text) < 50 Then PDF1.LeftMarginCM = LeftMarginCM.Text
    If Val(LeftMarginCM.Text) = 0 Then PDF1.LeftMarginCM = "0"
    If AutoSize Then AutoSizeButton_Click
End Sub

Private Sub NarrowBarWidth_Change()
    If Val(NarrowBarWidth.Text) > 0 And Val(NarrowBarWidth.Text) < 10 Then
        PDF1.NarrowBarCM = NarrowBarWidth.Text
        If PDF1.NarrowBarCM <> "0.03" And PDF1.NarrowBarCM <> "0.06" Then MsgBox "Notice: Because Windows desktops normally display images at 96 DPI, Narrow Bar Width settings that are not in .03CM increments may not display accurately. However, the image will be accurate when you generate a graphic image or print the image because the WMF image rasters to the print device. You may have to manually adjust the height and width in twips when using settings other than .03CM or .06CM."
    End If
    If AutoSize Then AutoSizeButton_Click
End Sub

Private Sub PrintDirect_Click()
    'Print text directly to the printer, start at position x=2048
    Printer.CurrentX = 2048
    Printer.Print "IDAutomation.com, Inc. VB Example"
    Printer.CurrentX = 2048
    Printer.Print "Printing the barcode at X=2048, Y=1024"
    'Print the barcode directly to the printer, start at position x=2048,y=1024
    Printer.PaintPicture PDF1.Picture, 2048, 1024
    'Eject the page...
    Printer.EndDoc
End Sub

Private Sub PDFColumns_Change()
    If Val(PDFColumns.Text) > 0 And Val(PDFColumns.Text) < 31 Then PDF1.PDFColumns = PDFColumns.Text
    If AutoSize Then AutoSizeButton_Click
End Sub

Private Sub PDFErrorCorrectionLevel_Change()
    If Val(PDFErrorCorrectionLevel.Text) > 0 And Val(PDFErrorCorrectionLevel.Text) < 8 Then PDF1.PDFErrorCorrectionLevel = PDFErrorCorrectionLevel.Text
    If AutoSize Then AutoSizeButton_Click
End Sub

Private Sub SaveToFile_Click()
    'MsgBox "Your barcode will be saved to the file [SavedBarCode.wmf] in the program directory."
    If SaveAs.Text = "" Then SaveAs.Text = "SavedBarCode.wmf"
    'PDF1.SaveBarCode SaveAs.Text
    PDF1.SaveEnhWMF SaveAs.Text
End Sub



Private Sub TopMarginCM_Change()
    If Val(TopMarginCM.Text) > 0 And Val(TopMarginCM.Text) < 10 Then PDF1.TopMarginCM = TopMarginCM.Text
    If Val(TopMarginCM.Text) = 0 Then PDF1.TopMarginCM = "0"
    If AutoSize Then AutoSizeButton_Click
End Sub

Private Sub Truncate_Click()
    If Truncate Then
        PDF1.Truncated = 1
    Else
        PDF1.Truncated = 0
    End If
End Sub

Private Sub W2NRatio_Change()
    If Val(W2NRatio.Text) > 0 And Val(W2NRatio.Text) < 20 Then PDF1.XtoYRatio = W2NRatio.Text
    If AutoSize Then AutoSizeButton_Click
End Sub
