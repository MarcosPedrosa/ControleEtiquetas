VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PDF417 Encoder for Windows; Version 4.1; Copyright IDAutomation.com, Inc."
   ClientHeight    =   7785
   ClientLeft      =   825
   ClientTop       =   1365
   ClientWidth     =   10560
   Icon            =   "Barcode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   120
      Picture         =   "Barcode.frx":030A
      ScaleHeight     =   975
      ScaleWidth      =   7335
      TabIndex        =   26
      Top             =   6600
      Width           =   7335
   End
   Begin VB.CheckBox chkApplyTilde 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Apply Tilde"
      Height          =   210
      Left            =   4155
      TabIndex        =   25
      Top             =   3450
      Width           =   3705
   End
   Begin VB.CheckBox ForceBin 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Text Compaction Mode (reduces size for text only - ASCII 9,10,13 and 32-128)"
      Height          =   210
      Left            =   4140
      TabIndex        =   24
      Top             =   3720
      Width           =   6345
   End
   Begin VB.CheckBox DisplayFont 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Show string below as selected font (changes the font property of the text box below)"
      Height          =   210
      Left            =   4140
      TabIndex        =   23
      Top             =   4395
      Value           =   1  'Checked
      Width           =   6225
   End
   Begin VB.ComboBox PointSize 
      Height          =   315
      ItemData        =   "Barcode.frx":20B1
      Left            =   135
      List            =   "Barcode.frx":20B3
      TabIndex        =   21
      Top             =   3480
      Width           =   2400
   End
   Begin VB.ComboBox FontSelection 
      Height          =   315
      ItemData        =   "Barcode.frx":20B5
      Left            =   120
      List            =   "Barcode.frx":20B7
      TabIndex        =   18
      Top             =   3000
      Width           =   2415
   End
   Begin VB.TextBox RowSpecify 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   240
      Left            =   4140
      MaxLength       =   2
      TabIndex        =   16
      Top             =   3150
      Width           =   465
   End
   Begin VB.TextBox ColumnSpecify 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   240
      Left            =   4140
      MaxLength       =   2
      TabIndex        =   14
      Top             =   2880
      Width           =   465
   End
   Begin VB.CheckBox Truncate 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Truncate (creates truncated symbol, must be chosen before generating)"
      Height          =   180
      Left            =   4140
      TabIndex        =   13
      Top             =   2655
      Width           =   5805
   End
   Begin VB.ComboBox EccLevel 
      Height          =   315
      ItemData        =   "Barcode.frx":20B9
      Left            =   4140
      List            =   "Barcode.frx":20BB
      TabIndex        =   11
      Top             =   2295
      Width           =   915
   End
   Begin VB.CommandButton Tutor 
      Caption         =   "Application Tutorial"
      Height          =   315
      Left            =   720
      TabIndex        =   8
      Top             =   360
      Width           =   2415
   End
   Begin VB.CheckBox CopyToClip 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Copy Output to Clipboard"
      Height          =   180
      Left            =   4140
      TabIndex        =   1
      Top             =   3945
      Value           =   1  'Checked
      Width           =   3195
   End
   Begin VB.CommandButton WebsiteVisit 
      Caption         =   "Order or Renew License"
      Height          =   315
      Left            =   720
      TabIndex        =   7
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton LicenseAgreement 
      Caption         =   "License Agreement"
      Height          =   315
      Left            =   720
      TabIndex        =   9
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Exit 
      Caption         =   "Exit"
      Height          =   315
      Left            =   720
      TabIndex        =   10
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton CreatePDF 
      Caption         =   "Print / Generate PDF417 Code"
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   3990
      Width           =   3015
   End
   Begin VB.CheckBox PrintBarcode 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Print Barcode (requires font to be installed)"
      Height          =   210
      Left            =   4140
      TabIndex        =   2
      Top             =   4170
      Width           =   3705
   End
   Begin VB.TextBox PrintableString 
      Height          =   1890
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4650
      Width           =   10380
   End
   Begin VB.TextBox TextToEncode 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1890
      Left            =   4140
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   315
      Width           =   6315
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Font Size"
      Height          =   225
      Left            =   2640
      TabIndex        =   22
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Font Printing Options:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   135
      TabIndex        =   20
      Top             =   2760
      Width           =   2280
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Font"
      Height          =   225
      Left            =   2640
      TabIndex        =   19
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Specify Minimum # of Rows:  (leave blank for automatic)"
      Height          =   225
      Left            =   4635
      TabIndex        =   17
      Top             =   3150
      Width           =   5340
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Specify Columns:  (leave blank for automatic)"
      Height          =   225
      Left            =   4635
      TabIndex        =   15
      Top             =   2880
      Width           =   5340
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Error Correction Level:  (leave blank for automatic and recommended level)"
      Height          =   225
      Left            =   5085
      TabIndex        =   12
      Top             =   2385
      Width           =   5340
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "String to print / symbol:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   5
      Top             =   4395
      Width           =   3825
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Data To Encode in Barcode:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4140
      TabIndex        =   4
      Top             =   90
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim I As Integer
Dim F As Integer
Dim EcLevel As Integer
Dim DataToPrint As String
Dim DataToEncode As String
Dim OnlyCorrectData As String
Dim Printable_string As String
Dim Truncated As Integer
Dim TotalRows As Integer
Dim TotalColumns As Integer
Dim msg As String
Dim Proportional As Integer
Dim PDFMode As Integer

Private Declare Function GetActiveWindow Lib _
 "user32" () As Long

Private Declare Function ShellExecute Lib _
 "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal _
    lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    Const SW_SHOWNORMAL = 1

Sub RunFile(ByVal mFile As String, mFilePath As String, RunStyle As Integer)

    Dim temp As Long
    Dim msg As String
    Dim x As Long

    temp = GetActiveWindow()
    x = ShellExecute(temp, "Open", mFile, "", mFilePath, RunStyle)
    
    If x < 32 Then
        Select Case x
            Case 0
                msg = "The file could not be run due to insufficient system memory or a corrupt program file"
            Case 2
                msg = "File Not Found"
            Case 3
                msg = "Invalid Path"
            Case 5
                msg = "Sharing or protection error"
            Case 6
                msg = "Separate data segments are required for each task "
            Case 8
                msg = "Insufficient memory to run the program"
            Case 10
                msg = "Incorrect Windows version"
            Case 11
                msg = "Invalid Program File"
            Case 12
                msg = "Program file requires a different operating System "
            Case 13
                msg = "Program requires MS-DOS 4.0"
            Case 14
                msg = "Unknown program file type"
            Case 15
                msg = "Windows prgram does not support protected memory mode"
            Case 16
                msg = "Invalid use of data segments when loading a second instance of a program"
            Case 19
                msg = "Attempt to run a compressed program file"
            Case 20
                msg = "Invalid dynamic link library"
            Case 21
                msg = "Program requires Windows 32-bit extensions"
            Case 31
                msg = "No application found for this file"
        End Select

        MsgBox msg, vbOKOnly + vbCritical + vbSystemModal, "Error Message"

    End If
    
End Sub


Private Sub ColumnSpecify_Change()
PrintableString.Text = ""
End Sub

Private Sub DisplayFont_Click()
If FontSelection.Text = "" Then
    MsgBox "Please select a PDF417 font from the [Font Printing Options Menu] drop down box first."
Else
    PrintableString.FontSize = "8"
    PrintableString.Font = FontSelection.Text
    If DisplayFont.Value = "0" Then PrintableString.Font = "MS Sans Serif"
    If DisplayFont.Value = "0" Then PrintableString.FontSize = "8"
End If

If DisplayFont.Value = "0" Then PrintableString.Font = "MS Sans Serif"
If DisplayFont.Value = "0" Then PrintableString.FontSize = "8"
'PrintableString.Text = ""
End Sub

Private Sub EccLevel_Change()
PrintableString.Text = ""
End Sub

Private Sub Exit_Click()
    End
End Sub

Private Sub FontSelection_Change()
PrintableString.Text = ""
If DisplayFont.Value = "1" Then
    PrintableString.Font = FontSelection.Text
End If
End Sub

Private Sub ForceBin_Click()
'MsgBox "NOTE: Encoding with binary compaction may increase the size of the symbol. Use this option only if your are encoding binary data or the text you are encoding is not encoding properly."
PrintableString.Text = ""
End Sub

Private Sub Form_Load()

' setup combo box
EccLevel.AddItem ""
EccLevel.AddItem "1"
EccLevel.AddItem "2"
EccLevel.AddItem "3"
EccLevel.AddItem "4"
EccLevel.AddItem "5"
EccLevel.AddItem "6"
EccLevel.AddItem "7"
EccLevel.AddItem "8"
    
Dim DetectedFont As String
    
'Detect number of installed barcode fonts.
For F = 0 To Printer.FontCount - 1
    DetectedFont = Printer.Fonts(F)
    If DetectedFont = "IDAutomationPDF417n3" Then
        FontSelection.AddItem "IDAutomationPDF417n3"
        FontSelection.Text = "IDAutomationPDF417n3"
        PointSize.Text = "8"
        DisplayFont.Value = 1
    End If
    If DetectedFont = "IDAutomationPDF417n4" Then FontSelection.AddItem "IDAutomationPDF417n4"
    If DetectedFont = "IDAutomationPDF417n5" Then FontSelection.AddItem "IDAutomationPDF417n5"
Next F

Show
PointSize.AddItem "2"
PointSize.AddItem "3"
PointSize.AddItem "4"
PointSize.AddItem "6"
PointSize.AddItem "8"
PointSize.AddItem "10"
PointSize.AddItem "12"

End Sub


Private Sub LicenseAgreement_Click()
    Dim iret As Long
    On Error GoTo URL_Err
    
    ' Open URL into the default internet browser
    iret = ShellExecute(Me.hwnd, _
        vbNullString, _
  "http://idautomation.com/licenses.html", _
        vbNullString, _
  "c:\", SW_SHOWNORMAL)

    Exit Sub

URL_Err:

    MsgBox "There was an ERROR?", _
        vbOKOnly + vbExclamation + vbSystemModal, _
  "Open URL Error"

    Err = 0



End Sub

Private Sub CreatePDF_Click()
' This module is Copyright, IDAutomation.com, Inc. 2000.  All rights reserved.
' For more info visit http://www.BizFonts.com or http://www.IDAutomation.com.com
'
' The purpose of this code is to format data for the PDF417 Font from the PDF417 DLL.
' After data is formatted, it is sent to the application where the PDF417 font is selected.
'
' Get data from user, this is the DataToEncode
DataToEncode = TextToEncode.Text
Printable_string = ""
DataToPrint = ""
If Mid(FontSelection.Text, 1, 1) = "s" Then MsgBox "You have selected an evaluation version of the PDF417 font. The evaluation version of this font may be used for testing purposes only. If printing proportionally, the word DEMO will replace the start symbol and dashes will be intermittently placed in the symbol making it difficult to scan. Evaluation version fonts always begin with the letter [s]."

'Determine the EcLevel
If EccLevel.Text = "" Then
    EcLevel = "0"
Else
    EcLevel = EccLevel.Text
End If

'Determine row and column limits
TotalColumns = "0"
TotalRows = "0"
If ColumnSpecify.Text <> "" Then TotalColumns = ColumnSpecify.Text
If RowSpecify.Text <> "" Then TotalRows = RowSpecify.Text

Truncated = "0"
If Truncate.Value = 1 Then Truncated = "1"

PDFMode = "0"
If ForceBin.Value = 1 Then PDFMode = "1"

'Format the Output by calling the DLL
'Make sure to add the COM Server DLL to the references!!
Dim PDF417FontEncoder As PDF417Lib.PDF
Set PDF417FontEncoder = New PDF
Dim TildeValue As Integer
If chkApplyTilde.Value = vbChecked Then
    TildeValue = 1
Else
    TildeValue = 0
End If
PDF417FontEncoder.FontEncode DataToEncode, EcLevel, TotalColumns, TotalRows, Truncated, PDFMode, TildeValue, Printable_string

'Print the barcode
Printer.FontSize = 10
'Check if font is selected & print selected one
If FontSelection.Text <> "" Then Printer.Font = FontSelection.Text
If PointSize.Text <> "" Then Printer.FontSize = PointSize.Text
If FontSelection.Text = "" Then Printer.Font = "IDAutomationPDF417n3"
If PrintBarcode.Value = 1 Then Printer.Print Printable_string

'print the human readable text below the barcode
Printer.FontSize = 8
Printer.Font = "Times New Roman"
If PrintBarcode.Value = 1 Then Printer.Print Chr(13) & Chr(10) & DataToEncode

'Send the barcode string to the clipboard
If CopyToClip.Value = 1 Then Clipboard.Clear
If CopyToClip.Value = 1 Then Clipboard.SetText Printable_string

'Print what is in the printer buffer
If PrintBarcode.Value = 1 Then Printer.EndDoc

'Display PrintableString codes in textbox
PrintableString.Text = Printable_string

End Sub

Private Sub PrintBarcode_Click()
    If FontSelection.Text = "" Then MsgBox "Please select a font from the font selection drop down box below. You will be unable to print PDF417 barcodes until the PDF417 font is installed."
'PrintableString.Text = ""
End Sub

Private Sub Proportion_Click()
MsgBox "NOTE: Encoding PDF417 proportionally will decrease computation time, however, the proportional encoding will not properly rasterize at all point sizes on all printers or computer monitors. Do not print proportionally if you notice that the right side of the symbol is not aligned. "
PrintableString.Text = ""
End Sub

Private Sub RowSpecify_Change()
PrintableString.Text = ""
End Sub

Private Sub TextToEncode_Change()
PrintableString.Text = ""
End Sub

Private Sub Truncate_Click()
PrintableString.Text = ""
End Sub

Private Sub Tutor_Click()
    Dim iret As Long
    On Error GoTo URL_Err
    
    ' Open URL into the default internet browser
    iret = ShellExecute(Me.hwnd, _
        vbNullString, _
  "http://www.idautomation.com/fonts/pdf417encoder/#Tutorial", _
        vbNullString, _
  "c:\", SW_SHOWNORMAL)

    Exit Sub

URL_Err:

    MsgBox "There was an ERROR?", _
        vbOKOnly + vbExclamation + vbSystemModal, _
  "Open URL Error"

    Err = 0


End Sub

Private Sub WebsiteVisit_Click()
    Dim iret As Long
    On Error GoTo URL_Err
    
    ' Open URL into the default internet browser
    iret = ShellExecute(Me.hwnd, _
        vbNullString, _
  "http://idautomation.com/fonts/pdf417/#ORDER", _
        vbNullString, _
  "c:\", SW_SHOWNORMAL)

    Exit Sub

URL_Err:

    MsgBox "There was an ERROR?", _
        vbOKOnly + vbExclamation + vbSystemModal, _
  "Open URL Error"

    Err = 0

End Sub
