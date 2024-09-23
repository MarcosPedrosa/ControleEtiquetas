VERSION 5.00
Object = "{D2F3FE5B-BCF5-4274-9EE3-8E402EAD4E12}#1.0#0"; "TBarCode6.ocx"
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form Form2 
   Caption         =   "RepMainForm"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9285
   LinkTopic       =   "Form2"
   ScaleHeight     =   7110
   ScaleWidth      =   9285
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   6135
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   9255
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "CRForm.frx":0000
      Top             =   120
      Width           =   3255
   End
   Begin TBARCODE6LibCtl.TBarCode6 TBarCode61 
      Height          =   615
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   3135
      _cx             =   5530
      _cy             =   1085
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
      LastError       =   "The operation completed successfully. "
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
   Begin VB.Label Label1 
      Caption         =   "Sample: Use of TBarCode OCX within Crystal Reports 8"
      Height          =   495
      Left            =   6960
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9240
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CrystalReport1

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
CRViewer1.ReportSource = Report
CRViewer1.ViewReport
Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
'CRViewer1.Top = 0
CRViewer1.Top = Me.ScaleY(20, vbMillimeters)

CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth

End Sub

