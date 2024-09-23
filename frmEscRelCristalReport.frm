VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmEscRelCristalReport 
   Caption         =   "Relatorios"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   11610
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   7905
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   12345
      lastProp        =   500
      _cx             =   21775
      _cy             =   13944
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmEscRelCristalReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CRViewer1_PrintButtonClicked(UseDefault As Boolean)
'If UseDefault = True Then
'   MsgBox "true"
'   UseDefault = False
'   Me.CommonDialog1.ShowPrinter
'   UseDefault = True
'   Me.CRViewer1.PrintReport
'   Me.CRViewer1.Name
'Else
'   MsgBox "false"
'End If
End Sub

Private Sub Form_Activate()
Me.CRViewer1.Top = Me.CRViewer1.Top + 100
Me.CRViewer1.Width = Me.Width - 1000
Me.CRViewer1.Left = 100
Me.CRViewer1.Height = Me.ScaleHeight - 1000
End Sub

Private Sub Form_Load()
Me.Top = 0
Me.Left = 0
End Sub
