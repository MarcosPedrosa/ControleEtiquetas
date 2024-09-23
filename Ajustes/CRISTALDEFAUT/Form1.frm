VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   7230
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type PRINTER_INFO_4
   pPrinterName As Long
   pServerName As Long
   Attributes As Long
End Type

Private Const SIZEOFPRINTER_INFO_4 = 12
Private Const HWND_BROADCAST As Long = &HFFFF&
Private Const WM_WININICHANGE As Long = &H1A
Private Const PRINTER_LEVEL4 = &H4
Private Const PRINTER_ENUM_LOCAL = &H2

Private Declare Function EnumPrinters Lib 'winspool.drv' Alias 'EnumPrintersA' (ByVal Flags As Long, ByVal Name As String, ByVal Level As Long, pPrinterEnum As Any, ByVal cbBuffer As Long, pcbNeeded As Long, pcReturned As Long) As Long
Private Declare Function SendNotifyMessage Lib 'user32' Alias 'SendNotifyMessageA' (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetDefaultPrinter Lib 'winspool.drv' Alias 'SetDefaultPrinterA' (ByVal pszPrinter As String) As Long
Private Declare Function lstrcpyA Lib 'kernel32' (ByVal RetVal As String, ByVal ptr As Long) As Long
Private Declare Function lstrlenA Lib 'kernel32' (ByVal ptr As Any) As Long
Private Declare Function GetProfileString Lib 'kernel32' Alias 'GetProfileStringA' (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Private Sub Form_Load()
   Command1.Enabled = List1.ListIndex > -1
   Call EnumPrintersWinNTPlus
   Label1.Caption = List1.ListCount & ' impressoras instaladas'
End Sub

Private Sub Command1_Click()
   SetDefaultPrinter List1.List(List1.ListIndex)
   Call SendNotifyMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal 'windows')
End Sub

Private Sub List1_Click()
   Command1.Enabled = List1.ListIndex > -1
End Sub

Private Function EnumPrintersWinNTPlus() As Long
   Dim cbRequired As Long
   Dim cbBuffer As Long
   Dim ptr() As PRINTER_INFO_4
   Dim nEntries As Long
   Dim cnt As Long
  
   List1.Clear
  
   Call EnumPrinters(PRINTER_ENUM_LOCAL, vbNullString, PRINTER_LEVEL4, 0, 0, cbRequired, nEntries)
   ReDim ptr((cbRequired  SIZEOFPRINTER_INFO_4))
   cbBuffer = cbRequired
   If EnumPrinters(PRINTER_ENUM_LOCAL, vbNullString, PRINTER_LEVEL4, ptr(0), cbBuffer, cbRequired, nEntries) Then
      For cnt = 0 To nEntries - 1
        List1.AddItem GetStrFromPtrA(ptr(cnt).pPrinterName)
        If GetStrFromPtrA(ptr(cnt).pPrinterName) = GetDefaultPrinter.DeviceName Then
         List1.ListIndex = List1.ListCount - 1
        End If
      Next cnt
   Else
      List1.AddItem 'Error enumerating printers.'
   End If
   EnumPrintersWinNTPlus = nEntries
End Function

Private Function GetStrFromPtrA(ByVal lpszA As Long) As String
   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
End Function

Private Function GetDefaultPrinter() As Printer
  Dim strBuffer As String * 254
  Dim iRetValue As Long
  Dim strDefaultPrinterInfo As String
  Dim tblDefaultPrinterInfo() As String
  Dim objPrinter As Printer

  iRetValue = GetProfileString('windows', 'device', ',,,', strBuffer, 254)
  strDefaultPrinterInfo = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
  tblDefaultPrinterInfo = Split(strDefaultPrinterInfo, ',')
  For Each objPrinter In Printers
    If objPrinter.DeviceName = tblDefaultPrinterInfo(0) Then
      Exit For
    End If
  Next

If objPrinter.DeviceName <> tblDefaultPrinterInfo(0) Then Set objPrinter = Nothing
Set GetDefaultPrinter = objPrinter
End Function
