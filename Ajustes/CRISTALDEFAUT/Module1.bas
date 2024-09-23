Attribute VB_Name = "Module1"

Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey _
As Long, ByVal lpSubKey As String, ByVal dwReserved As Long, ByVal samDesired _
As Long, phkResult As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName$, ByVal lpdwReserved As Long, _
lpdwType As Long, lpData As Any, lpcbData As Long) As Long

Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Public Const HKEY_CURRENT_CONFIG As Long = &H80000005

Public sImpressoraInmetro As String


Private Const SW_SHOWNORMAL As Long = 1

Private Declare Function ConnectToPrinterDlg Lib "winspool.drv" _
   (ByVal hwnd As Long, ByVal flags As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
   (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
    
