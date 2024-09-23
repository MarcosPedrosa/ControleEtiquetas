VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} CrystalReport1 
   ClientHeight    =   10650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12885
   OleObjectBlob   =   "CrystalReport1.dsx":0000
End
Attribute VB_Name = "CrystalReport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Section3_Format(ByVal pFormattingInfo As Object)

' Simulate data binding
' by saving the bar code as bitmap with data from fldArticleID
' then reload it to a picture control

On Error Resume Next

Dim nWidth
Dim path
Dim fso
    
    path = "c:\temp\CR8_" & Me.fldArticleID.Value & ".bmp"
    
    Form2.TBarCode61.Text = Me.fldArticleID.Value
    Form2.TBarCode61.PrintDataText = False
    nWidth = Form2.TBarCode61.CountModules * 3 'adapt width to number of graphical modules
    
    Form2.TBarCode61.SaveImage path, eIMBmp, nWidth, 100, 96, 96
    Me.pictBarcode.SetOleLocation (path)

End Sub

