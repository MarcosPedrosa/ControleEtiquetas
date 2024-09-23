VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} CrystalReport1 
   ClientHeight    =   10650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12885
   OleObjectBlob   =   "CrystalReport4.dsx":0000
End
Attribute VB_Name = "CrystalReport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Section3_Format(ByVal pFormattingInfo As Object)
    'create the bar code for each record set
    Dim data As String
    data = Field1.Value
    Set Me.Picture1.FormattedPicture = Form1.BarcodeGenerate(data, Me.Picture1.Width, Me.Picture1.Height).Image
End Sub
