Attribute VB_Name = "IDAutomation_PDF417_Macro"
Option Compare Database
Option Explicit

'*********************************************************************
'*  VB Function for IDAutomation PDF417 Fonts via ActiveX DLL
'*  Copyright, IDAutomation.com, Inc. 2000-2006. All rights reserved.
'*
'*  This function is only to be used with the
'*  IDAutomation PDF417 Font and Encoder
'*  http://www.idautomation.com/fonts/pdf417/
'*
'*  You may incorporate our Source Code in your application
'*  only if you own a valid license from IDAutomation.com, Inc.
'*  for the associated font and the copyright notices are not
'*  removed from the source code.
'*
'*  Distributing our source code or fonts outside your
'*  organization requires a Developer License.
'*********************************************************************

Public Function IDAutomation_PDF417(DataToEncode As String, Optional EcLevel As Integer = 0, Optional TotalColumns As Integer = 0, Optional TotalRows As Integer = 0, Optional Truncated As Integer = 0, Optional PDFMode As Integer = 0, Optional ApplyTilde As Integer = 0) As String
    ' NOTE: Before this function will work you may have to add the
    ' DLL reference by choosing Tools - References and choose
    ' "IDAutomation PDF417 Barcode"
    Dim PDF417FontEncoder As PDF417Lib.PDF
    Set PDF417FontEncoder = New PDF
    PDF417FontEncoder.FontEncode DataToEncode, EcLevel, TotalColumns, TotalRows, Truncated, PDFMode, ApplyTilde, IDAutomation_PDF417
End Function

