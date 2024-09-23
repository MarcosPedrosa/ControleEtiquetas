VERSION 5.00
Begin VB.UserControl HttpService 
   BackColor       =   &H8000000D&
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1710
   ScaleWidth      =   2475
End
Attribute VB_Name = "HttpService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const m_ksProperty_Default              As String = ""

Private m_sHost                                 As String
Private m_nPort                                 As Long
Private m_sPath                                 As String
Private m_dctQueryStringParameters              As Scripting.Dictionary

Private m_sOutput                               As String

' Ensure that all parts of the query string are deleted.
Public Sub ClearQueryString()

    Set m_dctQueryStringParameters = New Scripting.Dictionary

End Sub

' Executes "GET" method for URL.
Public Function Get_() As String

    ' Read in data from URL. UserControl_AsyncReadComplete will fire when finished.
    UserControl.AsyncRead "http://" & m_sHost & ":" & CStr(m_nPort) & "" & m_sPath & "?" & GetQueryString(), vbAsyncTypeByteArray, m_ksProperty_Default, vbAsyncReadSynchronousDownload
'    UserControl.AsyncRead "http://" & m_sHost & ":" & CStr(m_nPort) & "" & m_sPath & "" & GetQueryString(), vbAsyncTypeByteArray, m_ksProperty_Default, vbAsyncReadSynchronousDownload

'    UserControl.AsyncRead ("http://" & m_sHost & ":" & CStr(m_nPort) & "" & m_sPath & "?" & GetQueryString() & ",2")

    ' Return the contents of the buffer.
    Get_ = m_sOutput

    ' Clear down state.
    m_sOutput = vbNullString

End Function

' Returns query string based on dictionary.
Private Function GetQueryString() As String

    Dim vName                                   As Variant
    Dim sQueryString                            As String

    For Each vName In m_dctQueryStringParameters
        sQueryString = sQueryString & CStr(vName) & "=" & m_dctQueryStringParameters.Item(vName) & "&"
    Next vName
    
    If Len(sQueryString) > 0 Then
       GetQueryString = Left$(sQueryString, Len(sQueryString) - 1)
    End If
End Function

' Sets the remote host.
Public Property Let Host(ByVal the_sValue As String)

    m_sHost = the_sValue

End Property

' Sets the directory and filename part of the URL.
Public Property Let Path(ByVal the_sValue As String)

    m_sPath = the_sValue

End Property

' Sets the port number for this request.
Public Property Let Port(ByVal the_nValue As Long)

    m_nPort = the_nValue

End Property

' Sets a name/value pair in the query string. Supports duplicate names.
Public Property Let QueryStringParameter(ByVal the_sName As String, ByVal the_sValue As String)

    m_dctQueryStringParameters.Item(the_sName) = the_sValue

End Property

' Fired when the download is complete.
Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)

    ' Gets the data from the internet transfer.
    m_sOutput = StrConv(AsyncProp.Value, vbUnicode)

End Sub

Private Sub UserControl_Initialize()

    ' Initialises the scripting dictionary.
    Set m_dctQueryStringParameters = New Scripting.Dictionary

End Sub
