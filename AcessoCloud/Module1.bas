Attribute VB_Name = "Module1"
'
'Option Explicit
'
''* Tools->References
'' MSScriptControl      Microsoft Script Control 1.0    C:\Windows\SysWOW64\msscript.ocx
'' Scripting            Microsoft Scripting Runtime     C:\Windows\SysWOW64\scrrun.dll
'' MSXML2               Microsoft XML, v6.0             C:\Windows\SysWOW64\msxml6.dll
'
'
'Function CodeInjectingJson() As String
'    CodeInjectingJson = "{""foo"":""bar"", a:(function(){(new ActiveXObject('Scripting.FileSystemObject'))." & _
'                        "GetSpecialFolder(2).CreateTextFile('random.txt')." & _
'                        "Write('Use JSON.parse instead');})()}"
'End Function
'
'Private Sub Proof_That_ScriptControl_Eval_Executes_Injected_Javascript()
'
'    Dim FSO As Scripting.FileSystemObject
'    Set FSO = New Scripting.FileSystemObject
'
'    If FSO.FileExists(Environ$("USERPROFILE") & "\AppData\Local\Temp\random.txt") Then
'        FSO.DeleteFile Environ$("USERPROFILE") & "\AppData\Local\Temp\random.txt"
'    End If
'
'    Debug.Assert Not FSO.FileExists(Environ$("USERPROFILE") & "\AppData\Local\Temp\random.txt")
'
'    Dim oSC As ScriptControl
'    Set oSC = SC
'
'    Stop
'    Dim obj As Object
'    Set obj = oSC.Eval("(" & CodeInjectingJson & ")")
'    Stop
'
'    Debug.Assert FSO.FileExists(Environ$("USERPROFILE") & "\AppData\Local\Temp\random.txt")
'    End ' <- need this to kill references and garbage collect
'End Sub
'
'Private Sub Attempt_To_Parse_Injected_Code_With_JSON_Parse_Throws_Error()
'
'    Dim FSO As Scripting.FileSystemObject
'    Set FSO = New Scripting.FileSystemObject
'
'    If FSO.FileExists(Environ$("USERPROFILE") & "\AppData\Local\Temp\random.txt") Then
'        FSO.DeleteFile Environ$("USERPROFILE") & "\AppData\Local\Temp\random.txt"
'    End If
'
'
'    Debug.Assert Not FSO.FileExists(Environ$("USERPROFILE") & "\AppData\Local\Temp\random.txt")
'
'    Dim oSC As ScriptControl
'    Set oSC = SC
'
'    'Stop
'    Dim objSafelyParsed As Object
'    Set objSafelyParsed = SC.Run("JSON_parse", CodeInjectingJson)
'
'
'    'Details of error thrown {
'    ' Err.Number=5022,
'    ' Err.Description="Exception thrown and not caught",
'    ' Err.Source="Microsoft JScript runtime error"
'    ' }
'
'    '* search for '5022' on https://docs.microsoft.com/en-us/scripting/javascript/reference/javascript-run-time-errors
'    '* links to  https://docs.microsoft.com/en-us/scripting/javascript/misc/exception-thrown-and-not-caught
'    '* also https://referencesource.microsoft.com/#Microsoft.JScript/Microsoft/JScript/JSError.cs,a0c5ae7e7c2dd23c,references
'    Stop
'End Sub
