'
' This module is a home for package- and application-non-specific code that may
' be used to read or write Windows Registry keys/values.
'
' Macro Type:               Module
' Using:                    Declarations
'                           Subs_Util
'
' Contents
' fRegKeyExists             returns Boolean value indicating existence of registry key
' fRegKeyRead               returns value of specified registry key
'
'-----------------------------------------------------------------------------------------

Function funRegKeyExists(strRegKey)
    ' Checks for the existence of the specified registry key, returns Boolean value

    On Error Resume Next

    Dim objWS
    Dim strErrMsg

    Set objWS = CreateObject("WScript.Shell")
    If Err.Number = 0 Then
	MsgBox "Successfully created WScript object"
	objWS.RegRead strRegKey
	If Err.Number = 0 Then
	    'registry key found
	    funRegKeyExists = True
	Else
	    'registry key not found
	    funRegKeyExists = False
	    Err.Clear
	End If
    Else
	strErrMsg = "ERROR: Could not create WScript object"
	Subs_Util.subDisplayErrorMsg(strErrMsg)
    End If

End Function

'-----------------------------------------------------------------------------------------

Function funRegKeyRead(strRegKey) As String
    ' Returns the value of the specified registry key

    On Error Resume Next

    Dim objWS
    Dim strErrMsg

    Set objWS = CreateObject("WScript.Shell")
    If Err.Number = 0 Then
	MsgBox "Successfully created WScript object"
	objWS.RegRead strRegKey
	If Err.Number = 0 Then
	    'registry key found
	    funRegKeyRead = objWS.RegRead(strRegKey)
	Else
	    'registry key not found
	    funRegKeyExists = False
	    Err.Clear
	End If
    Else
	strErrMsg = "ERROR: Could not create WScript object"
	Subs_Util.subDisplayErrorMsg(strErrMsg)
    End If

End Function
