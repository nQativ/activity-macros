'
' This module is a home for package- and application-non-specific code that may
' be used to perform file-system actions.
'
' Macro Type:               Module
' Using:                    Declarations
'                           Subs_Util
'
' Contents
' fFileExists               returns Boolean value after checking for existence of file
' fSpecialFolderPath        returns path to the specified Windows 'special' folder
'
'-----------------------------------------------------------------------------------------

Function funFileExists(strFile)
    'Checks for existence of file or folder, returns Boolean
    'Expects strFile to be string including full path to the file or folder

    If Dir(strFile) <> "" Then
	funFileExists = True
    Else
	funFileExists = False
    End If

End Function

'--------------------------------------------------------------------------------------------------------------

Function funSpecialFolderPath(strSpecialFolder)
    ' Returns the path to the specified special folder as a string
    ' Examples:
    '   strPath = fSpecialFolderPath("MyDocuments")
    '   strPath = fSpecialFolderPath("Desktop")
    '   strPath = fSpecialFolderPath("AppData")  'user-specific, roaming tree
    ' For a list of special folders, run ListSpecialFolders()

    On Error Resume Next

    Dim oWSHShell
    Dim strErrMsg
    Dim strPath

    Set oWSHShell = CreateObject("WScript.Shell")
    strPath = oWSHShell.SpecialFolders(strSpecialFolder)

    If Err.Number = 0 Then
        'Found it
        'MsgBox "Entering Err.Number = 0 condition"
        funSpecialFolderPath = strPath
        Set oWSHShell = Nothing
    Else
        'Found it not
        strErrMsg = "ERROR in fSpecialFolderPath:"
        'MsgBox "Entering Err.Number <> 0 condition" & vbcr & strErrMsg
        Subs_Util.subReportErrMsg(strErrMsg)
    End If

End Function
