'
' This module is a home for package- and application-non-specific code that may
' be used to perform file-system actions.
'
' Macro Type:               Module
' Using:                    Declarations
'
' Contents
' ListSpecialFolders        Lists Windows special folders in a MsgBox
'
'-----------------------------------------------------------------------------------------

Sub subListSpecialFolders()
  ' Lists Windows special folders in a MsgBox box

    Dim objWSHShell

    Set objWSHShell = CreateObject("WScript.Shell")
    strPaths = ""

    For Each Item In oWSHShell.SpecialFolders
        If strPaths = "" Then
            strPaths = Item
        Else
            strPaths = strPaths & vbCr & Item
        End If
        'Wscript.Echo Item
    Next

    MsgBox (strPaths)

    Set objWSHShell = Nothing

End Sub
