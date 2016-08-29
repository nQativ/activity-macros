'
' This module is a home for package-non-specific code that may be used to gather
' or create information related to the Activity application
'
' Macro Type:               Module
' Using:                    Constants
'                           Declarations
'                           Functions_FileSystem
'
' Contents
' fActAppDataFold           creates Activity application data folder if it does not exist,
'                           returns path to it
'
'-----------------------------------------------------------------------------------------

Function funActAppDataFold()
    ' Create folder in which to store Activity-related files created
    ' by code in this module.

    On Error Resume Next

    Dim strAppDataPath
    Dim blnFileExists

    strAppDataPath = Functions_FileSystem.funSpecialFolderPath("AppData")
    strAppDataPath = strAppDataPath & Constants.bs & "nQativ"
    blnFileExists = Functions_FileSystem.funFileExists(strAppDataPath)
    If blnFileExists Is Not True Then
        'MsgBox "nQativ folder does not exist...creating"
        MkDir strAppDataPath
    End If
    strAppDataPath = strAppDataPath & Constants.bs & "Activity"
    blnFileExists = Functions_FileSystem.funFileExists(strAppDataPath)
    If blnFileExists Is Not True Then
        'MsgBox "Activity folder does not exist...creating"
        MkDir strAppDataPath
    End If
    funActAppDataFold = strAppDataPath

End Function
