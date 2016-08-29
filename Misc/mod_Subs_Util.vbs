'
' This module is a home for package- and application-non-specific code that may
' be broadly useful to others writing Activity macros
'
' Macro Type:               Module
' Using:                    Declarations
'
' Contents
' DisplayErrMsg             Displays as much detail as possible about errors
'
'-----------------------------------------------------------------------------------------

Sub subDisplayErrMsg(strMessage)
    ' Display custom message and information from VBScript Err Object

    On Error Resume Next

    Dim strError

    strError = VbCrLf & strMessage & vbCR & _
               "Number (dec) : " & Err.Number & vbCR & _
               "Number (hex) : &H" & Hex(Err.Number) & vbCR & _
               "Source       : " & Err.Source & vbCR & _
               "Description  : " & Err.Description
    Err.Clear
    MsgBox strError

End Sub
