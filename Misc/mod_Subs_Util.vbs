'
' This module is a home for package- and application-non-specific code that may
' be broadly useful to others writing Activity macros
'
' Macro Type:               Module
' Using:                    Declarations
'
' Contents
' subDisplayErrMsg             Displays as much detail as possible about errors
' subMacroAddMsg               Constructs and adds a new message using MacroProcess
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

Sub subMacroAddMsg(strBegEnd, strSection)
    ' Constructs and adds a new Macro message using MacroProcess.AddMessage

    Select Case UCase(strBegEnd)
        Case "BEG"
            MacroProcess.AddMessage("Begin section " & Constants.ap & strSection & Constants.ap)
        Case "END"
            MacroProcess.AddMessage("End section " & Constants.ap & strSection & Constants.ap)
        Case Else
            MacroProcess.AddMessage("Current section is " & Constants.ap & strSection & Constants.ap)
    End Select

End Sub
