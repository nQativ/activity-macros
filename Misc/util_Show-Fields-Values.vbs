' BEG 'Initialization' tab contents ------------------------------------------------------
'
' Returns values of all fields and detail fields available on the selected Activity entity
' Fails nicely when Detail does not exist for the selected item(s)
'
'
' Macro Type:      Record Loop
' Enable Results:  checked
' Shortcut:        Shift+Ctrl+F
' Record Type:     Generic (based on Run From)
' Run From:        wherever you like - press F3 to see choices
' Using:           blank
'

' Why to use:
' You want to know the names and values of the fields (main OR detail) on the selected
' Activity entity (ex: Vendor, Time Log Detail)

' How to use:
' - Make the macro available for the entity you are interested in by adding it to the
'   'Run From' field
' - Navigate to the folder containing that kind of Activity entity
' - Run the macro
' - Review the list of fields and values in the message window that appears after the
'   macro is finished running
'
' ----------------------------------------------------------------------------------------

On Error Resume Next

Const cl = ":"
Const sp = " "
Const sq = "'"
blnHasDetail = True  ' Set initial value; alter programmatically if no Detail exists

Sub sSepLine(intLength)
    ' Inserts a series of dashes of length intLength in the message window
    ' Makes it easier to distinguish between records when called at either end of a loop
    MacroProcess.AddMessage(String(intLength, "-"))
End Sub

Sub sErrMsg(strSection)
    ' Sends information about an error to the message window and clears the error so
    ' that processing can continue
    If Err.Number <> 0 Then
        MacroProcess.AddMessage("Error handler: " & strSection)
        MacroProcess.AddMessage("Error number: " & Err.Number)
        MacroProcess.AddMessage("Error source: " & Err.Source)
        MacroProcess.AddMessage("Error description: " & Err.Description)
        Err.Clear
        sSepLine(40)
    End If
End Sub

Set objFields = RecordLoop.Data.Fields
Set objDetailFields = RecordLoop.Data.Detail.Fields
If Err.Number <> 0 Then
    sErrMsg("Tried to access RecordLoop.Data.Detail.Fields")
    blnHasDetail = False
End If

' BEG 'Script' tab contents --------------------------------------------------------------
For Each varField in objFields
    MacroProcess.AddMessage(varField.Name & _
                            cl & sp & sq & _
                            varField.Value & _
                            sq)
Next

sSepLine(60)
'MacroProcess.AddMessage("blnHasDetail: " & blnHasDetail)
'sSepLine(60)

If blnHasDetail = False Then
    MacroProcess.AddMessage("Detail data does not exist for this entity.")
Else
    RecordLoop.Data.Detail.First
    Do Until RecordLoop.Data.Detail.IsEndOfFile
        For Each varDetailField in objDetailFields
            MacroProcess.AddMessage(varDetailField.Name & _
                                    cl & sp & sq & _
                                    varDetailField.Value & _
                                    sq)
        Next
        sSepLine(40)
        RecordLoop.Data.Detail.Next
    Loop
End If

' BEG 'Finalization' tab contents --------------------------------------------------------
Set objFields = Nothing
Set objDetailFields = Nothing
