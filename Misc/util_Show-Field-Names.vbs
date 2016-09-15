' BEG 'Initialization' tab contents ------------------------------------------------------
'
' Provides list of fields available on the selected Activity entity
'
'
' Macro Type:      Record Loop
' Enable Results:  checked
' Record Type:     Generic (based on Run From)
' Run From:        wherever you like - press F3 to see choices
' Using:           blank
'

' Why to use:
' You want to know the names of the fields (main OR detail) on the selected Activity
' entity (ex: Vendors, Time Log Detail)

' How to use:
' - Make the macro available for the entity you are interested in by adding it to the
'   'Run From' field
' - Navigate to the folder containing that kind of Activity entity
' - Run the macro
' - Observe the list of field names in the pop-up window
'
' ----------------------------------------------------------------------------------------

On Error Resume Next

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

For Each varField in objFields
    MacroProcess.AddMessage(varField.Name)
Next

sSepLine(60)
'MacroProcess.AddMessage("blnHasDetail: " & blnHasDetail)
'sSepLine(60)

If blnHasDetail = False Then
    MacroProcess.AddMessage("Detail data does not exist for this entity.")
Else
    For Each varDetailField in objDetailFields
        MacroProcess.AddMessage(varDetailField.Name)
    Next
End If

' BEG 'Script' tab contents --------------------------------------------------------------
' See the 'Initialization' tab
'
' For this particular macro, I put the 'ShowFields' statements on the Initialization tab
' because I really only want to see the field names once. If those statements were on this
' tab, they would be executed for every selected record.
'
' This isn't such a big deal when running it from the folder--just be careful what you
' choose. But when you run it from the Macro form using F9, the default behavior appears
' to be to select ALL items in the chosen 'Run From' folder. By moving the 'ShowFields'
' statements to the Initialization tab, at least the values will only pop up once, even
' if the process still loops through every item in the folder anyway.
'
' So why not just use a General macro? Because I would have to know every possible folder
' on which I might want to run the macro and build that into the macro. By using a
' RecordLoop macro, I can write generic code that will run anywhere it is legal to use a
' RecordLoop macro.

' BEG 'Finalization' tab contents --------------------------------------------------------
Set objFields = Nothing
Set objDetailFields = Nothing
