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
' entity (ex: Vendor, Time Log Detail)

' How to use:
' - Make the macro available for the entity you are interested in by adding it to the
'   'Run From' field
' - Navigate to the folder containing that kind of Activity entity
' - Run the macro
' - Observe the list of field names in the pop-up window
'
' ----------------------------------------------------------------------------------------

On Error Resume Next

RecordLoop.Data.Fields.ShowFields
If Err.Number = 0 Then
    MacroProcess.AddMessage("Displayed fields.")
Else
    MacroProcess.AddMessage("Did not display fields.")
    Err.Clear
End If

RecordLoop.Data.Detail.Fields.Showfields
If Err.Number = 0 Then
    MacroProcess.AddMessage("Displayed detail fields.")
Else
    MacroProcess.AddMessage("Did not display detail fields (probably because they did not exist for the chosen folder.")
    Err.Clear
End If

' BEG 'Script' tab contents --------------------------------------------------------------
' See the 'Initialization' tab
'
' Don't put the RecordLoop statement here or it will pop the field-listing
' window as many times as there are records in the Record Loop

' BEG 'Finalization' tab contents --------------------------------------------------------
' empty
