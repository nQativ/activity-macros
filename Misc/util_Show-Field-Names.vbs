' BEG 'Initialization' tab contents ------------------------------------------------------
'
' Provides list of fields available on the selected Activity entity
'
'
' Macro Type:      Record Loop
' Enable Results:  unchecked
' Record Type:     Generic (based on Run From)
' Run From:        wherever you like - press F3 to see choices
'                  Note: Can be run from multiple locations
' Using:           na
'

' Why to use:
' You want to know the names of the fields on the selected Activity entity (ex: Vendor)

' How to use:
' - Make the macro available for the entity you are interested in by adding that entity
'   to the 'Run From' field
' - Navigate to the folder containing that kind of Activity entity
' - Run the macro
' - Observe the list of field names in the pop-up window


' BEG 'Script' tab contents --------------------------------------------------------------
RecordLoop.Data.Fields.ShowFields

' BEG 'Finalization' tab contents --------------------------------------------------------
' empty