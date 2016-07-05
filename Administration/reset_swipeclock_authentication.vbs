' 
' Resets the calculated Web Password to match the shared secret password in the self-serve configuration.
' Note: cSharedSecret must match the shared secret in the settings.php file.
'
' Macro Type:     RecordLoop
' Record Type:    Authorized Users (Activity Company)
' Enable Results: unchecked
'

const cSharedSecret = "Shar3dS3cr3t"

RecordLoop.Data.Edit
RecordLoop.Data.Fields("WebPasswordText").Value = LCase(RecordLoop.Data.Fields("WebUsername").Value) & ":" & "!" & cSharedSecret
RecordLoop.Data.Fields("EnforcePasswordComplexity").Value = False
RecordLoop.Data.Fields("EnforcePasswordExpiration").Value = False
RecordLoop.Data.Save
RecordLoop.AddMessage "Reset the Activity Self-Serve for SwipeClock username: " & RecordLoop.Data.Fields("WebUsername").Value