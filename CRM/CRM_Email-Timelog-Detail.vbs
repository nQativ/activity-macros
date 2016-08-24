' BEG 'Initialization' tab contents ------------------------------------------------------
'
' Constructs a Microsoft Outlook e-mail message containing CRM Timelog information.
'
'
' Macro Type:      Record Loop
' Enable Results:  unchecked
' Record Type:     Specific - Time Log Detail (CRM)
' Run From:        Time Log Detail (CRM)
' Using:           Constants
'                  Note: Could eliminate reliance on module 'Constants' by replacing
'                  references to 'Constants.sp' with '" "' (ignore single quotes)
'

' Why to use:
' You have entered information in several timelog entries and need to provide
' the information to someone else. You could send them a message telling them
' to go read your Timelog Details (probably not) OR you could re-type a summary
' of the information while referring back to your notes (scattered across
' timelogs - not ideal) OR you could copy and paste from each timelog entry
' into an email message and send it as-is or use the information to create a
' summary OR...you could run this macro that basically automates the copy/paste
' operation while also providing information that can help you find those
' timelog details again, if necessary.

' How to use:
' - Select one or more Timelog Detail records
' - Run the macro
' - Review the generated Outlook e-mail message
' - Click Send

' Uncommenting either of the following 'RecordLoop' lines will provide field
' names that could be used in variations of this script
'RecordLoop.Data.Fields.ShowFields          'Shows fields on TimeLog
'RecordLoop.Data.Detail.Fields.ShowFields   'Shows fields on TimeLogDetail

'Initialize variable that will contain body of e-mail message
sBody = ""

'Instantiate Outlook objects
Set oOL  = CreateObject("Outlook.Application")
Set oMsg = oOL.CreateItem(olMailItem)

' BEG 'Script' tab contents --------------------------------------------------------------
Set oTL = RecordLoop.Data          ' TimeLog data
Set oTLD = RecordLoop.Data.Detail  ' TimeLogDetail data

sDetHead = "Time Log Entry: " & _
           oTL.Fields("CMPersonnel").Value & _
           Constants.sp & _
           oTL.Fields("TimeLogDate").Value & _
           Constants.sp & _
           oTLD.Fields("StartTime").Value
sDetDesc = oTLD.Fields("Description").Value

If sBody = "" Then
  sBody = sDetHead & vbCR & sDetDesc & vbCR
Else
  sBody = sBody & vbCR & "----------" & vbCR & vbCR & sDetHead & vbCR & sDetDesc & vbCR
End If

' BEG 'Finalization' tab contents --------------------------------------------------------
'- Provide a default Subject (can be edited when message is displayed later)
sSubj = Left(oTLD.Fields("Description").Value, 80)
'Previously allowed default Subject to be edited via InputBox (now in Outlook msg)
'sSubj = _
'  InputBox( _
'    "Enter subject of e-mail message:", _
'    "E-Mail Subject", _
'    sSubj)

'- Create an Outlook e-mail message and display it so the user can address, edit
With oMsg
  .Subject = sSubj
  .Body = sBody
  .Save
End With
oMsg.Display

Set oOL  = Nothing
Set oMsg = Nothing
