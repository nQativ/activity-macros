' -- Create Offer pdf from Word template. Email & attach to Note
dim file_path, word_document, pdf_filename,tmp_folder ' As Strings
dim fso, wsh ' As Objects

' -- Find the location of the Temp directory for this user
if IsEmpty(tmp_folder) then
   set wsh = CreateObject("WScript.Shell")
   tmp_folder = wsh.ExpandEnvironmentStrings("%Temp%")
end if
' --- Use %TMP% to locate output Temp folder for the current user

' -- Set reference for Activity Automation template file (Word doc with Content Control objects)
file_path = "\\compu-share\data\CSMS\ActivityMacros"
document_name = "AA-Employee_HealthInsOffer"

word_document = document_name & ".docm"
pdf_filename = tmp_folder & "\" & document_name  & ".pdf"
tmp_filename = tmp_folder & "\" & document_name  & ".txt"

' -- Create a Scripting.Dictionary to hold the values for the content controls.
Dim dict
set dict = CreateObject("Scripting.Dictionary")

'-ACA Record Data
set aca_data = RecordLoop.Data
aca_data.Edit

'-Validate ACA Record
if (aca_data.Fields("Type").Value <> "Result") then
  Err.Raise vbObjectError, "MacroSource", "The ACA Record Type must be Result."
end if
if IsNull(aca_data.Fields("NotifyNote").ValueInternal) then
  Err.Raise vbObjectError, "MacroSource", "The ACA Record does not have a Notify Note."
end if
if not IsNull(aca_data.Fields("NotifiedNote").ValueInternal) then
  Err.Raise vbObjectError, "MacroSource", "The ACA Record already has a Notified Note."
end if

RecordLoop.AddMessage "PREmployee GUID on this ACARecord" + aca_data.Fields("PREmployee").ValueInternal

Dim premployee 
Set premployee = company.Payroll.PREmployee

premployee.Locate aca_data.Fields("PREmployee").ValueInternal

' - Load the content control values.
dict.Add "ccEmployeeCode" , premployee.Code
dict.Add "ccEmployeeName", premployee.FirstNameFirst
dict.Add "ccEmployeeBirthDate", premployee.BirthDate
dict.Add "ccACAStabilityBegin", aca_data.Fields("StabilityPeriodBegin").Value
'dict.Add "ccACAStabilityEnd", aca_data.Fields("StabilityPeriodEnd").Value

' - Set the name of the file to create.
dict.Add "pdfFilename", pdf_filename

' -- Serialize the Scripting.Dictionary values into the file: tmp_filename.
if IsEmpty(fso) then
   set fso = CreateObject("Scripting.FileSystemObject")
end if
dim f, key
set f = fso.CreateTextFile(tmp_filename,true)
for each key in dict.Keys
    f.Writeline(key & ":" & dict.Item(key))
next
f.Close()

' -- Run Word and execute the "CreatePDF" macro.
if IsEmpty(wsh) then
	set wsh = CreateObject("WScript.Shell")
end if
wsh.Run "winword " & file_path & "\" & Word_document & " /q /mCreatePDF /cmd," & tmp_filename, 1, True

if (fso.FileExists(pdf_filename)) then

    sendToEmail = ""
    if (PREmployee.DistributionEmail = "H") then
        sendToEmail = PREmployee.HomeEmail
    elseif (PREmployee.DistributionEmail = "W") then
        sendToEmail = PREmployee.WorkEmail
    end if

    if (sendToEmail <> "") then
'-Email the pdf
    	RecordLoop.AddMessage "Emailing " + pdf_filename
        Set objEmail  = CreateObject("Outlook.Application")
        Set EmailItem = objEmail.CreateItem(olMailItem)

        With EmailItem
            .To   = sendToEmail
            .Subject = "ACA Health Coverage Offer"
            .Body = "Please complete attached option form and return to the Human Resources department."
            .Attachments.Add pdf_filename
            .Send
        End With
        note_desc = "Offer sent via Email to " + sendToEmail
    Else
'-Print the pdf
    	RecordLoop.AddMessage "Printing " + pdf_filename
        CreateObject("Shell.Application").Namespace(0).ParseName(pdf_filename).InvokeVerbEx("Print")
        note_desc = "Offer sent via Print"
    end if

'-Create a new Note
    note_type = "ACANotified"

    set note_data = Company.Payroll.PRNote
    note_data.New
    note_data.Fields("PRNoteType").Value = note_type
    note_data.Fields("NoteDate").Value = now
    note_data.Fields("Description").Value = note_desc
    note_data.Fields("References/Employees").Items.Add aca_data.Fields("PREmployee").ValueInternal
    note_data.Fields("References/Records").Items.Add aca_data.Fields("RecordNumber").Value
    note_data.Save

    note_data.Fields("Attachments").Items.Add pdf_filename

    RecordLoop.Results.Add note_data
    RecordLoop.AddMessage note_type + " Note created for " + premployee.FirstNameFirst
    RecordLoop.AddMessage "and attached " & document_name & " to this new note."

    '-Update the selected ACA Record
    aca_data.Fields("IsLocked").Value = False
    aca_data.Fields("NotifiedNote").ValueInternal = note_data.Fields("PRNote").Value
    aca_data.Fields("IsLocked").Value = True
    aca_data.Save

    '-Clear the Recall Date from the "NotifyNote" note
    note_data.Locate aca_data.Fields("NotifyNote").Value
    if (CStr(note_data.Fields("Number").Value) = aca_data.Fields("NotifyNote").Value) then 
        note_data.Edit
        note_data.Fields("RecallDate").Value = Null

        note_data.Save
    else
        RecordLoop.AddMessage "Did not update the Recall Date"
    end if

else
    RecordLoop.AddMessage "Error: The Offer document was not created."
End if