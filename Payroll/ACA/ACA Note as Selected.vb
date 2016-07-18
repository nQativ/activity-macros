'-ACA Record Data
set aca_data = RecordLoop.Data
aca_data.Edit

'-Validate
if (aca_data.Fields("Type").Value <> "Result") then
  Err.Raise vbObjectError, "MacroSource", "The ACA Record Type must be Result."
end if
if IsNull(aca_data.Fields("NotifyNote").ValueInternal) then
  Err.Raise vbObjectError, "MacroSource", "The ACA Record does not have a Notify Note."
else
  set note_data = Company.Payroll.PRNote
  note_data.Locate aca_data.Fields("NotifyNote").ValueInternal
  if (note_data.Fields("PRNoteType").Value <> "ACANotifyOffer") then
    Err.Raise vbObjectError, "MacroSource", "The ACA Record does not have an Offer Notify Note."
  end if
end if
if not IsNull(aca_data.Fields("ResponseNote").ValueInternal) then
  Err.Raise vbObjectError, "MacroSource", "The ACA Record already has a Response."
end if

'-Create a new PR Note record:
note_type = "ACASelected"
note_desc = "Selected plan"

set note_data = Company.Payroll.PRNote
note_data.New
note_data.Fields("PRNoteType").Value = note_type
note_data.Fields("NoteDate").Value = now
note_data.Fields("Description").Value = note_desc
note_data.Fields("References/Employees").Items.Add aca_data.Fields("PREmployee").ValueInternal
note_data.Fields("References/Records").Items.Add aca_data.Fields("PRACARecord").Value
note_data.Save

RecordLoop.Results.Add note_data
RecordLoop.AddMessage note_type + " Note created"

'-Update the selected ACA Record:
aca_data.Fields("IsLocked").Value = False
aca_data.Fields("ResponseNote").ValueInternal = note_data.Fields("PRNote").Value
aca_data.Fields("IsLocked").Value = True
aca_data.Save