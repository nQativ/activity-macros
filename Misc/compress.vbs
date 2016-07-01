'Script to reduce the size of pdf attachments
'   A GhostScript command is executed on each PDF. The output of the command ends up
'   replacing the original pdf attachment. The original attachment is either renamed
'   or removed based on settings below.

'Requirements:
'   GhostScript 9.19 - https://github.com/ArtifexSoftware/ghostpdl-downloads/releases/download/gs919/gs919w32.exe

'Change the following variables as needed
'---------------------------------------------------------------------------------------------------------------
' createBackup: 
'   true  - save the original as "<backupFilenamePrefix>-<filename>"
'   false - do not keep the original file
const createBackup = true

'backupFilenamePrefix
'   if createBackup is set to true, backup filenames will be prefixed with the value of the variable
const backupFilenamePrefix = "backup"

' pdfQuality:
'   "/screen"  - selects low-resolution output similar to the Acrobat Distiller "Screen Optimized" setting.
'   "/ebook"   - selects medium-resolution output similar to the Acrobat Distiller "eBook" setting. (recommended)
'   "/printer" - selects output similar to the Acrobat Distiller "Print Optimized" setting.
const pdfQuality = "/ebook"
'---------------------------------------------------------------------------------------------------------------

dim data, fso
set data = RecordLoop.Data
set fso = CreateObject("Scripting.FileSystemObject")

count = data.Attachments.Count
if count > 0 then
    For i=1 to data.Attachments.Count
        attachment = LCase(data.Attachments.item(i))
        if fso.GetExtensionName(attachment) = "pdf" then
            'save out to a temp file
            temp = fso.GetAbsolutePathName(data.Attachments.SaveAsTempFile(i))
            success = ProcessAttachment(temp, attachment, i)
            if success then 
                RecordLoop.AddMessage "Processed: " & attachment
            else
                RecordLoop.AddMessage "Unable to process " & attachment
            end if
        end if
    Next
end if

'CreateCompressedFile - Process the input pdf with GhostScript
'   input  = full path of the file that will be used for compression
'   output = full path and name of output file
function CreateCompressedFile(input, output)
    Dim objShell
    Set objShell = CreateObject("WScript.Shell") 
    objShell.run "cmd /c", 0 , True
    objShell.run """C:\Program Files (x86)\gs\gs9.19\bin\gswin32c.exe""" & _
        " -sDEVICE=pdfwrite" & _
        " -dCompatibilityLevel=1.4" & _
        " -dPDFSETTINGS=" & pdfQuality & _
        " -dNOPAUSE" & _
        " -dQUIET" & _
        " -dDetectDuplicateImages" & _ 
        " -dCompressFonts=true" & _
        " -r150" & _ 
        " -dBATCH" & _
        " -sOutputFile=""" & output & """ """ & input & """", 0, True
    CreateCompressedFile = fso.FileExists(output)
end function

'ProcessAttachment
'   input          = full path of the file that will be used for compression
'   attachmentName = name of the attachment being processed
'   attachment     = index of the attachment being processed
function ProcessAttachment(input, attachmentName, attachment)
    output = fso.BuildPath(fso.GetParentFolderName(input), attachmentName)
'   try to get a compressed version of our input pdf  
    compressed = CreateCompressedFile(input, output)
    if compressed then
        if createBackup then
            data.Attachments.Rename data.Attachments.Add(input), backupFilenamePrefix & "-" & attachmentName
        end if
        data.Attachments.Replace attachment, output
    end if
    ProcessAttachment = compressed
end function