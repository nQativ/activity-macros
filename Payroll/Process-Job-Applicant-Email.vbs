'
' Macro Type:                 General
' Enable Results:             unchecked
' Run From:                   Employees (Payroll), Macros (Activity Company)
' Using:                      Constants
'                             Declarations
'                             Functions_ActivityInfo
'                             Functions_FileSystem
'                             Functions_Util
'                             Subs_FileSystem
'                             Subs_Util
'
' The purpose of this macro is to provide an automated way to take job-applicant
' materials from Outlook and put them in Activity. The current version of this macro
' performs the following actions:
'   - Creates new Employee record, setting the Status to "Applicant"
'   - Creates new Note containing the text of their e-mail message
'   - Attaches files from e-mail to the new Note after standardizing the filenames
'
' TODO Verify that each attachment is actually useful for our purposes
'      Some attachments are hidden or embedded, and we probably do not want them. It is
'      possible to distinguish between these and other 'useful' attachments, but the
'      script currently does not.
'
'-----------------------------------------------------------------------------------------

'BEG code needed outside of Activity environment
'Dim Activity
'Dim Company

'Also need active reference to automation library (MS Office VBA environment, etc.)
'Set Activity = New Activity
'Activity.ServerAddress = "appserver3"
'Activity.Connect

'Set Company = Activity.Companies("Demo CNX")
'Company.Connect
'END code needed outside of Activity environment

On Error Resume Next

' Dimension variables
Dim objFSO                     ' File System object
Dim objOutlookApp              ' Outlook Application object
Dim objOutlookSel              ' Outlook Selection object
Dim objOutlookMsg              ' Outlook Message object
Dim objOutlookAttachments      ' Outlook Message Attachments object
Dim objDivisionFind            ' Activity
Dim objActivityDivisions       ' Activity Divisions (Custom Data)
Dim objActivityPositions       ' Activity PR Positions
Dim objActivityEE              ' Activity PR Employee
Dim objActivityPRNote          ' Activity PR Note

Dim i
Dim n

Dim strSection                 ' String to use to identify each section (debug) (string)
Dim strMessage                 ' Message to use in conjunction with 'DisplayErrMsg' (string)
Dim strAppDataFolder           ' Path to Application Data folder (string)
Dim datRDateTime               ' Outlook message received DateTime (DateTime)
Dim strNameFirst               ' Applicant's first name (string)
Dim strNameLast                ' Applicant's last name (string)
Dim strDivision                ' Division (string)
Dim strPositionPar             ' Parameter used in selection of Position (string)
Dim strPosition                ' Position code (string)
Dim strEECode                  ' Employee Code (string)
Dim strPRNoteTypePar           ' Parameter used in selection of Note Type (string)
Dim strPRNoteType              ' PR Note Type (string)
Dim strPRNoteDesc              ' PR Note Description (string)
Dim lngAttachCount             ' Count of Outlook attachment(s) (long)
Dim strAttachPrompt            ' Attachment Prompt (string)
Dim strAttachThis              ' AttachThis string (Y or N) (string)
Dim strAttachFilenameOrig      ' Original filename of attachment (string)
Dim strAttachType              ' Attachment Type selected via Find (string)
Dim intPeriod                  ' Location of '.' in strAttachFilenameOrig (integer)
Dim strAttachFilenameOrigExt   ' Original filename extension of attachment (string)
Dim strAttachFilenameNew       ' New filename for attachment (string)
Dim strAttachFullPathNew       ' Full path to newly renamed file to be attached to Activity Note (string)

strSection = "Create FSO"
Subs_Util.subMacroAddMsg "BEG", strSection
' Must write Outlook attachment(s) to disk in order to rename them, so need file system object

    Set objFSO = CreateObject("Scripting.FileSystemObject")

strSection = "Locate/Create AppData Folder"
Subs_Util.subMacroAddMsg "BEG", strSection
' Locate (and create, if necessary) place to put temporary files

    strAppDataFolder = Functions_ActivityInfo.funActAppDataFold()

    MacroProcess.AddMessage("strAppDataFolder: " & strAppDataFolder)

strSection = "Prepare to use Outlook information"
Subs_Util.subMacroAddMsg "BEG", strSection

    Set objOutlookApp = CreateObject("Outlook.Application")
    Set objOutlookSel = objOutlookApp.ActiveExplorer.Selection

    MacroProcess.AddMessage("The number of selected Outlook items is " & objOutlookSel.Count)

strSection = "Process selected Outlook item(s)"
Subs_Util.subMacroAddMsg "BEG", strSection

    ' Verify that at least one Outlook item is selected
    If objOutlookSel.Count = 0 Then
        MsgBox("Exiting -- no Outlook items selected. Select one or more relevant items and try again.")
        WScript.Quit  'This does not work but I don't know why
    End If

strSection = "For-Each Loop: Collect information about selected messages"
Subs_Util.subMacroAddMsg "BEG", strSection

    For Each objOutlookMsg In objOutlookSel

        If Not objOutlookMsg.MessageClass = "IPM.Note" Then
            MsgBox ("Skipping item because it is not an Outlook e-mail item")
        Else

            strSection = "Get received date & time from Outlook message"
            Subs_Util.subMacroAddMsg "BEG", strSection

                datRDateTime = objOutlookMsg.ReceivedTime
                MacroProcess.AddMessage("datRDateTime: " & datRDateTime)
                MacroProcess.AddMessage("Formatted datRDateTime: " & _
                                        Functions_Util.funDateTimeStamp(datRDateTime))

            strSection = "Prompt user for applicant names"
            Subs_Util.subMacroAddMsg "BEG", strSection

                strNameFirst = Trim( _
                    InputBox("Enter the applicant's first name:", _
                             "First Name", _
                             ""))
                MacroProcess.AddMessage("strNameFirst: " & strNameFirst)

                strNameLast = Trim( _
                    InputBox("Enter the applicant's last name:", _
                             "Last Name", _
                             ""))
                MacroProcess.AddMessage("strNameLast: " & strNameLast)

            strSection = "Prompt user to select Activity information"
            Subs_Util.subMacroAddMsg "BEG", strSection

                ' Prompt user for Division to which applicant is applying
                strDivision = Company.FindCode("Administration", _
                                               "Custom Data - Divisions")
                MacroProcess.AddMessage("strDivision: " & strDivision)

                ' Prompt user for Position for which applicant is applying
                ' Filter Positions by Division
                ' NOTE: There is a potential problem with filtering by Division because
                ' there is a set of Positions with a 'CSI' prefix but there is no 'CSI'
                ' Division.
                strPositionPar = _
                    "<p>" + _
                    "<Lookup Value='" + strDivision + "*' />" + _
                    "</p>"
                strPosition = _
                    Company.FindCode("Human Resources", _
                                     "Positions", _
                                     strPositionPar)
                MacroProcess.AddMessage("strPosition: " & strPosition)

                ' Remove the Division & dash from Position code
                strPosition = Replace(strPosition, strDivision & Constants.da, "")
                MacroProcess.AddMessage("strPosition after Replace: " & strPosition)

            strSection = "Prepare to work with Activity Employees"
            Subs_Util.subMacroAddMsg "BEG", strSection

                Set objActivityEE = Company.Payroll.PREmployee
                MsgBox "Activity will check to see if an Employee record already exists for this applicant." & vbCr & vbCr & _
                       "If one or more possible matches exist, the next window you see will list them. Select one (if appropriate) and click the 'OK' button to continue." & vbCr & vbCr & _
                       "If none of the listed records are a match, click the 'Cancel' button to continue. A new Employee record will be created." & vbCr & vbCr & _
                       "Similarly, if no records are listed, click the 'Cancel' button to continue. A new Employee record will be created.",, _
                       "-------------------- N O T I C E --------------------"

            strSection = "Try to determine if this employee already exists"
            Subs_Util.subMacroAddMsg "BEG", strSection

                strEEParam = _
                    "<p>" + _
                    "<Filter Name='? Employee Name' Type='Built-In'>" + _
                    "<Parameter Name='Name_First_Name_First' Value='" + _
                    strNameFirst + Constants.sp + strNameLast + "'/>" + _
                    "</Filter>" + _
                    "</p>"
                strEECode = _
                    Company.FindCode("Payroll", _
                                     "Employees", _
                                     strEEParam)
                MacroProcess.AddMessage("strEECode: " & strEECode)

                If strEECode = "" Then
                    ' No existing employee, so...
                    ' Create the Employee Code
                    strEECode = UCase( _
                        Left(strNameFirst, 1) & _
                        Left(strNameLast, 1)) & _
                        Functions_Util.funRandomNumber(0, 9999, 4, True)
                    MacroProcess.AddMessage("strEECode inside 'Create' condition: " & _
                                            strEECode)

                    ' Create the Activity Employee
                    objActivityEE.New
                    objActivityEE.FirstName = strNameFirst
                    objActivityEE.LastName = strNameLast
                    objActivityEE.Code = strEECode
                    objActivityEE.Save
                    MacroProcess.AddMessage("PREmployee.Value: " & _
                                            objActivityEE.Fields("PREmployee").Value)
                Else
                    ' Existing employee, so...
                    ' Locate the existing employee record (already has code, names)
                    objActivityEE.Locate(strEECode)
                End If

            strSection = "Set the Activity Employee's Status"
            Subs_Util.subMacroAddMsg "BEG", strSection

                Set objActivityEEStatuses = objActivityEE.EmployeeStatuses

                objActivityEE.Edit
                objActivityEEStatuses.Insert
                ' Note: DateTimes are stored with date as integer and time as decimal
                '       To get just the Date part, extract the integer, like so:
                objActivityEEStatuses.EffectiveDate = Int(datRDateTime)
                ' Hard-code Status but trap the error if the code does not exist, like so:
                objActivityEEStatuses.PRStatusCode = "Applicant"
                If Err.Number <> 0 Then
                    objActivityEEStatuses.PRStatusCode = _
                        Company.FindCode("Payroll", _
                                         "Statuses")
                    Err.Clear
                End If
                objActivityEEStatuses.Post
                objActivityEE.Save

                ' Add current EE record to Results
                MacroProcess.Results.Add objActivityEE

            strSection = "Create and attach Note to Employee"
            Subs_Util.subMacroAddMsg "BEG", strSection

                  strPRNoteTypePar = _
                    "<p>" + _
                    "<Lookup Value='*App*' />" + _
                    "</p>"
                strPRNoteType = _
                    Company.FindCode("Payroll", _
                                     "Note Types", _
                                     strPRNoteTypePar)
                MacroProcess.AddMessage("strPRNoteType: " & strPRNoteType)

                strPRNoteDesc = "Job application correspondence"

                MacroProcess.AddMessage("strPRNoteDesc: " & strPRNoteDesc)

                Set objActivityPRNote = Company.Payroll.PRNote
                objActivityPRNote.New
                objActivityPRNote.Fields("PRNoteType").Value = strPRNoteType
                objActivityPRNote.Fields("NoteDate").Value = datRDateTime
                objActivityPRNote.Fields("RecallDate").Value = datRDateTime + 7
                objActivityPRNote.Fields("Description").Value = strPRNoteDesc
                objActivityPRNote.Fields("NoteText").Value = objOutlookMsg.Body
                objActivityPRNote.Fields("References/Employees").Items.Add _
                                         objActivityEE.Fields("PREmployee").Value
                objActivityPRNote.Save

            strSection = "Count Outlook attachments"
            Subs_Util.subMacroAddMsg "BEG", strSection

                Set objOutlookAttachments = objOutlookMsg.Attachments
                lngAttachCount = objOutlookAttachments.Count
                MacroProcess.AddMessage("Number of attachments: " & lngAttachCount)

                If lngAttachCount > 0 Then
                    ' We need to use a count down loop for removing items from a
                    ' collection. Otherwise, the loop counter gets confused and only every
                    ' other item is processed.

                    strSection = "For-Next Loop: Outlook attachments"
                    Subs_Util.subMacroAddMsg "BEG", strSection

                        For n = lngAttachCount To 1 Step -1
                            ' TODO Verify that this is a useful attachment (not hidden/embedded)
                            ' If valid then:

                            strAttachFilenameOrig = _
                                objOutlookAttachments.Item(n).FileName
                            MacroProcess.AddMessage("strAttachFilenameOrig: " & _
                                                    strAttachFilenameOrig)

                            strAttachPrompt = _
                                "Current e-mail attachment: " & vbCr & _
                                strAttachFilenameOrig & vbCr & vbCr & _
                                "Do you want to attach this file? (Y or N)?"
                            strAttachThis = UCase( _
                                InputBox(strAttachPrompt, _
                                         "Attach This File?", _
                                         "Y"))

                            If strAttachThis = "Y" Then
                                strSection = "If-Then: strAttachThis = Y"
                                Subs_Util.subMacroAddMsg "BEG", strSection

                                    ' Set Attachment Type based on Custom Data
                                    strAttachTypePar = _
                                        "<p>" + _
                                        "<Lookup Value='Application*' />" + _
                                        "</p>"
                                    strAttachType = _
                                        Company.FindCode( _
                                            "Administration", _
                                            "Custom Data - Attachment Types", _
                                            strAttachTypePar)
                                    strAttachType = _
                                        Replace(strAttachType, "Application-", "")
                                    MacroProcess.AddMessage("strAttachType: " & _
                                                            strAttachType)

                                    ' Get the filename extension of the attachment
                                    intPeriod = InStr(strAttachFilenameOrig, ".")
                                    If intPeriod <> 0 Then
                                        strAttachFilenameOrigExt = _
                                            Mid(strAttachFilenameOrig, intPeriod)
                                        MacroProcess.AddMessage(_
                                            "strAttachFilenameOrigExt: " & _
                                            strAttachFilenameOrigExt)
                                    Else
                                        strAttachFilenameOrigExt = ""
                                        MacroProcess.AddMessage( _
                                            "strAttachFilenameOrigExt: " & _
                                            strAttachFilenameOrigExt)
                                    End If

                                    ' Create a new filename based on information we've gathered
                                    strAttachFilenameNew = _
                                        strNameLast & _
                                        strNameFirst & _
                                        Constants.us & _
                                        Trim(strDivision) & _
                                        Constants.us & _
                                        Trim(strPosition) & _
                                        Constants.us & _
                                        Trim(strAttachType) & _
                                        Constants.us & _
                                        Functions_Util.funDateTimeStamp(datRDateTime) & _
                                        strAttachFilenameOrigExt
                                    MacroProcess.AddMessage("strAttachFilenameNew (before): " & _
                                                            strAttachFilenameNew)

                                    strAttachFilenameNew = InputBox( _
                                        "Accept the generated filename or edit as necessary. Then press 'Enter'", _
                                        "Attachment Filename", _
                                        strAttachFilenameNew)
                                    MacroProcess.AddMessage("strAttachFilenameNew (after): " & _
                                                            strAttachFilenameNew)

                                    ' Write attachment to disk using new filename
                                    strAttachFullPathNew = _
                                        strAppDataFolder & _
                                        Constants.bs & _
                                        strAttachFilenameNew
                                    objOutlookAttachments.Item(n).SaveAsFile _
                                        strAttachFullPathNew
                                    MacroProcess.AddMessage("strAttachFullPathNew: " & _
                                                            strAttachFullPathNew)

                                    ' Attach to Note and save Note
                                    objActivityPRNote.Fields("Attachments").Items.Add _
                                        strAttachFullPathNew
                                    objActivityPRNote.Save

                                    ' Delete file from temporary folder
                                    objFSO.DeleteFile strAttachFullPathNew
                                    'Else
                                    'Skip this attachment
                            End If
                            'Else
                                'skip this attachment
                            'End If
                        Next
                End If
        End If
    Next

' Kill other objects
    Set objActivityPRNote = Nothing
    Set objActivityEE = Nothing
    Set objActivityPositions = Nothing
    Set objActivityDivisions = Nothing
    Set objOutlookAttachments = Nothing
    Set objOutlookMsg = Nothing
    Set objOutlookSel = Nothing
    Set objOutlookApp = Nothing
    Set objFSO = Nothing
