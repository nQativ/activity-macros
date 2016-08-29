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
' TODO Implement 'Enable Results' feature that is now available for General macros
'      Two potential applications:
'      1) Write information there for use in debugging (instead of using MsgBox, etc)
'      2) Want to show user what the macro did
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

' Create FSO
' Must write Outlook attachment(s) to disk in order to rename them, so need file system object
    strSection = "Create FSO"
    Set objFSO = CreateObject("Scripting.FileSystemObject")

' Locate (and create, if necessary) place to put temporary files
    strSection = "Locate/Create AppData Folder"
    strAppDataFolder = Functions_ActivityInfo.funActAppDataFold()
    'MsgBox "strAppDataFolder: " & strAppDataFolder

' Prepare to use Outlook data
    Set objOutlookApp = CreateObject("Outlook.Application")
    Set objOutlookSel = objOutlookApp.ActiveExplorer.Selection
    'MsgBox "The number of selected Outlook items is " & objOutlookSel.Count

' Process the selected Outlook item(s)
  ' Verify that at least one Outlook item is selected
    If objOutlookSel.Count = 0 Then
        MsgBox("Exiting -- no Outlook items selected. Select one or more relevant items and try again.")
        WScript.Quit  'This does not seem to work
    End If

' Collect information about selected messages
    For Each objOutlookMsg In objOutlookSel
        If Not objOutlookMsg.MessageClass = "IPM.Note" Then
            MsgBox ("Skipping item because it is not an Outlook e-mail item")
        Else
            ' Get the received date & time of the message
            datRDateTime = objOutlookMsg.ReceivedTime
            'MsgBox "datRDateTime: " & datRDateTime
            'MsgBox "Formatted datRDateTime: " & Functions_Util.funDateTimeStamp(datRDateTime)

            ' Prompt user for the applicant's first name
            strNameFirst = _
                         Trim( _
                              InputBox("Enter the applicant's first name:", _
                                       "First Name", _
                                       ""))
            'MsgBox "strNameFirst: " & strNameFirst

            ' Prompt user for the applicant's last name
            strNameLast = _
                        Trim( _
                             InputBox("Enter the applicant's last name:", _
                                      "Last Name", _
                                      ""))
            'MsgBox "strNameLast: " & strNameLast

            ' Prompt user for Division to which applicant is applying
            strDivision = Company.FindCode("Administration", _
                                           "Custom Data - Divisions")
            'MsgBox "strDivision: " & strDivision

            ' Prompt user for Position for which applicant is applying
            ' Filter Positions by Division
            ' NOTE: There is a potential problem with filtering by Division because there
            '       is a set of Positions with a 'CSI' prefix but there is no 'CSI'
            '       Division.
            strPositionPar = "<p>" + _
                             "<Lookup Value='" + strDivision + "*' />" + _
                             "</p>"
            strPosition = Company.FindCode("Human Resources", _
                                           "Positions", _
                                           strPositionPar)
            'MsgBox "strPosition: " & strPosition

            ' Remove the Division & dash from Position code
            strPosition = Replace(strPosition, strDivision & Constants.da, "")
            'MsgBox "strPosition after Replace: " & strPosition

            ' Prepare to work with Activity Employees
            Set objActivityEE = Company.Payroll.PREmployee
            MsgBox "-------------------- N O T I C E --------------------" & vbCR & vbCR & _
		   "Activity will check to see if an Employee record already exists for this applicant." & vbCr & vbCr & _
		   "If one or more possible matches exist, the next window you see will list them. Select one (if appropriate) and click the 'OK' button to continue." & vbCr & vbCr & _
		   "If none of the listed records are a match, click the 'Cancel' button to continue. A new Employee record will be created." & vbCr & vbCr & _
		   "Similarly, if no records are listed, click the 'Cancel' button to continue. A new Employee record will be created."

            ' Try to determine if this employee already exists
            strEEParam = "<p>" + _
                         "<Filter Name='? Employee Name' Type='Built-In'>" + _
                         "<Parameter Name='Name_First_Name_First' Value='" + strNameFirst + Constants.sp + strNameLast + "'/>" + _
                         "</Filter>" + _
                         "</p>"
            strEECode = Company.FindCode("Payroll", _
                                         "Employees", _
                                         strEEParam)
            'MsgBox "strEECode: " & strEECode

            If strEECode = "" Then
                ' No existing employee, so...
                ' Create the Employee Code
                strEECode = _
                          UCase( _
                                Left(strNameFirst, 1) & _
                                Left(strNameLast, 1)) & _
                          Functions_Util.funRandomNumber(0, 9999, 4, True)
                'MsgBox "strEECode inside 'Create' condition: " & strEECode

                ' Create the Activity Employee
                objActivityEE.New
                objActivityEE.FirstName = strNameFirst
                objActivityEE.LastName = strNameLast
                objActivityEE.Code = strEECode
                objActivityEE.Save
                'MsgBox "PREmployee.Value: " & objActivityEE.Fields("PREmployee").Value
            Else
                ' Existing employee, so...
                ' Need to locate the existing employee record (already has code, names)
                objActivityEE.Locate(strEECode)
            End If

            ' Set the Activity Employee's Status
            Set objActivityEEStatuses = objActivityEE.EmployeeStatuses
            'MsgBox "Count: " & objActivityEEStatuses.Count
            objActivityEE.Edit
            objActivityEEStatuses.Insert
            ' Note: DateTimes are stored with date as integer and time as decimal
            '       To get just the Date part, extract the integer, like so:
            objActivityEEStatuses.EffectiveDate = Int(datRDateTime)
            ' Hard-code Status but trap the error if the code does not exist, like so:
            objActivityEEStatuses.PRStatusCode = "Applicant"
            If Err.Number <> 0 Then
                objActivityEEStatuses.PRStatusCode = Company.FindCode("Payroll", _
                                                                      "Statuses")
                Err.Clear
            End If
            objActivityEEStatuses.Post
            objActivityEE.Save


            ' Set the PR Note Type
            strPRNoteTypePar = "<p>" + _
                               "<Lookup Value='*App*' />" + _
                               "</p>"
            strPRNoteType = Company.FindCode("Payroll", _
                                             "Note Types", _
                                             strPRNoteTypePar)
            'MsgBox "strPRNoteType: " & strPRNoteType

            ' Set the PR Note Description
            strPRNoteDesc = "Job application correspondence"


            ' Create an Activity Note and attach it to the Activity Employee
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

            ' Count attachments and loop through them if there are any
            Set objOutlookAttachments = objOutlookMsg.Attachments
            lngAttachCount = objOutlookAttachments.Count
            'MsgBox "Number of attachments: " & lngAttachCount

            If lngAttachCount > 0 Then
                ' We need to use a count down loop for removing items from a collection.
                ' Otherwise, the loop counter gets confused and only every other item is
                ' processed.
                For n = lngAttachCount To 1 Step -1
                    ' TODO Verify that this is a useful attachment (not hidden/embedded)
                    ' If valid then:
                    ' Get the filename of the attachment
                    strAttachFilenameOrig = objOutlookAttachments.Item(n).FileName
                    'WScript.Echo "strAttachFilenameOrig:", strAttachFilenameOrig

                    ' Ask user if want to attach this file to the Activity Note
                    strAttachPrompt = "Current e-mail attachment: " & vbCr & _
                                      strAttachFilenameOrig & vbCr & vbCr & _
                                      "Do you want to attach this file? (Y or N)?"
                    strAttachThis = UCase( _
                                          InputBox(strAttachPrompt, _
                                                   "Attach This File?", _
                                                   "Y"))

                    If strAttachThis = "Y" Then
                        ' Set Attachment Type based on Custom Data - Attachment Types
                        strAttachTypePar = "<p>" + _
                                           "<Lookup Value='Application*' />" + _
                                           "</p>"
                        strAttachType = Company.FindCode("Administration", _
                                                         "Custom Data - Attachment Types", _
                                                         strAttachTypePar)
                        strAttachType = Replace(strAttachType, "Application-", "")
                        ' Get the filename extension of the attachment
                        intPeriod = InStr(strAttachFilenameOrig, ".")
                        If intPeriod <> 0 Then
                            strAttachFilenameOrigExt = Mid(strAttachFilenameOrig, intPeriod)
                            'WScript.Echo "strAttachFilenameOrigExt:", strAttachFilenameOrigExt
                        Else
                            strAttachFilenameOrigExt = ""
                            'WScript.Echo "strAttachFilenameOrigExt:", strAttachFilenameOrigExt
                        End If

                        ' Create a new filename based on information we've gathered
                        strAttachFilenameNew = strNameLast & _
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
                        'WScript.Echo "strAttachFilenameNew:", strAttachFilenameNew

                        ' Ask user if this name is ok, if not let them edit it?
                        ' Could get tedious for many attachments, though.
                        ' Could ask them if want to edit any attachment filenames for
                        ' (1) this applicant (would allow examining at least first filename for next applicant)
                        ' (2) this session (would eliminate further prompts for this session)

                        strAttachFilenameNew = _
                                             InputBox( _
                                                      "Accept the generated filename or edit as necessary. Then press 'Enter'", _
                                                      "Attachment filename", _
                                                      strAttachFilenameNew)

                        ' Write attachment to disk using new filename
                        strAttachFullPathNew = strAppDataFolder & Constants.bs & strAttachFilenameNew
                        objOutlookAttachments.Item(n).SaveAsFile strAttachFullPathNew

                        ' Attach to Note and save Note
                        objActivityPRNote.Fields("Attachments").Items.Add strAttachFullPathNew
                        objActivityPRNote.Save

                        ' Delete file from temporary folder
                        objFSO.DeleteFile strAttachFullPathNew
                        'Else
                        'Skip this attachment
                    End If
                    'Else
                    'skip this attachment
                    'End If  'Close "Is useful attachment" If-block
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
