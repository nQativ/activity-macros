'The purpose of this macro is to demonstrate how to pop a particular
'Find form, select an item, and return one of its field values to a
'variable. This value is then used to locate the data and add it to
'the Results, which are displayed at the end of the process. This
'can be useful to show the user the records that were created,
'edited, or whatever so that they can review them if they like.

' No filtering --------------------------------------------------------------------
strDivision = Company.FindCode("Administration", "Custom Data - Divisions")
MsgBox "The chosen Division is " & strDivision

' Prepare for next two Find operations
Set objPositions = Company.HumanResources.HRPosition

' Limit items by Lookup -----------------------------------------------------------
strParam = "<p>" + _
           "<Lookup Value='" + strDivision + "*' />" + _
           "</p>"
strPosition = Company.FindCode("Human Resources", _
                               "Positions", _
                               strParam)
objPositions.Locate(strPosition)
MacroProcess.Results.Add objPositions

' Limit items by parameter filter -------------------------------------------------
strParam = "<p>" + _
           "<Filter Name='? Position Code' Type='Shared'>" + _
           "<Parameter Name='Code' Value='" + strDivision + "'/>" + _
           "</Filter>" + _
           "</p>"
strPosition = Company.FindCode("Human Resources", _
                               "Positions", _
                               strParam)
objPositions.Locate(strPosition)
MacroProcess.Results.Add objPositions
