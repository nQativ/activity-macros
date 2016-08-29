'The purpose of this macro is to demonstrate how to pop a particular
'Find form, select an item, and return one of its field values to a
'variable. Normally, that value would be used in subsequent actions,
'but in this demo, it is only displayed in a message box.

' No filtering --------------------------------------------------------------------
strDivision = Company.FindCode("Administration", "Custom Data - Divisions")
MsgBox "The chosen Division is " & strDivision

' Limit items by Lookup -----------------------------------------------------------
strParam = "<p>" + _
           "<Lookup Value='" + strDivision + "*' />" + _
           "</p>"
strPosition = Company.FindCode("Human Resources", _
                               "Positions", _
                               strParam)
MsgBox "The chosen Position is " & strPosition

' Limit items by parameter filter -------------------------------------------------
strParam = "<p>" + _
           "<Filter Name='? Position Code' Type='Shared'>" + _
           "<Parameter Name='Code' Value='" + strDivision + "'/>" + _
           "</Filter>" + _
           "</p>"
strPosition = Company.FindCode("Human Resources", _
                               "Positions", _
                               strParam)
MsgBox "The chosen Position is " & strPosition
