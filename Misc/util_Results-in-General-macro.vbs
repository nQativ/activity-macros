Set objEmployees = Company.Payroll.PREmployee

objEmployees.New
objEmployees.FirstName = "George"
objEmployees.LastName = "Hammond"
objEmployees.Code = "GH0001"
objEmployees.Save
MacroProcess.AddMessage("Created new employee " & _
                        objEmployees.FirstName & " " & _
                        objEmployees.LastName & " (" & _
                        objEmployees.Code & ")")
MacroProcess.Results.Add objEmployees
