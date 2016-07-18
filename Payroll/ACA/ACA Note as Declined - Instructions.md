#ACA Note as Declined - Instructions

Activity Macro:	**ACA Note as Declined**
Description:	Add ACADeclined Note to ACA Record
Macro Type:		Record Loop
Record Type:	ACA Records (Payroll)

For each ACA Record selected this macro does the following:

1. Validation (Displays error and skips record when not valid)
    1. Must be an ACA Result record, not a Designation record
    1. Must have an Offer Notify Note referenced
    1. Must not already have a Response Note
2. Create ACADeclined Note records with a description of **Declined coverage**
3. Update the reference field of the ACA Result record with the GUID of the **Response** Note record that was just created

Setup Instructions:

1. Create an Activity Macro record on the company database as specified at the top of this document and then copy the Visual Basic code from **ACA Note as Declined.vb**

HR Staff Instructions for how to use this Activity Macro:

1. Click the "**Waiting Response**" link on the **ACA Monthly Operations** dashboard gadget
2. Select the employees from the filtered view of ACA Result Records that **have Declined** the offer for insurance coverage
3. Run the Activity Macro named **ACA Note as Declined**
4. Attach any documentation from the employee that confirms their choice to decline to the newly created ACADeclined Note records
