'
' This module is a home for package- and application-non-specific code that may
' be broadly useful to others writing Activity macros.
'
' Macro Type:               Module
' Using:                    Constants
'                           Declarations
'
' Contents
' fRandomNumber             return random number of specified length w/in specified range
'
' ----------------------------------------------------------------------------------------

Function funRandomNumber(intMin, intMax, intLength, blnLeadingZeros)
    ' Returns random number of specified length within specified range,
    ' with or without leading zeros.

    Randomize
    If blnLeadingZeros = True Then
        funRandomNumber = _
            Right(String(intLength, "0") + ((intMax - intMin + 1) * Rnd + intMin), _
                  intLength)
    Else
        funRandomNumber = _
            Trim( _
                 Int(((intMax - intMin + 1) * Rnd + intMin)))
    End If

End Function

' ----------------------------------------------------------------------------------------

Function funDateTimeStamp(datDateTime)
    ' Returns a date-time stamp in the following format:
    ' YYYYMMDD-hhmm

    funDateTimeStamp = Year(datDateTime) & _
                       Right("0" & Month(datDateTime),2)  & _
                       Right("0" & Day(datDateTime),2)  & "-" & _  
                       Right("0" & Hour(datDateTime),2) & _
                       Right("0" & Minute(datDateTime),2)

End Function
