'
' Provides some constants for use in Windows Registry related scripts
'
'
' Macro Type:      Module
' Using:           na
'
'--------------------------------------------------------------------------------------------------------------

'Useful when you want to enumerate keys from a registry object created with WMI:
'Set oReg = _
'  GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
'  strComputer & _
'  "\root\default:StdRegProv")
'oReg.EnumKey HKEY_CURRENT_USER, <path as string>, <subkeys as array>
'For Each sSubkey in aSubkeys
'  whatever
'Next

Const HKEY_CLASSES_ROOT   = &H80000000
Const HKEY_CURRENT_USER   = &H80000001
Const HKEY_LOCAL_MACHINE  = &H80000002
Const HKEY_USERS          = &H80000003

'--------------------------------------------------------------------------------------------------------------

'Useful when you want to enumerate values from a registry object created with WMI:
'oReg.EnumValues HKEY_CURRENT_USER, <path as string>, <value names as array>, <value types as array>
'For i = LBound(aValueNames) To UBound(aValueNames)
'  sValueName = aValueNames(i)
'  Select Case aValueTypes(i)
'    Case REG_SZ                'Show a REG_SZ value
'       oReg.GetStringValue HKEY_CURRENT_USER, <path as string>, sValueName, sValue
'       wscript.echo "  " & sValueName & " (REG_SZ) = " & sValue
'    Case REG_EXPAND_SZ         'Show a REG_EXPAND_SZ value'
'	'etc
'    . . .
'  End Select

Const REG_SZ        = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY    = 3
Const REG_DWORD     = 4
Const REG_MULTI_SZ  = 7
