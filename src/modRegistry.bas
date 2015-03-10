Option Compare Database
'Option Explicit

' This is the contents of FAQ222-92 from Tek-Tips Forums, written by forum member MikeLacey
' See http://www.tek-tips.com/faqs.cfm?fid=92 for the original text

' J Barnett - have added comments to each of the functions explaining what they do.


'API Function and Constant Declarations
'--------------------------------------


'***Declare the value data types
Global Const REG_SZ As Long = 1 '***Registry string
Global Const REG_DWORD As Long = 4 '***Registry number (32-bit number)

'***Declare the keys that should exist.
'***Typically applications will put information under HKEY_CURRENT_USER
Global Const HKEY_CLASSES_ROOT = &H80000000
Global Const HKEY_CURRENT_USER = &H80000001
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const HKEY_USERS = &H80000003

'***Errors
Global Const ERROR_NONE = 0
Global Const ERROR_BADDB = 1
Global Const ERROR_BADKEY = 2
Global Const ERROR_CANTOPEN = 3
Global Const ERROR_CANTREAD = 4
Global Const ERROR_CANTWRITE = 5
Global Const ERROR_OUTOFMEMORY = 6
Global Const ERROR_INVALID_PARAMETER = 7
Global Const ERROR_ACCESS_DENIED = 8
Global Const ERROR_INVALID_PARAMETERS = 87
Global Const ERROR_NO_MORE_ITEMS = 259

'***Gives all users full access to the key
Global Const KEY_ALL_ACCESS = &H3F

Private Const STANDARD_RIGHTS_READ = &H20000
Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const SYNCHRONIZE = &H100000
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or _
                        KEY_QUERY_VALUE Or _
                        KEY_ENUMERATE_SUB_KEYS Or _
                        KEY_NOTIFY) And _
                        (Not SYNCHRONIZE))

'***Creates a key that is persistent
Global Const REG_OPTION_NON_VOLATILE = 0

Global gstrAppVersion As String

'***Registry API declarations
Declare Function RegCloseKey Lib "advapi32.dll" ( _
    ByVal hKey As Long _
) As Long

Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal Reserved As Long, _
    ByVal lpClass As String, _
    ByVal dwOptions As Long, _
    ByVal samDesired As Long, _
    ByVal lpSecurityAttributes As Long, _
    phkResult As Long, _
    lpdwDisposition As Long _
) As Long

Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long _
) As Long

Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    ByVal lpData As String, _
    lpcbData As Long _
) As Long

Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
String, ByVal lpReserved As Long, lpType As Long, lpData As _
Long, lpcbData As Long) As Long

Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
As Long, lpcbData As Long) As Long

Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
String, ByVal cbData As Long) As Long

Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
ByVal cbData As Long) As Long


Public Function SetValueEx( _
    ByVal hKey As Long, _
    sValueName As String, _
    lType As Long, _
    vValue As Variant _
) As Long
    'aexcode: Wrapper function for writing values to the Registry. Do Not Call Directly.

    '*** Called By: SetKeyValue
    '*** Description: Wrapper function around the registry API calls
    '*** RegSetValueExString/Long. Determines if the value
    '*** is a string or a long and calls the appropriate API.
    '*** Return Value: Returns the API call's return value, which is its
    '*** status (successful, error).

    Dim lValue As Long
    Dim sValue As String

    Select Case lType
        '***String
        Case REG_SZ
            sValue = vValue
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        '***32-bit number
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
    End Select
    
End Function

Private Function QueryValueEx( _
    ByVal lhKey As Long, _
    ByVal szValueName As String, _
    vValue As Variant _
) As Long
'aexcode: Code to read value from the Registry based on data; do not call directly.

'*** Called By: QueryValue
'*** Description: Wrapper function around the registry API calls to
'*** RegQueryValueExLong and RegQueryValueExString.
'*** Determines size and type of data to be read.
'*** Determines if the value is a string or a long
'*** and calls the appropriate API.
'*** Return Value: Returns the API call's return value, which is its
'*** status (successful, error). The parameter vValue
'*** contains the value queried.

Dim cch As Long
Dim lrc As Long
Dim lType As Long
Dim lValue As Long
Dim sValue As String

On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    
    If lrc <> ERROR_NONE Then Error 5
    
    Select Case lType
        ' For strings
        Case REG_SZ:
            sValue = String(cch, 0)
            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                If Mid(sValue, cch, 1) = Chr(0) Then
                vValue = Left$(sValue, cch - 1) ' get rid of trailing AsciiZ
            Else
                vValue = Left$(sValue, cch)
            End If
            Else
                vValue = Empty
            End If
            ' For DWORDS
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            'all other data types not supported
            lrc = -1
    End Select

QueryValueExExit:
    QueryValueEx = lrc

    Exit Function

QueryValueExError:
    Resume QueryValueExExit ' Hmmmm
End Function

Public Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
    'aexcode: Code to create a new Registry key. Not used in MDB Doc.

'***With this procedure a call of
'*** CreateNewKey "TestKey", HKEY_CURRENT_USER
'***will create a key called TestKey immediately under HKEY_CURRENT_USER.
'***Calling CreateNewKey like this
'*** CreateNewKey "TestKey\SubKey1\SubKey2", HKEY_CURRENT_USER
'***will create a three-nested keys beginning with TestKey immediately under
'***HKEY_CURRENT_USER, Subkey1 subordinate to TestKey, and SubKey3 under
'***SubKey2.

    '*** Called by: your own code to create keys
    '*** Description: Wrapper around the RegCreateKeyEx API call.

    Dim hNewKey As Long 'handle to the new key
    Dim lRetVal As Long 'result of the RegCreateKeyEx function

    lRetVal = RegCreateKeyEx( _
        lPredefinedKey, _
        sNewKeyName, _
        0&, _
        vbNullString, _
        REG_OPTION_NON_VOLATILE, _
        KEY_ALL_ACCESS, _
        0&, _
        hNewKey, _
        lRetVal _
    )
    
    RegCloseKey hNewKey

End Sub

Public Sub SetKeyValue( _
    ByVal lpParentKey As Long, _
    sKeyName As String, _
    sValueName As String, _
    vValueSetting As Variant, _
    lValueType As Long _
)
'aexcode: Code to write a value to the registry. Not used by MDB Doc.

'*** Called By: Your code when you want to set a KeyValue
'*** Description: Opens the key you want to set, calls the wrapper
'*** function SetValueEx, and closes key.
'*** ADD ERROR HANDLING!!

Dim lRetVal As Long 'result of the SetValueEx function
Dim hKey As Long 'handle of open key

    'open the specified key
    lRetVal = RegOpenKeyEx(lpParentKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    ' write the value
    lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    ' close the key
    RegCloseKey (hKey)
    
End Sub

Public Function QueryValue( _
    ByVal lpParentKey As Long, _
    sKeyName As String, _
    sValueName As String _
) As Variant
'aexcode: Code to read a value from the Registry. Used for the references section
'*** Called By: Your code when you want to set a read a KeyValue
'*** Description: Opens the key you want to set, calls the wrapper
'*** function QueryValueEx, closes key.
'*** Return Value: The value you are querying
'*** ADD ERROR HANDLING!!

    Dim lRetVal As Long 'result of the API functions
    Dim hKey As Long 'handle of opened key
    Dim vValue As Variant 'setting of queried value

    ' open the key
    lRetVal = RegOpenKeyEx(lpParentKey, sKeyName, 0, KEY_READ, hKey)
    ' get the value
    lRetVal = QueryValueEx(hKey, sValueName, vValue)
    ' close the key
    RegCloseKey (hKey)
    
    QueryValue = vValue
    
End Function

Public Sub TestRegModule()
'aexcode: Demo code for testing writing to the Registry. Not used in MDB Doc.
'
' This procedure demonstrates the use of this module
' All code below here is demo code
'
    CreateKeyDriver
    Debug.Print "Created key Cosmic Example and TestVals"
    
    SetStringValueDriver
    Debug.Print "Created the TestStringVal"
    
    SetNumberValueDriver
    Debug.Print "Created the TestNumVal"
    
    'ReadValueDriver

    'End ' Commented out

End Sub

Sub CreateKeyDriver()
'aexcode: Demo code to demonstrate creating keys.  Not used in MDB Doc.
'
' demo code
'

'*** Calls the CreateNewKey procedure
'*** Description: Creates Cosmic Example key and TestVals subkey under
'*** HKEY_CURRENT_USER\Software\VB and VBA Program Settings
'*** If Software and/or VB and VBA Program Settings do not
'*** exist, they are created.
'*** Usage: Use this as an example of how you would use the CreateNewKey
'*** procedure.

    Dim sNewKey As String
    Dim lPredefinedKeyValue As Long

    sNewKey = "Software\VB and VBA Program Settings\Cosmic Example\TestVals"
    lPredefinedKeyValue = HKEY_CURRENT_USER
    
    CreateNewKey sNewKey, lPredefinedKeyValue
    
End Sub

Sub SetStringValueDriver()
'aexcode: Demo code to demonstrate writing a string value. Not used in MDB Doc.
'
' demo code
'

'*** Calls the SetKeyValue procedure
'*** Description: Sets the value TestStringVal under the
'*** Cosmic Example\TestVals key and sets it to
'*** VB App Created. If it doesn't exist, it creates it.
'*** Usage: Use this as an example of how you would use the SetKeyValue
'*** procedure.

    Dim sKey As String '***Key under which to create the value
    Dim sValue As String '***Value name to set
    Dim vSetting As Variant '***What to set the Value to
    Dim sType As Long '***Value type -- string or number

    sKey = "Software\VB and VBA Program Settings\Cosmic Example\TestVals"
    sValue = "TestStringVal"
    vSetting = "VB App Created"
    sType = REG_SZ

    SetKeyValue HKEY_CURRENT_USER, sKey, sValue, vSetting, sType

End Sub
Sub SetNumberValueDriver()
'aexcode: Demo code to demonstrate creating numeric values in the Registry. Not used in MDB Doc.
'
' demo code
'

'*** Calls the SetKeyValue procedure
'*** Description: Sets the value TestNumVal under the
'*** Cosmic Example\TestVals key and sets it to 5.
'*** If it doesn't exist, it creates it.
'*** Usage: Use this as an example of how you would use the SetKeyValue
'*** procedure.

    Dim sKey As String '***Key under which to create the value
    Dim sValue As String '***Value name to set
    Dim vSetting As Variant '***Wht to set the Value to
    Dim sType As Long '***Value type -- string or number

    sKey = "Software\VB and VBA Program Settings\Cosmic Example\TestVals"
    sValue = "TestNumVal"
    vSetting = 5
    sType = REG_DWORD
    
    SetKeyValue HKEY_CURRENT_USER, sKey, sValue, vSetting, sType

End Sub