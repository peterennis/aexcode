Option Compare Database
Option Explicit

Public Const LOCAL_PREF_TABLE_NAME As String = "USysMDBDocLocalPreferences"

' Custom type definition, used for array of preferences (far quicker looking through this than reading from table)

Private Type Preference
    strPrefName As String
    strPrefValue As String
    strNotes As String
    blnCanOverride As Boolean
End Type

Dim preferences() As Preference

Public Function mdbdGetPreference(strPreferenceName As String) As String
    'MDBDOC: This function will retrieve the specified local preference from the preferences table. If there isn't one, it will retrieve it from the global prefs table.
    ' Function: mdbdGetPreference
    ' Scope:    Public
    ' Parameters: strPreferenceName - input string; Returns: Preference value.
    ' Author:   John Barnett
    ' Date:     9, 26 December 2003. Amended 30 April 2005.
    ' Description: This function will retrieve the specified preference from the local preferences table. If none are found, it will get it from the global preferences. If not, it will return an empty string.
    ' This ignores the preference "Halt on errors" deliberately - if a preference value can't be retrieved, its a
    ' serious problem for the application.  Error is bubbled up to the main application for handling.
    
    ' The following string will be easily detected if not found
    Const GET_PREFERENCE_ERROR = ""
    
    Dim intMax As Integer
    Dim intCount As Integer
    Dim blnFound As Boolean
    
    intMax = UBound(preferences())  ' Loop the preferences array
    blnFound = False
    
    For intCount = 0 To intMax
        If preferences(intCount).strPrefName = strPreferenceName Then
            blnFound = True
            mdbdGetPreference = preferences(intCount).strPrefValue ' Retrieve the data when found
            Exit For
        End If
    Next
    
    If blnFound = False Then        ' Preference not found in array, display error message box
        MsgBox APP_NAME & " Error in mdbdGetPreference. Preference requested was: " & strPreferenceName & vbCrLf _
            & "preference requested not found", vbOKOnly + vbCritical, APP_NAME
        Err.Raise vbObjectError + 1000, "mdbdGetPreference", "Preference " & strPreferenceName & " not found"
    End If
    
Exit_GetPreference:
    Exit Function

End Function

Public Function mdbdLoadPreferences() As Integer
    'MDBDOC: This function will retrieve the preferences from the preferences table and load them into an array.
    ' Function: mdbdLoadPreferences
    ' Scope:    Public
    ' Parameters: None
    ' Returns:  Integer - Number of parameters retrieved.
    ' Author:   John Barnett
    ' Date:     10 December 2003.
    ' Description: This function will load the preferences from the preferences table in the addin into the array.
    ' Cannot use DLookup here in a loop because it needs to access the codedb's tables, and DLookup will refer to the currentDB tables.

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim intCount As Integer


    On Error GoTo Err_MDBDLoadPreferences
    
    intCount = 0    ' Count of number of preferences
    
    ' Read the data
    Set db = CodeDb
    Set rs = db.OpenRecordset("Select PreferenceName, PreferenceValue, CanOverride, Notes FROM USysMDBDocPreferences ORDER BY PreferenceName;", dbReadOnly)
    rs.MoveLast
    
    ' Now we know number of prefs, resize array to correct dimensions
    ReDim preferences(rs.RecordCount)
    rs.MoveFirst
    
    ' now load them
    Do While Not rs.EOF
        preferences(intCount).strPrefName = rs!PreferenceName
        If Not IsNull(rs!PreferenceValue) Then preferences(intCount).strPrefValue = rs!PreferenceValue Else preferences(intCount).strPrefValue = ""
        If Not IsNull(rs!Notes) Then preferences(intCount).strNotes = rs!Notes Else preferences(intCount).strNotes = ""
        If IsNull(rs!CanOverride) Then preferences(intCount).blnCanOverride = False Else preferences(intCount).blnCanOverride = rs!CanOverride
        intCount = intCount + 1 ' Increment counter
        rs.MoveNext
    Loop
    rs.Close
    Set db = Nothing
    mdbdLoadPreferences = intCount ' return number of preferences loaded.

Exit_MDBDLoadPreferences:
    Exit Function
    
Err_MDBDLoadPreferences:
    Select Case Err.Number
        Case 3078
            Resume Next
        Case Else
            MsgBox Err.Number & " " & Err.Description
            mdbdLoadPreferences = 0
    End Select

End Function