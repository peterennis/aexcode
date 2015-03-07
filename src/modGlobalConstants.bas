Option Compare Database
Option Explicit

' Application name - for status bar messages, message boxes etc.
' Use this if you want to change the name of the application. Note that the only place this won't change it
' is in the preferences text and source code comments.
Public Const APP_NAME As String = "MDB Doc"

Public Const ACCESS_2007_VERSION As Integer = 12

' Global constants for on/off Preferences; used for decision making throughout the software.

' If you want to use Y/N instead, change these constants and use SQL update statements to modify the
' values in the table.
Public Const PREFERENCE_DISABLED As String * 1 = "0" ' Constant meaning on/off preference is disabled. Only used for error checking, but important nonetheless
Public Const PREFERENCE_ENABLED As String * 1 = "1" ' Constant meaning on/off preference is enabled

' Constants that relate to the stylesheets
Public Const STYLESHEET_ERROR As Integer = 0        ' 0 - Error detected (not able to load a file with text inclusion; invalid file etc)
Public Const STYLESHEET_PATH As Integer = 1         ' 1 - Included using path method
Public Const STYLESHEET_INCLUDE As Integer = 2      ' 2 - Included text in output file
    