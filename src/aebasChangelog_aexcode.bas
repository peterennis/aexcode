Option Compare Database
Option Explicit

' Constants for settings of "aexcode"
Public Const gstrPROJECT_AEXCODE As String = "aexcode"

Private Const gstrVERSION_AEXCODE As String = "2.0.0"
Private Const gstrDATE_AEXCODE As String = "March 6, 2015"

Public Const THE_SOURCE_FOLDER = "C:\ae\aexcode\src\"
Public Const THE_XML_FOLDER = "C:\ae\aexcode\src\xml\"
'

Public Function getMyVersion() As String
    On Error GoTo 0
    getMyVersion = gstrVERSION_AEXCODE
End Function

Public Function getMyDate() As String
    On Error GoTo 0
    getMyDate = gstrDATE_AEXCODE
End Function

Public Function getMyProject() As String
    On Error GoTo 0
    getMyProject = gstrPROJECT_AEXCODE
End Function

Public Sub AEXCODE_EXPORT(Optional ByVal varDebug As Variant)

    On Error GoTo PROC_ERR

    If Not IsMissing(varDebug) Then
        aegitClassTest varDebug:="varDebug", varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER, varXmlData:=THE_XML_FOLDER
    Else
        aegitClassTest varSrcFldr:=THE_SOURCE_FOLDER, varXmlFldr:=THE_XML_FOLDER, varXmlData:=THE_XML_FOLDER
    End If

PROC_EXIT:
    Exit Sub

PROC_ERR:
    MsgBox "Erl=" & Erl & " Error " & Err.Number & " (" & Err.Description & ") in procedure AEXCODE_EXPORT"
    Resume Next

End Sub

'=============================================================================================================================
' Tasks:
' %010 -
' %009 -
' %008 -
' %007 -
' %006 -
' %005 -
' %004 -
' %003 -
' %002 -
' %001 -
' Issues:
' #010 -
' #009 -
' #008 -
' #007 -
' #006 -
' #005 -
' #004 -
' #003 -
' #002 -
' #001 -
'=============================================================================================================================
'
'
'20150306 - v200
    ' Load mdbdoc v161 as accdb for Access 2013, export and push to GitHub aexcode branch 161
    ' Update license and author details on GitHub, set this project to aexcode, merge to main branch
    ' Import aegit 1.2.9 code, create AEXCODE_EXPORT and procedures in this module
    ' Set version to 2.0.0, run first export and commit