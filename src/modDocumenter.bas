Option Compare Database
Option Explicit

Private Const QUOTE_MARK As Integer = 34    ' ASCII code for double quote mark - to help understand code
Private Const OBJECT_CLOSED As Integer = 0  ' (Unofficial) SysCmd constant for object being closed (not in the standard enum results)

' Return to contents list item from the output file.
Private Const RETURN_TO_CONTENTS As String = "<p><a class=""contentslist"" href=""#mdbContentsList"">Click here to return to the contents list</a></p>"

Public Function mdbdProcessDatabase(strFileName As String) As Boolean
    'MDBDOC: Function that does the bulk of the processing.
    ' Function: ProcessDatabase
    ' Scope:    Global
    ' Parameters: strFilename (string) - the path/filename to output the HTML to.
    ' Return Value: Boolean - success/fail
    ' Author:   John Barnett
    ' Date:     1-15 July 2001.  Amended 22-24 Oct and 01 Nov 2003 (fixes minor HTML output bugs).
    ' Updated Dec 2003 to output to XHTML rather than HTML 4, and comply with Bobby requirements for v1.3
    ' Updated Sep 2006 to add support for Data Access Pages (DAPs from Access 2000 and later).
    ' Updated October 2007 to: - retrive name from constant, quote mark constant and
    ' application version from the preference in the preferences table.
    ' Updated Feb 2008 to add facility to use FormatSQL preference in v1.42
    ' Description: The ProcessDatabase function processes the current database and outputs the
    ' result to an HTML file specified as the strFilename parameter.
    ' Updated April 2011 for v1.45 to:
    '   - fix omission of /head tag
    '   - fix tag nesting error in header of references
    ' Updated February-April 2014 for v1.50 to:
    '   - move to HTML 5, including:
    '   - amendment of DOCTYPE and META content-type
    '   - removal of summary tag on table elements
    '   - replacement of text-align:right with class="rightNum" and CSS class
    '   - removed meta copyright item in line with HTML5 specification
    '   - replaced name values with IDs
    ' Called by: cmdProcess_Click subroutine on frmStartup.

    Dim db As DAO.Database            ' DAO Database object, used for access to all subcontainers plus DB properties
    Dim cnt As DAO.Container          ' Container object, used for accessing objects in a container (forms/reports/macros)
    Dim qdf As DAO.QueryDef           ' QueryDef object, used for accessing queries in the database
    Dim doc As DAO.Document           ' Document object, used for enumerating members of a container
    Dim tdf As DAO.TableDef           ' Tabledef object, used for enumerating tables in the database
    Dim mdl As Module                 ' Module object, used for enumerating modules in the database
    Dim fld As DAO.Field              ' Field object, used for enumerating fields within tables (table detail only)
    Dim prp As DAO.Property           ' Property object, used for accessing database properties
    Dim rel As DAO.Relation           ' Relation object, used for looking at table relationships (1:1, 1:M links etc)
    Dim CBR As CommandBar             ' CommandBar object, used for accessing commandbars in the database
    Dim dap As DataAccessPage         ' DataAccessPage object, used for accessing DAPs in the database
    Dim intReferenceCount As Integer  ' Integer variable used for looping through references.

    Dim strHTML As String             ' Temporary storage for building up HTML.
    Dim intModcount As Integer        ' Variable for looping through modules in processing
    Dim intCount As Integer           ' Counter used for looping through form and report collections
    Dim clsfh As mdbdclsFileHandle    ' File handler for writing the output to.
    Dim blnWrittenHeader As Boolean   ' Flag indicating if a table header has been written (or not) for the current section
    Dim blnOpened As Boolean          ' Flag indicating if MDB Doc opened a particular object or not
    Dim strPreference As String       ' Holds value of the most recently retrieved preference.
    Dim intStylesheetResult As Integer ' Result of Stylesheet function
    Dim blnVBEVisible As Boolean       ' Is VBA editor window visible?
    
    Dim blnFormatSQL As Boolean       ' Use "Format SQL" routine (from preference)
    Dim blnNewlineInDAPConnection As Boolean ' insert newline with semi colon in a DAP connection string
    
    Dim strHaltOnErrors As String     ' "Halt on Errors" preference value
    
    Call mdbdCloseObjects             ' close objects prior to starting

    
    On Error Resume Next

    ' Retrieve value of "Halt on Errors" preference
    strHaltOnErrors = mdbdGetPreference("HaltOnErrors")
    
    If (strHaltOnErrors = PREFERENCE_ENABLED) Or (strHaltOnErrors = PREFERENCE_DISABLED) Then
        ' Valid value detected, continue with processing...
        
        blnVBEVisible = Application.VBE.MainWindow.Visible
        
        ' Turn off screen echo to avoid distractions
        DoCmd.Echo False, APP_NAME & " is executing."
        
        On Error Resume Next
        
        ' obtain the basic database information
        Set db = CurrentDb
        Set clsfh = New mdbdclsFileHandle
        With clsfh
            .Filename = strFileName ' configure the output file handler
            .FileMode = "W"
            .OpenFile

            ' put in generic database information for the HTML header section
            ' For more information about writing HTML/XHTML code please see http://www.w3schools.com/xhtml/default.asp
            
            .WriteData "<!DOCTYPE html>"
            .WriteData "<html lang=""en"">" ' HTML header section.
            .WriteData "<head>"
            .WriteData "<title>Database Information for " & db.Name & "</title>"
            .WriteData "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />"
            .WriteData "<meta name=""Generator"" content=""" & APP_NAME & " " & mdbdGetPreference("MDBDocVersion") & """ />"
    
            ' Meta Description tag
            strPreference = mdbdGetPreference("MetaDescription")
            If Len(strPreference & "") > 0 Then
                .WriteData "<meta name=""description"" content=""" & strPreference & """ />"
            End If
    
            ' Meta Keywords
            strPreference = mdbdGetPreference("MetaKeywords")
            If Len(strPreference & "") > 0 Then
                .WriteData "<meta name=""keywords"" content=""" & strPreference & """ />"
            End If
    
            ' Meta Author
            strPreference = mdbdGetPreference("MetaAuthor")
            If Len(strPreference & "") > 0 Then
                .WriteData "<meta name=""author"" content=""" & strPreference & """ />"
            End If

            ' Now do the rest of the stylesheet data
            intStylesheetResult = LoadStylesheets(clsfh, strHaltOnErrors)  ' load the stylesheet information.
            If (intStylesheetResult = STYLESHEET_ERROR) And (strHaltOnErrors = PREFERENCE_ENABLED) Then
                MsgBox "Warning! Error processing Stylesheet data. Please check the file at the end, continuing...", vbOKOnly + vbInformation
            End If
            
            ' This is the end of the Head section of the output file.
            .WriteData "</head>"
            ' Start of the Body section
            .WriteData "<body>"
    
            ' Title at the top of the body section giving the path to the database processed and date/time of processing.
            .WriteData "<h1 id=""mdbHeader"">Database information for file: " & db.Name & ", created " & str(Now()) & "</h1>"
            
            ' The project details
            .WriteData "<table class=""mdbtable"">"
            .WriteData "<caption>Key data</caption>"
            .WriteData "<thead>"
            .WriteData "<tr><th scope=""col"">Setting</th><th scope=""col"">Value</th></tr>"
            .WriteData "</thead>"
            .WriteData "<tbody>"
            .WriteData "<tr><td>Application name:</td><td>" & Application.GetOption("Project Name") & "</td></tr>"
            .WriteData "<tr><td>Is it compiled?</td><td>" & Application.IsCompiled & "</td></tr>"
            .WriteData "<tr><td>Does it have a broken reference?</td><td>" & Application.BrokenReference & "</td></tr>"
            .WriteData "</tbody>"
            .WriteData "</table>"
            
            ' The contents list from the output file
            .WriteData "<h2 id=""mdbContentsList"">Contents</h2>"
            .WriteData "<nav role=""navigation"">"
            .WriteData "<ul>"
            If mdbdGetPreference("ProcessDBProperties") = PREFERENCE_ENABLED Then .WriteData "<li><a href=" & Chr$(QUOTE_MARK) & "#" & "mdbDBProps" & Chr$(QUOTE_MARK) & ">Database Properties</a></li>"
            If mdbdGetPreference("ProcessTableSummary") = PREFERENCE_ENABLED Then .WriteData "<li><a href=" & Chr$(QUOTE_MARK) & "#" & "mdbTables" & Chr$(QUOTE_MARK) & ">Tables</a></li>"
            If mdbdGetPreference("ProcessTableDetail") = PREFERENCE_ENABLED Then .WriteData "<li><a href=" & Chr$(QUOTE_MARK) & "#" & "mdbQueries" & Chr$(QUOTE_MARK) & ">Queries</a></li>"
            If mdbdGetPreference("ProcessQueries") = PREFERENCE_ENABLED Then .WriteData "<li><a href=" & Chr$(QUOTE_MARK) & "#" & "mdbForms" & Chr$(QUOTE_MARK) & ">Forms</a></li>"
            If mdbdGetPreference("ProcessReports") = PREFERENCE_ENABLED Then .WriteData "<li><a href=" & Chr$(QUOTE_MARK) & "#" & "mdbReports" & Chr$(QUOTE_MARK) & ">Reports</a></li>"
            If mdbdGetPreference("ProcessModuleDetail") = PREFERENCE_ENABLED Or mdbdGetPreference("ProcessModuleSummary") = PREFERENCE_ENABLED Then .WriteData "<li><a href=" & Chr$(QUOTE_MARK) & "#" & "mdbModules" & Chr$(QUOTE_MARK) & ">Modules</a></li>"
            If mdbdGetPreference("ProcessFormModules") = PREFERENCE_ENABLED Then .WriteData "<li><a href=" & Chr$(QUOTE_MARK) & "#" & "mdbFormModules" & Chr$(QUOTE_MARK) & ">Form Modules</a></li>"
            If mdbdGetPreference("ProcessReportModules") = PREFERENCE_ENABLED Then .WriteData "<li><a href=" & Chr$(QUOTE_MARK) & "#" & "mdbReportModules" & Chr$(QUOTE_MARK) & ">Report Modules</a></li>"
            If mdbdGetPreference("ProcessMacros") = PREFERENCE_ENABLED Then .WriteData "<li><a href=" & Chr$(QUOTE_MARK) & "#" & "mdbMacros" & Chr$(QUOTE_MARK) & ">Macros</a></li>"
            If mdbdGetPreference("ProcessRelationships") = PREFERENCE_ENABLED Then .WriteData "<li><a href=" & Chr$(QUOTE_MARK) & "#" & "mdbRelationships" & Chr$(QUOTE_MARK) & ">Relationships</a></li>"
            If mdbdGetPreference("ProcessCommandBars") = PREFERENCE_ENABLED Then .WriteData "<li><a href=" & Chr$(QUOTE_MARK) & "#" & "mdbCommandBars" & Chr$(QUOTE_MARK) & ">Command Bars</a></li>"
            If mdbdGetPreference("ProcessDAPS") = PREFERENCE_ENABLED Then .WriteData "<li><a href=" & Chr$(34) & "#" & "mdbDataAccessPages" & Chr$(34) & ">Data Access Pages</a></li>"
            If mdbdGetPreference("ProcessReferences") = PREFERENCE_ENABLED Then .WriteData "<li><a href=" & Chr$(QUOTE_MARK) & "#" & "mdbReferences" & Chr$(QUOTE_MARK) & ">References</a></li>"
            If mdbdGetPreference("ProcessRibbons") = PREFERENCE_ENABLED Then .WriteData "<li><a href=" & Chr$(QUOTE_MARK) & "#" & "mdbRibbons" & Chr$(QUOTE_MARK) & ">Ribbons</a></li>"
            If mdbdGetPreference("ProcessImpExpSpecs") = PREFERENCE_ENABLED Then .WriteData "<li><a href=" & Chr$(QUOTE_MARK) & "#" & "mdbImpExpSpecs" & Chr$(QUOTE_MARK) & ">Import/Export Specifications</a></li>"
            .WriteData "</ul>"
            .WriteData "</nav>"
        End With
        DoEvents
    
        ' ******** Database properties ********
        ' Useful for basic information about the database itself, plus the version of the file that it is using.
    
        If mdbdGetPreference("ProcessDBProperties") = PREFERENCE_ENABLED Then
            SysCmd acSysCmdSetStatus, APP_NAME & ": Processing database properties"
            clsfh.WriteData "<h2 id=""mdbDBProps"">Contents</h2>"
            clsfh.WriteData "<table class=""mdbtable"">"
            clsfh.WriteData "<caption>Database Properties</caption>"
            clsfh.WriteData "<thead>"
            clsfh.WriteData "<tr><th scope=""col"">Name</th><th scope=""col"">Type</th><th scope=""col"">Value</th></tr>"
            clsfh.WriteData "</thead>"
            clsfh.WriteData "<tbody>"
            SysCmd acSysCmdSetStatus, APP_NAME & ": Processing properties"
            For Each prp In db.Properties
                clsfh.WriteData "<tr><td>" & prp.Name & "</td><td>" & mdbdGetPropertyType(prp) & "</td><td>" & IIf(prp.Type <> 0, prp, "") & "</td></tr>"
            Next prp
            clsfh.WriteData "</tbody>"
            clsfh.WriteData "</table>"
            clsfh.WriteData RETURN_TO_CONTENTS
            SysCmd acSysCmdClearStatus
        End If
        DoEvents
        
        ' ******** Tables (summary) ********
        blnWrittenHeader = False
        If mdbdGetPreference("ProcessTableSummary") = PREFERENCE_ENABLED Then
            SysCmd acSysCmdSetStatus, APP_NAME & ": Processing tables"
            clsfh.WriteData "<h2 id=""mdbTables"">Tables</h2>"
           
            strPreference = mdbdGetPreference("IncludeUSysTables")
            If strPreference = PREFERENCE_ENABLED Then ' include User system tables
                For Each tdf In db.TableDefs
                    If Left$(tdf.Name, 4) <> "MSys" Then
                        If blnWrittenHeader = False Then
                            clsfh.WriteData "<table class=""mdbtable"">"
                            clsfh.WriteData "<caption>Table Summary Information</caption>"
                            clsfh.WriteData "<thead>"
                            clsfh.WriteData "<tr><th scope=""col"">Name</th><th scope=""col"">Created</th><th scope=""col"">Last Updated</th><th scope=""col"">Record Count</th><th scope=""col"">Type</th><th scope=""col"">Description</th></tr>"
                            clsfh.WriteData "</thead>"
                            clsfh.WriteData "<tbody>"
                            blnWrittenHeader = True
                        End If
                        clsfh.WriteData "<tr><td><a href=" & Chr$(QUOTE_MARK) & "#" & tdf.Name & Chr$(QUOTE_MARK) & ">" & tdf.Name & "</a></td><td>" & tdf.DateCreated & "</td><td>" & tdf.LastUpdated & "</td><td class=""rightNum"">" & IIf(tdf.RecordCount = -1, "Unknown", tdf.RecordCount) & "</td><td>" & mdbdGetTableType(tdf) & "</td><td >" & mdbdGetDescription(tdf) & "</td></tr>"
                    End If
                Next tdf
            Else ' don't include user system tables
                For Each tdf In db.TableDefs
                    ' If table name does not start "USys" or "MSys"
                    If Left(tdf.Name, 4) <> "USys" And Left(tdf.Name, 4) <> "MSys" Then
                        If blnWrittenHeader = False Then
                            clsfh.WriteData "<table class=""mdbtable"">"
                            clsfh.WriteData "<caption>Table Summary Information</caption>"
                            clsfh.WriteData "<thead>"
                            clsfh.WriteData "<tr><th scope=""col"">Name</th><th scope=""col"">Created</th><th scope=""col"">Last Updated</th><th scope=""col"">Record Count</th><th scope=""col"">Type</th><th scope=""col"">Description</th></tr>"
                            clsfh.WriteData "</thead>"
                            clsfh.WriteData "<tbody>"
                            blnWrittenHeader = True
                        End If
                        clsfh.WriteData "<tr><td><a href=" & Chr$(QUOTE_MARK) & "#" & Replace(tdf.Name, " ", "%20") & Chr$(QUOTE_MARK) & ">" & tdf.Name & "</a></td><td>" & tdf.DateCreated & "</td><td>" & tdf.LastUpdated & "</td><td class=""rightNum"">" & IIf(tdf.RecordCount = -1, "Unknown", tdf.RecordCount) & "</td><td >" & mdbdGetTableType(tdf) & "</td><td >" & mdbdGetDescription(tdf) & "</td></tr>"
                        blnOpened = True
                    End If
                Next tdf
            End If
            clsfh.WriteData "</tbody>"
            clsfh.WriteData "</table>"
            clsfh.WriteData "<br />"
            If blnWrittenHeader = False Then
                clsfh.WriteData "<p>MDB Doc is configured to not display USys tables</p>" & RETURN_TO_CONTENTS
            End If
            SysCmd acSysCmdClearStatus
        End If
        DoEvents
        
        strPreference = mdbdGetPreference("ProcessTableDetail")
        If strPreference = PREFERENCE_ENABLED Then
            ' ******** Fields (table detail) ********
            ' there is no need to have safeguards here in case there are no fields, because it is not possible
            ' to have a table with no fields in Access.
            For Each tdf In db.TableDefs
                SysCmd acSysCmdSetStatus, APP_NAME & ": Processing table " & tdf.Name
                If Left$(tdf.Name, 4) <> "MSys" Then
                    If ((Left$(tdf.Name, 4) = "USys" And strPreference = PREFERENCE_ENABLED) Or (strPreference = PREFERENCE_ENABLED)) Then
                        clsfh.WriteData "<table class=""mdbtable"" id=" & Chr$(QUOTE_MARK) & Replace(tdf.Name, " ", "%20") & Chr$(QUOTE_MARK) & ">"
                        clsfh.WriteData "<caption>List of fields within the " & tdf.Name & " table</caption>"
                        clsfh.WriteData "<thead>"
                        clsfh.WriteData "<tr><th scope=""col"">Fieldname</th><th scope=""col"">Default Value</th><th scope=""col"">Data Type</th><th scope=""col"">Required</th><th scope=""col"">Is PK?</th><th scope=""col"">Description</th></tr>"
                        clsfh.WriteData "</thead>"
                        clsfh.WriteData "<tbody>"
                        For Each fld In tdf.Fields
                            clsfh.WriteData "<tr><td>" & fld.Name & "</td><td>" & fld.DefaultValue & "</td><td>" & mdbdGetFieldType(fld, fld.size) & "</td><td>" & fld.Required & "</td><td>" & mdbdIsPK(tdf, fld) & "</td><td>" & mdbdGetDescription(fld) & "</td></tr>"
                        Next fld
                        clsfh.WriteData "</tbody>"
                        clsfh.WriteData "</table>"
                        clsfh.WriteData "<br />"
                        SysCmd acSysCmdClearStatus
                    End If
                End If
            Next tdf
            clsfh.WriteData RETURN_TO_CONTENTS
        End If
        DoEvents
        
        ' ******** Queries ********
        If mdbdGetPreference("ProcessQueries") = PREFERENCE_ENABLED Then
            
            ' Added for 1.42 to include SQL formatting routines
            If mdbdGetPreference("FormatSQL") = PREFERENCE_ENABLED Then
                blnFormatSQL = True
            Else
                blnFormatSQL = False
            End If

            clsfh.WriteData "<h2 id=""mdbQueries"">Queries</h2>"
            If db.QueryDefs.Count > 0 Then
                blnWrittenHeader = False
                SysCmd acSysCmdSetStatus, APP_NAME & ": Processing Queries"
                For Each qdf In db.QueryDefs
                    If Left$(qdf.Name, 3) <> "~sq" Then ' Comment out this line to include temporary queries
                        If blnWrittenHeader = False Then
                            clsfh.WriteData "<table class=""mdbtable"">"
                            clsfh.WriteData "<caption>Query Information</caption>"
                            clsfh.WriteData "<thead>"
                            clsfh.WriteData "<tr><th scope=""col"">Name</th><th scope=""col"">Type</th><th scope=""col"">SQL</th><th scope=""col"">Description</th></tr>"
                            clsfh.WriteData "</thead>"
                            clsfh.WriteData "<tbody>"
                            blnWrittenHeader = True
                        End If ' Comment out this line to include temporary queries
                        
                        If blnFormatSQL = True Then
                            clsfh.WriteData "<tr><td>" & qdf.Name & "</td><td>" & mdbdGetQueryType(qdf) & "</td><td><code>" & FormatSQL(mdbdReplaceSpecialChars(qdf.sql), False) & "</code></td><td>" & mdbdGetDescription(qdf) & "</td></tr>"
                        Else
                            clsfh.WriteData "<tr><td>" & qdf.Name & "</td><td>" & mdbdGetQueryType(qdf) & "</td><td><code>" & mdbdReplaceSpecialChars(qdf.sql) & "</code></td><td>" & mdbdGetDescription(qdf) & "</td></tr>"
                        End If
                    End If
                Next qdf
                If blnWrittenHeader = True Then
                    clsfh.WriteData "</tbody>"
                    clsfh.WriteData "</table>"
                Else
                    clsfh.WriteData "<p>There are no queries in this database.</p>"
                End If
                SysCmd acSysCmdClearStatus
            Else
                clsfh.WriteData "<p>There are no queries in this database.</p>"
            End If
            clsfh.WriteData RETURN_TO_CONTENTS
        End If
        DoEvents

        ' ******** Forms ********
        If mdbdGetPreference("ProcessForms") = PREFERENCE_ENABLED Then
            clsfh.WriteData "<h2 id=""mdbForms"">Forms</h2>"
            Set cnt = db.Containers!Forms
            If cnt.Documents.Count > 0 Then
                SysCmd acSysCmdSetStatus, APP_NAME & ": Processing forms"
                clsfh.WriteData "<table class=""mdbtable"">"
                clsfh.WriteData "<caption>Information about forms within this database</caption>"
                clsfh.WriteData "<thead>"
                clsfh.WriteData "<tr><th scope=""col"">Name</th><th scope=""col"">Created</th><th scope=""col"">Last Updated</th><th scope=""col"">Has Module</th><th scope=""col"">Record Source</th><th scope=""col"">Description</th></tr>"
                clsfh.WriteData "</thead>"
                clsfh.WriteData "<tbody>"
                For Each doc In cnt.Documents
                    blnOpened = False
                    If SysCmd(acSysCmdGetObjectState, acForm, doc.Name) = OBJECT_CLOSED Then
                        DoCmd.OpenForm doc.Name, acDesign, windowmode:=acHidden
                        blnOpened = True
                    End If
                    For intCount = 0 To Forms.Count - 1
                        If Forms(intCount).Name = doc.Name Then
                            If Forms(doc.Name).Properties("HasModule") = True Then
                                clsfh.WriteData "<tr><td><a href=" & Chr$(QUOTE_MARK) & "#" & Replace(Forms(intCount).Module, " ", "%20") & Chr$(QUOTE_MARK) & ">" & doc.Name & "</a></td><td>" & doc.DateCreated & "</td><td>" & doc.LastUpdated & "</td><td>" & Forms(intCount).Properties("HasModule") & "</td><td>" & mdbdReplaceSpecialChars(Forms(intCount).Properties("RecordSource")) & "</td><td>" & mdbdGetDescription(doc) & "</td></tr>"
                            Else
                                clsfh.WriteData "<tr><td>" & doc.Name & "</td><td>" & doc.DateCreated & "</td><td>" & doc.LastUpdated & "</td><td>" & Forms(intCount).Properties("HasModule") & "</td><td>" & mdbdReplaceSpecialChars(Forms(intCount).Properties("RecordSource")) & "</td><td>" & mdbdGetDescription(doc) & "</td></tr>"
                            End If
                        End If
                    Next intCount
                    If blnOpened = True Then DoCmd.Close acForm, doc.Name, acSavePrompt
                Next doc
                clsfh.WriteData "</tbody>"
                clsfh.WriteData "</table>"
                SysCmd acSysCmdClearStatus
            Else
                clsfh.WriteData "<p>There are no forms in this database.</p>"
            End If
            clsfh.WriteData RETURN_TO_CONTENTS
        End If
        DoEvents

        ' ******** Reports ********
        ' Note for v1.4x - October 2007.
        ' Access 2002 and 2003 have a WindowMode property on DoCmd.OpenReport that permits opening as a hidden window.
        ' However this isn't present in earlier versions.  In the interests of cross version compatability I have not
        ' used this functionality.
        '
        ' However, if you use 2002 or later exclusively, you may want to put this option on as it lowers the amount of stuff
        ' opened on the screen.  I can't put this on a preference based on the Access version number as otherwise it won't compile
        ' in Access 2000.
        
        ' Note: for MDB Doc 1.60 by default this is now on, as it is compatible with Access 2007 or later.
        
        If mdbdGetPreference("ProcessReports") = PREFERENCE_ENABLED Then
            clsfh.WriteData "<h2 id=""mdbReports"">Reports</h2>"
            Set cnt = db.Containers!Reports

            If cnt.Documents.Count > 0 Then
                SysCmd acSysCmdSetStatus, APP_NAME & ": Processing reports"
                clsfh.WriteData "<table class=""mdbtable"">"
                clsfh.WriteData "<caption>Information about the reports within this database</caption>"
                clsfh.WriteData "<thead>"
                clsfh.WriteData "<tr><th scope=""col"">Name</th><th scope=""col"">Created</th><th scope=""col"">Last Updated</th><th scope=""col"">Has Module</th><th scope=""col"">Record Source</th><th scope=""col"">Description</th></tr>"
                clsfh.WriteData "</thead>"
                clsfh.WriteData "<tbody>"
                For Each doc In cnt.Documents
                    blnOpened = False
                    If SysCmd(acSysCmdGetObjectState, acReport, doc.Name) = OBJECT_CLOSED Then
                        blnOpened = True
                        DoCmd.OpenReport doc.Name, acViewDesign, windowmode:=acHidden
                    End If
                    For intCount = 0 To Reports.Count - 1
                        If Reports(intCount).Name = doc.Name Then
                            If Reports(intCount).Properties("HasModule") = True Then
                                clsfh.WriteData "<tr><td><a href=" & Chr$(QUOTE_MARK) & "#" & Replace(Reports(intCount).Module, " ", "%20") & Chr$(QUOTE_MARK) & ">" & doc.Name & "</a></td><td>" & doc.DateCreated & "</td><td>" & doc.LastUpdated & "</td><td>" & Reports(intCount).Properties("HasModule") & "</td><td>" & mdbdReplaceSpecialChars(Reports(intCount).Properties("RecordSource")) & "</td><td>" & mdbdGetDescription(doc) & "</td></tr>"
                            Else
                                clsfh.WriteData "<tr><td>" & doc.Name & "</td><td>" & doc.DateCreated & "</td><td>" & doc.LastUpdated & "</td><td>" & Reports(intCount).Properties("HasModule") & "</td><td>" & mdbdReplaceSpecialChars(Reports(intCount).Properties("RecordSource")) & "</td><td>" & mdbdGetDescription(doc) & "</td></tr>"
                            End If
                        End If
                    Next intCount
                    If blnOpened = True Then DoCmd.Close acReport, doc.Name
                Next doc
                clsfh.WriteData "</tbody>"
                clsfh.WriteData "</table>"
                SysCmd acSysCmdClearStatus
            Else
                clsfh.WriteData "<p>There are no reports in this database.</p>"
            End If
            clsfh.WriteData RETURN_TO_CONTENTS
        End If
        DoEvents

        ' Modules are a bit more tricky than other types. This is because not only are there standard modules accessed from the
        ' tab on the database window, but there are modules behind forms and reports as well.  Each different type is being
        ' documented in a different section below.
        ' In addition, some of the module properties I want are only available if you open the module.

        ' ******** Normal/Class modules ********
        If mdbdGetPreference("ProcessModuleSummary") = PREFERENCE_ENABLED Then
            SysCmd acSysCmdSetStatus, APP_NAME & ": Processing modules"
            clsfh.WriteData "<h2 id=""mdbModules"">Modules</h2>"
            Set cnt = db.Containers!Modules
            If cnt.Documents.Count > 0 Then
                clsfh.WriteData "<h2>Ordinary and Class Modules</h2>" ' headers
                clsfh.WriteData "<table class=""mdbtable"">"
                clsfh.WriteData "<caption>Details of the ordinary and class modules within this database</caption>"
                clsfh.WriteData "<thead>"
                clsfh.WriteData "<tr><th scope=""col"">Name</th><th scope=""col"">Created</th><th scope=""col"">Last Updated</th><th scope=""col"">Type</th><th scope=""col"">No. Lines</th><th scope=""col"">Description</th></tr>"
                clsfh.WriteData "</thead>"
                clsfh.WriteData "<tbody>"
                For Each doc In cnt.Documents
                   ' The rest of the HTML is generated below as it needs to access the module itself rather than the document object
                    blnOpened = False
                    If SysCmd(acSysCmdGetObjectState, acModule, doc.Name) = OBJECT_CLOSED Then
                        blnOpened = True
                        DoCmd.OpenModule doc.Name  ' Some of the module properties we want are only available if it is open, so we have to open it...
                    End If
                    
                    ' Fix for v1.5.1 so that those that only have declaration lines don't get a hyperlink ID
                    If Modules(doc.Name).CountOfLines = Modules(doc.Name).CountOfDeclarationLines Then
                        strHTML = "<tr><td>" & doc.Name & "</td><td>" & doc.DateCreated & "</td><td>" & doc.LastUpdated & "</td>"
                    Else
                        strHTML = "<tr><td><a href=" & Chr$(QUOTE_MARK) & "#" & Replace(doc.Name, " ", "%20") & Chr$(QUOTE_MARK) & ">" & doc.Name & "</a></td><td>" & doc.DateCreated & "</td><td>" & doc.LastUpdated & "</td>"
                    End If

                    If Modules(doc.Name) = doc.Name Then
                        ' find the correct module and create the HTML line for it...
                        strHTML = strHTML & "<td>" & IIf(Modules(doc.Name).Type = acClassModule, "Class Module", "Normal Module") & "</td><td class=""rightNum"">" & Modules(doc.Name).CountOfDeclarationLines + Modules(doc.Name).CountOfLines & "</td><td>" & mdbdGetDescription(doc) & "</td></tr>"
                    End If
                    If blnOpened = True Then DoCmd.Close acModule, doc.Name, acSaveNo
                    clsfh.WriteData strHTML ' ... save it to the file
                Next doc
                clsfh.WriteData "</tbody>"
                clsfh.WriteData "</table>"
                clsfh.WriteData RETURN_TO_CONTENTS
            Else
                clsfh.WriteData "<p>This database has no modules in it.</p>"
            End If
        End If
        DoEvents

        If mdbdGetPreference("ProcessModuleDetail") = PREFERENCE_ENABLED Then
        ' ******** Normal/Class module detail (function/sub/property routines) ********
            For Each doc In cnt.Documents
                blnOpened = False
                If SysCmd(acSysCmdGetObjectState, acModule, doc.Name) = OBJECT_CLOSED Then
                    blnOpened = True
                    DoCmd.OpenModule doc.Name ' Open each module
                End If
                For intModcount = 0 To Modules.Count - 1
                    If doc.Name = Modules(intModcount).Name Then ' find the right module
                        mdbdListCodeBlocks Modules(intModcount), clsfh ' and write the data out
                    End If
                Next intModcount
                If blnOpened = True Then DoCmd.Close acModule, doc.Name ' and close the module
            Next doc
            clsfh.WriteData RETURN_TO_CONTENTS
            SysCmd acSysCmdClearStatus
        End If
        DoEvents

        ' ******** Form modules (ie modules behind forms) ********
        ' To get to form modules, we must open the form which it relates to, and then open the module
        ' that is behind it, but only if there is a module.

        If mdbdGetPreference("ProcessFormModules") = PREFERENCE_ENABLED Then
            Set cnt = db.Containers!Forms
            clsfh.WriteData "<h2 id=""mdbFormModules"">Form Modules</h2>"
            SysCmd acSysCmdSetStatus, APP_NAME & ": Processing form modules"
            blnWrittenHeader = False
    
            strHTML = ""
            For intCount = 0 To Forms.Count - 1 ' Loop through each form in the database
                blnOpened = False
                If SysCmd(acSysCmdGetObjectState, acForm, Forms(intCount).Name) = OBJECT_CLOSED Then
                    blnOpened = True
                    DoCmd.OpenForm Forms(intCount).Name, acDesign, windowmode:=acHidden  ' and open it in design mode.
                End If
                If Forms(intCount).Properties("HasModule") = True Then ' Check to see if there is a module
                    DoCmd.OpenModule Forms(intCount).Module ' and open the related module
                    If blnWrittenHeader = False Then ' Write the header information
                        clsfh.WriteData "<table class=""mdbtable"">"
                        clsfh.WriteData "<caption>Details of Form Modules within this database</caption>"
                        clsfh.WriteData "<thead>"
                        clsfh.WriteData "<tr><th scope=""col"">Form Name</th><th scope=""col"">Module Name</th><th scope=""col"">No. Lines</th></tr>"
                        clsfh.WriteData "</thead>"
                        clsfh.WriteData "<tbody>"
                        blnWrittenHeader = True
                    End If
                    strHTML = "<tr><td><a id=" & Chr$(QUOTE_MARK) & Replace(Forms(intCount).Name, " ", "%20") & Chr$(QUOTE_MARK) & "</a></td><td>" & Forms(intCount).Form.Module & "</td><td class=""rightNum"">" & Modules(Forms(intCount).Module).CountOfDeclarationLines + Modules(Forms(intCount).Module).CountOfLines & "</td></tr>"
                    DoCmd.Close acModule, Forms(intCount).Module
                    If blnOpened = True Then DoCmd.Close acForm, Forms(intCount).Name
                    clsfh.WriteData strHTML
                End If
            Next intCount
            If blnWrittenHeader = True Then
                clsfh.WriteData "</tbody>"
                clsfh.WriteData "</table>"
            End If
            DoEvents
            
            ' ******** Form module detail (function/sub/property routines) ********
            For Each doc In cnt.Documents
                blnOpened = False
                If SysCmd(acSysCmdGetObjectState, acForm, doc.Name) = OBJECT_CLOSED Then
                    blnOpened = True
                    DoCmd.OpenForm doc.Name, acDesign
                End If
                If Forms(doc.Name).Properties("HasModule") = True Then
                    DoCmd.OpenModule Forms(doc.Name).Module
                    mdbdListCodeBlocks Forms(doc.Name).Module, clsfh
                    DoCmd.Close acModule, Forms(doc.Name).Module
                End If
                If blnOpened = True Then DoCmd.Close acForm, Forms(doc.Name).Name
            Next doc

            clsfh.WriteData RETURN_TO_CONTENTS
            SysCmd acSysCmdClearStatus
        End If
        DoEvents

        ' ******** Report modules (ie modules behind reports) ********
        ' To get to report modules, we must open the report which it relates to, and then open the module
        ' that is behind it, but only if there is a module.
    
        If mdbdGetPreference("ProcessReportModules") = PREFERENCE_ENABLED Then
            blnWrittenHeader = False
            clsfh.WriteData "<h2 id=""mdbReportModules"">Report Modules</h2>"
            SysCmd acSysCmdSetStatus, APP_NAME & ": Processing Report modules"

            Set cnt = db.Containers!Reports
            strHTML = ""
            For Each doc In cnt.Documents
                blnOpened = False
                If SysCmd(acSysCmdGetObjectState, acReport, doc.Name) = OBJECT_CLOSED Then
                    blnOpened = True
                    DoCmd.OpenReport doc.Name, acViewDesign
                End If
                If Reports(doc.Name).Properties("HasModule") = True Then
                    
                    If blnWrittenHeader = False Then
                        clsfh.WriteData "<table class=""mdbtable"">"
                        clsfh.WriteData "<caption>Details of reports within this database</caption>"
                        clsfh.WriteData "<thead>"
                        clsfh.WriteData "<tr><th scope=""col"">Report Name</th><th scope=""col"">Module Name</th><th scope=""col"">No. Lines</th></tr>"
                        clsfh.WriteData "</thead>"
                        clsfh.WriteData "<tbody>"
                        blnWrittenHeader = True
                    End If

                    DoCmd.OpenModule Reports(doc.Name).Module
                    For intModcount = 0 To Modules.Count - 1
                        If Modules(intModcount).Name = Reports(intCount).Module Then
                            strHTML = "<tr><td>" & Reports(intCount).Name & "</td><td>" & Reports(intCount).Module & "</td><td class=""rightNum"">" & Modules(intModcount).CountOfDeclarationLines + Modules(intModcount).CountOfLines & "</td></tr>"
                        End If
                    Next intModcount
                    clsfh.WriteData strHTML
                    DoCmd.Close acModule, Reports(intCount).Module
                End If
                If blnOpened = True Then DoCmd.Close acReport, doc.Name
            Next doc
            If blnWrittenHeader = True Then
                clsfh.WriteData "</tbody>"
                clsfh.WriteData "</table>"
            End If

            For Each doc In cnt.Documents
                blnOpened = False
                If SysCmd(acSysCmdGetObjectState, acReport, doc.Name) = OBJECT_CLOSED Then
                    DoCmd.OpenReport doc.Name, acViewDesign
                    blnOpened = True
                End If
                If Reports(doc.Name).Properties("HasModule") = True Then
                    DoCmd.OpenModule Reports(doc.Name).Module
                    mdbdListCodeBlocks Reports(doc.Name).Module, clsfh
                    DoCmd.Close acModule, Reports(doc.Name).Module
                End If
                If blnOpened = True Then DoCmd.Close acReport, doc.Name
            Next doc
            
            If blnWrittenHeader = False Then clsfh.WriteData "<p>There are no report modules in this database.</p>"
            
            clsfh.WriteData RETURN_TO_CONTENTS
            SysCmd acSysCmdClearStatus
        End If
        DoEvents

        ' ******** Macros ********
        If mdbdGetPreference("ProcessMacros") = PREFERENCE_ENABLED Then
            clsfh.WriteData "<h2 id=""mdbMacros"">Macros</h2>"
            Set cnt = db.Containers!Scripts
            If cnt.Documents.Count > 0 Then
                SysCmd acSysCmdSetStatus, APP_NAME & ": Processing macros"
                clsfh.WriteData "<table class=""mdbtable"">"
                clsfh.WriteData "<caption>Details of Macros within this database</caption>"
                clsfh.WriteData "<thead>"
                clsfh.WriteData "<tr><th scope=""col"">Name</th><th scope=""col"">Date Created</th><th scope=""col"">Last Updated</th><th scope=""col"">Description</th></tr>"
                clsfh.WriteData "</thead>"
                clsfh.WriteData "<tbody>"
                For Each doc In cnt.Documents
                    clsfh.WriteData "<tr><td>" & doc.Name & "</td><td>" & doc.DateCreated & "</td><td>" & doc.LastUpdated & "</td><td>" & mdbdGetDescription(doc) & "</td></tr>"
                Next doc
                clsfh.WriteData "</tbody>"
                clsfh.WriteData "</table>"
                SysCmd acSysCmdClearStatus
            Else
                clsfh.WriteData "<p>There are no macros in this database.</p>"
            End If
            clsfh.WriteData RETURN_TO_CONTENTS
        End If
        DoEvents

        ' ******** Relationships ********
        If mdbdGetPreference("ProcessRelationships") = PREFERENCE_ENABLED Then
            clsfh.WriteData "<h2 id=""mdbRelationships"">Relationships</h2>"
            If db.Relations.Count > 0 Then
                SysCmd acSysCmdSetStatus, APP_NAME & ": Processing relationships"
                clsfh.WriteData "<table class=""mdbtable"">"
                clsfh.WriteData "<caption>Details of table relationships within this database</caption>"
                clsfh.WriteData "<thead>"
                clsfh.WriteData "<tr><th scope=""col"">Primary Name</th><th scope=""col"">Foreign Name</th><th scope=""col"">Description</th></tr>"
                clsfh.WriteData "</thead>"
                clsfh.WriteData "<tbody>"
                For Each rel In db.Relations
                    clsfh.WriteData "<tr><td>" & rel.Table & "." & rel.Fields(0).Name & "</td><td>" & rel.ForeignTable & "." & rel.Fields(0).ForeignName & "</td><td>" & mdbdGetRelType(rel) & "</td></tr>"
                Next rel
                clsfh.WriteData "</tbody>"
                clsfh.WriteData "</table>"
                SysCmd acSysCmdClearStatus
            Else
                clsfh.WriteData "<p>There are no relationships in this database.</p>"
            End If
            clsfh.WriteData RETURN_TO_CONTENTS
        End If
        DoEvents

        ' ******** Command Bars ********
        If mdbdGetPreference("ProcessCommandBars") = PREFERENCE_ENABLED Then
            clsfh.WriteData "<h2 id=""mdbCommandBars"">Command Bars</h2>"
            clsfh.WriteData "<p>Please note that " & APP_NAME & " only displays information about custom command bars.</p>"
            If Application.CommandBars.Count > 0 Then
                SysCmd acSysCmdSetStatus, APP_NAME & ": Processing Command Bars"
                blnWrittenHeader = False
        
                For Each CBR In Application.CommandBars
                    If CBR.BuiltIn = False Then
                        If blnWrittenHeader = False Then
                            clsfh.WriteData "<table class=""mdbtable"">"
                            clsfh.WriteData "<caption>Details of custom CommandBars within this database</caption>"
                            clsfh.WriteData "<thead>"
                            clsfh.WriteData "<tr><th scope=""col"">Name</th><th scope=""col"">Visible</th><th scope=""col"">NameLocal</th></tr>"
                            clsfh.WriteData "</thead>"
                            clsfh.WriteData "<tbody>"
                            blnWrittenHeader = True
                        End If
                        clsfh.WriteData "<tr><td>" & CBR.Name & "</td><td>" & CBR.Visible & "</td><td>" & CBR.NameLocal & "</td></tr>"
                    End If
                Next CBR
    
                If blnWrittenHeader = True Then
                    clsfh.WriteData "</tbody>"
                    clsfh.WriteData "</table>"
                Else
                    clsfh.WriteData "<p>There are no custom command bars in this database.</p>"
                End If
                SysCmd acSysCmdClearStatus
            Else
                clsfh.WriteData "<p>There are no custom command bars in this database.</p>"
            End If
            clsfh.WriteData RETURN_TO_CONTENTS
        End If
        DoEvents

        ' *********** Data Access Pages (new in 1.31 for Access 2000/XP/2003)
        ' Not in acc97 version as the feature isn't supported.
    
        If mdbdGetPreference("ProcessDAPS") = PREFERENCE_ENABLED Then
            blnWrittenHeader = False
            ' Process Data Access Pages
            clsfh.WriteData "<h2 id=""mdbDataAccessPages"">Data Access Pages</h2>"
        
            If mdbdGetPreference("FormatDAPConnection") = PREFERENCE_ENABLED Then
                blnNewlineInDAPConnection = True
            Else
                blnNewlineInDAPConnection = False
            End If
            
            Set cnt = db.Containers!DataAccessPages
        
            If cnt.Documents.Count > 0 Then
                SysCmd acSysCmdSetStatus, APP_NAME & ": Processing Data Access Pages"
                For Each doc In cnt.Documents
                    If Left$(doc.Name, 4) <> "~TMP" Then
                        ' Exclude temporary files
                        If blnWrittenHeader = False Then
                            clsfh.WriteData "<table class=""mdbtable"">"
                            clsfh.WriteData "<caption>Details of Data Access Pages within this database</caption>"
                            clsfh.WriteData "<thead>"
                            clsfh.WriteData "<tr><th scope=""col"">Name</th><th scope=""col"">Date Created</th><th scope=""col"">Last Updated</th><th scope=""col"">Connection String</th><th scope=""col"">Description</th></tr>"
                            clsfh.WriteData "</thead>"
                            clsfh.WriteData "<tbody>"
                            blnWrittenHeader = True
                        End If
                        DoCmd.OpenDataAccessPage doc.Name, acDataAccessPageDesign
                        Set dap = DataAccessPages(doc.Name)
                        If blnNewlineInDAPConnection = True Then
                            ' replace semi colons in connection string with semi colon and <br> tag to force new line
                            clsfh.WriteData "<tr><td>" & dap.Name & "</td><td>" & doc.DateCreated & "</td><td>" & doc.LastUpdated & "</td><td>" & mdbdReplace2(dap.ConnectionString, ";", ";<br>") & "</td><td>" & mdbdGetDescription(doc) & "</td></tr>"
                        Else
                            clsfh.WriteData "<tr><td>" & dap.Name & "</td><td>" & doc.DateCreated & "</td><td>" & doc.LastUpdated & "</td><td>" & dap.ConnectionString & "</td><td>" & mdbdGetDescription(doc) & "</td></tr>"
                        End If
                        DoCmd.Close acDataAccessPage, doc.Name, acSaveNo
                    End If
                Next doc
                If blnWrittenHeader = True Then
                    clsfh.WriteData "</tbody>"
                    clsfh.WriteData "</table>"
                End If
            Else
                clsfh.WriteData "<p>There are no Data Access Pages in this database.</p>"
            End If
            SysCmd acSysCmdClearStatus
            clsfh.WriteData RETURN_TO_CONTENTS
        End If
        DoEvents
    
        ' ******** References ********
        ' new in 1.4x
        ' Although references are not strictly speaking part of an Access database, they do
        ' have an effect on the features that can be called from VBA routines in external libraries.
        ' Newer versions of specific modules add features that mean problematic use in earlier versions
        ' eg ADO 2.6 adds the ability to run remote stored procedures with named parameters, but this isn't
        ' possible in earlier versions unless you build your own query string and execute it.
        ' Also DAO 3.6 adds features over 3.51.
        
        If mdbdGetPreference("ProcessReferences") = PREFERENCE_ENABLED Then
            clsfh.WriteData "<h2 id=""mdbReferences"">References</h2>"
            If Application.References.Count > 0 Then
                SysCmd acSysCmdSetStatus, APP_NAME & ": Processing References"
                blnWrittenHeader = False
        
                For intReferenceCount = 1 To Application.References.Count
                    If blnWrittenHeader = False Then
                        clsfh.WriteData "<table class=""mdbtable"">"
                        clsfh.WriteData "<caption>Details of VBA references within this database</caption>"
                        clsfh.WriteData "<thead>"
                        clsfh.WriteData "<tr><th scope=""col"">Name</th><th scope=""col"">Full Path</th><th scope=""col"">Reference Broken</th></tr>"
                        clsfh.WriteData "</thead>"
                        clsfh.WriteData "<tbody>"
                        blnWrittenHeader = True
                    End If
                    clsfh.WriteData "<tr><td>" & GetGuidDescription(Application.References(intReferenceCount).GUID, Application.References(intReferenceCount).Major, _
                        Application.References(intReferenceCount).Minor) & "</td><td>" _
                            & Application.References(intReferenceCount).FullPath & "</td><td>" & Application.References(intReferenceCount).IsBroken & "</td></tr>"
                Next intReferenceCount
    
                If blnWrittenHeader = True Then
                    clsfh.WriteData "</tbody>"
                    clsfh.WriteData "</table>"
                Else
                    ' Given the defaults I'd hardly ever expect this section to get executed, but there's always a first time.
                    clsfh.WriteData "<p>There are no references in this database.</p>"
                End If
                SysCmd acSysCmdClearStatus
            Else
                clsfh.WriteData "<p>There are no references in this database.</p>"
            End If
            clsfh.WriteData RETURN_TO_CONTENTS
        End If
        DoEvents

        ' ********  Ribbons ********
        ' New in version 1.60, this is why it requires Access 2007 or later.
        ' There are two ways of setting ribbons:
        ' 1) the ribbon XML code can be stored in a table called USysRibbons under a given name, and it can be confirmed
        ' through code by explicitly assigning the XML to the relevant
        ' or 2) VBA can be assigned the XML directly.
        ' This particular code will only cover case 1.
        If mdbdGetPreference("ProcessRibbons") = PREFERENCE_ENABLED Then
            clsfh.WriteData "<h2 id=""mdbRibbons"">Ribbons</h2>"
            If LoadRibbons(clsfh, strHaltOnErrors) = -1 Then clsfh.WriteData "<p>There are no ribbons in this database</p>"
            clsfh.WriteData RETURN_TO_CONTENTS
        End If
        
        ' ******** Import/Export specifications ********
        ' New in version 1.60, this is why it requires Access 2007 or later.
        If mdbdGetPreference("ProcessImpExpSpecs") = PREFERENCE_ENABLED Then
            clsfh.WriteData "<h2 id=""mdbImpExpSpecs"">Import/Export Specifications</h2>"
            If CurrentProject.ImportExportSpecifications.Count > 0 Then
                SysCmd acSysCmdSetStatus, APP_NAME & ": Processing Import/Export specifications"
                blnWrittenHeader = False
                For intReferenceCount = 0 To CurrentProject.ImportExportSpecifications.Count
                    If blnWrittenHeader = False Then
                        clsfh.WriteData "<table class=""mdbtable"">"
                        clsfh.WriteData "<caption>Details of Import/Export specifications within this database</caption>"
                        clsfh.WriteData "<thead>"
                        clsfh.WriteData "<tr><th scope=""col"">Name</th><th>XML</th></tr>"
                        clsfh.WriteData "</thead>"
                        clsfh.WriteData "<tbody>"
                        blnWrittenHeader = True
                    End If
                    clsfh.WriteData "<tr><td>" & mdbdReplaceSpecialChars(CurrentProject.ImportExportSpecifications(intReferenceCount).Name) & "</td><td><code>" & mdbdReplaceSpecialChars(CurrentProject.ImportExportSpecifications(intReferenceCount).XML) & "</code></td></tr>"
                Next intReferenceCount
    
                If blnWrittenHeader = True Then
                    clsfh.WriteData "</tbody>"
                    clsfh.WriteData "</table>"
                Else
                    clsfh.WriteData "<p>There are no Import/Export Specifications in this database.</p>"
                End If
                SysCmd acSysCmdClearStatus
            Else
                clsfh.WriteData "<p>There are no Import/Export Specifications in this database.</p>"
            End If
            clsfh.WriteData RETURN_TO_CONTENTS
        End If
        DoEvents

        ' ******** Append the footer to the end of the report. ********
        clsfh.WriteData "<footer>"
        clsfh.WriteData "<p>This Database information report was generated using " & APP_NAME & " version " & mdbdGetPreference("MDBDocVersion") & ".  " & APP_NAME & " is Copyright &copy; 2007-14 John Barnett.</p>"
        clsfh.WriteData "</footer>"
        clsfh.WriteData "</body>"
        clsfh.WriteData "</html>"
        
        ' ******** Now tidy up ********
        clsfh.CloseFile
        DoCmd.Echo True
        Set db = Nothing
        SysCmd acSysCmdClearStatus
        
        ' Close VBA editor window if it was opened by this function
        If blnVBEVisible = False Then
            Application.VBE.MainWindow.Visible = False
        End If
        ' ******** Ask user if they want to view the report. ********
        If MsgBox(APP_NAME & " has saved its report to " & strFileName & ". Would you like to view it now?", vbYesNo + vbQuestion) = vbYes Then
            Application.FollowHyperlink Address:=strFileName, addhistory:=False
        End If
    Else    ' Halt On Preference has invalid value
        MsgBox "Error in preference section: HaltOnErrors does not exist or has invalid value. Please check your preferences screen. This has not been run", vbOKOnly + vbCritical, APP_NAME
    End If
    
End Function

Private Function mdbdGetPropertyType(prp As DAO.Property) As String
    'MDBDOC: Returns descriptive property type given supplied property.
    ' Function: GetPropertyType
    ' Scope:    Private
    ' Parameters: prp (property) - the property to examine
    ' Return Value: String - descriptive value of property.
    ' Author:   John Barnett
    ' Date:     4 July 2001.
    ' Description: GetPropertyType returns a descriptive version of the data type of a property object.
    ' Called by: Database properties section of ProcessDatabase

    ' These values have been taken straight out of the Access help file.
    Select Case prp.Type
        Case vbEmpty
            mdbdGetPropertyType = "Empty"
        Case vbNull
            mdbdGetPropertyType = "Yes/No"
        Case vbInteger
            mdbdGetPropertyType = "Integer"
        Case vbLong
            mdbdGetPropertyType = "Long Integer"
        Case vbSingle
            mdbdGetPropertyType = "Single"
        Case vbDouble
            mdbdGetPropertyType = "Double"
        Case vbCurrency
            mdbdGetPropertyType = "Currency"
        Case vbDate
            mdbdGetPropertyType = "Date/Time"
        Case vbVariant
            mdbdGetPropertyType = "Memo"
        Case vbByte
            mdbdGetPropertyType = "Byte"
        Case dbText
            mdbdGetPropertyType = "String"
        Case vbObject
            mdbdGetPropertyType = "OLE Object"
        Case vbBoolean
            mdbdGetPropertyType = "Boolean"
        Case 10
            mdbdGetPropertyType = "String"
        Case dbGUID
            mdbdGetPropertyType = "GUID"
        Case Else
            mdbdGetPropertyType = "Unknown property type " & prp.Type
    End Select

End Function

Private Function mdbdGetFieldType(fld As DAO.Field, intSize As Integer) As String
    'MDBDOC: Returns descriptive field type from supplied field.
    ' Function: mdbdGetFieldType
    ' Scope:    Private
    ' Parameters: fld (Field) - the field to examine
    ' Return Value: String - descriptive value of field.
    ' Author:   John Barnett
    ' Date:     5 July 2001.
    ' Description: mdbdGetFieldType returns a descriptive version of the data type of a field object.
    ' Called by: Fields section of ProcessDatabase

    Select Case fld.Type
        Case dbText
            mdbdGetFieldType = "Text (" & intSize & ")"
        Case dbBigInt
            mdbdGetFieldType = "Big Integer"
        Case dbBinary
            mdbdGetFieldType = "Binary"
        Case dbBoolean
            mdbdGetFieldType = "Boolean"
        Case dbByte
            mdbdGetFieldType = "Byte"
        Case dbChar
            mdbdGetFieldType = "Char"
        Case dbCurrency
            mdbdGetFieldType = "Currency"
        Case dbDate
            mdbdGetFieldType = "Date / Time"
        Case dbDecimal
            mdbdGetFieldType = "Decimal"
        Case dbDouble
            mdbdGetFieldType = "Double"
        Case dbFloat
            mdbdGetFieldType = "Float"
        Case dbGUID
            mdbdGetFieldType = "Guid"
        Case dbInteger
            mdbdGetFieldType = "Integer"
        Case dbLong
            mdbdGetFieldType = "Long"
        Case dbLongBinary
            mdbdGetFieldType = "Long Binary (OLE Object)"
        Case dbMemo
            mdbdGetFieldType = "Memo"
        Case dbNumeric
            mdbdGetFieldType = "Numeric"
        Case dbSingle
            mdbdGetFieldType = "Single"
        Case dbTime
            mdbdGetFieldType = "Time"
        Case dbTimeStamp
            mdbdGetFieldType = "Time Stamp"
        Case dbVarBinary
            mdbdGetFieldType = "VarBinary"
        Case Else
            mdbdGetFieldType = "Unknown field type " & fld.Type
    End Select

    ' If the field has specific attributes, override the ones specified above.
    If fld.Attributes And dbAutoIncrField Then
        mdbdGetFieldType = "Auto Number"
    End If
    If fld.Attributes And dbHyperlinkField Then
        mdbdGetFieldType = "Hyperlink"
    End If

End Function

Private Function mdbdGetQueryType(qry As DAO.QueryDef) As String
    'MDBDOC: Returns descriptive query type.
    ' Function: mdbdGetQueryType
    ' Scope:    Private
    ' Parameters: qry (Query) - the query to examine
    ' Return Value: String - descriptive value of query.
    ' Author:   John Barnett
    ' Date:     5 July 2001.
    ' Description: GetQueryType returns a descriptive version of a querydef (stored query) object..
    ' Called by: Queries section of ProcessDatabase

    ' This function is a simple SELECT case on the query type.  The constants are predefined - search the access help
    ' for them.
    Select Case qry.Type
        Case dbQAction
            mdbdGetQueryType = "Action"
        Case dbQAppend
            mdbdGetQueryType = "Append"
        Case dbQCompound
            mdbdGetQueryType = "Compound"
        Case dbQCrosstab
            mdbdGetQueryType = "Crosstab"
        Case dbQDDL
            mdbdGetQueryType = "Data definition"
        Case dbQDelete
            mdbdGetQueryType = "Delete"
        Case dbQMakeTable
            mdbdGetQueryType = "Make Table"
        Case dbQProcedure
            mdbdGetQueryType = "Procedure (ODBCDirect workspaces only)"
        Case dbQSelect
            mdbdGetQueryType = "Select"
        Case dbQSetOperation
            mdbdGetQueryType = "Union"
        Case dbQSPTBulk  'Used with dbQSQLPassThrough to specify a query that doesn't return records (Microsoft Jet workspaces only).
            mdbdGetQueryType = "SPTBulk - a pass through query that doesn't return a result set"
        Case dbQSQLPassThrough
            mdbdGetQueryType = "Pass-through (Microsoft Jet workspaces only)"
        Case dbQUpdate
            mdbdGetQueryType = "Update"
        Case Else
            mdbdGetQueryType = "Unknown Query type " & qry.Type
    End Select

End Function

Private Function mdbdGetTableType(tdf As DAO.TableDef) As String
    'MDBDOC: Returns descriptive table type.
    ' Function: mdbdGetTableType
    ' Scope:    Private
    ' Parameters: tbl (Table) - the table to examine
    ' Return Value: String - descriptive value of table.
    ' Author:   John Barnett
    ' Date:     6 July 2001.
    ' Description: GetTableType returns a descriptive version of a tabledef object..
    ' Called by: Tables section of ProcessDatabase

    ' This function can't be done as a Select Case like others above because these
    ' options are not mutually exclusive.

    ' Set description to empty
    mdbdGetTableType = ""

    If tdf.Attributes And dbAttachExclusive Then
        mdbdGetTableType = "Attached for exclusive use"
    End If
    If tdf.Attributes And dbAttachSavePWD Then
        mdbdGetTableType = mdbdGetTableType & "Attached (connection information saved)"
    End If
    If tdf.Attributes And dbSystemObject Then
        mdbdGetTableType = mdbdGetTableType & "MS Access System Object"
    End If
    If tdf.Attributes And dbHiddenObject Then
        mdbdGetTableType = mdbdGetTableType & "Hidden object"
    End If
    If tdf.Attributes And dbAttachedTable Then
        mdbdGetTableType = mdbdGetTableType & "Attached table (non ODBC)"
    End If
    If tdf.Attributes And dbAttachedODBC Then
        mdbdGetTableType = mdbdGetTableType & "Attached table via ODBC"
    End If
    If tdf.Attributes = 0 Then
        mdbdGetTableType = "Normal table"
        
    End If
End Function

Private Function mdbdGetRelType(rel As DAO.Relation) As String
    'MDBDOC: Returns descriptive relationship type.
    ' Function: mdbdGetRelType
    ' Scope:    Private
    ' Parameters: rel (Relation) - the relationship to examine
    ' Return Value: String - descriptive value of relationship.
    ' Author:   John Barnett
    ' Date:     8 July 2001.
    ' Description: GetRelType returns a descriptive version of a Relationship object..
    ' Called by: Relationships section of ProcessDatabase

    ' This function cannot be written using a SELECT CASE because the options are not
    ' mutually exclusive.
    If rel.Attributes And dbRelationUnique Then
        mdbdGetRelType = " 1:1 Relationship"
    End If

    If rel.Attributes And dbRelationDontEnforce Then
        mdbdGetRelType = mdbdGetRelType & " Referential integrity not enforced"
    End If

    If rel.Attributes And dbRelationUpdateCascade Then
        mdbdGetRelType = mdbdGetRelType & " Cascade updates"
    End If

    If rel.Attributes And dbRelationDeleteCascade Then
        mdbdGetRelType = mdbdGetRelType & " Cascade deletes"
    End If

    mdbdGetRelType = Trim$(mdbdGetRelType)
End Function

Private Function mdbdIsPK(tdf As DAO.TableDef, fld As DAO.Field) As Boolean
    'MDBDOC: Returns True/False depending whether a specific field is part of a specific tables primary key.
    ' Function: mdbdIsPK
    ' Scope:    Private
    ' Parameters: tdf - Tabledef; fld Field
    ' Return Value: Boolean - True if fld is in primary key of table tdf, false otherwise
    ' Author:   John Barnett
    ' Date:     10 July 2001.
    ' Description: IsPK returns a True/False value depending whether field fld is part of the primary key of table pk.
    ' Called by: Fields section of ProcessDatabase

    Dim Idx As Index

    mdbdIsPK = False
    For Each Idx In tdf.Indexes
        ' Loop through each of the table indexes
        If InStr(Idx.Fields, fld.Name) Then
            ' if the field name is in the index and the index is the primary one then it is true
            If Idx.Primary = True Then
                mdbdIsPK = True
                Exit For
            End If
        End If
    Next Idx
End Function

Private Function mdbdGetDescription(obj As Object) As String
    'MDBDOC: Returns the Description property of an object.
    ' Function: mdbdGetDescription
    ' Scope:    Private
    ' Parameters: obj - any object
    ' Return Value: String - the value of the description property.
    ' Author:   John Barnett
    ' Date:     10 July 2001 amended 01 November 2003.
    ' Description: mdbdGetDescription returns the description property of the object passed as a parameter.
    ' Called by: Tables, Fields, Macros, Reports and Forms sections of ProcessDatabase

    On Error Resume Next

    mdbdGetDescription = mdbdReplaceSpecialChars(obj.Properties("Description") & "")
End Function

Public Function StartMDBDoc() As Boolean
    'MDBDOC: Function called from Addins menu to start MDB Doc running.
    ' Function: StartMDBDoc
    ' Scope:    Public
    ' Parameters: None
    ' Return Value: None
    ' Author:   John Barnett
    ' Date:     5 July 2001.
    ' Description: This is called from the Addins menu as registered in the USysRegInfo table. All it does is open the Startup form.
    ' Called by: Addins menu.
    
    DoCmd.OpenForm "frmMdbdocStartup", acNormal, windowmode:=acDialog
    StartMDBDoc = True
End Function

Private Function mdbdListCodeBlocks(mdl As Module, outfile As mdbdclsFileHandle) As String
    'MDBDOC: Provides a list of code blocks within a module and writes them to an output file.
    ' Function: mdbdListCodeBlocks
    ' Scope:    Private
    ' Parameters: mdl - Module to process; outfile - mdbdclsfilehandle to process data with.
    ' Return Value: String
    ' Author:   John Barnett
    ' Date:     22 July 2001.
    ' Description: mdbdListCodeBlocks parses the source code of a module and lists out the
    ' names, prototypes & return values of sub, function and property procedures.
    ' Called by: Modules, Form Modules and Report Modules section of mdbdProcessDatabase.
    '
    ' As this module is quite complex, extra comments are included below indicating how it works.

    ' Dimensions of array for holding variables
    Const ARRAY_MIN_SIZE As Integer = 1
    Const ARRAY_MAX_SIZE As Integer = 25

    Dim strCurrentLine As String                ' Current line of code being processed
    Dim strDescription As String                ' MDBDOC Description tag from current module
    Dim intCount As Integer                     ' Counter - current line throughout the module
    Dim intStartLine As Integer                 ' Line of module that current module starts (used for calculating length)
    Dim intStart As Integer                     ' First line of module to start processing at
    Dim blnWrittenHeader As Boolean             ' Has a table header been written for this module?
    Dim blnProcHeader As Boolean                ' Is the current line of code a procedure header?
    Dim intLoop As Integer
    Dim strHTML As String                       ' HTML fragment being dynamically built.

    Static strStartCombinations(ARRAY_MIN_SIZE To ARRAY_MAX_SIZE) As String ' Create static array as this can be used many times.

    Dim strCommentTag As String                 ' Value of Comment tag from Preferences
    Dim intCTagLength As Integer                ' Length (in characters) of comment tag
    
    ' Insert the data into the array
    strStartCombinations(1) = "Property Get" ' Property Get routines
    strStartCombinations(2) = "Public Property Get"
    strStartCombinations(3) = "Public Static Property Get"
    strStartCombinations(4) = "Private Static Property Get"
    strStartCombinations(5) = "Private Property Get"

    strStartCombinations(6) = "Property Let" ' Property Let routines
    strStartCombinations(7) = "Public Property Let"
    strStartCombinations(8) = "Public Static Property Let"
    strStartCombinations(9) = "Private Property Let"
    strStartCombinations(10) = "Private Static Property Let"

    strStartCombinations(11) = "Property Set" ' Property Set routines
    strStartCombinations(12) = "Public Property Set"
    strStartCombinations(13) = "Public Static Property Set"
    strStartCombinations(14) = "Private Property Set"
    strStartCombinations(15) = "Private Static Property Set"

    strStartCombinations(16) = "Function" ' Functions
    strStartCombinations(17) = "Public Function"
    strStartCombinations(18) = "Private Function"
    strStartCombinations(19) = "Public Static Function"
    strStartCombinations(20) = "Private Static Function"

    strStartCombinations(21) = "Sub" ' Subs
    strStartCombinations(22) = "Public Sub"
    strStartCombinations(23) = "Private Sub"
    strStartCombinations(24) = "Public Static Sub"
    strStartCombinations(25) = "Private Static Sub"

    ' Retrieve the comment tag from preferences and compute its length (so it knows where to start looking from for the text)
    strCommentTag = UCase$(mdbdGetPreference("CommentTag"))
    intCTagLength = Len(strCommentTag)
    
    ' Means of starting in case there are no declaration lines in the module.
    ' Only likely if "Require variable declaration" is switched off.
    If mdl.CountOfDeclarationLines > 0 Then
        intStart = mdl.CountOfDeclarationLines
    Else
        intStart = 1
    End If

    ' Loop through the code in the module, line by line avoiding the declaration section
    For intCount = intStart To mdl.CountOfLines
        If Len(mdl.Lines(intCount, 1) & "") > 0 Then
            strCurrentLine = mdl.Lines(intCount, 1) & "" ' Grab the current line
            strCurrentLine = Trim$(strCurrentLine)  ' remove any leading / trailing spaces
        Else
            strCurrentLine = ""
        End If
    
        ' Now retrive MDBDoc's procedure/sub/function/property comments.
        ' use CommentTag preference to change name if needed
    
        If UCase$(Left$(strCurrentLine, intCTagLength)) = strCommentTag Then
            ' comments start 'MDBDOC: by default but could be different (introduced 1.4)
            strDescription = Trim(Mid$(strCurrentLine, intCTagLength + 1))
        End If
    
        If Left$(strCurrentLine, 1) <> "'" And Left$(strCurrentLine, 3) <> "Rem" Then
            ' Ignore comments
        
            Do While Right$(strCurrentLine, 1) = "_" ' If continuation char at end of line - grab continual until none left.
                ' grab the next line
                strCurrentLine = Left$(strCurrentLine, Len(strCurrentLine) - 1)  ' strip it off
                intCount = intCount + 1
                strCurrentLine = strCurrentLine & mdl.Lines(intCount, 1) ' and grab the next one
            Loop
            blnProcHeader = False ' Set default of line header =false

            ' but now find out what type it is by comparing it to each element in the array in turn.
            ' Update for 1.4.x - this section has been greatly simplified, resulting in quite a performance improvement.
            
            For intLoop = ARRAY_MIN_SIZE To ARRAY_MAX_SIZE
                If Left$(strCurrentLine, Len(strStartCombinations(intLoop))) = strStartCombinations(intLoop) Then
                    ' Match found, so set the header line and record the start line of the current procedure
                    blnProcHeader = True
                    intStartLine = intCount
                End If
            Next intLoop

            ' If this is a header, then write one if its not been written before.
            If (blnProcHeader = True) Then
                If blnWrittenHeader = False Then
                    outfile.WriteData "<table class=""mdbtable"" id=" & Chr$(QUOTE_MARK) & Replace(mdl.Name, " ", "%20") & Chr$(QUOTE_MARK) & ">"
                    outfile.WriteData "<caption>Code routines within the module " & mdl.Name & "</caption>"
                    outfile.WriteData "<thead>"
                    outfile.WriteData "<tr><th scope=""col"">Prototype</th><th scope=""col"">No. Lines</th><th scope=""col"">Description</th></tr>"
                    outfile.WriteData "</thead>"
                    outfile.WriteData "<tbody>"
                    blnWrittenHeader = True
                End If
                strHTML = "<tr><td>" & mdbdReplaceSpecialChars(strCurrentLine) & "</td>"  ' Start the line to write off
            End If
            
            ' If the current line is an end of a sub/function/property then complete the data and write it out.
            If Left$(strCurrentLine, 7) = "End Sub" Or Left$(strCurrentLine, 12) = "End Function" Or Left$(strCurrentLine, 12) = "End Property" Then
                ' Can't just Left$ (3) = "End" here because of End If / End Loop / End Select / End Type etc.
                outfile.WriteData strHTML & "<td class=""rightNum"">" & intCount - intStartLine - 1 & "</td><td>" & mdbdReplaceSpecialChars(strDescription) & "</td></tr>"
                strDescription = "" ' Written out this data so blank out description in case next proc doesn't have one.
            End If
        End If
    Next intCount
    
    If blnWrittenHeader = True Then
        outfile.WriteData "</tbody>"
        outfile.WriteData "</table>"
    End If
    outfile.WriteData "<br>"
End Function

Private Sub mdbdCloseObjects()
    'MDBDOC: This sub closes any open objects prior to running
    ' Sub: mdbdCloseObjects
    ' Scope:    Private
    ' Parameters: None
    ' Author:   John Barnett
    ' Date:     24 October 2003.
    '           Amended 9 February 2008 to loop from 0 to intcount in forms and reports; bug
    ' Description: Closes any open objects, prompting for saving. this is because open objects before MDBDoc starts running can cause problems.
    ' Called by: ProcessDatabase, before it starts "processing"

    Dim intCount As Integer

    ' Close all open forms
    If Forms.Count > 0 Then
        For intCount = 0 To Forms.Count - 1
            DoCmd.Close acForm, Forms(intCount).Name, acSavePrompt
        Next intCount
    End If

    ' Close all open reports
    If Reports.Count > 0 Then
        For intCount = 0 To Reports.Count - 1
            DoCmd.Close acReport, Reports(intCount).Name, acSavePrompt
        Next intCount
    End If

End Sub

Public Function mdbdReplace2(Expression As String, Find As String, Replace As String, Optional Start As Long = 1) As String
    'MDBDOC: Replace2 function taken from Tek-Tips database used for search/replace of special characters with their escape sequences.
    ' Author: John Barnett, handle jrbarnett
    ' In response to: http://www.tek-tips.com/viewthread.cfm?SQID=565966&SPID=700&page=1
    ' Date: 4th June 2003
    ' Purpose: A Find and replace VBA function for Access 97 and 2000, which didn't have the Replace() function of 2002(XP).
    ' although it does work in 2002 as well.  It implements the mandatory functionality of the 2002 function plus the optional start, but not the Count of replacements and Binary Compare method.
    ' It has the same function header, so can be dropped in as a replacement.
    ' Called by mdbdReplaceSpecialChars to do the brute force search/replace work programmatically.

    Dim strResult As String ' variable to store result
    Dim intPosition As Integer ' variable to store current position in Expression
    Dim intStartPos As Integer ' variable to store current starting position within 'expression'

    If IsMissing(Start) Then
        intStartPos = 1 ' start position not supplied; so set it to beginning
    Else
        intStartPos = Start ' otherwise set it and...
        strResult = Left$(Expression, Start - 1) ' start by copying over first chars before start position
    End If

    intPosition = InStr(Expression, Find) ' locate first occurrence of 'find' data
    ' Remember that intPosition will = 0 if no occurrences are found.

    Do While intPosition > 0
        strResult = strResult & Mid$(Expression, intStartPos, (intPosition - intStartPos)) ' copy everything over from it that hasn't been copied yet
        intStartPos = intPosition + Len(Find) ' increase the pointer by the length of the "to find" data so it won't find the current occurrence
        strResult = strResult & Replace ' add the replacement data
        intPosition = InStr(intPosition + Len(Find), Expression, Find) ' and reset the position for the new start point
    Loop

    ' In case we aren't changing the very last part of the string...
    If intStartPos < Len(Expression) Then
        ' copy over rest to result
        strResult = strResult & Mid$(Expression, intStartPos)
    End If
    mdbdReplace2 = strResult ' and return a value
End Function

Public Function mdbdReplaceSpecialChars(strInputString As String) As String
    'MDBDOC: This function replaces special punctuation characters with their HTML escape equivalent
    ' Function: mdbdReplaceSpecialChars
    ' Scope:    Private
    ' Parameters: strInputString - input string
    ' Author:   John Barnett
    ' Date:     1 November 2003. Amended October 2007 to add support for trademark and (R) symbol in v1.4.
    ' Description: Repeatedly calls the mdbdReplace2 function (also written for tek-tips) to replace occurrences of special characters with their HTML escape equivalents.
    ' Called by: various parts of ProcessDatabase, for formatting description fields, SQL code, MDBDoc user comments etc. Basically anywhere where user defined text output could be written to the output file.

    Dim strTemp As String

    strTemp = mdbdReplace2(strInputString, "&", "&amp;") ' Ampersand symbol
    strTemp = mdbdReplace2(strTemp, "<", "&lt;")    ' Less than symbol
    strTemp = mdbdReplace2(strTemp, ">", "&gt;")    ' Greater than symbol
    strTemp = mdbdReplace2(strTemp, "©", "&copy;")  ' Copyright symbol
    strTemp = mdbdReplace2(strTemp, "™", "&#153;")  ' Trademark symbol
    strTemp = mdbdReplace2(strTemp, "®", "&#34;")   ' (R) reserved symbol

    mdbdReplaceSpecialChars = strTemp
End Function

Private Function LoadStylesheets(clsfh As mdbdclsFileHandle, strHaltOnError As String) As Integer
    'MDBDOC: This function loads the stylesheet specified in the preferences and writes the meta tag in the header.
    ' Function: LoadStylesheets
    ' Scope:    Private
    ' Parameters: clsfh - clsfilehandle object used to write the data out to.
    ' Author:   John Barnett
    ' Date:     10 December 2003; better documentation added 13 Oct 2007.
    ' Description: This function loads the stylesheet specified in the preferences and writes the meta tag in the header.
    ' Returns:  0 - Error somewhat; not included
    '           1 - Included path to relative link (pref: "P" - Path) or
    '           2 - included contents of stylesheet (preference set: I - include)
    
    Const STYLESHEET_PATH_PREFERENCE As String = "P" ' Preference for including path to css file
    Const STYLESHEET_INCL_PREFERENCE As String = "I" ' Preference for including css file contents
    
    Dim strStylesheet As String                     ' Path to Stylesheet on disk (from preference)
    Dim intHandle As Integer                        ' File handle (for including contents)
    Dim strCode As String                           ' Variable for temporarily holding stylesheet code
    Dim strMode As String                           ' Mode of handling stylesheet (from preferences)
    
    strStylesheet = mdbdGetPreference("StyleSheetPath")
    If Len(strStylesheet & "") > 0 Then
        ' Stylesheet path set (if not, ignore)
        strStylesheet = mdbdGetPreference("StyleSheetPath")
        strMode = UCase$(mdbdGetPreference("StyleSheetMode")) ' Retrieve stylesheet mode (include link or text)
        If strMode = STYLESHEET_PATH_PREFERENCE Then
            ' P for include path to relative link.
            ' Can't check for existence of file as it could be one of:
            ' * A URL (http://intranet/css/file.css) - most likely
            ' * Drive mapping (eg i:\intranet\file.css), less likely in an intranet environment but possible for files on disk.
            ' * A UNC path (eg \\server\intranet\file.css) - not likely for web apps
            ' As only the first of these could be easily checked, as the others need to be done with reference to any specified output path
            ' I have elected to take the easy option and not explicitly check this option.
            
            clsfh.WriteData "<link rel=""stylesheet"" type=""text/css"" href=""" & strStylesheet & """ title=""" & APP_NAME & " Stylesheet"" media=""screen, print"" />"
            LoadStylesheets = STYLESHEET_PATH
        ElseIf strMode = STYLESHEET_INCL_PREFERENCE Then
            
            ' Include mode (I) selected.  Load the contents of the file into the output HTML file.
            If Dir(strStylesheet) <> "" Then
                ' File found at location specified in preference
                intHandle = FreeFile
                Open strStylesheet For Binary Access Read As #intHandle
                ' This uses the technique demonstrated by MichaelRed in a particular thread
                ' It was years ago and I have lost the reference (will add it in if I find it). Apologies Michael.
                strCode = String$(LOF(intHandle), Space(1)) ' Set a string variable of the spaces length equivalent to the number of characters in the file
                Get #intHandle, 1, strCode ' Now read that number of bytes from the file (ie whole contents)
                Close #intHandle            ' Close the file, and write the data into the output file.
                clsfh.WriteData "<style type=""text/css"" title=""MDB Doc Stylesheet"" media=""screen, print"">" & vbCrLf & strCode & "</style>"
                LoadStylesheets = STYLESHEET_INCLUDE ' Return success code
            Else
                ' Stylesheet blank but include mode set, can't continue.
                LoadStylesheets = STYLESHEET_ERROR
                If strHaltOnError = PREFERENCE_ENABLED Then MsgBox "Cannot find stylesheet file " & strStylesheet & " specified in preferences to include in output.", vbOKOnly + vbCritical
            End If
        Else
            ' Invalid stylesheet mode detected in preferences, display error
            If strHaltOnError = PREFERENCE_ENABLED Then MsgBox "Invalid stylesheet mode specified in preferences.", vbOKOnly + vbCritical
            LoadStylesheets = STYLESHEET_ERROR
        End If
    End If

End Function

Public Function StartLocalPreferences() As Boolean
    'MDBDOC: Function called from Addins menu to start MDB Doc running.
    ' Function: StartMDBDoc
    ' Scope:    Public
    ' Parameters: None
    ' Return Value: None
    ' Author:   John Barnett
    ' Date:     26 Dec 2003.
    ' Description: This is called from the Addins menu as registered in the USysRegInfo table. All it does is open the local preferences form.
    ' Called by: Addins menu.
    
    DoCmd.OpenForm "frmLocalPreferences", acNormal
End Function

Private Function GetGuidDescription(strGuid As String, strMajorVersion As String, strMinorVersion As String) As String
    'MDBDOC: This function returns the string description of a specific GUID and version number.
    ' Function: GetGuidDescription
    ' Scope: Public
    ' Parameters: strGuid - GUID of the reference, strMajorVersion - Major version number, strMinorVersion - Minor version number
    ' Return value: String - the description of the Guid (from the Windows Registry)
    ' Date: 13 October 2007
    ' This requires the API functions wrapped in modRegistry to work.
    ' Called from: mdbdProcessDatabase

    ' It is used in the functionality to retrieve the description of a reference.
    ' This is stored in the Registry as:
    ' HKEY_CLASSES_ROOT\TypeLib\<guid>\<majorversion>.<minorversion>
    GetGuidDescription = QueryValue(HKEY_CLASSES_ROOT, "TypeLib\" & strGuid & "\" & strMajorVersion & "." & strMinorVersion, "")
End Function

Public Function FormatSQL(strSQL As String, blnNewlinePerField As Boolean) As String
'MDBDOC: SQL Formatting function for Access VBA, J Barnett, Oct 2007.
' ******************************************************
' SQL Formatting Function by J Barnett,
' Date: 26-27 October 2007
' Language: VBA (Access). Requires Replace() function so needs Acc2000 + but could use alternative
' Purpose: Formats an SQL statement. Puts a new line between each keyword (in the strSQLKeywords() array)
' and optionally (if blnNewlineperfield = True) it puts a new line at each comma (for example a field list in SELECT or GROUP BY clauses).
'
' Parameters: strSQL - the SQL string
'             blnNewLinePerField - flag indicating whether to put a new line at each comma or not (field lists in select, group by etc)
' Returns:    String - the formatted SQL
' Version 1.0 - 26 October 2007 - first out of beta
'         1.01 - 27 October 2007 - added support for action queries (insert/update/delete)
' Known limitations:
' * Designed to work with JET SQL, so if non Jet SQL code (eg Transact SQL or PL/SQL used in a pass through query)
'   were passed in, no idea how it would make of the  more in depth parts of that language
' * Doesn't support DDL queries (CREATE/ALTER/DROP), just returns the original code
' * Long lines of code with lots of join statements in still stay as long lines with join statements in.
'   Because of the multitude of ways these can be written, tidying this up is going to be awkward
' ******************************************************

    ' Minimum and maximum dimensions of array, use these as boundaries in loops later on
    Const ARRAY_MIN_SIZE As Integer = 1
    Const ARRAY_MAX_SIZE As Integer = 15

    Dim strSQLKeywords(ARRAY_MIN_SIZE To ARRAY_MAX_SIZE) As String          ' Array of SQL keywords

    Dim strTempSQL As String            ' Temp variable for storing output
    Dim strTemp1 As String, strTemp2 As String ' Temp for storing parts of code (before and after where the CrLf character needs to go)
    
    Dim intKWPos As Integer             ' Position of next keyword (ie current keyword +1)
    Dim intCurrentKW As Integer         ' Position of current keyword
    
    Dim intKWPosn As Integer            ' Temp working variable, holds start position of current keyword
    Dim intOldTemp As Integer           ' Previous temp working variable, hold start position of previous keyword
    
    Dim intCount As Integer             ' Integer
    
    ' Populate array with keywords to search
    ' These are likely to occur in this order, but they don't have to (eg subqueries, union queries etc)
    
    strSQLKeywords(1) = "PARAMETERS"
    strSQLKeywords(2) = "TRANSFORM"
    strSQLKeywords(3) = "SELECT"
    strSQLKeywords(4) = "FROM"
    strSQLKeywords(5) = "WHERE"
    strSQLKeywords(6) = "GROUP BY"
    strSQLKeywords(7) = "HAVING"
    strSQLKeywords(8) = "UNION"
    strSQLKeywords(9) = "ORDER BY"
    strSQLKeywords(10) = "PIVOT"
    strSQLKeywords(11) = "INSERT"
    strSQLKeywords(12) = "VALUES" ' for Insert statements
    strSQLKeywords(13) = "UPDATE"
    strSQLKeywords(14) = "SET"
    strSQLKeywords(15) = "DELETE"
    
    ' Check for unsupported query types and abort (ie return original code) if found...
    ' DDL statements need object type after create/alter/drop to be specified (eg create table tblName)
    ' which gives an extra
    
    If InStr(strSQL, " CREATE") Then
        ' Can't handle DDL queries - create
        FormatSQL = strSQL
        Exit Function
    End If
    
    If InStr(strSQL, " ALTER") Then
        FormatSQL = strSQL ' alter
        Exit Function
    End If
    
    If InStr(strSQL, " DROP") Then
        FormatSQL = strSQL ' drop
        Exit Function
    End If
    
    strTempSQL = strSQL     ' Start with current query
        
    ' Loop through each keyword and replace with capital equivalent of any found
    
    For intCurrentKW = ARRAY_MIN_SIZE To ARRAY_MAX_SIZE
            
        ' Is the current keyword in the array present in the query?
        intKWPosn = InStr(strSQL, strSQLKeywords(intCurrentKW))
        If (intKWPosn > 0) Then
        
            ' Keyword Found, convert it to upper case and insert line break between it and rest of characters
            strTempSQL = Replace(strTempSQL, strSQLKeywords(intCurrentKW), vbCrLf & UCase$(strSQLKeywords(intCurrentKW)))
        
        End If
    
    Next intCurrentKW
       
    If blnNewlinePerField Then
            ' Insert CRLF and <br> tag between sections and fields in the code
            ' Loop from intOldTemp to intTemp step -1 and replace commas with comma vbCrLf
            ' Examples: field list in select clause, group by clause etc.
            
          intCount = Len(strTempSQL)
          strTempSQL = strTempSQL & vbCr ' last character carriage return just to ensure it gets triggered on the first loop
          
          Do While intCount > 0
            If Mid$(strTempSQL, intCount, 1) = "," Then
                If Mid$(strTempSQL, intCount + 1, 1) <> vbCr Then
                    strTemp1 = Left$(strTempSQL, intCount)
                    strTemp2 = Right$(strTempSQL, Len(strTempSQL) - intCount)
                    strTempSQL = strTemp1 & vbCrLf & "<br />" & strTemp2
                End If
            End If
            intCount = intCount - 1
        Loop
            
    End If
    
    ' Now convert aggregate functions to capitals
    
    If InStr(strSQL, " SUM") > 0 Then strTempSQL = Replace(strTempSQL, " sum", " SUM")   ' Sum
    If InStr(strSQL, " AVG") > 0 Then strTempSQL = Replace(strTempSQL, " avg", " AVG")   ' Average
    If InStr(strSQL, " MIN") > 0 Then strTempSQL = Replace(strTempSQL, " min", " MIN")   ' Minimum
    If InStr(strSQL, " MAX") > 0 Then strTempSQL = Replace(strTempSQL, " max", " MAX")   ' Maximum
    If InStr(strSQL, " COUNT") > 0 Then strTempSQL = Replace(strTempSQL, " count", " COUNT")   ' Count
    If InStr(strSQL, " STDEV") > 0 Then strTempSQL = Replace(strTempSQL, " stdev", " STDEV")   ' Standard Deviation (StDev)
    
    ' Convert DISTINCT and DISTINCTROW to capitals
    If InStr(strSQL, " DISTINCT ") > 0 Then strTempSQL = Replace(strTempSQL, " DISTINCT ", " DISTINCT ")
    ' Note trailing space so it doesn't pick up the word as part of DISTINCTROW (you would end up with DISTINCTrow if this was the case)
    
    If InStr(strSQL, " DISTINCTROW ") > 0 Then strTempSQL = Replace(strTempSQL, " DISTINCTROW ", " DISTINCTROW ")
    
    ' Convert boolean operators (and/or/not) to capitals
    If InStr(strTempSQL, " AND ") > 0 Then strTempSQL = Replace(strTempSQL, " AND ", " AND ")
    If InStr(strTempSQL, " OR ") > 0 Then strTempSQL = Replace(strTempSQL, " OR ", " OR ")
    If InStr(strTempSQL, " NOT ") > 0 Then strTempSQL = Replace(strTempSQL, " NOT ", " NOT ")
    
    ' Convert Like and Exists to capitals
    If InStr(strTempSQL, " LIKE ") > 0 Then strTempSQL = Replace(strTempSQL, " LIKE ", " LIKE ")
    If InStr(strTempSQL, " EXISTS ") > 0 Then strTempSQL = Replace(strTempSQL, " EXISTS ", " EXISTS ")
    
    ' Now replace the ASC/DESC qualifiers for order by arguments with capitals; no need to insert new line as they will be last
    If InStr(strTempSQL, " ASC") > 0 Then strTempSQL = Replace(strTempSQL, " asc", " ASC")
    If InStr(strTempSQL, " DESC") > 0 Then strTempSQL = Replace(strTempSQL, " DESC", " DESC")
    
    FormatSQL = strTempSQL
End Function

Private Function LoadRibbons(clsfh As mdbdclsFileHandle, strHaltOnError As String) As Integer
' MDBDOC: Function for handling Ribbons
' Function: LoadRibbons
' Scope: Private
' Parameters: clsfh - File handle, strHaltOnError - string
' Return value: Integer - the number of ribbons included or -1 if not enabled (table missing)
' Date: August 2014
' Called from: mdbdProcessDatabase

' It is used in the functionality to handle ribbon detail
' Ribbons are stored in the USysRibbons table. Absence of the table means no ribbons (although they could be
' built in memory or loaded from disk this is the simplest solution).
    
    Dim intExitCode As Integer
    Dim intCount As Integer
    Dim intErrno As Integer
    
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    
    On Error GoTo err_LoadRibbons

    intExitCode = DCount("*", "USysRibbons")
    intErrno = Err.Number
    
    On Error Resume Next
    If intErrno = 3078 Then
        intExitCode = -1    ' table not found
    Else
        clsfh.WriteData "<table class=""mdbtable"">"
                        clsfh.WriteData "<caption>Details of Ribbons within this database</caption>"
                        clsfh.WriteData "<thead>"
                        clsfh.WriteData "<tr><th scope=""col"">ID</th><th scope=""col"">Name</th><th>XML</th></tr>"
                        clsfh.WriteData "</thead>"
                        clsfh.WriteData "<tbody>"
   
        Set db = CurrentDb
        Set rst = db.OpenRecordset("SELECT id, RibbonName, RibbonXML FROM USysRibbons ORDER BY RibbonName")
        Do While Not rst.EOF
            clsfh.WriteData "<tr><td>" & rst!Id & "</td><td>" & rst!RibbonName & "</td><td><code>" & mdbdReplaceSpecialChars(rst!RibbonXML) & "</code></td></tr>"
            rst.MoveNext
        Loop
        intExitCode = rst.RecordCount
        rst.Close
        Set db = Nothing
        clsfh.WriteData "</tbody>"
        clsfh.WriteData "</table>"
    End If
    
Exit_LoadRibbons:
    LoadRibbons = intExitCode
    Exit Function
    
err_LoadRibbons:
    Select Case Err.Number
        Case 3078 ' table not found, so no ribbons - return -1 error state
            intExitCode = -1
    Case Else
         clsfh.WriteData "<table class=""mdbtable"">"
                        clsfh.WriteData "<caption>Details of Ribbons within this database</caption>"
                        clsfh.WriteData "<thead>"
                        clsfh.WriteData "<tr><th scope=""col"">ID</th><th scope=""col"">Name</th><th scope=""col"">Ribbon XML</th></tr>"
                        clsfh.WriteData "</thead>"
                        clsfh.WriteData "<tbody>"
    End Select
    Resume Exit_LoadRibbons:
    
End Function