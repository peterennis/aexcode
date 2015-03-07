Version =20
VersionRequired =20
Checksum =-1164872733
Begin Form
    AutoResize = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6916
    DatasheetFontHeight =10
    ItemSuffix =23
    Left =165
    Top =2400
    Right =6750
    Bottom =4740
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9b785614191ae240
    End
    Caption ="MDB Doc"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    PrtDevMode = Begin
        0x00000000a43c2600705c0100380000000b0000000000000030ac170c30ac170c ,
        0x010400059c000c00138d0100010009009a0b34086400010000002c0102000100 ,
        0x000000000000413400002000530065007200690066000000000000008c2b2600 ,
        0x39d1f87500000000000000000000000000000000000000000000000001000000 ,
        0x000000000100000001000000000000000000000000000000000000004d444550 ,
        0x0010000001000000
    End
    PrtDevNames = Begin
        0x08001f0028000100000000000000000000000000000000000000000000000000 ,
        0x0000000000000000444f50373a0000000000
    End
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    AllowLayoutView =0
    DatasheetGridlinesColor12 =12632256
    PrtDevModeW = Begin
        0x00002700f57144777ab38b04feffffffd33c4077fe3c4077e8000000f8000000 ,
        0x5280cb0e5080cb0ea016cb753800ed06090300000000f60d0800000014000000 ,
        0x01040005dc000c00138d0100010009009a0b34086400010000002c0102000100 ,
        0x000000000000410034000000000062008477270063814077380162003800ed06 ,
        0x3c00ed06f571447732b08b04feffffff165942770b4fc875c0772700684fc875 ,
        0x00006200c4770000000000000000000000000000000000000000000001000000 ,
        0x000000000100000001000000000000000000000000000000000000004d444550 ,
        0x0010000001000000
    End
    PrtDevNamesW = Begin
        0x04001b0024000100000000000000000000000000000000000000000000000000 ,
        0x0000000000000000000000000000000000000000000000000000000000000000 ,
        0x000000000000000044004f00500037003a0000000000000000000000
    End
    NoSaveCTIWhenDisabled =1
    Begin
        Begin Label
            BackStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
            Width =1701
            Height =1701
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
        End
        Begin Section
            Height =1871
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    BackStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =465
                    Top =120
                    Width =5775
                    Height =450
                    FontSize =18
                    BackColor =-2147483633
                    ForeColor =-2147483640
                    Name ="lblMDBDoc"
                    Caption ="MDB Doc version"
                    FontName ="Arial"
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =113
                    Top =1077
                    Height =448
                    TabIndex =2
                    Name ="cmdProcess"
                    Caption ="Process Database"
                    StatusBarText ="Click to process the current database."
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =1927
                    Top =1077
                    Height =448
                    TabIndex =3
                    Name ="cmdCancel"
                    Caption ="Cancel"
                    StatusBarText ="Click here to close MDB Doc"
                    OnClick ="[Event Procedure]"

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin TextBox
                    OverlapFlags =215
                    AccessKey =79
                    Left =1013
                    Top =737
                    Width =4716
                    Name ="txtOutputFile"
                    StatusBarText ="Enter path to where you would like the output file to be placed."
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =120
                            Top =735
                            Width =975
                            Height =240
                            Name ="lblOutputFile"
                            Caption ="&Output File:"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5782
                    Top =737
                    Width =351
                    Height =291
                    TabIndex =1
                    Name ="cmdSelect"
                    Caption ="Select File"
                    StatusBarText ="Click to pick filename to save output to"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaad000000000000add0330880000330da ,
                        0xa0330880000330add0330880000330daa0330000000330add0333333333330da ,
                        0xa0330000000330add030fffffff030daa030fccccff030add030ffcccff030da ,
                        0xa03dfccccff000add0dacccfcff070daadacccadad0000addacccadadadadada ,
                        0xadacadadadadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin Image
                    BackStyle =1
                    Left =113
                    Top =113
                    Width =450
                    Height =510
                    BackColor =-2147483633
                    Name ="imgMDBDoc"
                    PictureData = Begin
                        0x280000001b0000001c0000000100040000000000c0010000120b0000120b0000 ,
                        0x10000000000000000000000020202000282828004040400080808000a0a0a000 ,
                        0xb0b0b000b8b8b800c0c0c000ffffff001010280078d8d8001888880020f0f000 ,
                        0x00ffff00ff0000006666aaaaaaaaaaaaaaaaaaaaaaa0f27f666aafffffffffff ,
                        0xfffffffffff0f27f66afaffffffffffffffffffffff0f27f6affafffffffffff ,
                        0xfff1234ffff0f27fafffaffffffffffffff0c747fff0f27faf9fafffffffffff ,
                        0xf00e81b7fff0f27faf9faffffffffffff0e80db7fff0f27faf9fafffff000000 ,
                        0x0c70ed4ffff0f27faf9faffff1ccce0ce70eeffffff0f27faf9fafff4cceeeee ,
                        0x80ee0ffffff0f27faf9fafff4ceecee70ee0fffffff0f27faf9fafff4ceeecee ,
                        0xce0ffffffff0f27faf9fafff5440eeceec0ffffffff0f27faf9fafff4d080eec ,
                        0xee0ffffffff0f27faf9fafff4ce040eece0ffffffff0f27faf9faffff1ce04ee ,
                        0xec0ffffffff0f27faf9fafffff0ce4eec0fffffffff0f27faf9faffffff33533 ,
                        0x3ffffffffff0f27faf9faffffffffffffffffffffff0f27faf9fafffffffffff ,
                        0xfffffffffff0f27faf9faffffaa9aaa9aa9aa9aafff0f27faf9faffff9999999 ,
                        0x99999999fff0f27faf9faffffffffffffffffffffff0f27faf9fafffffffffff ,
                        0xfffffffffff0f27faffaaaaaaaaaaaaaaaaaaaaaaaa0f27fafa5666666666666 ,
                        0x666666666660f27faa66767676767676767676766660f27f6aaaaaaaaaaaaaaa ,
                        0xaaaaaaaa6660f27f000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    Picture ="MDBDOC5.bmp"

                    TabIndex =6
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3685
                    Top =1077
                    Height =448
                    TabIndex =4
                    Name ="cmdPreferences"
                    Caption ="Preferences"
                    OnClick ="[Event Procedure]"
                    GUID = Begin
                        0xc6ffeaf537468343af874145fc8327a5
                    End

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =5499
                    Top =1077
                    Width =680
                    Height =448
                    TabIndex =5
                    Name ="cmdAbout"
                    Caption ="About"
                    OnClick ="[Event Procedure]"
                    GUID = Begin
                        0x9b17e08fc162e446b54c748b587ae2be
                    End

                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub cmdAbout_Click()
    'MDBDOC: Event Handler for About Button on the Startup form.
    MsgBox APP_NAME & " is an open source Documentation addin for Microsoft Access 97-2010. " & vbCrLf & _
                    "It is Copyright John Barnett released under the GNU General Public Software License version 3." & vbCrLf & vbCrLf & _
                    "Product Homepage: http://mdbdoc.sourceforge.net/" & vbCrLf & _
                    "GNU General Public Software license: http://www.gnu.org/licenses/, " & vbCrLf & _
                    "a copy of which is included in the file GPL.TXT with this application.", vbOKOnly + vbInformation, APP_NAME
End Sub

Private Sub cmdCancel_Click()
    'MDBDOC: Event handler for Cancel button on MBD Doc startup form.
    DoCmd.Close acForm, Me.Name, acSaveNo
End Sub

Private Sub cmdProcess_Click()
    'MDBDOC: Event handler for Process button on MBD Doc startup form.
    ' Amended 09 Dec 2003 to send files to overwrite to recycle bin rather than delete.
    Dim strOutputFile As String
    Dim strPreference As String
  
    Me.Visible = False
    strOutputFile = Me.txtOutputFile
    If Len(strOutputFile & "") > 0 Then
        Me.Repaint
        If Dir(strOutputFile) <> "" Then
            If MsgBox("The file " & strOutputFile & " already exists. Overwrite?", vbYesNo + vbQuestion) = vbYes Then
                If mdbdGetPreference("DeleteToRecycleBin") = PREFERENCE_ENABLED Then
                    RecycleFile strOutputFile
                Else
                    Kill strOutputFile
                End If
                DoCmd.Close acForm, Me.Name
                mdbdProcessDatabase strOutputFile
            Else
                Me.Visible = True
            End If
        Else
            DoCmd.Close acForm, Me.Name
            mdbdProcessDatabase strOutputFile
        End If
    Else
        MsgBox "Please enter an output path and filename.", vbOKOnly + vbInformation
        DoCmd.Close acForm, Me.Name
    End If
End Sub

Private Sub cmdSelect_Click()
    'MDBDOC: Event handler for "Select File" button on MBD Doc startup form.
    Dim strOutput As String
    
    strOutput = PrepareSavefile
    If Len(strOutput & "") > 0 Then
        Me.txtOutputFile = strOutput
    End If
End Sub

Private Sub Form_Open(Cancel As Integer)
'MDBDOC: Form Open event - sets up form with default output location and checks that database is not in MDE format.
' Amended 30 August 2014 to check Access version >= 12.0 (Access 2007) for MDB Doc 1.60.
    Dim strOutputFile As String
    Dim strPreference As String

    Dim strVersion As String
    
    If SysCmd(acSysCmdRuntime) = True Then
        MsgBox APP_NAME & " cannot run under the Runtime version of Microsoft Access", vbOKOnly + vbInformation
        Exit Sub
    End If
    
    ' Check for Access 2007 version
    If (Int(Application.Version) < ACCESS_2007_VERSION) Then
        MsgBox APP_NAME & " " & strVersion & " requires Microsoft Access 2007 or later. Please use version 1.52 from the product website."
        DoCmd.Close acForm, Me.Name
        Exit Sub
    End If
    
    If Not mdbdIsMDE(CurrentDb) Then
        mdbdLoadPreferences
        Me.txtOutputFile = SetOutputFilename
    Else
        MsgBox APP_NAME & " cannot process an MDE database", vbInformation + vbOKOnly
        DoCmd.Close acForm, Me.Name ' use this rather than Cancel =true as it stops an ugly "OpenForm action cancelled" error msg.
        Exit Sub
    End If

    Me.lblMDBDoc.Caption = APP_NAME & " Copyright John Barnett"
End Sub

Private Sub txtOutputFile_AfterUpdate()
    'MDBDOC: Checks to disable the "Process" button if the txtOutputfile text box is empty.
    If Len(Me.txtOutputFile & "") > 0 Then
        cmdProcess.Enabled = True
    Else
        cmdProcess.Enabled = False
    End If
End Sub

Private Function mdbdIsMDE(db As DAO.Database) As Boolean
    'MDBDOC: Function to determine if a particular database is in MDE format or not.
    ' Function: mdbdIsMDE
    ' Scope:    Private
    ' Parameters: db - DAO database object
    ' Return Value: Boolean - True if the db database is an MDE, false otherwise.
    ' Author:   John Barnett
    ' Date:     21 July 2001, amended 23 October 2003.
    ' Description: IsMDE returns True/False indicating if the database supplied as a parameter is an MDE file
    ' Called by: cmdProcess_click routine in frmStartup.
    '
    ' It works on the fact that an MDE database has a property of "MDE" added
    ' with a value of "T".  This is far more reliable than checking the file extension.
    
    mdbdIsMDE = False

    On Error Resume Next

    mdbdIsMDE = (db.Properties("MDE") = "T")

End Function

Private Sub cmdPreferences_Click()
    'MDBDOC: Sub to open the preferences form.
    On Error GoTo Err_cmdPreferences_Click

    Me.Visible = False
    DoCmd.OpenForm "frmPreferences"

Exit_cmdPreferences_Click:
    Exit Sub

Err_cmdPreferences_Click:
    MsgBox Err.Description
    Resume Exit_cmdPreferences_Click
    
End Sub

Private Function SetOutputFilename() As String
    'MDBDOC: Function to generate the default output filename, based on the database name and the user preferences.
    Dim strPreference As String
    Dim intCount As Integer
    Dim strOutputfilename As String

    Dim intDotPosn As Integer
    
    ' Retrieve default output path preference.
    strPreference = mdbdGetPreference("DefaultOutputPath")
    If Len(strPreference & "") = 0 Then
        strOutputfilename = CurrentDb.Name
        ' Default to current DB's path if not specified
    Else
        For intCount = Len(CurrentDb.Name) To 1 Step -1
            ' find the position of the last backslash
            If Mid$(CurrentDb.Name, intCount, 1) = "\" Then Exit For
        Next
        ' If \ is the last character of the path ...
        If Right$(strPreference, 1) = "\" Then
            ' Output filename = default path + current database name
            strOutputfilename = strPreference & Mid$(CurrentDb.Name, intCount + 1)
        Else
            ' Otherwise the output file is the current path plus the db name
            strOutputfilename = strPreference & "\" & Mid$(CurrentDb.Name, intCount + 1)
        End If
    End If
    
    ' Retrieve default extension from the preferences setup
    strPreference = mdbdGetPreference("OutputFileExtension") ' get extension
    
    ' Remove the extension from the existing output path (num chars since last dot. 3 for mdb/mda; 5 for accdb)
    For intDotPosn = Len(strOutputfilename) To 1 Step -1
        
        If Mid$(strOutputfilename, intDotPosn, 1) = "." Then
            ' remove anything after this
            strOutputfilename = Left$(strOutputfilename, intDotPosn)
            Exit For
        End If
    Next intDotPosn
    
    ' and add the retrieved preference
    strOutputfilename = strOutputfilename & strPreference
    
    ' If Convert to lower case preference set, then do  it
    ' mostly useful for web servers that have case sensitive file handling
    If mdbdGetPreference("ConvertToLowerCase") = PREFERENCE_ENABLED Then strOutputfilename = LCase$(strOutputfilename)
    SetOutputFilename = strOutputfilename
End Function