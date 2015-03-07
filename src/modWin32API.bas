Option Compare Database
Option Explicit

Private Declare Function GetOpenFilename Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFilename Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type


Public Function PrepareOpenfile() As String
    'MDBDOC: Function to prepare the common dialog in Open file mode.
    Dim OpenFile As OPENFILENAME
    
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = Application.hWndAccessApp
    'openfile.hInstance = application.in
    OpenFile.lpstrFilter = "HTML Files (*.htm)" & Chr(0) & "*.htm" & Chr(0) ' Drop down list of filters
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = "C:\"    ' Initial directory, leave blank for My Documents
    OpenFile.lpstrTitle = "Open File"   ' Title of the dialog box
    OpenFile.flags = 0
    If GetOpenFilename(OpenFile) = 1 Then
        PrepareOpenfile = Mid$(OpenFile.lpstrFile, 1, InStr(OpenFile.lpstrFile, Chr(0)) - 1)
    Else
        PrepareOpenfile = ""
    End If
End Function

Sub RecycleFile(sFile As String)
    'MDBDOC: Code to delete a file to the Recycle bin from Chip Pearson's Excel site.
    ' this code and the other functions necessary to delete a file to the recycle bin
    ' came from Chip Pearson's site at http://www.cpearson.com/excel/Recycle.htm
    Dim FileOperation As SHFILEOPSTRUCT
    Dim lReturn As Long
    Dim sFileName As String

    Const FO_DELETE = &H3
    Const FOF_ALLOWUNDO = &H40
    Const FOF_NOCONFIRMATION = &H10

    With FileOperation
        .wFunc = FO_DELETE
        .pFrom = sFile
        .fFlags = FOF_ALLOWUNDO
        '
        ' OR if you want to suppress the "Do You want
        ' to delete the file" message, use
        '
        .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION

    End With

    lReturn = SHFileOperation(FileOperation)

End Sub

Public Function PrepareSavefile() As String
    'MDBDOC: Function to prepare the common dialog in Save file mode. Used by the button on the startup form.
    Dim OpenFile As OPENFILENAME
    
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = Application.hWndAccessApp
    
    ' Filters for file dialog
    OpenFile.lpstrFilter = "HTML Files (*.htm)" & Chr(0) & "*.htm" & Chr(0)
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = ""           ' Use default directory
    OpenFile.lpstrTitle = "Save to File"    ' Title
    OpenFile.flags = 0
    If GetSaveFilename(OpenFile) = 1 Then
        PrepareSavefile = Mid$(OpenFile.lpstrFile, 1, InStr(OpenFile.lpstrFile, Chr(0)) - 1)
    Else
        PrepareSavefile = ""
    End If
End Function