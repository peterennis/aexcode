Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' These are constants for the mode in which files can be opened works.
' When in use they are stored in m_filemode.
Private Const FileModeAppend = "A" ' Append mode
Private Const FileModeClosed = "C" ' Closed file (ie not active)
Private Const FileModeRead = "R" ' Read only
Private Const FileModeWrite = "W" ' Write

Private m_filename As String ' member variable for the filename
Private m_handle As Integer  ' VB file handle
Private m_filemode As String * 1 ' file mode - Read/Write/Append/Closed from above
Private m_curdata As String  ' current row of data
    
Public Property Get Filename() As String
    'MDBDOC: Filename Get property - retrieves "Filename" property.
    Filename = m_filename ' Get routines are used for retreiving values of member variables
End Property

Public Property Let Filename(fName As String)
    'MDBDOC: Filename Let property - allows it to be set.
    m_filename = fName
End Property

Public Property Get FileMode() As String
    'MDBDOC:FileMode Get property - allows it to be retrived
    FileMode = m_filemode
End Property

Private Property Get Filenumber() As Integer
    'MDBDOC: FileNumber property - allows it to be retrieved.
    Filenumber = m_handle
End Property

Private Property Let Filenumber(num As Integer)
    'MDBDOC: FileNumber property - allows it to be set.
    m_handle = num
End Property

Public Property Let FileMode(mode As String)
    'MDBDOC: FileMode property - allows it to be set. Includes basic validation.
    If mode = FileModeRead Or mode = FileModeWrite Or mode = FileModeAppend Then
        m_filemode = mode
    End If
End Property

Public Function OpenFile()
    'MDBDOC: Function to open file.
    
    m_handle = FreeFile
    Select Case m_filemode
        Case FileModeRead ' read only
            Open Me.Filename For Input As #m_handle
        Case FileModeWrite ' Write
            Open Me.Filename For Output As #m_handle
        Case FileModeAppend ' Append
            Open Me.Filename For Append As #m_handle
    End Select
    OpenFile = m_handle ' return handle as check
End Function

Public Function WriteData(data As String)
    'MDBDOC: Function to write data to the file. Includes validation to stop it being written to a closed file, or one opened for read only access.
    If Me.FileMode = FileModeWrite Or Me.FileMode = FileModeAppend Then
        ' can't write to input files
        Print #m_handle, data
        m_curdata = data
    End If
End Function

Public Function ReadData() As String
    'MDBDOC: Function to read data from a file and return it.
    Dim strTmp As String
    If Me.FileMode = FileModeRead Then
        ' can't read from write/append files
        Input #m_handle, strTmp
        ReadData = strTmp
        m_curdata = strTmp
    End If
End Function

Public Function CloseFile()
    'MDBDOC: Function to close a file.
    Close #m_handle
    m_filemode = FileModeClosed
End Function

Private Sub Class_Initialize()
    'MDBDOC: Initialise feature for the class - will set things up as class is initialised.
    m_filemode = FileModeClosed ' closed - can't read or write
    m_handle = -1 ' invalid handle
End Sub

Public Function IsOpen() As Boolean
    'MDBDOC: Function to determine if a file is open or not.
    IsOpen = m_filemode <> FileModeClosed
End Function

Private Sub Class_Terminate()
    'MDBDOC: Function to tidy up on closing down the class.
    If Me.IsOpen Then Me.CloseFile
End Sub

Public Function AtEOF() As Boolean
    'MDBDOC: Function to determine if the file is at the end of the file.
    AtEOF = EOF(m_handle)
End Function