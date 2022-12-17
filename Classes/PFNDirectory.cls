VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PFNDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#If VBA7 = 0 Then
'    Public Enum LongPtr
'        [_]
'    End Enum
#End If

Private Enum FileAttributeConstants
    FILE_ATTRIBUTE_READONLY = 1                    ' (0x1)      '    A file that is read-only. Applications can read the file, but cannot write to it or delete it. This attribute is not honored on directories. For more information, see You cannot view or change the Read-only or the System attributes of folders in Windows Server 2003, in Windows XP, in Windows Vista or in Windows 7.
    FILE_ATTRIBUTE_HIDDEN = 2                      ' (0x2)      '    The file or directory is hidden. It is not included in an ordinary directory listing.
    FILE_ATTRIBUTE_SYSTEM = 4                      ' (0x4)      '    A file or directory that the operating system uses a part of, or uses exclusively.
    FILE_ATTRIBUTE_DIRECTORY = 16                  ' (0x10)     '     The handle that identifies a directory.
    FILE_ATTRIBUTE_ARCHIVE = 32                    ' (0x20)     '    A file or directory that is an archive file or directory. Applications typically use this attribute to mark files for backup or removal .
    FILE_ATTRIBUTE_DEVICE = 64                     ' (0x40)     '    This value is reserved for system use.
    FILE_ATTRIBUTE_NORMAL = 128                    ' (0x80)     '    A file that does not have other attributes set. This attribute is valid only when used alone.
    FILE_ATTRIBUTE_TEMPORARY = 256                 ' (0x100)    '    A file that is being used for temporary storage. File systems avoid writing data back to mass storage if sufficient cache memory is available, because typically, an application deletes a temporary file after the handle is closed. In that scenario, the system can entirely avoid writing the data. Otherwise, the data is written after the handle is closed.
    FILE_ATTRIBUTE_SPARSE_FILE = 512               ' (0x200)    '    A file that is a sparse file.
    FILE_ATTRIBUTE_REPARSE_POINT = 1024            ' (0x400)    '    A file or directory that has an associated reparse point, or a file that is a symbolic link.
    FILE_ATTRIBUTE_COMPRESSED = 2048               ' (0x800)    '    A file or directory that is compressed. For a file, all of the data in the file is compressed. For a directory, compression is the default for newly created files and subdirectories.
    FILE_ATTRIBUTE_OFFLINE = 4096                  ' (0x1000)   '    The data of a file is not available immediately. This attribute indicates that the file data is physically moved to offline storage. This attribute is used by Remote Storage, which is the hierarchical storage management software. Applications should not arbitrarily change this attribute.
    FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = 8192      ' (0x2000)   '    The file or directory is not to be indexed by the content indexing service.
    FILE_ATTRIBUTE_ENCRYPTED = 16384               ' (0x4000)   '    A file or directory that is encrypted. For a file, all data streams in the file are encrypted. For a directory, encryption is the default for newly created files and subdirectories.
    FILE_ATTRIBUTE_INTEGRITY_STREAM = 32768        ' (0x8000)   '    The directory or user data stream is configured with integrity (only supported on ReFS volumes). It is not included in an ordinary directory listing. The integrity setting persists with the file if it's renamed. If a file is copied the destination file will have integrity set if either the source file or destination directory have integrity set.
                                                                '    Windows Server 2008 R2, Windows 7, Windows Server 2008, Windows Vista, Windows Server 2003 and Windows XP: This flag is not supported until Windows Server 2012.
    FILE_ATTRIBUTE_VIRTUAL = 65536                 ' (0x10000)  '    This value is reserved for system use.
    FILE_ATTRIBUTE_NO_SCRUB_DATA = 131072          ' (0x20000)  '    The user data stream not to be read by the background data integrity scanner (AKA scrubber). When set on a directory it only provides inheritance. This flag is only supported on Storage Spaces and ReFS volumes. It is not included in an ordinary directory listing.
                                                                '    Windows Server 2008 R2, Windows 7, Windows Server 2008, Windows Vista, Windows Server 2003 and Windows XP: This flag is not supported until Windows 8 and Windows Server 2012.
    FILE_ATTRIBUTE_RECALL_ON_OPEN = 262144         ' (0x40000)  '    This attribute only appears in directory enumeration classes (FILE_DIRECTORY_INFORMATION, FILE_BOTH_DIR_INFORMATION, etc.). When this attribute is set, it means that the file or directory has no physical representation on the local system; the item is virtual. Opening the item will be more expensive than normal, e.g. it will cause at least some of it to be fetched from a remote store.
    FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS = 4194304 ' (0x400000) '    When this attribute is set, it means that the file or directory is not fully present locally. For a file that means that not all of its data is on local storage (e.g. it may be sparse with some data still in remote storage). For a directory it means that some of the directory contents are being virtualized from another location. Reading the file / enumerating the directory will be more expensive than normal, e.g. it will cause at least some of the file/directory content to be fetched from a remote store. Only kernel-mode callers can set this bit.
End Enum

Private Const INVALID_HANDLE_VALUE As Long = -1
'Private Const MAX_PATH As Long = 260
Private Const MAX_PATH As Long = 2048 '32768 '260

'https://learn.microsoft.com/en-us/windows/win32/api/minwinbase/ns-minwinbase-filetime
'Contains a 64-bit value representing the number of 100-nanosecond intervals since January 1, 1601 (UTC).
Private Type FILETIME
    dwLowDateTime      As Long
    dwHighDateTime     As Long
End Type

'https://learn.microsoft.com/en-us/windows/win32/api/minwinbase/ns-minwinbase-win32_find_dataw
Private Type WIN32_FIND_DATAW
    dwFileAttributes As Long
    ftCreationTime   As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime  As FILETIME
    'nFileSizeHigh    As Long
    'nFileSizeLow     As Long
    nFileSize        As Currency
    dwReserved0      As Long
    dwReserved1      As Long
    'cFileName        As String * MAX_PATH 'Byte '
    cFileName(MAX_PATH) As Byte '
    'cAlternateFileName(14) As Integer
    'dwFileType As Long        '; // Obsolete. Do not use.
    'dwCreatorType  As Long    '; // Obsolete. Do not use
    'wFinderFlags  As Integer  ' WORD     ; // Obsolete. Do not use
End Type
'} WIN32_FIND_DATAW, *PWIN32_FIND_DATAW, *LPWIN32_FIND_DATAW;

'Private Type WIN32_FIND_DATA
'    dwFileAttributes   As Long
'    ftCreationTime     As FILETIME
'    ftLastAccessTime   As FILETIME
'    ftLastWriteTime    As FILETIME
'    nFileSizeHigh      As Long
'    nFileSizeLow       As Long
'    dwReserved0        As Long
'    dwReserved1        As Long
'    cFileName          As String * MAX_PATH
'    cAlternate         As String * 14
'End Type

'https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-findfirstfilew
Private Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByRef lpFindFileData As WIN32_FIND_DATAW) As LongPtr

'https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-findnextfilew
Private Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As LongPtr, ByRef lpFindFileData As WIN32_FIND_DATAW) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-findclose
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As LongPtr) As Long

Private m_Root  As PFNDirectory
Private m_Name  As String
Private m_Dirs  As List 'Of PFNDirectory 'is it possible to have 2 directorys of the same name in one directory? no? so a name is unique and m_Dirs  could be hashed?
Private m_Files As List 'Of PathFileName 'is it possible to have 2 files      of the same name in one directory? no? so a name is unique and m_Files could be hashed?
'

Private Sub Class_Initialize()
    Set m_Dirs = MNew.List(vbObject)  ', TypeName(Me))     Of PFNDirectory
    Set m_Files = MNew.List(vbObject) ', TypeName(m_Root)) Of PathFileName
End Sub

Friend Sub New_(aRoot As PFNDirectory, Name As String)
    Set m_Root = aRoot:       m_Name = Name
End Sub

Public Property Get Root() As PFNDirectory 'PathFileName
    Set Root = m_Root
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Friend Property Let Name(ByVal Value As String)
    m_Name = Value
    'also rename directory on disk? nope not now
End Property

Public Property Get Key() As String
    Key = m_Name
End Property

Public Function ToStr() As String
    Dim s As String
    If Not m_Root Is Nothing Then s = m_Root.ToStr
    ToStr = s & m_Name & "\"
End Function

Public Function ToPFN() As PathFileName
    Set ToPFN = New PathFileName: ToPFN.Value = Me.ToStr
End Function

Public Property Get Count() As Long
    Dim c As Long
    If Not m_Dirs Is Nothing Then c = m_Dirs.Count
    If Not m_Files Is Nothing Then c = c + m_Files.Count
    If c = 0 Then c = QuickCount
    Count = c
End Property

Public Property Get Directories() As List 'Of PFNDirectory
    Set Directories = m_Dirs
End Property

Public Property Get Files() As List 'Of PathFileName
    Set Files = m_Files
End Property

Public Function Populate() As Boolean
    'New_ and Populate has to be in different functions because they work at different times!
    If m_Dirs.Count = 0 And m_Files.Count = 0 Then
        Populate = Update
    Else
        Populate = True
    End If
End Function

Friend Function Update() As Boolean
    'New_ and Populate has to be in different functions because they work at different times!
Try: On Error GoTo Catch
    Dim wfd As WIN32_FIND_DATAW
    Dim s As String: s = Me.ToStr & "*"
    's = "\\?\" & s
    's = s & vbNullChar
    'Debug.Print s
    Dim hnd As LongPtr: hnd = FindFirstFileW(StrPtr(s), wfd)
    
    If hnd = INVALID_HANDLE_VALUE Then
        'MsgBox "Invalid handle value from FindfirstFileW()"
        'Exit Function
        GoTo Catch
    End If
    Dim dirs As List: Set dirs = MNew.List(vbObject) ', TypeName(Me))     Of PFNDirectory
    Dim fils As List: Set fils = MNew.List(vbObject) ', TypeName(m_Root)) Of PathFileName
    Dim f As PathFileName
    Dim d As PFNDirectory
    
    Do
        If (wfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            'nd = nd + 1
            s = wfd.cFileName
            s = MString.Trim0(s)
            Set d = MNew.PFNDirectory(Me, s) 'MNew.PathFileName(s))
            'If s = "." Then d.Root.Value = "."
            dirs.Add d
        Else
            'nf = nf + 1
            s = wfd.cFileName
            s = MString.Trim0(s)
            Set f = MNew.PathFileName(MString.Trim0(wfd.cFileName))
            fils.Add f
        End If
    Loop While FindNextFileW(hnd, wfd) <> 0
    'Debug.Print nd
    'Debug.Print nf
    Set m_Dirs = dirs
    Set m_Files = fils
    Update = True
    GoTo Finally
Catch:
    ErrHandler "Update", s, hnd
Finally:
    FindClose hnd
End Function

Private Function QuickCount() As Long
    'just count we do not store anything
Try: On Error GoTo Catch
    Dim wfd As WIN32_FIND_DATAW
    Dim s As String: s = Me.ToStr & "*"
    Dim hnd As LongPtr: hnd = FindFirstFileW(StrPtr(s), wfd)
    If hnd = INVALID_HANDLE_VALUE Then GoTo Catch
    Dim nd As Long, nf As Long
    Do
        If (wfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = FILE_ATTRIBUTE_DIRECTORY Then
            nd = nd + 1
        Else
            nf = nf + 1
        End If
    Loop While FindNextFileW(hnd, wfd) <> 0
    QuickCount = nd + nf
    GoTo Finally
Catch:
    ErrHandler "QuickCount", s, hnd
Finally:
    FindClose hnd
End Function

'copy this same function to every class, form or module
'the name of the class or form will be added automatically
'in standard-modules the function "TypeName(Me)" will not work, so simply replace it with the name of the Module
' v #################### v '     Local ErrHandler      ' v #################### v '
Private Function ErrHandler(ByVal FuncName As String, _
                            Optional ByVal AddInfo As String, _
                            Optional WinApiError, _
                            Optional bLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKCancel, _
                            Optional bRetry As Boolean) As VbMsgBoxResult

    If bRetry Then

        ErrHandler = MessErrorRetry(TypeName(Me), FuncName, AddInfo, WinApiError, bErrLog)

    Else

        ErrHandler = MessError(TypeName(Me), FuncName, AddInfo, WinApiError, bLoud, bErrLog, vbDecor)

    End If

End Function

