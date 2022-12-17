VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PFNDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements PFNDirectory
Private Const DRIVE_UNKNOWN     As Long = 0 'The drive type cannot be determined.
Private Const DRIVE_NO_ROOT_DIR As Long = 1 'The root path is invalid; for example, there is no volume mounted at the specified path.
Private Const DRIVE_REMOVABLE   As Long = 2 'The drive has removable media; for example, a floppy drive, thumb drive, or flash card reader.
Private Const DRIVE_FIXED       As Long = 3 'The drive has fixed media; for example, a hard disk drive or flash drive.
Private Const DRIVE_REMOTE      As Long = 4 'The drive is a remote (network) drive.
Private Const DRIVE_CDROM       As Long = 5 'The drive is a CD-ROM drive.
Private Const DRIVE_RAMDISK     As Long = 6 'The drive is a RAM disk.

Private m_DL  As Integer 'String * 1 'the drive letter like C,D,E,F,G,H...
Private m_DT  As Long   'the drive type one value from above

'https://docs.microsoft.com/en-us/windows/win32/api/fileapi/nf-fileapi-getdrivetypew
Private Declare Function GetDriveTypeW Lib "kernel32" (ByVal nDrive As LongPtr) As Long

Private m_Base As PFNDirectory
'Private m_Dirs  As List 'Of PFNDirectory 'is it possible to have 2 directories of the same name in one directory? no! so m_dirs could be hashed
'Private m_Files As List 'Of PathFileName 'is it possible to have 2 files       of the same name in one directory? no! so m_Files could be hashed

Friend Sub New_(ByVal aLetter As String)
    'm_DL = Left$(aLetter, 1) & ":\"
    m_DL = AscW(Left(aLetter, 1))
    
    'Laufwerkstyp ermitteln
    Dim s As String: s = Me.DriveLetter & ":" ' m_DL '& vbNullChar
    m_DT = GetDriveTypeW(StrPtr(s))
    Set m_Base = MNew.PFNDirectory(Nothing, s)
End Sub

Public Property Get Name() As String
    Name = Me.DriveLetter & ":" & "\"
End Property

Public Property Get DriveLetter() As String
    DriveLetter = ChrW(m_DL)
End Property

Public Property Get DriveType() As Long
    DriveType = m_DT
End Property

Public Function DriveTypeToStr() As String
    Dim s As String
    Select Case m_DT
    Case DRIVE_UNKNOWN:      s = "Unknown drive"
    Case DRIVE_NO_ROOT_DIR:  s = "No root directory"
    Case DRIVE_CDROM:        s = "CD-ROM drive"
    Case DRIVE_FIXED:        s = "Fixed harddisk" 'HDD or SSD
    Case DRIVE_RAMDISK:      s = "RAM-Disk"
    Case DRIVE_REMOTE:       s = "Network-drive"
    Case DRIVE_REMOVABLE:    s = "Removeable"
    End Select
    DriveTypeToStr = s
End Function

Public Function ToStr() As String
    'ToStr = m_DL & " " & DriveTypeToStr
    ToStr = Me.Name & " " & DriveTypeToStr
End Function

Public Function Populate() As Boolean
    'New_ and Populate has to be in different functions because they work at different times!
    Populate = m_Base.Populate
End Function

Public Property Get Directories() As List 'Of PFNDirectory
    Set Directories = m_Base.Directories
End Property

Public Property Get Files() As List 'Of PathFileName
    Set Files = m_Base.Files
End Property

' v ############################## v '    Implements PFNDirectory   ' v ############################## v '
Private Property Get PFNDirectory_Count() As Long
    PFNDirectory_Count = m_Base.Count
End Property

Private Property Get PFNDirectory_Root() As PFNDirectory
    Set PFNDirectory_Root = Nothing 'a drive has no root
End Property

Private Property Get PFNDirectory_Name() As String
    PFNDirectory_Name = Me.Name
End Property

Private Property Get PFNDirectory_Key() As String
    PFNDirectory_Key = Me.Name
End Property

Private Function PFNDirectory_Populate() As Boolean
    PFNDirectory_Populate = Me.Populate
End Function

Private Function PFNDirectory_ToPFN() As PathFileName
    Set PFNDirectory_ToPFN = New PathFileName
    PFNDirectory_ToPFN.Value = Me.Name
End Function

Private Function PFNDirectory_ToStr() As String
    PFNDirectory_ToStr = Me.Name
End Function

Private Property Get PFNDirectory_Files() As List
    Set PFNDirectory_Files = Me.Files
End Property

Private Property Get PFNDirectory_Directories() As List
    Set PFNDirectory_Directories = Me.Directories
End Property

