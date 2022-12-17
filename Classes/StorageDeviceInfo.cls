VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StorageDeviceInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const OPEN_EXISTING                 As Long = 3&
Private Const FILE_SHARE_READ               As Long = &H1&
Private Const FILE_SHARE_WRITE              As Long = &H2&
Private Const GENERIC_READ                  As Long = &H80000000
Private Const IOCTL_STORAGE_QUERY_PROPERTY  As Long = &H2D1400

Public Enum STORAGE_BUS_TYPE
    BusTypeUnknown = 0
    BusTypeScsi
    BusTypeAtapi
    BusTypeAta
    BusType1394
    BusTypeSsa
    BusTypeFibre
    BusTypeUsb
    BusTypeRAID
    BusTypeMaxReserved = &H7F
End Enum

Private Enum STORAGE_PROPERTY_ID
    StorageDeviceProperty = 0
    StorageAdapterProperty
    StorageDeviceIdProperty
End Enum

Private Enum STORAGE_QUERY_TYPE
    PropertyStandardQuery = 0
    PropertyExistsQuery
    PropertyMaskQuery
    PropertyQueryMaxDefined
End Enum

Private Type STORAGE_PROPERTY_QUERY
    PropertyId            As STORAGE_PROPERTY_ID
    QueryType             As STORAGE_QUERY_TYPE
    AdditionalParameters  As Byte
End Type

Private Type STORAGE_DEVICE_DESCRIPTOR
    Version               As Long
    Size                  As Long
    DeviceType            As Byte
    DeviceTypeModifier    As Byte
    RemovableMedia        As Byte
    CommandQueueing       As Byte
    VendorIdOffset        As Long
    ProductIdOffset       As Long
    ProductRevisionOffset As Long
    SerialNumberOffset    As Long
    BusType               As Integer
    RawPropertiesLength   As Long
    RawDeviceProperties   As Byte
End Type

Private Type DEVICE_INFORMATION
    Valid                 As Boolean
    BusType               As STORAGE_BUS_TYPE
    Removable             As Boolean
    VendorID              As String
    ProductID             As String
    ProductRevision       As String
End Type

Private Declare Sub RtlMoveMemory Lib "kernel32" (ByRef pDst As Any, ByRef pSrc As Any, ByVal cb As Long)

Private Declare Function lstrcpyW Lib "kernel32" (ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As Long

Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long

Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As LongPtr) As LongPtr

'https://docs.microsoft.com/en-us/windows/win32/api/ioapiset/nf-ioapiset-deviceiocontrol
'https://docs.microsoft.com/en-us/windows/win32/devio/calling-deviceiocontrol
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As LongPtr, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, lpOverlapped As Any) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long

Private m_DevInfo As DEVICE_INFORMATION

Friend Sub New_(strDrive As String)
    
    Dim hDrive          As LongPtr
    Dim udtQuery        As STORAGE_PROPERTY_QUERY
    Dim dwOutBytes      As Long
    Dim lngResult       As Long
    Dim btBuffer(9999)  As Byte
    Dim udtOut          As STORAGE_DEVICE_DESCRIPTOR
    
    hDrive = CreateFileW(StrPtr("\\.\" & Left$(strDrive, 1) & ":"), 0, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)

    If hDrive = -1 Then Exit Sub
    
    'With udtQuery
    '    .PropertyId = STORAGE_PROPERTY_ID.StorageDeviceProperty '0
    '    .QueryType = STORAGE_QUERY_TYPE.PropertyStandardQuery   '0
    'End With
    
    lngResult = DeviceIoControl(hDrive, IOCTL_STORAGE_QUERY_PROPERTY, udtQuery, LenB(udtQuery), btBuffer(0), UBound(btBuffer) + 1, dwOutBytes, ByVal 0&)
    
    If lngResult Then
        
        RtlMoveMemory udtOut, btBuffer(0), Len(udtOut)
        
        With m_DevInfo
            
            .Valid = True
            .BusType = udtOut.BusType
            .Removable = CBool(udtOut.RemovableMedia)
            
            If udtOut.ProductIdOffset > 0 Then _
                .ProductID = StringCopy(VarPtr(btBuffer(udtOut.ProductIdOffset)))
            If udtOut.ProductRevisionOffset > 0 Then _
                .ProductRevision = StringCopy(VarPtr(btBuffer(udtOut.ProductRevisionOffset)))
            If udtOut.VendorIdOffset > 0 Then
                .VendorID = StringCopy(VarPtr(btBuffer(udtOut.VendorIdOffset)))
            End If
            
        End With
    Else
        m_DevInfo.Valid = False
    End If
    
    CloseHandle hDrive

End Sub

Private Function StringCopy(ByVal pBuffer As LongPtr) As String
    Dim tmp As String
    tmp = Space(lstrlenW(ByVal pBuffer))
    lstrcpyW ByVal tmp, ByVal pBuffer
    StringCopy = Trim$(tmp)
End Function

Public Property Get IsRemoveable() As Boolean
    IsRemoveable = m_DevInfo.Removable
End Property

Public Property Get IsValid() As Boolean
    IsValid = m_DevInfo.Valid
End Property

Public Property Get BusType() As STORAGE_BUS_TYPE
    BusType = m_DevInfo.BusType
End Property

Public Property Get BusType_ToStr() As String
    Dim s As String
    Select Case m_DevInfo.BusType
    Case BusTypeUnknown:     s = "Unknown"
    Case BusTypeScsi:        s = "Scsi"
    Case BusTypeAtapi:       s = "Atapi"
    Case BusTypeAta:         s = "Ata"
    Case BusType1394:        s = "1394"
    Case BusTypeSsa:         s = "Ssa"
    Case BusTypeFibre:       s = "Fibre"
    Case BusTypeUsb:         s = "Usb"
    Case BusTypeRAID:        s = "RAID"
    Case BusTypeMaxReserved: s = "Other"
    Case Else:               s = "something else"
    End Select
    BusType_ToStr = s
End Property

Public Property Get VendorID() As String
    VendorID = m_DevInfo.VendorID
End Property

Public Property Get ProductID() As String
    ProductID = m_DevInfo.ProductID
End Property

Public Property Get ProductRevision() As String
    ProductRevision = m_DevInfo.ProductRevision
End Property

