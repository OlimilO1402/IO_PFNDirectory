Attribute VB_Name = "MPC"
Option Explicit
'Module MPC read: My-PC

'ActiveVB.de: VB 5/6-Tipp 0012: "Verfügbare Laufwerke, einschl. der Netzlaufwerke erkennen"
'http://www.activevb.de/tipps/vb6tipps/tipp0012.html
'VB 5/6-Tipp 0704: "Laufwerke (insb. USB-Laufwerke) identifizieren" von Arne Elster
'http://www.activevb.de/tipps/vb6tipps/tipp0704.html

Private Declare Function GetLogicalDriveStringsW Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As LongPtr) As Long

Private m_Drives As List 'Of PFNDirectory(PFNDrive)

Public Property Get LogicalDrives() As List 'Of PFNDrive
    If m_Drives Is Nothing Then
        Set m_Drives = MNew.List(vbObject)
        Populate
    End If
    Set LogicalDrives = m_Drives
End Property

Private Sub Populate()
    Dim lBuffer As Long:   lBuffer = 64
    Dim sBuffer As String: sBuffer = Space$(lBuffer)
    Dim lResult As Long:   lResult = GetLogicalDriveStringsW(lBuffer, StrPtr(sBuffer))
    Dim sDrives As String: sDrives = Left$(sBuffer, lResult - 1)
    Dim i As Long, sDrive As String
    Dim sa() As String: sa = Split(sDrives, vbNullChar)
    For i = 0 To UBound(sa)
        sDrive = sa(i)
        If Len(sDrive) Then
            m_Drives.Add MNew.PFNDrive(sDrive)
        End If
    Next
'    Do While i < Len(sBuffer)
'        i = InStr(sBuffer, vbNullChar)
'        If i = 0 Then Exit Do
'        sDrive = Left$(sBuffer, i - 1)
'        sBuffer = Mid$(sBuffer, i + 1, Len(sBuffer))
'    Loop
End Sub
