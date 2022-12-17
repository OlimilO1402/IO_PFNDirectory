Attribute VB_Name = "MNew"
Option Explicit

Public Function List(Of_T As EDataType, _
                     Optional ArrColStrTypList, _
                     Optional ByVal IsHashed As Boolean = False, _
                     Optional ByVal Capacity As Long = 32, _
                     Optional ByVal GrowRate As Single = 2, _
                     Optional ByVal GrowSize As Long = 0) As List
    Set List = New List: List.New_ Of_T, ArrColStrTypList, IsHashed, Capacity, GrowRate, GrowSize
End Function

Public Function PathFileName(ByVal aPathOrPFN As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathOrPFN, aFileName, aExt
End Function

Public Function PFNDirectory(aRoot As PFNDirectory, Name As String) As PFNDirectory
    Set PFNDirectory = New PFNDirectory: PFNDirectory.New_ aRoot, Name
End Function

Public Function PFNDirectoryP(PFN As PathFileName) As PFNDirectory
    If Not PFN.IsPath Then Exit Function
    Dim root As PFNDirectory: Set root = MNew.PFNDrive(PFN.Drive)
    Dim i As Long
    For i = 0 To PFN.PathCount - 1
        Set root = MNew.PFNDirectory(root, PFN.PathI(i))
    Next
    Set PFNDirectoryP = root
End Function

Public Function PFNDrive(ByVal aLetter As String) As PFNDrive
    Set PFNDrive = New PFNDrive: PFNDrive.New_ aLetter
End Function

