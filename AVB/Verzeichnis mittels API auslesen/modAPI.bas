Attribute VB_Name = "modAPI"
Option Explicit

Private Declare Function FindFirstFile Lib "kernel32" _
                  Alias "FindFirstFileA" ( _
                         ByVal lpFileName As String, _
                         ByRef lpFindFileData As WIN32_FIND_DATA) As Long
        
Private Declare Function FindNextFile Lib "kernel32" _
                  Alias "FindNextFileA" ( _
                         ByVal hFindFile As Long, _
                         ByRef lpFindFileData As WIN32_FIND_DATA) As Long
        
Private Declare Function FindClose Lib "kernel32" (ByVal _
                         hFindFile As Long) As Long


Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260

Public Type FILETIME
  dwLowDateTime      As Long
  dwHighDateTime     As Long
End Type

Public Type WIN32_FIND_DATA
  dwFileAttributes   As Long
  ftCreationTime     As FILETIME
  ftLastAccessTime   As FILETIME
  ftLastWriteTime    As FILETIME
  nFileSizeHigh      As Long
  nFileSizeLow       As Long
  dwReserved0        As Long
  dwReserved1        As Long
  cFileName          As String * MAX_PATH
  cAlternate         As String * 14
End Type

Public Type SearchResult
   Folders()         As WIN32_FIND_DATA
   Files()           As WIN32_FIND_DATA
   FolderCount       As Long
   FileCount         As Long
End Type

Public Function GetFiles(ByVal SearchPattern As String, ByRef Result As SearchResult) As Boolean
   Dim FileName      As String
   Dim FolderCount   As Long
   Dim FileCount     As Long
   Dim hSearch       As Long
   Dim FindData      As WIN32_FIND_DATA
   
   hSearch = FindFirstFile(SearchPattern, FindData)
   If hSearch = INVALID_HANDLE_VALUE Then
      ' Suche ist fehlgeschlagen (bspw. keine Suchergebnisse). False zurückgeben und Funktion verlassen.
      GetFiles = False
      Exit Function
   End If
   
   ReDim Result.Files(100)
   ReDim Result.Folders(100)
   
   Do
      If (FindData.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) <> 0 Then
         ' Es handelt sich um einen Ordner
         Result.Folders(FolderCount) = FindData
         FolderCount = FolderCount + 1
         
         ' Muss zusätzlicher Platz geschaffen werden?
         If FolderCount > UBound(Result.Folders) Then
            ReDim Preserve Result.Folders(FolderCount + 100)
         End If
      Else
         ' Es handelt sich um eine Datei
         Result.Files(FileCount) = FindData
         FileCount = FileCount + 1
         
         ' Muss zusätzlicher Platz geschaffen werden?
         If FileCount > UBound(Result.Files) Then
            ReDim Preserve Result.Files(FileCount + 100)
         End If
      End If
   Loop While FindNextFile(hSearch, FindData) <> 0
   
   ' Such-Handle schließen
   Call FindClose(hSearch)
   
   
   If FolderCount > 0 Then
      ' Wenn Ordner gefunden wurden, Array zurechtschneiden
      ReDim Preserve Result.Folders(FolderCount - 1)
   Else
      ' Wenn keine Ordner gefunden wurden, Array löschen
      Erase Result.Folders
   End If
   
   If FileCount > 0 Then
      ' Wenn Dateien gefunden wurden, Array zurechtschneiden
      ReDim Preserve Result.Files(FileCount - 1)
   Else
      ' Wenn keine Dateien gefunden wurden, Array löschen
      Erase Result.Files
   End If
   
   ' Anzahl gefundener Ordner und Dateien setzen
   Result.FolderCount = FolderCount
   Result.FileCount = FileCount
   
   GetFiles = True
End Function
