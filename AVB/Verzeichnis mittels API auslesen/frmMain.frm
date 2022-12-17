VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "www.activevb.de"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox lstResults 
      Height          =   5130
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Dim FoundFiles As modAPI.SearchResult
   Dim FileName   As String
   Dim n          As Long
   
   If modAPI.GetFiles("C:\*.*", FoundFiles) Then
      Me.Caption = CStr(FoundFiles.FileCount) & " Dateien und " & CStr(FoundFiles.FolderCount) & " Ordner gefunden."
      
      For n = 0 To FoundFiles.FolderCount - 1
         FileName = Left$(FoundFiles.Folders(n).cFileName, InStr(1, FoundFiles.Folders(n).cFileName, vbNullChar) - 1)
         Call lstResults.AddItem("\" & FileName)
      Next n
      
      For n = 0 To FoundFiles.FileCount - 1
         FileName = Left$(FoundFiles.Files(n).cFileName, InStr(1, FoundFiles.Files(n).cFileName, vbNullChar) - 1)
         Call lstResults.AddItem(FileName)
      Next n
   Else
      Me.Caption = "Suche ist fehlgeschlagen/Keine Suchergebnisse"
   End If
End Sub
