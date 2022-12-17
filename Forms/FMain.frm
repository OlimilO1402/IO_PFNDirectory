VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "PFNDirectory"
   ClientHeight    =   6660
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10815
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox PBStatusbar 
      BorderStyle     =   0  'Kein
      Height          =   360
      Left            =   0
      ScaleHeight     =   360
      ScaleWidth      =   10815
      TabIndex        =   4
      Top             =   6255
      Width           =   10815
      Begin VB.TextBox TBResizer 
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         Height          =   300
         Left            =   10440
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Beides
         TabIndex        =   5
         Top             =   120
         Width           =   300
      End
      Begin VB.Label LblStatusbar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "   "
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   315
      End
   End
   Begin VB.PictureBox PBPath 
      Height          =   375
      Left            =   3240
      ScaleHeight     =   315
      ScaleWidth      =   7395
      TabIndex        =   2
      Top             =   0
      Width           =   7455
      Begin VB.Label LblPath 
         AutoSize        =   -1  'True
         Caption         =   "    "
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   420
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   0
      Width           =   3255
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   10695
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpenFolder 
         Caption         =   "OpenFolder"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Root As PFNDirectory
Private bInit As Boolean

Private Sub Form_Load()
    bInit = True
    MPC.LogicalDrives.ToListbox Combo1
    Combo1.ListIndex = 0
    bInit = False
End Sub

Private Sub Form_Resize()
    Dim L As Single
    Dim T As Single: T = List1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T - PBStatusbar.Height
    If W > 0 And H > 0 Then List1.Move L, T, W, H
    
    L = 0: T = Me.ScaleHeight - PBStatusbar.Height
    W = Me.ScaleWidth: H = PBStatusbar.Height
    If W > 0 And H > 0 Then PBStatusbar.Move L, T, W, H
    
    L = PBStatusbar.Width - TBResizer.Width
    T = PBStatusbar.Height - TBResizer.Height
    W = 285: H = 285
    If W > 0 And H > 0 Then TBResizer.Move L, T, W, H
End Sub

Private Sub List1_DblClick()
    Dim i As Long:   i = List1.ListIndex
    Dim s As String: s = List1.List(i)
    Dim T As String: T = Left(s, 4)
    Dim f As String: f = Mid(s, 5)
    If T = "[||]" Then
        'OK it's a directory, now we change into this directory
        Dim nd As PFNDirectory: Set nd = m_Root.Directories(i)
        'now nd is a child of the Parent Root
        'do not use IIF here because it's cumbersome, If() would be OK though
        If nd.Name = ".." Then
            'but now we want to go back
            'so we want the parent of m_Root
            Set m_Root = m_Root.Root
        Else
            Set m_Root = nd
        End If
        Debug.Print m_Root.ToPFN.Value
        Dim mp As MousePointerConstants: mp = Me.MousePointer
        Me.MousePointer = vbArrowHourglass
        m_Root.Populate
        UpdateView
        Me.MousePointer = mp
    ElseIf T = "[|o]" Then
        'this is a file, so we must decide what to do with it
        'if it's an exe we start the program, or if any other file we start the associated program and open the file in it
        'or maybe we do a preview window or we just show file properties
    End If
    
End Sub

Private Sub mnuFileOpenFolder_Click()
    Dim OFD As New OpenFolderDialog
    If OFD.ShowDialog = vbCancel Then Exit Sub
    Dim s As String: s = OFD.Folder
    If Len(s) = 0 Then
        MsgBox "Empty string"
        Exit Sub
    End If
    Set m_Root = MNew.PFNDirectoryP(MNew.PathFileName(OFD.Folder))
    m_Root.Populate
    UpdateView
End Sub

Private Sub Combo1_Click()
    Dim i As Long: i = Combo1.ListIndex
    If i < 0 Or Combo1.ListCount - 1 < i Then Exit Sub
    Dim s As String: s = Combo1.List(i)
    Dim drv As PFNDrive: Set drv = MPC.LogicalDrives.Item(i)
    Set m_Root = drv
    If Not m_Root.Populate Then MsgBox "Drive is not ready!"
    UpdateView
End Sub

Sub UpdateView()
    If m_Root Is Nothing Then Exit Sub
    Dim i As Long
    List1.Clear
    List1.Enabled = False
    List1.Visible = False
    Dim dc As Long: dc = m_Root.Directories.Count
    Dim fc As Long: fc = m_Root.Files.Count
    Dim pre As String
    Dim p As PFNDirectory
    For i = 0 To dc - 1
        Set p = m_Root.Directories.Item(i)
        pre = "[||] "
        List1.AddItem pre & p.Name
        'if it's open pre is "[//] "
    Next
    'List2.Clear
    Dim f As PathFileName
    For i = 0 To fc - 1
        Set f = m_Root.Files.Item(i)
        pre = "[o|] "
        List1.AddItem pre & f.Value
    Next
    LblStatusbar.Caption = "Directories: " & dc & "; Files: " & fc
    LblPath.Caption = m_Root.ToStr
    List1.Visible = True
    List1.Enabled = True
End Sub
