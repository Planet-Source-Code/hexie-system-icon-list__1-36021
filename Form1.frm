VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Icons"
   ClientHeight    =   5865
   ClientLeft      =   1500
   ClientTop       =   1530
   ClientWidth     =   6915
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   5865
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtScanPath 
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Text            =   "c:\"
      Top             =   4980
      Width           =   3135
   End
   Begin VB.CommandButton cmdGetIcons 
      Caption         =   "Get Icons"
      Height          =   375
      Left            =   5820
      TabIndex        =   2
      Top             =   4980
      Width           =   975
   End
   Begin VB.TextBox txtOutput 
      Height          =   4755
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   60
      Width           =   6795
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5490
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   12144
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Scan Path:"
      Height          =   195
      Left            =   1620
      TabIndex        =   4
      Top             =   5040
      Width           =   855
   End
   Begin VB.Image imgIconSmall 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   960
      Top             =   4980
      Width           =   435
   End
   Begin VB.Image imgIconLarge 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   120
      Top             =   4920
      Width           =   615
   End
   Begin ComctlLib.ImageList imlSystemSmall 
      Left            =   2820
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList imlSystemLarge 
      Left            =   3480
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":27A2
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents sysIcons As clsSystemIcons
Attribute sysIcons.VB_VarHelpID = -1



Public Sub OutputText(TextToAdd As String)
    Dim l&, temp$, Prefix$, a$

    With Form1
        a = .txtOutput.Text
        Prefix = ""
        Select Case TextToAdd
            Case vbCrLf
                'Skip adding the time to a line feed.
                temp = TextToAdd + vbCrLf
            Case Else
                'Normal text processing.
                temp = Prefix + TextToAdd + vbCrLf
        End Select
        a = temp + a
        'Take care of overflow.
        l = Len(a)
        If l > 32768 Then
            'Remove some data from the beginning of the text box to make room for the new data.
            a = Left$(a, l - 32768) ' & temp
            l = Len(a)
        End If
        .txtOutput.Text = a
    End With
End Sub







Private Sub ScanFolder(sPath As String)
    Dim hFind As Long
    Dim FindDat As WIN32_FIND_DATA
    Dim sPathOnly As String
    
    If Right$(sPath, 1) = "\" Then
        sPathOnly = sPath
    Else
        sPathOnly = sPath & "\"
    End If
    
    'Get the First filename in the folder.
    hFind = FindFirstFile(sPathOnly & "*.*", FindDat)
    If hFind = INVALID_HANDLE_VALUE Then
        Exit Sub
    Else
        Do
            If (FindDat.dwFileAttributes And FILE_ATTRIBUTE_ALL) = FindDat.dwFileAttributes Then
                sysIcons.AddIcon sPathOnly & FindDat.cFileName
            End If
        Loop While FindNextFile(hFind, FindDat)
    End If
    FindClose hFind
End Sub


Private Sub cmdGetIcons_Click()
    Dim n%, temp$
    Dim LocalCollection As Collection
    
    cmdGetIcons.Enabled = False
    'Scan the path, this populates the icon collection for that path.
    temp = txtScanPath.Text
    Form1.Caption = "Scanning " & temp
    Screen.MousePointer = vbHourglass
    ScanFolder temp
    Screen.MousePointer = vbDefault
    
    'Get a collection roster of all the icons in the image list (so far).
    Set LocalCollection = New Collection
    sysIcons.ReturnRoster LocalCollection, True
    For n = 1 To LocalCollection.Count
        OutputText LocalCollection.Item(n)
        sysIcons.AssignToPicture LocalCollection.Item(n), imgIconLarge, True
        sysIcons.AssignToPicture LocalCollection.Item(n), imgIconSmall, False
        
        'Wait 1 second so the user can see this icon.
        temp = Now
        While DateDiff("s", temp, Now) < 1
            DoEvents
        Wend
    Next
    OutputText ""
    OutputText "-- end of roster --"
    Form1.Caption = "System Icons"
    
    Set LocalCollection = Nothing
    cmdGetIcons.Enabled = True
End Sub

Private Sub Form_Load()
    'Create the system icons class.
    Set sysIcons = New clsSystemIcons
    
    'Get rid of those unsightly borders. (Oh my!).
    'Note: I leave the borders on so I can see them on the form, but at runtime, disappear them.
    imgIconLarge.BorderStyle = 0
    imgIconSmall.BorderStyle = 0
End Sub



Private Sub Form_Unload(Cancel As Integer)
    Set sysIcons = Nothing
End Sub











Private Sub sysIcons_Error(Message As String)
    OutputText Message
End Sub

Private Sub sysIcons_Status(Message As String)
    OutputText Message
End Sub




