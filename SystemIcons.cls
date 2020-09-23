VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSystemIcons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'System Icons class.
'The concept here is that you pass it filename or folder.
'The icon is extracted from the system and stored in the large and small image lists.
'As you browse through your system, you can pass it more file names, and the icons build up in the list.
'Duplicates generate an error on add, handled by the resume next, so you never have duplicates in the list.
'NOTE: There is no taking care of count overflow here, this is an example. (I think image lists have 256 limit).
'To view the icons, you can get the roster and assign to picture.
'ToDo: You may want to make AssignTOListView and AssignToTreeView.

'Event Declarations:
Public Event Status(Message As String)
Public Event Error(Message As String)

'Our internal refernces to the imagelists, that are bound to the form image lists. (See class initialize).
Private imlLarge As ImageList
Private imlSmall As ImageList

'For getting shell file info.
Private Const MAX_PATH = 260
Private Type SHFILEINFO
        hIcon As Long                                   '  out: icon
        iIcon As Long                                   '  out: icon index
        dwAttributes As Long                            '  out: SFGAO_ flags
        szDisplayName As String * MAX_PATH              '  out: display name (or path)
        szTypeName As String * 80                       '  out: type name
End Type

Private Const SHGFI_ICON = &H100                         '  get icon
Private Const SHGFI_DISPLAYNAME = &H200                  '  get display name
Private Const SHGFI_TYPENAME = &H400                     '  get type name
Private Const SHGFI_ATTRIBUTES = &H800                   '  get attributes
Private Const SHGFI_ICONLOCATION = &H1000                '  get icon location
Private Const SHGFI_EXETYPE = &H2000                     '  return exe type
Private Const SHGFI_SYSICONINDEX = &H4000                '  get system icon index
Private Const SHGFI_LINKOVERLAY = &H8000                 '  put a link overlay on icon
Private Const SHGFI_SELECTED = &H10000                   '  show icon in selected state
Private Const SHGFI_LARGEICON = &H0                      '  get large icon
Private Const SHGFI_SMALLICON = &H1                      '  get small icon
Private Const SHGFI_OPENICON = &H2                       '  get open icon
Private Const SHGFI_SHELLICONSIZE = &H4                  '  get shell size icon
Private Const SHGFI_PIDL = &H8                           '  pszPath is a pidl
Private Const SHGFI_USEFILEATTRIBUTES = &H10             '  use passed dwFileAttribute

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

'For creating pictures on the fly.
Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lpPictDesc As PictDesc, riid As Guid, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long




Public Function AssignToPicture(Key As String, Destination As Control, AsLarge As Boolean) As Integer
    Err.Clear
    On Error Resume Next
    If AsLarge Then
        Destination.Picture = imlLarge.ListImages(LCase$(Key)).Picture
    Else
        Destination.Picture = imlSmall.ListImages(LCase$(Key)).Picture
    End If
    Select Case Err.Number
        Case 0
            AssignToPicture = True
        Case Else
            AssignToPicture = False
            RaiseEvent Error("Error [" & Err.Description & "] while fetching [" & Key & "].")
    End Select
    On Error GoTo 0
End Function

Private Function GetIcon(Key As String, hIcon As Long, AsLarge As Boolean) As Integer
    Dim result%
    Dim NewPic As Picture
    
    If (hIcon <> 0) Then
        Set NewPic = IconToPicture(hIcon, result)
        If result Then
            'If we got the icon, then create a VB picture object from it.
            Err.Clear
            On Error Resume Next
            Select Case Err.Number
                Case 0
                    'Add this icon to the ImageList.
                    If AsLarge Then
                        imlLarge.ListImages.Add , LCase$(Key), NewPic
                    Else
                        imlSmall.ListImages.Add , LCase$(Key), NewPic
                    End If
                Case Else
                
            End Select
            On Error GoTo 0
        End If
        GetIcon = True
    Else
        GetIcon = False
    End If
End Function
Public Function AddIcon(FileItem As String) As Integer
    Dim n%, temp$, LocalFileItem$
    Dim SFI As SHFILEINFO                       'Shell File Info structure.
    Dim cbSFI As Long                           'size of SFI
    Dim lFlags As Long                          'ShFileInfo uFlags param
    
    LocalFileItem = StripNull(FileItem)
    If FileExists(LocalFileItem) Then n = True
    If FolderExists(LocalFileItem) Then n = True
    If n Then
        'Get large icon.
        cbSFI = Len(SFI)
        lFlags = SHGFI_DISPLAYNAME Or SHGFI_SYSICONINDEX Or SHGFI_TYPENAME Or SHGFI_ICON Or SHGFI_LARGEICON
        SHGetFileInfo LocalFileItem, 0&, SFI, cbSFI, lFlags
        temp = StripNull(SFI.szTypeName)
        If Len(temp) Then
            GetIcon Trim$(temp), SFI.hIcon, True
        End If
        
        'Get small icon.
        lFlags = SHGFI_DISPLAYNAME Or SHGFI_SYSICONINDEX Or SHGFI_TYPENAME Or SHGFI_ICON Or SHGFI_SMALLICON
        SHGetFileInfo LocalFileItem, 0&, SFI, cbSFI, lFlags
        temp = StripNull(SFI.szTypeName)
        If Len(temp) Then
            GetIcon Trim$(temp), SFI.hIcon, False
        End If
    Else
        RaiseEvent Error("File or path not found [" & FileItem & "].")
    End If
End Function


Private Function FileExists(FileName As String) As Integer
    Dim n%

    If Len(FileName) = 0 Then
        FileExists = False
        Exit Function
    End If
    'See if a local disk file exists.
    On Error GoTo NoFileHandles        'We could be out of handles?
Recover_Handles:
     n = FreeFile
    On Error GoTo FileNotFound         'Otherwise this means file does not exist.
     Open FileName For Input As #n
     FileExists = True
Recover_FileNotFound:
     On Error GoTo 0
     Close #n
     Exit Function
     
FileNotFound:
    Select Case Err.Number
        Case 55
            'File already open which means it does exists.
            FileExists = True
        Case Else
            FileExists = False
     End Select
     Resume Recover_FileNotFound
     
NoFileHandles:
    'We are out of handles, so free some up.
    Close
    DoEvents
    Resume Recover_Handles
End Function
Private Function FolderExists(FullPath As String) As Integer
    Dim n%
  
    'See if a local disk folder exists.
    On Error GoTo Error_FolderNotFound
    If (GetAttr(FullPath) And 16) = 16 Then
        'Use bitwise comparison to make sure temp is a directory.
        n = True
    End If
Resume_FolderNotFound:
    On Error GoTo 0
    FolderExists = n
    Exit Function

Error_FolderNotFound:
    n = False
    Resume Resume_FolderNotFound
End Function

Private Function IconToPicture(ByVal hIcon As Long, result As Integer) As IPicture
    Dim NewPic As Picture
    Dim PicConv As PictDesc
    Dim IGuid As Guid
    
    If hIcon = 0 Then
        Set IconToPicture = Nothing
        result = False
    Else
        PicConv.cbSizeofStruct = Len(PicConv)
        PicConv.picType = vbPicTypeIcon
        PicConv.hImage = hIcon
        
        ' Fill in magic IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
        With IGuid
            .Data1 = &H7BF80980
            .Data2 = &HBF32
            .Data3 = &H101A
            .Data4(0) = &H8B
            .Data4(1) = &HBB
            .Data4(2) = &H0
            .Data4(3) = &HAA
            .Data4(4) = &H0
            .Data4(5) = &H30
            .Data4(6) = &HC
            .Data4(7) = &HAB
        End With
        OleCreatePictureIndirect PicConv, IGuid, True, NewPic
        Set IconToPicture = NewPic
        result = True
    End If
End Function
Public Sub ReturnRoster(ResultCollection As Collection, AsLarge As Boolean)
    Dim m%, n%
    
    'Return a collection that contains all the names of the icons in the small or large imagelist.
    If AsLarge Then
        m = imlLarge.ListImages.Count
    Else
        m = imlSmall.ListImages.Count
    End If
    
    'Clear out the collection that was passed to us.
    'Programmers may or may NOT want this.
    While ResultCollection.Count > 0
        ResultCollection.Remove 1
    Wend
    
    For n = 1 To m
        If AsLarge Then
            ResultCollection.Add imlLarge.ListImages(n).Key
        Else
            ResultCollection.Add imlSmall.ListImages(n).Key
        End If
    Next
End Sub

Private Function StripNull(ByVal Item As String) As String
    Dim pos As Long
    
    pos = InStr(Item, Chr$(0))
    If pos Then
        'Character zero detected, chop it off.
        StripNull = Left$(Item, pos - 1)
    Else
        'No character zero, just return the string.
        StripNull = Item
    End If
End Function
Private Sub Class_Initialize()
    'Hard coded link to the form.
    Set imlLarge = Form1.imlSystemLarge
    Set imlSmall = Form1.imlSystemSmall
End Sub


Private Sub Class_Terminate()
    Set imlLarge = Nothing
    Set imlSmall = Nothing
End Sub


