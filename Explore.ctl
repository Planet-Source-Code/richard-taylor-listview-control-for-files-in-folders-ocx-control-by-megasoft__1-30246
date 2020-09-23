VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Explore 
   ClientHeight    =   3615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4815
   ScaleHeight     =   3615
   ScaleWidth      =   4815
   ToolboxBitmap   =   "Explore.ctx":0000
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   1080
      ScaleHeight     =   570
      ScaleWidth      =   570
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   1200
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox pic32 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   2280
      ScaleHeight     =   570
      ScaleWidth      =   570
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox pic16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3360
      ScaleHeight     =   330
      ScaleWidth      =   330
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   360
   End
   Begin MSComctlLib.ImageList iml32 
      Left            =   1800
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList iml16 
      Left            =   3120
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name:"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size:"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "File Type:"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Last Modified:"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Explore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Text
Option Explicit
Const MaxRecentFiles = 8

'Api for Emptying Recycle Bin
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hWnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Const SHERB_NOPROGRESSUI = &H2

'Icon Sizes in pixels
Private Const LARGE_ICON As Integer = 32
Private Const SMALL_ICON As Integer = 16
Private Const MAX_PATH = 260

Private Const ILD_TRANSPARENT = &H1       'Display transparent

'ShellInfo Flags
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'System icon index
Private Const SHGFI_LARGEICON = &H0       'Large icon
Private Const SHGFI_SMALLICON = &H1       'Small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400

Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO                   'As required by ShInfo
  hIcon As Long
  iIcon As Long
  dwAttributes As Long
  szDisplayName As String * MAX_PATH
  szTypeName As String * 80
End Type
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal x&, ByVal y&, ByVal Flags&) As Long

Private ShInfo As SHFILEINFO

Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As LARGE_INTEGER, lpTotalNumberOfBytes As LARGE_INTEGER, lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Dim lngResult As Long

Private Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=ListView1,ListView1,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=ListView1,ListView1,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=ListView1,ListView1,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Event Click() 'MappingInfo=ListView1,ListView1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=ListView1,ListView1,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when you press and release a mouse button and then press and release it again over an object."
'Default Property Values:
Const m_def_Selected = 0
'Property Variables:
Dim m_Selected As Boolean





Private Sub ShowIcons()
    On Error Resume Next
    Dim Item As ListItem
    With ListView1
      '.ListItems.Clear
      .Icons = iml32        'Large
      .SmallIcons = iml16   'Small
      For Each Item In .ListItems
        Item.Icon = Item.Index
        Item.SmallIcon = Item.Index
      Next
    End With
End Sub
Private Function GetIcon(FileName As String, Index As Long) As Long
    Dim hLIcon As Long, hSIcon As Long    'Large & Small Icons
    Dim imgObj As ListImage               'Single bmp in imagelist.listimages collection
    Dim r As Long
    'Get a handle to the small icon
    hSIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    'Get a handle to the large icon
    hLIcon = SHGetFileInfo(FileName, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
    'If the handle(s) exists, load it into the picture box(es)
    If hLIcon <> 0 Then
      'Large Icon
      With pic32
        Set .Picture = LoadPicture("")
        .AutoRedraw = True
        r = ImageList_Draw(hLIcon, ShInfo.iIcon, pic32.hdc, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
      'Small Icon
      With pic16
        Set .Picture = LoadPicture("")
        .AutoRedraw = True
        r = ImageList_Draw(hSIcon, ShInfo.iIcon, pic16.hdc, 0, 0, ILD_TRANSPARENT)
        .Refresh
      End With
      Set imgObj = iml32.ListImages.Add(Index, , pic32.Image)
      Set imgObj = iml16.ListImages.Add(Index, , pic16.Image)
    End If
End Function
Private Sub GetAllIcons()
    Dim Item As ListItem
    Dim FileName As String
    On Local Error Resume Next
    For Each Item In ListView1.ListItems
      FileName = Item.SubItems(1) & Item.Text
      GetIcon FileName, Item.Index
    Next
End Sub
Private Function SizeString(ByVal num_bytes As Double) As String
    Const SIZE_KB As Double = 1024
    Const SIZE_MB As Double = 1024 * SIZE_KB
    Const SIZE_GB As Double = 1024 * SIZE_MB
    Const SIZE_TB As Double = 1024 * SIZE_GB
    If num_bytes < SIZE_KB Then
        SizeString = Format$(num_bytes) & " bytes"
    ElseIf num_bytes < SIZE_MB Then
        SizeString = Format$(num_bytes / SIZE_KB, "0.00") & " KB"
    ElseIf num_bytes < SIZE_GB Then
        SizeString = Format$(num_bytes / SIZE_MB, "0.00") & " MB"
    Else
        SizeString = Format$(num_bytes / SIZE_GB, "0.00") & " GB"
    End If
End Function
Private Sub Initialise()
    On Local Error Resume Next
    'Break the link to iml lists
    ListView1.ListItems.Clear
    ListView1.Icons = Nothing
    ListView1.SmallIcons = Nothing
    'Clear the image lists
    iml32.ListImages.Clear
    iml16.ListImages.Clear
End Sub

Private Function ShellDelete(ParamArray vntFileName() As Variant) As Long
    Dim i As Integer
    Dim sFileNames As String
    Dim SHFileOp As SHFILEOPSTRUCT
    For i = LBound(vntFileName) To UBound(vntFileName)
        sFileNames = sFileNames & vntFileName(i) & vbNullChar
    Next
    sFileNames = sFileNames & vbNullChar
    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = sFileNames
        .fFlags = FOF_ALLOWUNDO
    End With
    ShellDelete = SHFileOperation(SHFileOp)
End Function

Private Function LargeIntegerToDouble(low_part As Long, high_part As Long) As Double
    Dim Result As Double
    Result = high_part
    If high_part < 0 Then
        Result = Result + 2 ^ 32
    End If
    Result = Result * 2 ^ 32
    Result = Result + low_part
    If low_part < 0 Then
        Result = Result + 2 ^ 32
    End If
    LargeIntegerToDouble = Result
End Function

Public Function GetFileSize(FileName As String)
On Error GoTo Oops:
Dim strSize As String
strSize = FileLen(FileName)
If strSize >= "1024" Then
    'Checks if its in KB, B, Or MB
    strSize = CCur(strSize / 1024 & "KiloBytes")
    Else
        If strSize >= "1048576" Then
            'Checks if you can put it as MB not KB
            strSize = CCur(strSize / (1024 * 1024)) & "KB"
            Else
                strSize = CCur(strSize) & "Bytes"
        End If
End If
GetFileSize = strSize
Exit Function
Oops:
GetFileSize = "Error"
    Resume
End Function
Public Function GetProperties(FileName As String)
On Error GoTo Oops:
Dim strProp As String
strProp = GetAttr(FileName)
If strProp = "64" Then
    strProp = "Alias"
End If
If strProp = "32" Then
    strProp = "Archive"
End If
If strProp = "16" Then
    strProp = "Folder"
End If
If strProp = "2" Then
    strProp = "File Hidden"
End If
If strProp = "0" Then
    strProp = "Normal"
End If
If strProp = "1" Then
    strProp = "Read-Only File"
End If
If strProp = "4" Then
    strProp = "System File"
End If
If strProp = "8" Then
    strProp = "Volume"
End If
GetProperties = strProp
Exit Function
Oops:
    GetProperties = "N/A"
    Resume
End Function
Public Function SetHide(FileName As String)
On Error Resume Next
SetAttr FileName, vbHidden
End Function
Public Function SetReadOnly(FileName As String)
    On Error Resume Next
    SetAttr FileName, vbReadOnly
End Function
Public Function SetSystemFile(FileName As String)
    On Error Resume Next
    SetAttr FileName, vbSystem
End Function
Public Function SetNormal(FileName As String)
    On Error Resume Next
    SetAttr FileName, vbNormal
End Function
Public Function GetExtension(FileName As String)
On Error Resume Next
Dim strExt As String
strExt = Right(FileName, 2)
If Left(strExt, 1) = "." Then
    strExt = Right(FileName, 1)
    Exit Function
Else
    strExt = Right(FileName, 3)
If Left(strExt, 1) = "." Then
    strExt = Right(FileName, 2)
    Exit Function
Else
    strExt = Right(FileName, 4)
If Left(strExt, 1) = "." Then
    strExt = Right(FileName, 3)
    Exit Function
Else
    strExt = Right(FileName, 5)
If Left(strExt, 1) = "." Then
    strExt = Right(FileName, 4)
    Exit Function
Else
    GetExtension = "Unknown"
End If
End If
End If
End If
End Function
Public Function GetDate(FileName As String)
    On Error Resume Next
    GetDate = FileDateTime(FileName)
End Function
Public Function Delete(FileName As String)
On Error GoTo Oops:
Kill FileName
Exit Function
Oops:
MsgBox "Could not Delete file"
End Function
Public Function Copy(Original As String, Destination As String)
On Error GoTo Oops:
FileCopy Original, Destination
Exit Function
Oops:
MsgBox "Error trying to Copy file"
End Function
Public Function Cut(Original As String, Destination As String)
On Error GoTo Oops:
FileCopy Original, Destination
Kill Original
Exit Function
Oops:
MsgBox "Error trying to move File"
Resume
End Function
Public Function CreateFolder(path As String)
On Error GoTo Oops:
MkDir path
Exit Function
Oops:
MsgBox "Error trying to create Folder"
Resume
End Function

Public Function DeleteFolder(path As String)
On Error GoTo Oops:
RmDir path
Exit Function
Oops:
MsgBox "Error trying to remove Folder"
Resume
End Function
Public Function Icon(FileName As String)
DestroyIcon lngIcon
lngIcon = ExtractIcon(App.hInstance, FileName, 0)
If lngIcon = 0 Then
    MsgBox "Error - No Icon."
Else
DrawIcon picIcon.hdc, 0, 0, lngIcon
Icon = picIcon.Picture
End If
End Function
Public Sub path(path As String)
Initialise
FillListView1WithFiles path
GetAllIcons
ShowIcons
File1.path = path
End Sub

Private Sub UserControl_Initialize()
path "C:\"
End Sub
Private Sub FillListView1WithFiles(ByVal path As String)
    Dim Item As ListItem
    Dim s As String
    path = CheckPath(path)    'Add '\' to end if not present
    s = Dir(path, vbNormal)
    Do While s <> ""
      Set Item = ListView1.ListItems.Add()
      Item.Key = path & s
      'Item.SmallIcon = "Folder"
      Item.Text = s
      Item.SubItems(1) = path
      s = Dir
    Loop
End Sub
Private Function CheckPath(ByVal path As String) As String
    If Right(path, 1) <> "\" Then
      CheckPath = path & "\"
    Else
      CheckPath = path
    End If
End Function
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ListView1,ListView1,-1,View
Public Property Get View() As ListViewConstants
Attribute View.VB_Description = "Returns/sets the current view of the ListView control."
    View = ListView1.View
End Property

Public Property Let View(ByVal New_View As ListViewConstants)
    ListView1.View() = New_View
    PropertyChanged "View"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    ListView1.View = PropBag.ReadProperty("View", 0)
    ListView1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    ListView1.MultiSelect = PropBag.ReadProperty("MultiSelect", True)
    m_Selected = PropBag.ReadProperty("Selected", m_def_Selected)
End Sub

Private Sub UserControl_Resize()
ListView1.Width = UserControl.Width
ListView1.Height = UserControl.Height
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("View", ListView1.View, 0)
    Call PropBag.WriteProperty("MousePointer", ListView1.MousePointer, 0)
    Call PropBag.WriteProperty("MultiSelect", ListView1.MultiSelect, True)
    Call PropBag.WriteProperty("Selected", m_Selected, m_def_Selected)
End Sub

Private Sub ListView1_Click()
    RaiseEvent Click
End Sub

Private Sub ListView1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ListView1,ListView1,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = ListView1.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    ListView1.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=ListView1,ListView1,-1,MultiSelect
Public Property Get MultiSelect() As Boolean
Attribute MultiSelect.VB_Description = "Returns/sets a value indicating whether a user can make multiple selections in the ListView control and how the multiple selections can be made."
    MultiSelect = ListView1.MultiSelect
End Property

Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)
    ListView1.MultiSelect() = New_MultiSelect
    PropertyChanged "MultiSelect"
End Property

Public Function Selected_File()
Selected_File = ListView1.SelectedItem
End Function

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Selected = m_def_Selected
End Sub

