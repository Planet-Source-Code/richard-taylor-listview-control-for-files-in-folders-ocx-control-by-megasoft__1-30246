Attribute VB_Name = "mod3"
Option Explicit

'Api's For getting computer information
Public Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

'Const's for getting computer information
Public Const KEY_QUERY_VALUE = &H1
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_DYN_DATA = &H80000006
Public Const RK_Processor = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
Public Const RK_Performance = "PerfStats\StatData"
Public Const RK_WIN32_OS = "SOFTWARE\Microsoft\Windows\CurrentVersion"
Public Const RK_WIN32_OS_NT = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"

'Api for Moving Files
Public Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

'Constants for saving last done
Global Const LISTVIEW_MODE0 = "View Large Icons"
Global Const LISTVIEW_MODE1 = "View Small Icons"
Global Const LISTVIEW_MODE2 = "View List"
Global Const LISTVIEW_MODE3 = "View Details"

'Api for shutting down and restarting the computer
Public Declare Function ExitWindowsEx Lib "User32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Public Const EWX_SHUTDOWN As Long = 1
Public Const EWX_REBOOT As Long = 2

'Api for formatting disk
Public Declare Function SHFormatDrive Lib "shell32" (ByVal hWnd As Long, ByVal Drive As Long, ByVal fmtID As Long, ByVal options As Long) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Global lngIcon
Global strProgram
Global strSaveIconFile

'Api for extracting the first icon in a dll or exe file
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function DrawIcon Lib "User32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function DestroyIcon Lib "User32" (ByVal hIcon As Long) As Long

'Api for creating directory and deleting folders/files
Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long
Declare Function CreateDirectory& Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpnewDirectory As String, lpSecurityAttributes As SECURITY_ATTRIBUTES)
Declare Function GetDesktopWindow& Lib "User32" ()

Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Const FO_MOVE As Long = &H1
Public Const FO_COPY As Long = &H2
Public Const FO_DELETE As Long = &H3
Public Const FO_RENAME As Long = &H4

Public Const FOF_MULTIDESTFILES As Long = &H1
Public Const FOF_CONFIRMMOUSE As Long = &H2
Public Const FOF_SILENT As Long = &H4
Public Const FOF_RENAMEONCOLLISION As Long = &H8
Public Const FOF_NOCONFIRMATION As Long = &H10
Public Const FOF_WANTMAPPINGHANDLE As Long = &H20
Public Const FOF_CREATEPROGRESSDLG As Long = &H0
Public Const FOF_ALLOWUNDO As Long = &H40
Public Const FOF_FILESONLY As Long = &H80
Public Const FOF_SIMPLEPROGRESS As Long = &H100
Public Const FOF_NOCONFIRMMKDIR As Long = &H200

Public Sub File_Delete(path As String)
    Dim FileOperation As SHFILEOPSTRUCT
    Dim lReturn As Long
    Dim sTempFilename As String * 100
    Dim sSendMeToTheBin As String
    sSendMeToTheBin = path
    With FileOperation
        .wFunc = FO_DELETE
        .pFrom = sSendMeToTheBin
        .fFlags = FOF_ALLOWUNDO
    End With
    lReturn = SHFileOperation(FileOperation)
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
    Dim lKey As Long
    Dim tmpVal As String
    Dim tmpKeySize As Long
    Dim tmpKeyType As Long
    Dim Counter As Integer
    tmpVal = String(1024, 0)
    tmpKeySize = 1024
    If RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_QUERY_VALUE, lKey) <> 0 Then
        GetKeyValue = ""
        RegCloseKey lKey
        Exit Function
    End If
    If RegQueryValueEx(lKey, SubKeyRef, 0, tmpKeyType, tmpVal, tmpKeySize) Then
        GetKeyValue = ""
        RegCloseKey lKey
        Exit Function
    End If
    If (Asc(Mid(tmpVal, tmpKeySize, 1)) = 0) Then
        tmpVal = Left(tmpVal, tmpKeySize - 1)
    Else
        tmpVal = Left(tmpVal, tmpKeySize)
    End If
    If tmpKeyType = 4 Then
        For Counter = Len(tmpVal) To 1 Step -1
            GetKeyValue = GetKeyValue + Hex(Asc(Mid(tmpVal, Counter, 1)))
        Next
        GetKeyValue = Format("&h" + GetKeyValue)
    Else
        GetKeyValue = tmpVal
    End If
    RegCloseKey lKey
End Function




