Attribute VB_Name = "WinLibRoutines"
'
' WinLibRoutines.bas - collection of windows library calls
'  taken from various locations on the internet as indicated
'
' Portions of code below taken from:
' http://www.mvps.org/access/api/api0004.htm
' Courtesy of Terry Kreft
Option Explicit

Private Const STARTF_USESHOWWINDOW& = &H1
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessID As Long
    dwThreadId As Long
End Type

#If VBA7 Then
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal _
    hHandle As Long, ByVal dwMilliseconds As Long) As Long
    
Private Declare PtrSafe Function CreateProcessA Lib "kernel32" (ByVal _
    lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
    lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As _
    PROCESS_INFORMATION) As Long
    
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal _
    hObject As Long) As Long
    
Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long
    
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

Public Declare PtrSafe Function TerminateProcess Lib "kernel32" _
    (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    
Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal Operation As String, _
  ByVal Filename As String, Optional ByVal Parameters As String, _
  Optional ByVal directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long

#Else
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
    hHandle As Long, ByVal dwMilliseconds As Long) As Long
    
Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
    lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
    lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
    ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
    lpStartupInfo As STARTUPINFO, lpProcessInformation As _
    PROCESS_INFORMATION) As Long
    
Private Declare Function CloseHandle Lib "kernel32" (ByVal _
    hObject As Long) As Long
    
Private Declare Function GetExitCodeProcess Lib "kernel32" _
    (ByVal hProcess As Long, lpExitCode As Long) As Long
    
Private Declare Function GetLastError Lib "kernel32" () As Long

Public Declare Function TerminateProcess Lib "kernel32" _
    (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    
Public Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal Operation As String, _
  ByVal Filename As String, Optional ByVal Parameters As String, _
  Optional ByVal Directory As String, _
  Optional ByVal WindowStyle As Long = vbMinimizedFocus _
  ) As Long
#End If

'Written: August 02, 2010
'Author:  Leith Ross
'Summary: Makes the UserForm resizable by dragging one of the sides. Place a call
'         to the macro MakeFormResizable in the UserForm's Activate event.
'Source: http://www.mrexcel.com/forum/excel-questions/485489-resize-userform.html

#If VBA7 Then
 Private Declare PtrSafe Function SetLastError _
   Lib "kernel32.dll" _
     (ByVal dwErrCode As Long) _
   As Long
   
 Public Declare PtrSafe Function GetActiveWindow _
   Lib "user32.dll" () As Long

 Private Declare PtrSafe Function GetWindowLong _
   Lib "user32.dll" Alias "GetWindowLongA" _
     (ByVal hWnd As Long, _
      ByVal nIndex As Long) _
   As Long
               
 Private Declare PtrSafe Function SetWindowLong _
   Lib "user32.dll" Alias "SetWindowLongA" _
     (ByVal hWnd As Long, _
      ByVal nIndex As Long, _
      ByVal dwNewLong As Long) _
   As Long
 
 Private Declare PtrSafe Function GetDC Lib "user32" _
    (ByVal hWnd As Long) As Long

 Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" _
    (ByVal hDC As Long, ByVal nIndex As Long) As Long

 Private Declare PtrSafe Function ReleaseDC Lib "user32" _
    (ByVal hWnd As Long, ByVal hDC As Long) As Long

#Else
 Private Declare Function SetLastError _
   Lib "kernel32.dll" _
     (ByVal dwErrCode As Long) _
   As Long
   
 Public Declare Function GetActiveWindow _
   Lib "user32.dll" () As Long

 Private Declare Function GetWindowLong _
   Lib "user32.dll" Alias "GetWindowLongA" _
     (ByVal hwnd As Long, _
      ByVal nIndex As Long) _
   As Long
               
 Private Declare Function SetWindowLong _
   Lib "user32.dll" Alias "SetWindowLongA" _
     (ByVal hwnd As Long, _
      ByVal nIndex As Long, _
      ByVal dwNewLong As Long) _
   As Long
   
 Private Declare Function GetDC Lib "User32" _
    (ByVal hwnd As Long) As Long

 Private Declare Function GetDeviceCaps Lib "gdi32" _
    (ByVal hDC As Long, ByVal nIndex As Long) As Long

 Private Declare Function ReleaseDC Lib "User32" _
    (ByVal hwnd As Long, ByVal hDC As Long) As Long
#End If

Private Const LOGPIXELSX = 88  'Pixels/inch in


' Portions taken from:
' http://www.kbalertz.com/kb_145679.aspx
   
   'Option Explicit

   Public Const REG_SZ As Long = 1
   Public Const REG_DWORD As Long = 4

   Public Const HKEY_CLASSES_ROOT = &H80000000
   Public Const HKEY_CURRENT_USER = &H80000001
   Public Const HKEY_LOCAL_MACHINE = &H80000002
   Public Const HKEY_USERS = &H80000003

   Public Const ERROR_NONE = 0
   Public Const ERROR_BADDB = 1
   Public Const ERROR_BADKEY = 2
   Public Const ERROR_CANTOPEN = 3
   Public Const ERROR_CANTREAD = 4
   Public Const ERROR_CANTWRITE = 5
   Public Const ERROR_OUTOFMEMORY = 6
   Public Const ERROR_ARENA_TRASHED = 7
   Public Const ERROR_ACCESS_DENIED = 8
   Public Const ERROR_INVALID_PARAMETERS = 87
   Public Const ERROR_NO_MORE_ITEMS = 259

   Public Const KEY_QUERY_VALUE = &H1
   Public Const KEY_SET_VALUE = &H2
   Public Const KEY_ALL_ACCESS = &H3F

   Public Const REG_OPTION_NON_VOLATILE = 0

   #If VBA7 Then
   Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" _
   (ByVal hKey As Long) As Long
   Declare PtrSafe Function RegCreateKeyEx Lib "advapi32.dll" Alias _
   "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
   As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
   As Long, phkResult As Long, lpdwDisposition As Long) As Long
   Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias _
   "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
   Long) As Long
   Declare PtrSafe Function RegQueryValueExString Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As String, lpcbData As Long) As Long
   Declare PtrSafe Function RegQueryValueExLong Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, lpData As _
   Long, lpcbData As Long) As Long
   Declare PtrSafe Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As Long, lpcbData As Long) As Long
   Declare PtrSafe Function RegSetValueExString Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
   String, ByVal cbData As Long) As Long
   Declare PtrSafe Function RegSetValueExLong Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
   ByVal cbData As Long) As Long
   #Else
   Declare Function RegCloseKey Lib "advapi32.dll" _
   (ByVal hKey As Long) As Long
   Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
   "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions _
   As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes _
   As Long, phkResult As Long, lpdwDisposition As Long) As Long
   Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
   "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
   ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As _
   Long) As Long
   Declare Function RegQueryValueExString Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As String, lpcbData As Long) As Long
   Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, lpData As _
   Long, lpcbData As Long) As Long
   Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias _
   "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
   String, ByVal lpReserved As Long, lpType As Long, ByVal lpData _
   As Long, lpcbData As Long) As Long
   Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As _
   String, ByVal cbData As Long) As Long
   Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
   "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
   ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, _
   ByVal cbData As Long) As Long
   #End If


   ' From https://social.msdn.microsoft.com/Forums/en-US/7d584120-a929-4e7c-9ec2-9998ac639bea/mouse-scroll-in-userform-listbox-in-excel-2010?forum=isvvba
'
'Enables mouse wheel scrolling in controls
'

#If Win64 Then
    Private Type POINTAPI
       XY As LongLong
    End Type
#Else
    Private Type POINTAPI
           X As Long
           Y As Long
    End Type
#End If

Private Type MOUSEHOOKSTRUCT
    Pt As POINTAPI
    hWnd As Long
    wHitTestCode As Long
    dwExtraInfo As Long
End Type

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" _
                                            Alias "FindWindowA" ( _
                                                            ByVal lpClassName As String, _
                                                            ByVal lpWindowName As String) As Long ' not sure if this should be LongPtr
    #If Win64 Then
        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" _
                                            Alias "GetWindowLongPtrA" ( _
                                                            ByVal hWnd As LongPtr, _
                                                            ByVal nIndex As Long) As LongPtr
    #Else
        Private Declare PtrSafe Function GetWindowLong Lib "user32" _
                                            Alias "GetWindowLongA" ( _
                                                            ByVal hWnd As LongPtr, _
                                                            ByVal nIndex As Long) As LongPtr
    #End If
    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" _
                                            Alias "SetWindowsHookExA" ( _
                                                            ByVal idHook As Long, _
                                                            ByVal lpfn As LongPtr, _
                                                            ByVal hmod As LongPtr, _
                                                            ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" ( _
                                                            ByVal hHook As LongPtr, _
                                                            ByVal nCode As Long, _
                                                            ByVal wParam As LongPtr, _
                                                           lParam As Any) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" ( _
                                                            ByVal hHook As LongPtr) As LongPtr ' MAYBE Long
    #If Win64 Then
        Private Declare PtrSafe Function WindowFromPoint Lib "user32" ( _
                                                            ByVal Point As LongLong) As LongPtr    '
    #Else
        Private Declare PtrSafe Function WindowFromPoint Lib "user32" ( _
                                                            ByVal xPoint As Long, _
                                                            ByVal yPoint As Long) As LongPtr    '
    #End If
    Private Declare PtrSafe Function GetCursorPos Lib "user32" ( _
                                                            ByRef lpPoint As POINTAPI) As LongPtr   'MAYBE Long
#Else
    Private Declare Function FindWindow Lib "user32" _
                                            Alias "FindWindowA" ( _
                                                            ByVal lpClassName As String, _
                                                            ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32.dll" _
                                            Alias "GetWindowLongA" ( _
                                                            ByVal hWnd As Long, _
                                                            ByVal nIndex As Long) As Long
    Private Declare Function SetWindowsHookEx Lib "user32" _
                                            Alias "SetWindowsHookExA" ( _
                                                            ByVal idHook As Long, _
                                                            ByVal lpfn As Long, _
                                                            ByVal hmod As Long, _
                                                            ByVal dwThreadId As Long) As Long
    Private Declare Function CallNextHookEx Lib "user32" ( _
                                                            ByVal hHook As Long, _
                                                            ByVal nCode As Long, _
                                                            ByVal wParam As Long, _
                                                           lParam As Any) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
                                                            ByVal hHook As Long) As Long
    Private Declare Function WindowFromPoint Lib "user32" ( _
                                                            ByVal xPoint As Long, _
                                                            ByVal yPoint As Long) As Long
    Private Declare Function GetCursorPos Lib "user32.dll" ( _
                                                            ByRef lpPoint As POINTAPI) As Long
#End If


'Attribute VB_Name = "CopyToClipboard"
#If VBA7 Then
Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
  ByVal dwBytes As LongPtr) As LongPtr
Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Declare PtrSafe Function EmptyClipboard Lib "user32" () As Long
Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
  ByVal lpString2 As Any) As Long
Declare PtrSafe Function SetClipboardData Lib "user32" (ByVal wFormat _
  As Long, ByVal hMem As LongPtr) As LongPtr
#Else
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
  ByVal dwBytes As Long) As Long
Declare Function CloseClipboard Lib "User32" () As Long
Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
Declare Function EmptyClipboard Lib "User32" () As Long
Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
  ByVal lpString2 As Any) As Long
Declare Function SetClipboardData Lib "User32" (ByVal wFormat _
  As Long, ByVal hMem As Long) As Long
#End If

'Attribute VB_Name = "MouseWheel"

Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const HC_ACTION As Long = 0
Private Const GWL_HINSTANCE As Long = (-6)
'Private Const WM_KEYDOWN As Long = &H100
'Private Const WM_KEYUP As Long = &H101
'Private Const VK_UP As Long = &H26
'Private Const VK_DOWN As Long = &H28
'Private Const WM_LBUTTONDOWN As Long = &H201
Dim n As Long
Private mCtl As MSForms.control
Private mbHook As Boolean
#If VBA7 Then
    Private mLngMouseHook As LongPtr
    Private mListBoxHwnd As LongPtr
#Else
    Private mLngMouseHook As Long
    Private mListBoxHwnd As Long
#End If

' for clipboard functions
Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096


' No VT_GUID available so must declare type GUID
' see https://stackoverflow.com/questions/45332357/ms-access-vba-error-run-time-error-70-permission-denied'
Private Type GUID_TYPE
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

#If VBA7 Then
 Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (guid As GUID_TYPE) As LongPtr
 Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (guid As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As LongPtr
#Else
 Declare  Function CoCreateGuid Lib "ole32.dll" (Guid As GUID_TYPE) As Long
 Declare  Function StringFromGUID2 Lib "ole32.dll" (Guid As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As Long
#End If

Function CreateGuidString()
    Dim guid As GUID_TYPE
    Dim strGuid As String
    Dim retValue As LongPtr
    Const guidLength As Long = 39 'registry GUID format with null terminator {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}
    retValue = CoCreateGuid(guid)
    If retValue = 0 Then
        strGuid = String$(guidLength, vbNullChar)
        retValue = StringFromGUID2(guid, StrPtr(strGuid), guidLength)
        If retValue = guidLength Then
            ' valid GUID as a string
            CreateGuidString = strGuid
        End If
    End If
End Function

Public Function GetUUID(Optional lowercase As Boolean = True, Optional parens As Boolean = False) As String
    ' default is lowercase without { and }
    GetUUID = CreateGuidString()
    ' strip off trailing null
    If Right(GetUUID, 1) = vbNullChar Then GetUUID = Left(GetUUID, Len(GetUUID) - 1)
    If lowercase Then GetUUID = LCase(GetUUID)
    If Not parens Then
        If Right(GetUUID, 1) = "}" Then GetUUID = Left(GetUUID, Len(GetUUID) - 1)
        If Left(GetUUID, 1) = "{" Then GetUUID = Right(GetUUID, Len(GetUUID) - 1)
    End If
    Debug.Print GetUUID
End Function
    
Public Function ShellWait(Pathname As String, Optional StartupDir As String, Optional WindowStyle As Long, Optional WaitTime As Long = -1) As Long
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim ret As Long
    Dim exitcode As Long
    Dim lastError As Long
    Dim retWait As Long
    
    ' Initialize the STARTUPINFO structure:
    With start
        .cb = Len(start)
        If Not IsMissing(WindowStyle) Then
            .dwFlags = STARTF_USESHOWWINDOW
            .wShowWindow = WindowStyle
        End If
    End With
    Dim sdir As String
    If IsMissing(StartupDir) Then
        sdir = ""
    Else
        sdir = StartupDir
    End If

    ' Start the shelled application:
    ret& = CreateProcessA(0&, Pathname, 0&, 0&, 1&, _
            NORMAL_PRIORITY_CLASS, 0&, sdir, start, proc)
    lastError& = GetLastError()
    If (ret& = 0) Then
        MsgBox "Could not start process: '" & Pathname & "'. GetLastError returned " & Str(lastError&)
        ShellWait = 1
        Exit Function
    End If
        
    ' Wait for the shelled application to finish:
    If WaitTime > 0 Then
        retWait& = WaitForSingleObject(proc.hProcess, WaitTime)
    Else
        retWait& = WaitForSingleObject(proc.hProcess, INFINITE)
    End If
    ' Get return value
    exitcode& = 1234
    ret& = GetExitCodeProcess(proc.hProcess, exitcode&)
    If (ret& = 0) Then
        lastError& = GetLastError()
        MsgBox "GetExitCodeProcess returned " + Str(ret&) + ", GetLastError returned " + Str(lastError&)
    End If
    ' Tidy up if time out
    If (retWait& = 258) Then
        ret& = TerminateProcess(proc.hProcess, 0)
    End If
    ' Close handle
    ret& = CloseHandle(proc.hProcess)
    ShellWait = exitcode&
End Function

Public Function Execute(CommandLine As String, StartupDir As String, Optional debugMode As Boolean = False, Optional WaitTime As Long = -1) As Long
    Dim RetVal As Long
    If debugMode Then
        ClipBoard_SetData CommandLine
        MsgBox CommandLine, , StartupDir
        RetVal = ShellWait(CommandLine, StartupDir, 1&, WaitTime)
    Else
        RetVal = ShellWait(CommandLine, StartupDir, , WaitTime)
    End If
    Execute = RetVal
End Function


'
Public Function lDotsPerInch() As Long
    ' The size of a pixel, in points
    ' constant for GetDeviceCaps
    Dim hDC As Long
    hDC = GetDC(0)
    lDotsPerInch = GetDeviceCaps(hDC, LOGPIXELSX)
    ReleaseDC 0, hDC
End Function

Public Function PointsPerPixel() As Double
    'The size of a pixel, in points
    Dim PointsPerInch As Long
    'A point is defined as 1/72 inches
    PointsPerInch = 72
    PointsPerPixel = PointsPerInch / lDotsPerInch()
End Function

Public Sub MakeFormResizable()
    Dim lStyle As Long
    Dim hWnd As Long
    Dim RetVal
    Const WS_THICKFRAME = &H40000
    Const GWL_STYLE As Long = (-16)
    hWnd = GetActiveWindow
    'Get the basic window style
    lStyle = GetWindowLong(hWnd, GWL_STYLE) Or WS_THICKFRAME
    'Set the basic window styles
    RetVal = SetWindowLong(hWnd, GWL_STYLE, lStyle)
    'Clear any previous API error codes
    SetLastError 0
    'Did the style change?
    If RetVal = 0 Then MsgBox "Unable to make UserForm Resizable."
End Sub

Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    Select Case lType
        Case REG_SZ
            sValue = vValue & Chr$(0)
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
        End Select
End Function

Public Function QueryValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String

    On Error GoTo QueryValueExError

    ' Determine the size and type of data to be read
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5

    Select Case lType
        ' For strings
        Case REG_SZ:
            sValue = String(cch, 0)

            lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
            If lrc = ERROR_NONE Then
                vValue = Left$(sValue, cch - 1)
            Else
                vValue = Empty
            End If
        ' For DWORDS
        Case REG_DWORD:
            lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
            If lrc = ERROR_NONE Then vValue = lValue
        Case Else
            'all other data types not supported
            lrc = -1
    End Select

QueryValueExExit:
    QueryValueEx = lrc
    Exit Function

QueryValueExError:
    Resume QueryValueExExit
End Function

Private Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
    Dim hNewKey As Long         'handle to the new key
    Dim lRetVal As Long         'result of the RegCreateKeyEx function

    lRetVal = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
              vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
              0&, hNewKey, lRetVal)
    RegCloseKey (hNewKey)
End Sub

Public Function GetRegistryValue(Hive, Keyname, Valuename, DefaultValue)
       ' returns Value of the registry Vlauename if it exists or user supplied Default Value of it does not
       Dim lRetVal As Long         'result of the API functions
       Dim hKey As Long         'handle of opened key
       Dim vValue As Variant      'setting of queried value

       lRetVal = RegOpenKeyEx(Hive, Keyname, 0, KEY_QUERY_VALUE, hKey)
       lRetVal = QueryValueEx(hKey, Valuename, vValue)
       RegCloseKey (hKey)
       
       If (lRetVal = 0) Then
           GetRegistryValue = vValue
       Else
           GetRegistryValue = DefaultValue
       End If
End Function

Public Sub SetRegistryValue(Hive, ByRef Keyname As String, ByRef Valuename As String, ValueType As Long, value As Variant)
    Dim lRetVal As Long         'result of the SetValueEx function
    Dim hKey As Long         'handle of open key
    
    'open the specified key
    lRetVal = RegOpenKeyEx(Hive, Keyname, 0, KEY_SET_VALUE, hKey)
    If (lRetVal = 0) Then
        lRetVal = SetValueEx(hKey, Valuename, ValueType, value)
        RegCloseKey (hKey)
    Else
        RegCloseKey (hKey)
        Dim MyKeyname As String
        MyKeyname = Keyname
        Dim MyPredefKey As Long
        MyPredefKey = Hive
        CreateNewKey MyKeyname, MyPredefKey
        lRetVal = RegOpenKeyEx(Hive, Keyname, 0, KEY_SET_VALUE, hKey)
        lRetVal = SetValueEx(hKey, Valuename, ValueType, value)
        RegCloseKey (hKey)
    
    End If
        
    If (lRetVal <> 0) Then
        MsgBox "Error saving registry key."
    End If
End Sub



     
Sub HookListBoxScroll(frm As Object, ctl As MSForms.control)
    Dim tPT As POINTAPI
    #If VBA7 Then
        Dim lngAppInst As LongPtr
        Dim hwndUnderCursor As LongPtr
    #Else
        Dim lngAppInst As Long
        Dim hwndUnderCursor As Long
    #End If
    GetCursorPos tPT
    #If Win64 Then
        hwndUnderCursor = WindowFromPoint(tPT.XY)
    #Else
        hwndUnderCursor = WindowFromPoint(tPT.X, tPT.Y)
    #End If
    If Not frm.ActiveControl Is ctl Then
           ctl.SetFocus
    End If
    If mListBoxHwnd <> hwndUnderCursor Then
        UnhookListBoxScroll
        Set mCtl = ctl
        mListBoxHwnd = hwndUnderCursor
        #If Win64 Then
            lngAppInst = GetWindowLongPtr(mListBoxHwnd, GWL_HINSTANCE)
        #Else
            lngAppInst = GetWindowLong(mListBoxHwnd, GWL_HINSTANCE)
        #End If
        ' PostMessage mListBoxHwnd, WM_LBUTTONDOWN, 0&, 0&
        If Not mbHook Then
            mLngMouseHook = SetWindowsHookEx( _
                                            WH_MOUSE_LL, AddressOf MouseProc, lngAppInst, 0)
            mbHook = mLngMouseHook <> 0
        End If
    End If
End Sub

Sub UnhookListBoxScroll()
    If mbHook Then
        Set mCtl = Nothing
        UnhookWindowsHookEx mLngMouseHook
        mLngMouseHook = 0
        mListBoxHwnd = 0
        mbHook = False
    End If
End Sub
#If VBA7 Then
Private Function MouseProc( _
                        ByVal nCode As Long, ByVal wParam As Long, _
                        ByRef lParam As MOUSEHOOKSTRUCT) As LongPtr
#Else
Private Function MouseProc( _
                        ByVal nCode As Long, ByVal wParam As Long, _
                        ByRef lParam As MOUSEHOOKSTRUCT) As Long
#End If
    Dim idx As Long
    Dim tPT As POINTAPI
    On Error GoTo errH
    If (nCode = HC_ACTION) Then
    GetCursorPos tPT
        #If Win64 Then
            ' I moved to ignoring the point returned in lParam because it may be in the wrong coordinates depending on DPI
            ' GetCursorPos gives consistent coordinates regradless.
            ' This may create some racing issues, but it seems to be working fine as far as I can tell...
            'If WindowFromPoint(lParam.Pt.XY) = mListBoxHwnd Then
            If WindowFromPoint(tPT.XY) = mListBoxHwnd Then
                If wParam = WM_MOUSEWHEEL Then
                    MouseProc = True
'                        If lParam.hWnd > 0 Then
'                            postMessage mListBoxHwnd, WM_KEYDOWN, VK_UP, 0
'                        Else
'                            postMessage mListBoxHwnd, WM_KEYDOWN, VK_DOWN, 0
'                        End If
'                        postMessage mListBoxHwnd, WM_KEYUP, VK_UP, 0
                    If TypeOf mCtl Is Frame Then
                        If lParam.hWnd > 0 Then idx = -10 Else idx = 10
                        idx = idx + mCtl.ScrollTop
                        If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
                            mCtl.ScrollTop = idx
                        End If
                    ElseIf TypeOf mCtl Is UserForm Then
                        If lParam.hWnd > 0 Then idx = -10 Else idx = 10
                        idx = idx + mCtl.ScrollTop
                        If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
                            mCtl.ScrollTop = idx
                        End If
                    ElseIf TypeOf mCtl Is TextBox Then
                        If lParam.hWnd > 0 Then idx = -3 Else idx = 3
                        idx = idx + mCtl.CurLine
                        If idx < 0 Then idx = 0
                        If idx > mCtl.LineCount - 1 Then idx = mCtl.LineCount - 1
                        mCtl.CurLine = idx
                    Else
                        If lParam.hWnd > 0 Then idx = -1 Else idx = 1
                        idx = idx + mCtl.ListIndex
                        If idx < 0 Then idx = 0
                        If idx > mCtl.ListCount - 1 Then idx = mCtl.ListCount - 1
                        mCtl.ListIndex = idx
                    End If
                Exit Function
                End If
            Else
                UnhookListBoxScroll
            End If
        #Else
            If WindowFromPoint(tPT.X, tPT.Y) = mListBoxHwnd Then
                If wParam = WM_MOUSEWHEEL Then
                    MouseProc = True
'                        If lParam.hWnd > 0 Then
'                            postMessage mListBoxHwnd, WM_KEYDOWN, VK_UP, 0
'                        Else
'                            postMessage mListBoxHwnd, WM_KEYDOWN, VK_DOWN, 0
'                        End If
'                        postMessage mListBoxHwnd, WM_KEYUP, VK_UP, 0
                    If TypeOf mCtl Is Frame Then
                        If lParam.hWnd > 0 Then idx = -10 Else idx = 10
                        idx = idx + mCtl.ScrollTop
                        If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
                            mCtl.ScrollTop = idx
                        End If
                    ElseIf TypeOf mCtl Is UserForm Then
                        If lParam.hWnd > 0 Then idx = -10 Else idx = 10
                        idx = idx + mCtl.ScrollTop
                        If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
                            mCtl.ScrollTop = idx
                        End If
                    Else
                        If lParam.hWnd > 0 Then idx = -1 Else idx = 1
                        idx = idx + mCtl.ListIndex
                        If idx >= 0 Then mCtl.ListIndex = idx
                    End If
                    Exit Function
                End If
            Else
                UnhookListBoxScroll
            End If
        #End If
    End If
    MouseProc = CallNextHookEx( _
                            mLngMouseHook, nCode, wParam, ByVal lParam)
    Exit Function
errH:
    Debug.Print "error"
    UnhookListBoxScroll
End Function
'#Else
'    Private Function MouseProc( _
'                            ByVal nCode As Long, ByVal wParam As Long, _
'                            ByRef lParam As MOUSEHOOKSTRUCT) As Long
'        Dim idx As Long
'        On Error GoTo errH
'        If (nCode = HC_ACTION) Then
'            If WindowFromPoint(lParam.Pt.X, lParam.Pt.Y) = mListBoxHwnd Then
'                If wParam = WM_MOUSEWHEEL Then
'                    MouseProc = True
''                    If lParam.hWnd > 0 Then
''                    postMessage mListBoxHwnd, WM_KEYDOWN, VK_UP, 0
''                    Else
''                    postMessage mListBoxHwnd, WM_KEYDOWN, VK_DOWN, 0
''                    End If
''                    postMessage mListBoxHwnd, WM_KEYUP, VK_UP, 0
'                    If TypeOf mCtl Is Frame Then
'                        If lParam.hWnd > 0 Then idx = -10 Else idx = 10
'                        idx = idx + mCtl.ScrollTop
'                        If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
'                            mCtl.ScrollTop = idx
'                        End If
'                    ElseIf TypeOf mCtl Is UserForm Then
'                        If lParam.hWnd > 0 Then idx = -10 Else idx = 10
'                        idx = idx + mCtl.ScrollTop
'                        If idx >= 0 And idx < ((mCtl.ScrollHeight - mCtl.Height) + 17.25) Then
'                            mCtl.ScrollTop = idx
'                        End If
'                    Else
'                        If lParam.hWnd > 0 Then idx = -1 Else idx = 1
'                        idx = idx + mCtl.ListIndex
'                        If idx >= 0 Then mCtl.ListIndex = idx
'                    End If
'                    Exit Function
'                End If
'            Else
'                UnhookListBoxScroll
'            End If
'        End If
'        MouseProc = CallNextHookEx( _
'        mLngMouseHook, nCode, wParam, ByVal lParam)
'        Exit Function
'errH:
'        UnhookListBoxScroll
'    End Function
'#End If




Function ClipBoard_SetData(MyString As String)
'PURPOSE: API function to copy text to clipboard
'SOURCE: www.msdn.microsoft.com/en-us/library/office/ff192913.aspx

#If VBA7 Then
   Dim hGlobalMemory As LongPtr, lpGlobalMemory As LongPtr, hClipMemory As LongPtr
#Else
   Dim hGlobalMemory As Long, lpGlobalMemory As Long, hClipMemory As Long
#End If

Dim X As Long

'Allocate moveable global memory
  hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

'Lock the block to get a far pointer to this memory.
  lpGlobalMemory = GlobalLock(hGlobalMemory)

'Copy the string to this global memory.
  lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

'Unlock the memory.
  If GlobalUnlock(hGlobalMemory) <> 0 Then
    MsgBox "Could not unlock memory location. Copy aborted."
    GoTo OutOfHere2
  End If

'Open the Clipboard to copy data to.
  If OpenClipboard(0&) = 0 Then
    MsgBox "Could not open the Clipboard. Copy aborted."
    Exit Function
  End If

'Clear the Clipboard.
  X = EmptyClipboard()

'Copy the data to the Clipboard.
  hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:
  If CloseClipboard() = 0 Then
    MsgBox "Could not close Clipboard."
  End If

End Function


