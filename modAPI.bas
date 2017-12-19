Attribute VB_Name = "modAPI"
Option Explicit
'LICENCE
'This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.

'    This program is copyright 2005 Alasdair King alasdair@webbie.org.uk

'I've put all of the constants and API function calls in this module - Alasdair - 20 September 2002
'And I've moved some of the out again - Wiser Alasdair, 6 June 2006.

'URLDownloadToFile
'Downloads a specified url to a local file
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

'SHGetSpecialFolderLocation
'Returns the Folder ID of the user's My Documents folder (or another folder indicated
'by CSIDL)
Public Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWnd As Long, ByVal nFolder As Long, ppidl As Long) As Long

'SHGetPathFromIDList
'Returns the path (string) from the folder ID obtained by SHGetSpecialFolderLocation
Public Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal Pidl As Long, ByVal pszPath As String) As Long

'Gets system time in milliseconds
Public Declare Function GetTickCount Lib "kernel32" () As Long


'CONSTANTS
' constants for Shell.NameSpace method -- these are the "special folders"
'contained in the Windows shell
Const CSIDL_DESKTOP = &H0 ' Desktop
Const CSIDL_INTERNET = &H1 ' The internet
Const CSIDL_PROGRAMS = &H2 ' Shortcuts in the Programs menu
Const CSIDL_CONTROLS = &H3 ' Control Panel
Const CSIDL_PRINTERS = &H4 ' Printers
Const CSIDL_PERSONAL = &H5 ' Shortcuts to Personal files
Const CSIDL_FAVORITES = &H6 ' Shortcuts to favorite folders
Const CSIDL_STARTUP = &H7 ' Shortcuts to apps that start at boot Time
Const CSIDL_RECENT = &H8 ' Shortcuts to recently used docs
Const CSIDL_SENDTO = &H9 ' Shortcuts for the SendTo menu
Const CSIDL_BITBUCKET = &HA ' Recycle Bin
Const CSIDL_STARTMENU = &HB ' User-defined items in Start Menu
Const CSIDL_DESKTOPDIRECTORY = &H10 ' Directory with all the desktop shortcuts
Const CSIDL_DRIVES = &H11 ' My Computer
Const CSIDL_NETWORK = &H12 ' Network Neighborhood virtual folder
Const CSIDL_NETHOOD = &H13 ' Directory containing objects in the network neighborhood
Const CSIDL_FONTS = &H14 ' Installed fonts
Const CSIDL_TEMPLATES = &H15 ' Shortcuts to document templates
Const CSIDL_COMMON_STARTMENU = &H16 ' Directory with items in the Start menu for all users
Const CSIDL_COMMON_PROGRAMS = &H17 ' Directory with items in the Programs menu for all users
Const CSIDL_COMMON_STARTUP = &H18 ' Directory with items in the StartUp submenu for all users
Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19 ' Directory with items on the desktop of all users
Const CSIDL_APPDATA = &H1A ' Folder for application-specific data
Const CSIDL_PRINTHOOD = &H1B ' Directory with references to printer links
Const CSIDL_ALTSTARTUP = &H1D ' (DBCS) Directory corresponding to user 's nonlocalized Startup program group
Const CSIDL_COMMON_ALTSTARTUP = &H1E ' (DBCS) Directory with Startup items for all users
Const CSIDL_COMMON_FAVORITES = &H1F ' Directory with all user's favorit items
Const CSIDL_INTERNET_CACHE = &H20 ' Directory for temporary internet Files
Const CSIDL_COOKIES = &H21 ' Directory for Internet cookies
Const CSIDL_HISTORY = &H22 ' Directory for Internet history items


Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
Public Const REG_SZ = 1 'Unicode nul terminated string
Public Const REG_BINARY = 3 'Free form binary
Public Const REG_DWORD = 4 '32-bit number
Public Const ERROR_SUCCESS = 0&
Public Const RESERVED_NULL = 0
Public Const MIIM_SUBMENU As Long = &H4
Public Const APINULL = 0&
Public Const EM_SCROLLCARET = &HB7
Public Const EM_LINEINDEX = &HBB
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINELENGTH = &HC1
Public Const EM_GETFIRSTVISIBLELINE = &HCE
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_GETLINE = &HC4
Public Const EM_LINESCROLL = &HB6
Public Const EM_SETSEL = &HB1
Public Const KEY_READ = &H20019                    '-- Permission for general read access.
Public Const KEY_QUERY_VALUE = &H1
Public Const MF_MENUBREAK As Long = &H40&  'indicates a vertical break
Public Const MIIM_STATE As Long = &H1
Public Const MIIM_ID As Long = &H2
Public Const MIIM_CHECKMARKS As Long = &H8
Public Const MIIM_TYPE As Long = &H10
Public Const MIIM_DATA As Long = &H20
Public Const MIIM_STRING As Long = &H40
Public Const MIIM_BITMAP As Long = &H80
Public Const MIIM_FTYPE As Long = &H100
Public Const FBYPOSITION_POSITION As Boolean = True
Public Const FBYPOSTION_IDENTIFIER As Boolean = False
Public Const MF_STRING As Long = &H0&
Public Const GWL_WNDPROC As Long = (-4)
Public Const WM_COMMAND = &H111 ' indicates that a command has been intercepted by the app
Public Const MF_BYPOSITION = &H400& ' indicates a menu item by position, not by name, for RemoveMenu
Public Const MF_DISABLED = &H2& ' used in setting menu items (InsertMenuItem etc) to indicate that an item is greyed
Public Const MF_GRAYED = &H1& ' allegedly does the same thing as MF_DISABLED

'TYPES
Public Type DllVersionInfo
   cbSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformID As Long
End Type

Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type


'METHODS


'For removing a menu item
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

'For finding the system time - used for timing
Public Declare Function TimeGetTime Lib "winmm.dll" () As Long

'For finding IE version
Public Declare Function DllGetVersion Lib "Shlwapi.dll" (dwVersion As DllVersionInfo) As Long
    'for finding IE version
    
'For changing to IE homepage
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

'This is declaring a Function called
'"PlaySound" and we tell VB that it's located in the
'Library called "winmm.dll". The word 'Alias' means that "PlaySound"
'is actually stored in the DLL as "PlaysoundA", and written in C++ but
'we can use the function in VB as "PlaySound".
'lpszname = file path, hmodule = 0 and dwflags = (Synchonous = 0/Asynchronously = 1)
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
'returns a long, but 0 = successful play

Public Declare Function RasEnumConnections Lib "RasApi32.DLL" Alias "RasEnumConnectionsA" (lpRasConn As Any, lpcb As Long, lpcConnections As Long) As Long

Public Declare Function RasHangUp Lib "RasApi32.DLL" Alias "RasHangUpA" (ByVal hRasConn As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function FolderRegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
    ' Note that if you declare the lpData parameter as String, you must pass it
    ' By Value.

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
    'must be called before FolderRegQueryEx

'for constructing the Favorites menu
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

'for processing the menu favorites
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function InsertMenuItem Lib "user32" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
    
Public Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long

Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Public Declare Function CreateMenu Lib "user32" () As Long

Public Declare Function CreatePopupMenu Lib "user32" () As Long

Public Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

Public Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

'Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

'for getting the locale ID for the current machine/user
Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

'Used to determine the (window handle of) the current active window in Windows.
Public Declare Function GetForegroundWindow Lib "user32" () As Long


Public Function ReadRegistryEntryNumber(hKey As String, regKey As String, itemKey As String) As Long
'returns the value of the registry key indicated by hKey and regKey
    On Error Resume Next
    Dim lengthState As Integer
    Dim result As Long
    Dim hKeyHandle As Long
    Dim keySize As Long
    Dim initialValue As Long
    'first get the original value so we can restore it when we exit
    'also opens the key for writing
    regKey = regKey & Chr(0)
    result = RegOpenKey( _
        hKey, _
        regKey, _
        hKeyHandle _
    )
    'Debug.Print "res1:" & result
    keySize = Len(initialValue)
    itemKey = itemKey & Chr(0)
    result = RegQueryValueEx(hKeyHandle, itemKey, RESERVED_NULL, REG_DWORD, _
        initialValue, keySize)
    'Debug.Print "res2:" & result
    'Debug.Print "Got from registry number: " & initialValue
    Call RegCloseKey(hKeyHandle)
    'If result <> ERROR_SUCCESS Then MsgBox result
    ReadRegistryEntryNumber = initialValue

End Function

Public Function ReadRegistryEntryString(hKey As String, regKey As String, itemKey As String) As String
'returns the value of the registry key indicated by hKey and regKey
'NOTE: be careful when parsing the results, it seems to produce an ANSI version
'overlaying a Unicode version, e.g. "HELLO L O "
    On Error Resume Next
    Dim lengthState As Integer
    Dim result As Long
    Dim hKeyHandle As Long
    Dim keySize As Long
    Dim initialValue As String
    'first get the original value so we can restore it when we exit
    'also opens the key for writing
    regKey = regKey & Chr(0)
    result = RegOpenKey( _
        hKey, _
        regKey, _
        hKeyHandle _
    )
    'Debug.Print "Result1:" & result
    keySize = 256
    initialValue = String(keySize, 0)
    itemKey = itemKey & Chr(0)
    result = RegQueryValueEx(hKeyHandle, itemKey, RESERVED_NULL, REG_SZ, _
        ByVal initialValue, keySize)
        'Debug.Print "Result2:" & result
    initialValue = Trim(initialValue)
    'Debug.Print "Got from registry: " & initialValue
    Call RegCloseKey(hKeyHandle)
    'If result <> ERROR_SUCCESS Then MsgBox result
    ReadRegistryEntryString = initialValue
End Function




Public Function GetUserDirectory() As String
'uses the Windows API to get the path for the user's home directory,
'My Documents.
    On Error Resume Next
    Dim path As String
    Dim result As Long
    Dim referenceID As Long
    
    'work out where to save them to:
    path = Space(260)
    result = modAPI.SHGetSpecialFolderLocation(0, CSIDL_PERSONAL, referenceID)
    result = modAPI.SHGetPathFromIDList(referenceID, path)
    'assertion: path now contains path to My Documents
    path = Trim(path)
    'take off final null character which trim has left behind
    path = Replace(path, Chr(0), "")
    'return
    GetUserDirectory = path
End Function


Public Function GetLocalApplicationDirectory() As String
'uses the Windows API to get the path for the local application directory,
'Local Settings/Application Data
    On Error Resume Next
    Dim path As String
    Dim result As Long
    Dim referenceID As Long
    
    'work out where to save them to:
    path = Space(260)
    result = modAPI.SHGetSpecialFolderLocation(0, CSIDL_LOCAL_APPDATA, referenceID)
    result = modAPI.SHGetPathFromIDList(referenceID, path)
    'assertion: path now contains path to My Documents
    path = Trim(path)
    'take off final null character which trim has left behind
    path = Replace(path, Chr(0), "")
    'return
    GetLocalApplicationDirectory = path

End Function

Public Function GetApplicationDirectory() As String
'uses the Windows API to get the path for the application directory,
'Application Data
    On Error Resume Next
    Dim path As String
    Dim result As Long
    Dim referenceID As Long
    
    'work out where to save them to:
    path = Space(260)
    result = modAPI.SHGetSpecialFolderLocation(0, CSIDL_APPDATA, referenceID)
    result = modAPI.SHGetPathFromIDList(referenceID, path)
    'assertion: path now contains path to My Documents
    path = Trim(path)
    'take off final null character which trim has left behind
    path = Replace(path, Chr(0), "")
    'return
    GetApplicationDirectory = path
End Function

Public Function GetUserHomeDirectory() As String
'uses the Windows API to get the path for the user's home directory,
'My Documents.
    On Error Resume Next
    Dim path As String
    Dim result As Long
    Dim referenceID As Long
    
    'work out where to save them to:
    path = Space(260)
    result = modAPI.SHGetSpecialFolderLocation(0, CSIDL_PERSONAL, referenceID)
    result = modAPI.SHGetPathFromIDList(referenceID, path)
    'assertion: path now contains path to My Documents
    path = Trim(path)
    'take off final null character which trim has left behind
    path = Replace(path, Chr(0), "")
    'return
    GetUserHomeDirectory = path
End Function

Public Function GetTempDirectory() As String
'uses the Windows API to get the path for the current temp directory
    On Error GoTo localHandler:
    GetTempDirectory = GetLocalApplicationDirectory
    'This is now ....Local Settings\Application Data, but we want
    '....Local Settings\Temp
    'I'm assuming that these aren't localised!
    GetTempDirectory = Replace(GetLocalApplicationDirectory, "Application Data", "Temp")
    Exit Function
localHandler:
    'MsgBox Err.Number & " " & Err.Description & vbNewLine & Err.Source, vbOKOnly, "Error: GetTempDirectory"
    Resume Next
End Function

Public Function GetCurrentDate() As String
'returns current date in RFC822 specification
    Dim datetime As String
    
    datetime = WeekdayName(Weekday(Date), True) & ", "
    datetime = datetime & Day(Date) & " "
    datetime = datetime & MonthName(Month(Date), True) & " "
    datetime = datetime & Right(Year(Date), 2) & " "
    datetime = datetime & DatePart("h", Now) & ":" & DatePart("n", Now)
    
    GetCurrentDate = datetime
End Function

