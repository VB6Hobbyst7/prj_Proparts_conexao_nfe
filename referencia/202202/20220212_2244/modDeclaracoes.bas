Attribute VB_Name = "modDeclaracoes"
Option Compare Database   'Usar ordem do banco de dados para comparações
Option Explicit
' Global variables
Global BuiltInToolbarsAvailable As String               ' The state of the toolbars when this app was started
Global Result As Variant                                ' Stores return value of functions
' Type RECT.
Type RECT
    left As Integer
    top As Integer
    right As Integer
    bottom As Integer
End Type

' Windows API Declarations.
#If VBA7 Then
    Declare PtrSafe Function SndPlaySound Lib "MMSystem" (ByVal Lpsound As String, ByVal Flag As Integer) As Integer
    Declare PtrSafe Function GetActiveWindow Lib "User" () As Integer
    Declare PtrSafe Function GetClassName Lib "User" (ByVal hWnd As Integer, ByVal stBuf$, ByVal cch As Integer) As Integer
    Declare PtrSafe Function GetDesktopWindow Lib "User" () As Integer
    Declare PtrSafe Function GetParent Lib "User" (ByVal hWnd As Integer) As Integer
    Declare PtrSafe Function GetWindowRect Lib "User" (ByVal hWnd As Integer, rc As RECT) As Integer
    Declare PtrSafe Function IsIconic Lib "User" (ByVal hWnd As Integer) As Integer
#Else
    Declare Function SndPlaySound Lib "MMSystem" (ByVal Lpsound As String, ByVal Flag As Integer) As Integer
    Declare Function GetActiveWindow Lib "User" () As Integer
    Declare Function GetClassName Lib "User" (ByVal hWnd As Integer, ByVal stBuf$, ByVal cch As Integer) As Integer
    Declare Function GetDesktopWindow Lib "User" () As Integer
    Declare Function GetParent Lib "User" (ByVal hWnd As Integer) As Integer
    Declare Function GetWindowRect Lib "User" (ByVal hWnd As Integer, rc As RECT) As Integer
    Declare Function IsIconic Lib "User" (ByVal hWnd As Integer) As Integer
#End If


' Constants used in the functions above.
Const SW_RESTORE = 9
Const GW_HWNDFIRST = 0
Const GW_HWNDLAST = 1
Const GW_HWNDNEXT = 2
Const GW_HWNDPREV = 3
Const GW_OWNER = 4
Const GW_CHILD = 5
Const LOGPIXELSX = 88
Const LOGPIXELSY = 90
Const MF_BYCOMMAND = &H0
Const MF_BYPOSITION = &H400
Const MF_ENABLED = &H0
Const MF_GRAYED = &H1
Const MF_DISABLED = &H2
Const MF_MENUBREAK = &H40
Const MF_CHECKED = &H8
Const MF_UNCHECKED = &H0
' Access window class
Const WC_ACCESS = "OMain"
' Message box types
Global Const MB_OKCANCEL = &H1
Global Const MB_ABORTRETRYIGNORE = &H2
Global Const MB_YESNOCANCEL = &H3
Global Const MB_YESNO = &H4
Global Const MB_RETRYCANCEL = &H5
' Message box icons
Global Const MB_ICONSTOP = &H10
Global Const MB_ICONQUESTION = &H20
Global Const MB_ICONEXCLAMATION = &H30
Global Const MB_ICONINFORMATION = &H40
' Message box default buttons
Global Const MB_DEFBUTTON1 = &H0
Global Const MB_DEFBUTTON2 = &H100
Global Const MB_DEFBUTTON3 = &H200
' Message box return values
Global Const MB_OK = 1
Global Const MB_CANCEL = 2
Global Const MB_ABORT = 3
Global Const MB_RETRY = 4
Global Const MB_IGNORE = 5
Global Const MB_YES = 6
Global Const MB_NO = 7
' Useful error constants
Global Const ERR_COMMANDNOTAVAILABLE = 2046
Global Const ERR_ACTIONCANCELLED = 2501
Global Const ERR_INVALIDREFTOFIELD = 2465
' Config IDs - used to lookup values in the [Config] table.
Global Const CONFIG_ID_VERSION = 1
Global Const CONFIG_ID_DEFAULTDIR = 2
Global Const CONFIG_ID_MYDATA_DB_NAME = 3
Global Const CONFIG_ID_SAMPDATA_DB_NAME = 4
Global Const CONFIG_ID_ATTACHED_TABLE_NAME = 5
Global Const CONFIG_ID_USERS_TABLE = 5
Global Const CONFIG_ID_APPLICATION_NAME = 6
Global Const CONFIG_ID_LOGINNAME_COLUMN = 7
Global Const CONFIG_ID_PASSWORD_COLUMN = 8
Global Const CONFIG_ID_HELPFILE_NAME = 9

#If VBA7 Then
    Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

#Else
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If



