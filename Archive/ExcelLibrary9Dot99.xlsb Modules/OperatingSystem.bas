Attribute VB_Name = "OperatingSystem"
Option Explicit
Option Base 1

Private Declare PtrSafe Function SHSetValue Lib "SHLWAPI.DLL" Alias "SHSetValueA" (ByVal hKey As Long, ByVal pszSubKey As String, ByVal pszValue As String, ByVal dwType As Long, pvData As String, ByVal cbData As Long) As Long
Private Declare PtrSafe Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As String, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare PtrSafe Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Private Declare PtrSafe Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Purpose   :    Returns the Windows user name of the person logged on in the active account
'Inputs    :
'Outputs   :    N/A
'Author    :    Richard Shepherd in his book Access 2010 VBA Macro Programming
'Date      :
'Notes     :
Public Function ReturnUserName()
    Dim strUser As String, x As Integer
    
    Let strUser = Space$(256)
    
    Let x = GetUserName(strUser, 256)
    Let strUser = RTrim(strUser)
    Let ReturnUserName = Left(strUser, Len(strUser) - 1)
End Function

'Purpose   :    Updates the specified environment variable.
'Inputs    :    sSettingName            The name of the environment variable.
'               sSettingValue           The new value of the environment variable.
'               [bSystemEnvironment]    If true sets a system environment variable,
'                                       else sets local user environment variable.
'Outputs   :    N/A
'Author    :    Andrew Baker
'Date      :    9/Jul/2001
'Notes     :
Public Sub EnvironmentSet(sSettingName As String, sSettingValue As String, Optional bSystemEnvironment As Boolean = False)
    Dim lRet As Long
    Const REG_EXPAND_SZ = 2, HWND_BROADCAST = &HFFFF&, WM_WININICHANGE = &H1A
    Const HKEY_CURRENT_USER = &H80000001, REG_SZ = 1
    Const SHREGSET_FORCE_HKCU = &H1
    Const SMTO_ABORTIFHUNG = &H2
    Const HKEY_LOCAL_MACHINE = &H80000002
    
    'Set the environment variable for the current process
    SetEnvironmentVariable sSettingName, sSettingValue
    
    If bSystemEnvironment = False Then
        'Set the local environment variable for all other processes (via registry)
        lRet = SHSetValue(HKEY_CURRENT_USER, "Environment", sSettingName, REG_EXPAND_SZ, ByVal CStr(sSettingValue), CLng(LenB(StrConv(sSettingValue, vbFromUnicode)) + 1))
    Else
        'Set the system environment variable for all other processes (via registry)
        lRet = SHSetValue(HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Control\Session Manager\Environment", sSettingName, REG_EXPAND_SZ, ByVal CStr(sSettingValue), CLng(LenB(StrConv(sSettingValue, vbFromUnicode)) + 1))
    End If
    
    'Send the environment update message (with a 5 sec timeout)
    Call SendMessageTimeout(HWND_BROADCAST, WM_WININICHANGE, 0, "Environment", SMTO_ABORTIFHUNG, 5000, lRet)
End Sub

'Purpose   :    Replace function for the VB Environ() function. Returns the CURRENT value of
'               an environment variable.
'Inputs    :    sSettingName            The name of the environment setting to return the value of.
'               [sDefaultValue]         The default value to return.
'Outputs   :    Returns the value of the environment variable, else returns the default value
'               if the value was not found.
'Author    :    Andrew Baker
'Date      :    9/Jul/2001
'Notes     :

Function EnvironmentGet(sSettingName As String, Optional sDefaultValue As String = "") As String
    Dim lEndPos As Long
    Dim sBuffer As String * 512
    lEndPos = GetEnvironmentVariable(sSettingName, sBuffer, Len(sBuffer))
    If lEndPos > 0 Then
        'Setting found, return it's value
        EnvironmentGet = Left$(sBuffer, lEndPos)
    Else
        'Return the default value
        EnvironmentGet = sDefaultValue
    End If
End Function

'reads the value for the registry key i_RegKey
'if the key cannot be found, the return value is ""
' Example: Debug.Print RegKeyRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable")
Function RegKeyRead(i_RegKey As String) As String
Dim myWS As Object

  On Error Resume Next
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'read key from registry
  RegKeyRead = myWS.RegRead(i_RegKey)
End Function

'returns True if the registry key i_RegKey was found
'and False if not
Function RegKeyExists(i_RegKey As String) As Boolean
Dim myWS As Object

  On Error GoTo ErrorHandler
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'try to read the registry key
  myWS.RegRead i_RegKey
  'key was found
  RegKeyExists = True
  Exit Function
  
ErrorHandler:
  'key was not found
  RegKeyExists = False
End Function

'sets the registry key i_RegKey to the
'value i_Value with type i_Type
'if i_Type is omitted, the value will be saved as string
'if i_RegKey wasn't found, a new registry key will be created
'Example: call RegKeySave("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyEnable",0)
Sub RegKeySave(i_RegKey As String, _
               i_Value As String, _
      Optional i_Type As String = "REG_SZ")
Dim myWS As Object

  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'write registry key
  myWS.RegWrite i_RegKey, i_Value, i_Type

End Sub

'deletes i_RegKey from the registry
'returns True if the deletion was successful,
'and False if not (the key couldn't be found)
Function RegKeyDelete(i_RegKey As String) As Boolean
Dim myWS As Object

  On Error GoTo ErrorHandler
  'access Windows scripting
  Set myWS = CreateObject("WScript.Shell")
  'delete registry key
  myWS.RegDelete i_RegKey
  'deletion was successful
  RegKeyDelete = True
  Exit Function

ErrorHandler:
  'deletion wasn't successful
  RegKeyDelete = False
End Function

Public Function GenGUID() As String
    Let GenGUID = Mid$(CreateObject("Scriptlet.TypeLib").GUID, 2, 36)
End Function


