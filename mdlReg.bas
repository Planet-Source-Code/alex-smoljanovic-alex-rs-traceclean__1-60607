Attribute VB_Name = "modReg"
Option Explicit

'***********************************************************************
'This application was developed for a
'PSC(Planet Source Code) User(s) on request.
'
'If you compile this application, please dont distribute it.
'However, feel free to use any of this code in you're own application(s).
'
'Alex Smoljanovic [Salex] 2005
'salex_software@shaw.ca, alexrs@gmail.com
'***********************************************************************

Const REG_SZ = 1
Const REG_BINARY = 3
Const ERROR_SUCCESS = 0&

Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_NO_MORE_ITEMS = 259&
Const BUFFER_SIZE As Long = 255

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal HKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal HKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal HKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal HKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long


Function DelEnumValues(hMainKey&, strPath$) As Integer
On Error GoTo errh
    Dim HKey&, sName$, sData$, Ret&, RetData&, Cnt&, KeyValues() As String, i&
    ReDim KeyValues(0)
    If RegOpenKey(hMainKey&, strPath$, HKey) = 0 Then
    
        sName = Space(BUFFER_SIZE)
        sData = Space(BUFFER_SIZE)
        Ret = BUFFER_SIZE
        RetData = BUFFER_SIZE
        
        While RegEnumValue(HKey, Cnt, sName, Ret, 0, ByVal 0&, ByVal sData, RetData) <> ERROR_NO_MORE_ITEMS
            If Ret Then
                ReDim Preserve KeyValues(Cnt)
                KeyValues(UBound(KeyValues)) = Trim(Left$(sName, Ret))
            End If
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            sData = Space(BUFFER_SIZE)
            Ret = BUFFER_SIZE
            RetData = BUFFER_SIZE
        Wend
        
        RegCloseKey HKey
        DelEnumValues = True
        For i = 0 To UBound(KeyValues)
            DelSetting hMainKey&, strPath$, KeyValues(i)
        Next i
        DelEnumValues = UBound(KeyValues) + 1
    Else
errh:
    DelEnumValues = False
    End If
End Function


Function RegQueryStringValue(ByVal HKey&, ByVal strKeyName$) As String
    Dim rLng&, rKeyType, buffer$, rKeyLength&, lBuffer&
    rLng& = RegQueryValueEx(HKey, strKeyName, 0, rKeyType, ByVal 0, rKeyLength)
    If rLng& = ERROR_SUCCESS Then
        If rKeyType = REG_SZ Then
            buffer$ = String(rKeyLength, Chr$(0))
            rLng& = RegQueryValueEx(HKey, strKeyName, 0, 0, ByVal buffer$, rKeyLength)
            If rLng& = ERROR_SUCCESS Then
                RegQueryStringValue = Left$(buffer$, InStr(1, buffer$, Chr$(0)) - 1)
            End If
        ElseIf rKeyType = REG_BINARY Then
            rLng& = RegQueryValueEx(HKey, strKeyName, 0, 0, lBuffer, rKeyLength)
            If rLng& = ERROR_SUCCESS Then
                RegQueryStringValue = CStr(lBuffer)
            End If
        End If
    End If
End Function

Function GetString(HKey&, strPath$, strValue$)
    Dim rRes&
    RegOpenKey HKey, strPath, rRes
    GetString = RegQueryStringValue(rRes, strValue)
    RegCloseKey rRes
End Function

Sub SaveString(HKey&, strPath$, strValue$, strData$)
    Dim rRes&
    RegCreateKey HKey, strPath, rRes
    RegSetValueEx rRes, strValue, 0, REG_SZ, ByVal strData, Len(strData)
    RegCloseKey rRes
End Sub

Sub SaveStringLong(HKey&, strPath$, strValue$, strData$)
    Dim rRes&
    RegCreateKey HKey, strPath, rRes
    RegSetValueEx rRes, strValue, 0, REG_BINARY, CByte(strData), 4
    RegCloseKey rRes
End Sub

Sub DelSetting(HKey&, strPath$, strValue$)
    Dim rRes&
    RegCreateKey HKey, strPath, rRes
    RegDeleteValue rRes, strValue
    RegCloseKey rRes
End Sub

