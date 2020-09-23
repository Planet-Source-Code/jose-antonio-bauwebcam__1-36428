VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase para el manejo del registro de Windows
'--> Implementa la cuatro funciones básicas sobre el registro de Windows
'--> <BLOCKQUOTE>
'--> Añadir claves
'--> Cambiar el valor de una clave
'--> Borrar una clave
'--> Preguntar por un valor de clave
'--> </BLOCKQUOTE>
'--> Basado en un módulo de Kenneth Ives (kenaso@home.com)
Option Explicit

Private Enum etTypeValue ' Constantes con los valores retornados por  las claves
  REG_NONE = 0                      ' No value type
  REG_SZ = 1                        ' nul terminated string
  REG_EXPAND_SZ = 2                 ' nul terminated string w/enviornment var
  REG_BINARY = 3                    ' Free form binary
  REG_DWORD = 4                     ' 32-bit number
  REG_DWORD_LITTLE_ENDIAN = 4       ' 32-bit number (same as REG_DWORD)
  REG_DWORD_BIG_ENDIAN = 5          ' 32-bit number
  REG_LINK = 6                      ' Symbolic Link (unicode)
  REG_MULTI_SZ = 7                  ' Multiple Unicode strings
  REG_RESOURCE_LIST = 8             ' Resource list in the resource map
  REG_FULL_RESOURCE_DESCRIPTOR = 9  ' Resource list in the hardware description
  REG_RESOURCE_REQUIREMENTS_LIST = 10
End Enum

Private Enum etAccessRights 'Derechos específicos de acceso al registro
  KEY_QUERY_VALUE = &H1
  KEY_SET_VALUE = &H2
  KEY_CREATE_SUB_KEY = &H4
  KEY_ENUMERATE_SUB_KEYS = &H8
  KEY_NOTIFY = &H10
  KEY_CREATE_LINK = &H20
  KEY_ALL_ACCESS = &H3F
End Enum

Public Enum etKeyLocation 'Claves primarias en el registro
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_USERS = &H80000003
  HKEY_PERFORMANCE_DATA = &H80000004
  HKEY_CURRENT_CONFIG = &H80000005
  HKEY_DYN_DATA = &H80000006
End Enum

Private Enum etError 'Valores de retorno (Códigos de error)
  ERROR_SUCCESS = 0
  ERROR_ACCESS_DENIED = 5
  ERROR_NO_MORE_ITEMS = 259
End Enum

Private Enum etOption 'Constantes de apertura y creación
  REG_OPTION_NON_VOLATILE = 0
  REG_OPTION_VOLATILE = &H1
End Enum

' Declaraciones para el acceso al registro
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
  
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
  
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
  
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
  
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Function ExistKey(ByVal hKey As etKeyLocation, ByVal strRegKeyPath As String) As Boolean
'--> Comprueba si una clave existe
'--> @param hKey HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'--> @param strRegKeyPath Path de la clave a comprobar
'--> @return true si la clave existe, false en caso contrario
' Syntax:
'    sKeyQuery = ExistKey(HKEY_CURRENT_USER, "Software\AAA-Registry Test\Products")
Dim lngKeyHandle As Long, lngRet As Long

  lngKeyHandle = 0
  lngRet = RegOpenKey(hKey, strRegKeyPath, lngKeyHandle)
  ExistKey = (lngKeyHandle <> 0)
  lngRet = RegCloseKey(lngKeyHandle) 'Cierra el registro
End Function

Public Function QueryKey(ByVal hKey As etKeyLocation, ByVal strRegKeyPath As String, _
                         ByVal strRegSubKey As String) As Variant
'--> Devuelve el valor de una subclave
'--> @param hKey HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'--> @param strRegKeyPath path de la clave a buscar
'--> @param strRegSubKey nombre de la clave a buscar
'
' Syntax:
'    sKeyQuery = QueryKey(HKEY_CURRENT_USER, "Software\AAA-Registry Test\Products", "StringTestData")
'
' Returns the key value of "StringTestData"
' --------------------------------------------------------------
Dim intPosition As Integer
Dim lngKeyHandle As Long, lngRet As Long
Dim lngDataType As Long, lngBufferSize As Long, lngBuffer As Long
Dim strBuffer As String
  
  QueryKey = ""
  lngKeyHandle = 0
  lngBufferSize = 0
  lngRet = RegOpenKey(hKey, strRegKeyPath, lngKeyHandle)
  If lngKeyHandle <> 0 Then 'Comprueba que exista el handle
    'Busca en el registro y determina el tipo de los datos
    lngRet = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, ByVal 0&, lngBufferSize)
    If lngKeyHandle <> 0 Then 'Comprueba que exista el handle
      Select Case lngDataType ' Dependiendo del tipo de los datos llama a una función u otra
        Case REG_SZ       ' String data
          strBuffer = Space(lngBufferSize)
          lngRet = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, 0&, ByVal strBuffer, lngBufferSize)
          If lngRet = ERROR_SUCCESS Then
            intPosition = InStr(1, strBuffer, Chr(0))  ' look for the first null char
            If intPosition > 0 Then
              QueryKey = Left(strBuffer, intPosition - 1)
            Else
              QueryKey = strBuffer
            End If
          End If
        Case REG_DWORD    ' Numeric data (Integer)
          lngRet = RegQueryValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                                 lngBuffer, 4&)  ' 4& = 4-byte word (long integer)
          If lngRet = ERROR_SUCCESS Then QueryKey = lngBuffer
        Case Else    'Desconocido
          QueryKey = ""
      End Select
    End If
  End If
  lngRet = RegCloseKey(lngKeyHandle) 'Cierra el handle
End Function

Public Sub CreateKeyValue(ByVal hKey As etKeyLocation, ByVal strRegKeyPath As String, _
                          ByVal strRegSubKey As String, vRegData As Variant)
'--> Crea el valor de una subclave
'--> @param hKey HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'--> @param strRegKeyPath path de la clave a buscar
'--> @param strRegSubKey nombre de la clave a modificar
'--> @param vRegData Valor del dato
'
' Syntax:
'    CreateKeyValue HKEY_CURRENT_USER, _
'                      "Software\AAA-Registry Test\Products", _
'                      "StringTestData", "22 Jun 1999"
Dim lngKeyHandle As Long, lngRet As Long
Dim lngDataType As Long, lngKeyValue As Long
Dim strKeyValue As String
  
  lngDataType = IIf(IsNumeric(vRegData), REG_DWORD, REG_SZ)
  lngRet = RegCreateKey(hKey, strRegKeyPath, lngKeyHandle)
  Select Case lngDataType
    Case REG_SZ 'Cadena
       strKeyValue = Trim(vRegData) & Chr(0)     'La cadena debe terminar en \0
       lngRet = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, _
                            ByVal strKeyValue, Len(strKeyValue))
    Case REG_DWORD 'Valor numérico
       lngKeyValue = CLng(vRegData)
       lngRet = RegSetValueEx(lngKeyHandle, strRegSubKey, 0&, lngDataType, lngKeyValue, 4&) ' 4& = 4-byte word (long integer)
  End Select
  lngRet = RegCloseKey(lngKeyHandle) 'Cierra la clave
End Sub

Public Function CreateKey(ByVal hKey As etKeyLocation, ByVal strRegKeyPath As String)
'--> Crea una naueva clave
'--> @param hKey HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'--> @param strRegKeyPath  Nombre de la clave a crear
' Syntax:
'   CreateKey HKEY_CURRENT_USER, "Software\AAA-Registry Test"
'   CreateKey HKEY_CURRENT_USER, "Software\AAA-Registry Test\Products"
' --------------------------------------------------------------
Dim lngKeyHandle As Long, lngRet As Long
  
  lngRet = RegCreateKey(hKey, strRegKeyPath, lngKeyHandle) 'Crea la clave, si ya existía no hace nada
  lngRet = RegCloseKey(lngKeyHandle)
End Function

Public Function DeleteKey(ByVal hKey As etKeyLocation, ByVal strRegKeyPath As String, _
                          ByVal strRegKeyName As String)
'--> Borra una clave y todas sus subclaves
'--> @param hKey HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'--> @strRegKeyPath path de la clave
'--> @strRegKeyName nombre de la clave a borrar
' Syntax:
'    DeleteKey HKEY_CURRENT_USER, "Software", "AAA-Registry Test"
Dim lngKeyHandle As Long, lngRet As Long
  
  If ExistKey(hKey, strRegKeyPath) Then
    lngRet = RegOpenKey(hKey, strRegKeyPath, lngKeyHandle)
    lngRet = RegDeleteKey(lngKeyHandle, strRegKeyName)
    lngRet = RegCloseKey(lngKeyHandle) 'Cierra el registro
  End If
End Function

Public Function DeleteSubKey(ByVal hKey As etKeyLocation, ByVal strRegKeyPath As String, _
                             ByVal strRegSubKey As String)
'--> Borra una subclave
'--> @ param hKey - HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE, HKEY_USERS, etc
'--> @ param strRegKeyPath path de la clave principal donde borrar
'--> @ param strRegSubKey nombre de la clave a borrar
' Syntax:
'    DeleteSubKey HKEY_CURRENT_USER, "Software\AAA-Registry Test\Products", "StringTestData"
Dim lngKeyHandle As Long, lngRet As Long
  
  If ExistKey(hKey, strRegKeyPath) Then
    lngRet = RegOpenKey(hKey, strRegKeyPath, lngKeyHandle) 'Obtiene el handle a la clave
    lngRet = RegDeleteValue(lngKeyHandle, strRegSubKey) 'Borra la subclave, si no existe lo ignora
    lngRet = RegCloseKey(lngKeyHandle) 'Cierra el registro para evitar corromperlo
  End If
End Function
