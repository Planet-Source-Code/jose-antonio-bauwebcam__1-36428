VERSION 5.00
Begin VB.UserControl Directorio 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   330
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Directorio.ctx":0000
   ScaleHeight     =   345
   ScaleWidth      =   330
   ToolboxBitmap   =   "Directorio.ctx":0102
End
Attribute VB_Name = "Directorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--> Control para especificar un directorio
'--> Utiliza las funciones de Api de Windows32 BrowseFolder
Option Explicit

Private Const WM_USER = &H400&
Private Const MAX_PATH = 260&
 
Private Type BrowseInfo 'Tipo para usar con SHBrowseForFolder
  hWndOwner As Long 'hWnd del formulario
  pIDLRoot As Long 'pID de la carpeta inicial
  pszDisplayName As String 'Nombre del item seleccionado
  lpszTitle As String 'Título a mostrar encima del árbol
  ulFlags As Long 'Flags con lo que de debe mostrar (tipo etBrowsing)
  lpfnCallback As Long 'Función CallBack que se llamará para cambiar el título, la selección, etc...
  lngParam As Long 'Información extra a pasar a la función Callback
  iImage As Long
End Type

Public Enum etBrowsing 'Tipo de ventana que aparecerá en el browser
  BIF_RETURNONLYFSDIRS = &H1& 'Encuentra una carpeta desde la seleccionada
  BIF_DONTGOBELOWDOMAIN = &H2& 'Para comenzar a encontrar un ordenador
  BIF_STATUSTEXT = &H4&
  BIF_RETURNFSANCESTORS = &H8&
  BIF_EDITBOX = &H10&
  BIF_VALIDATE = &H20& 'Insiste hasta encontrar un resultado válido (o CANCEL)
  BIF_BROWSEFORCOMPUTER = &H1000& 'Busca por ordenadores
  BIF_BROWSEFORPRINTER = &H2000& 'Busca por impresoras
  BIF_BROWSEINCLUDEFILES = &H4000& 'Busca por todo
End Enum

' Mensajes desde el browser
Private Const BFFM_INITIALIZED = 1 'Inicializado
Private Const BFFM_SELCHANGED = 2 'Selección cambiada
Private Const BFFM_VALIDATEFAILED = 3 'lngParam:szPath ret:1(cont),0(EndDialog)

' Mensajes al browser
Private Const BFFM_SETSTATUSTEXT = (WM_USER + 100)
Private Const BFFM_ENABLEOK = (WM_USER + 101)
Private Const BFFM_SETSELECTION = (WM_USER + 102)

Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpbi As BrowseInfo) As Long
Private Declare Sub CoTaskMemFree Lib "OLE32.DLL" (ByVal hMem As Long)
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lngParam As Any) As Long
Private Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
             
Private strBFFCaption As String 'Caption del árbol de directorios (para usar con la función CallBack)
Private strFolderIni As String 'Carpeta de inicio (para usar con la función CallBack)

Public Function BrowseFolderCallbackProc(ByVal hWndOwner As Long, ByVal uMSG As etBrowsing, _
                                         ByVal lngParam As Long, ByVal pData As Long) As Long
'--> Función CallBack para usar con la función BrowseForFolder
Dim szDir As String
    
  On Local Error Resume Next
  Select Case uMSG
    Case BFFM_INITIALIZED
        If Len(strFolderIni) Then 'Si hay un path de inicio comenzar por ahí
            szDir = strFolderIni & Chr$(0)
            Call SendMessage(hWndOwner, BFFM_SETSELECTION, 1&, ByVal szDir)
        End If
        ' Si se ha especificado el título de la ventana
        If Len(strBFFCaption) > 0 Then
            ' Cambiar el título de la ventana. de selección
            Call SetWindowText(hWndOwner, strBFFCaption)
        End If
    Case BFFM_SELCHANGED
        'Cuando se cambia el path cambia el de la ventana (en realidad no hace nada)
        szDir = String$(MAX_PATH, 0)
        If SHGetPathFromIDList(lngParam, szDir) Then
            Call SendMessage(hWndOwner, BFFM_SETSTATUSTEXT, 0&, ByVal szDir)
        End If
        Call CoTaskMemFree(lngParam)
  End Select
  Err = 0
  BrowseFolderCallbackProc = 0
End Function

Public Function rtnAddressOf(ByVal lngProc As Long) As Long
'--> Devuelve la dirección pasada como parámetro, se usará para asignar a una variable la dirección de una función, o procedimiento.
  rtnAddressOf = lngProc
End Function

Public Function BrowseForFolder(ByVal hWndOwner As Long, ByVal sPrompt As String, _
                                Optional ByVal sInitDir As String = "", _
                                Optional ByVal lFlags As Long = BIF_RETURNONLYFSDIRS, _
                                Optional ByVal sCaption As String = "") As String
'--> Muestra el diálogo de selección de directorios de Windows, si todo va bien, devuelve el directorio seleccionado, si se cancela, se devuelve una cadena vacía y se produce el error 32755
'--> @param hWndOwner El hWnd de la ventana
'--> @param sPrompt El título a mostrar encima del árbol
'--> @param sInitDir Opcionalmente el directorio de inicio
'--> @param lFlags se puede especificar lo que se podrá seleccionar BIF_BROWSEINCLUDEFILES, etc., por defecto es: BIF_RETURNONLYFSDIRS
'--> @param sCaption el Caption de la ventana
Dim iNull As Integer
Dim lpIDList As Long, lResult As Long
Dim sPath As String
Dim udtBI As BrowseInfo
    
    On Local Error Resume Next
    With udtBI
        .hWndOwner = hWndOwner
        .lpszTitle = sPrompt & vbNullChar
        .ulFlags = lFlags
        strBFFCaption = sCaption
        strFolderIni = sInitDir
        '.lpfnCallback = rtnAddressOf(AddressOf BrowseFolderCallbackProc) 'Indica la función CallBack, sólo es necesaria para cambiar el caption de la función
    End With
    Err = 0
    On Local Error GoTo 0
    lpIDList = SHBrowseForFolder(udtBI) 'Muestra la ventana de directorios
    If lpIDList Then 'Obtiene el directorio y le quita los \0 finales
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left$(sPath, iNull - 1)
    Else 'Si se ha cancelado devuelve un error y el directorio vacío
        sPath = ""
        With Err
            .Source = "MBrowseFolder::BrowseForFolder"
            .Number = 32755
            .Description = "Cancelada la operación de BrowseForFolder"
        End With
    End If
    BrowseForFolder = sPath
End Function

Private Sub UserControl_Resize()
  If Width <> 330 Then
    Width = 330
  ElseIf Height <> 330 Then
    Height = 330
  End If
End Sub
