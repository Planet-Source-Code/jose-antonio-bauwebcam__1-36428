VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase para control de FTP
Option Explicit

'Constantes para uso de las opciones de acceso a Internet almacenadas en el registro
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3
Private Const INTERNET_INVALID_PORT_NUMBER = 0

'Constantes con el tipo de tranferencia
Private Const FTP_TRANSFER_TYPE_UNKNOWN = &H0
Private Const FTP_TRANSFER_TYPE_ASCII = &H1
Private Const FTP_TRANSFER_TYPE_BINARY = &H2

Private Const INTERNET_FLAG_PASSIVE = &H8000000

Private Const MAX_PATH As Integer = 260 'Número máximo de caracteres en una ruta
Private Const ERROR_NO_MORE_FILES As Integer = 18

Private Const ERROR_INTERNET_EXTENDED_ERROR = 12003

'Tipo de acceso
Private Const INTERNET_SERVICE_FTP = 1 'Acceso mediante FTP
Private Const INTERNET_SERVICE_GOPHER = 2 'Acceso Gopher
Private Const INTERNET_SERVICE_HTTP = 3 'Acceso HTTP

Private Const INTERNET_FLAG_RELOAD = &H80000000
Private Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Private Const INTERNET_FLAG_MULTIPART = &H200000

'Mensajes de error
Private Const ERRCHANGEDIRSTR As String = "No es posible pasar al directorio %s. Puede que no exista o esté protegido"
Private Const ERRCONNECTERROR As String = "No es posible conectar a %s utilizando los parámetros de usuario y contraseña"
Private Const ERRNOCONNECTION As String = "No existe conexión con un sitio FTP"
Private Const ERRNODOWNLOAD As String = "No es posible descargar el fichero %s del servidor"
Private Const ERRNOUPLOAD As String = "No es posible subir el fichero %s al servidor"
Private Const ERRNORENAME As String = "No es posible cambiar el nombre al fichero %s del servidor"
Private Const ERRNODELETE As String = "No es posible borrar el fichero  %s del servidor"
Private Const ERRALREADYCONNECTED As String = "No es posible cambiar esta propiedad mientras está conectado al servidor FTP"
Private Const ERRFATALERROR As String = "No es posible obtener conectar con WinInet.dll !"
Private Const ERRNOCOMMAND As String = "No es posible enviar un comando al servidor"

Private Const SESSION As String = "clsFTP Instance" 'Identificador de sesión para Windows

Private Type FILETIME 'Tipo con la hora de fichero
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA 'Tipo para la búsqueda de datos
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

'Declaraciones para utilizar las APIS de winInet
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long
Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal lpszSearchFile As String, lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long
Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" (ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hFtpSession As Long, ByVal lpszLocalFile As String, ByVal lpszRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
Private Declare Function FtpCommand Lib "wininet.dll" Alias "FtpCommandA" (ByVal hFtpSession As Long, ByVal fExpectedReponse As Boolean, ByVal dwFlags As Long, ByVal lpszCommand As String, ByVal dwContext As Long, ByVal phFtpCommand As Long) As Boolean
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean
Private Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal sFileName As String, ByVal lAccess As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, ByVal lpszFileName As String) As Boolean
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal sstrExistingName As String, ByVal sstrNewName As String) As Boolean
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

'Tipos definidos por el programa
Public Enum FileFTPTransferType 'Tipo de transferencia de ficheros
  ftAscii = FTP_TRANSFER_TYPE_ASCII
  ftBinary = FTP_TRANSFER_TYPE_BINARY
End Enum

Private lngHndINet As Long 'Handle de la conexión
Private lngHndConnection As Long 'Hande de la conexión FTP

'Propiedades FTP
Private strHostName As String 'Dirección del host
Private strHostPort As Integer 'Puerto del host
Private strUser As String 'Usuario
Private strPassword As String 'Contraseña
Private strPath As String 'Directorio
Private blnPassiveConnection As Boolean 'Conexión FTP pasiva
Public colErrors As New colFTPErrors 'Colección con los errores

Public Property Let Host(ByVal strNewHostName As String)
'--> Cambia el nombre del servidor. Si está ya conectado lanza un error
  If lngHndConnection <> 0 Then
    addError errInvalidProperty, "clsFTP::Host", ERRALREADYCONNECTED
  End If
  strHostName = strNewHostName
End Property

Public Property Get Host() As String
'--> Obtiene el nombre del servidor
  Host = strHostName
End Property

Public Property Let Port(ByVal strNewHostPort As Integer)
'--> Cambia el puerto del servidor. Si está ya conectado devuelve un error
  If lngHndConnection <> 0 Then
    addError errInvalidProperty, "clsFTP::Port", ERRALREADYCONNECTED
  End If
  strHostPort = strNewHostPort
End Property

Public Property Get Port() As Integer
'--> Obtiene el puerto del servidor
  Port = strHostPort
End Property

Public Property Let User(ByVal strNewUser As String)
'--> Establece el usuario. Si ya está conectado devuelve un error
  If lngHndConnection <> 0 Then
    addError errInvalidProperty, "clsFTP::User", ERRALREADYCONNECTED
  End If
  strUser = strNewUser
End Property

Public Property Get User() As String
'--> Obtiene el nombre del usuario
  User = strUser
End Property

Public Property Let Password(ByVal strNewPassword As String)
'--> Establece la contraseña. Si ya está conectado devuelve un error
  If lngHndConnection <> 0 Then
    addError errInvalidProperty, "clsFTP::Password", ERRALREADYCONNECTED
  End If
  strPassword = strNewPassword
End Property

Public Property Get Password() As String
'--> Obtiene la contraseña
  Password = strPassword
End Property

Public Property Let PassiveConnection(ByVal blnNewPassiveConnection As Boolean)
'--> Establece el nuevo valor de conexión Pasiva
  If lngHndConnection <> 0 Then
    addError errInvalidProperty, "clsFTP::PassiveConnection", ERRALREADYCONNECTED
  End If
  blnPassiveConnection = blnNewPassiveConnection
End Property

Public Property Get PassiveConnection() As Boolean
'--> Obtiene si la conexión es pasiva
  PassiveConnection = blnPassiveConnection
End Property

Public Property Get Directory() As String
'--> Obtiene el directorio actual
  Directory = strPath
End Property

Public Property Let Directory(ByVal strDirectory As String)
'--> Establece el directorio actual. Sólo si ya está conectado
Dim strError As String

  On Error GoTo vbErrorHandler
    If lngHndConnection <> 0 Then
        RemoteChDir strDirectory
        strPath = strDirectory
    Else
        On Error GoTo 0
        addError errNotConnectedToSite, "clsFTP::Directory", ERRNOCONNECTION
    End If
    Exit Property

vbErrorHandler:
  addError errNoDirChange, "clsFTP::Directory", Err.Description
End Property

Private Sub addError(ByVal errType As errFtpErrors, ByVal strSource As String, ByVal strDescription As String)
'--> Añade un error a la colección
  colErrors.Add errType, strSource, strDescription
End Sub

Public Property Get Connected() As Boolean
'--> Comprueba si está conectado a un servidor FTP
  Connected = (lngHndConnection <> 0)
End Property

Public Function Connect(Optional strNewHost As String, Optional intNewPort As Integer = -1, _
                        Optional strNewUser As String, _
                        Optional strNewPassword As String) As Boolean
'--> Se conecta al servidor FTP
Dim strError As String

  On Error GoTo vbErrorHandler
  'Supone que no puede conectar
    Connect = False
  'Si ya estábamos conectados se lanza un error
    If lngHndConnection <> 0 Then
      On Error GoTo 0
      addError errInvalidProperty, "clsFTP::Connect", "Ya está conectado al servidor FTP " & strHostName
      Exit Function
    End If
  'Si se nos han pasado los datos como parámetros sobreescribimos los anteriores
    If Len(strNewHost) > 0 Then
      Host = strNewHost
    End If
    If intNewPort > 0 Then
      Port = intNewPort
    End If
    If Len(strNewUser) > 0 Then
      User = strNewUser
    End If
    If Len(strNewPassword) > 0 Then
      Password = strNewPassword
    End If
  'Conecta con el servidor
    If Len(Host) = 0 Then
      addError errInvalidProperty, "clsFTP::Connect", "No se ha definido la dirección del Host"
    End If
    lngHndConnection = InternetConnect(lngHndINet, strHostName, strHostPort, _
                                        strUser, strPassword, INTERNET_SERVICE_FTP, _
                                        IIf(PassiveConnection, INTERNET_FLAG_PASSIVE, 0), 0)
  'Comprueba los errores de conexión
    If lngHndConnection = 0 Then
        strError = Replace(ERRCONNECTERROR, "%s", strHostName)
        On Error GoTo 0
        addError errCannotConnect, "clsFTP::Connect", strError & vbCrLf & GetINETErrorMsg(Err.LastDllError)
    End If
  'Indica que ha realizado la conexión correctamente
    Connect = True
  'Sale de la función
    Exit Function

vbErrorHandler:
  addError Err.Number, "clsFTP::Connect", Err.Description
End Function

Public Function Disconnect() As Boolean
'--> Desconecta el ordenador del servidor FTP
  If lngHndConnection <> 0 Then
    InternetCloseHandle lngHndConnection
    lngHndConnection = 0
  Else
    addError errNotConnectedToSite, "clsFTP::Disconnect", ERRNOCONNECTION
  End If
End Function

Public Function Download(ByVal strServerFileName As String, ByVal strLocalFileName As String, _
                         Optional ftpTransferType As FileFTPTransferType = ftAscii) As Boolean
'--> Descarga un fichero utilizando el tipo de transferencia definido
Dim strError As String

  On Error GoTo vbErrorHandler
  'Supone que no puede descargar el fichero
    Download = False
  'Si no está conectado lanza un error
    If lngHndConnection = 0 Then
      On Error GoTo 0
      addError errNotConnectedToSite, "clsFTP::Download", ERRNOCONNECTION
    End If
  'Descarga el fichero
    If FtpGetFile(lngHndConnection, strServerFileName, strLocalFileName, False, INTERNET_FLAG_RELOAD, ftpTransferType, 0) <> 1 Then
      strError = ERRNODOWNLOAD
      strError = Replace(strError, "%s", strServerFileName)
      On Error GoTo 0
      addError errGetFileError, "clsFTP::Download", strError
    End If
  'Indica que se ha descargado el fichero correctamente
    Download = True
    Exit Function

vbErrorHandler:
  addError errGetFileError, "clsFTP::Download", Err.Description
End Function

Public Function Upload(ByVal strLocalFileName As String, ByVal strServerFileName As String, _
                       Optional ftpTransferType As FileFTPTransferType) As Boolean
'--> Envía un fichero al servidor
Dim strError As String
Dim lngReturn As Long

  On Error GoTo vbErrorHandler
  'Indica que no puede enviar el fichero
    Upload = False
  'Si no está conectado lanza un error
    If lngHndConnection = 0 Then
      On Error GoTo 0
      addError errNotConnectedToSite, "clsFTP::Upload", ERRNOCONNECTION
    End If
  'Envía el fichero y comprueba los errores
    lngReturn = FtpPutFile(lngHndConnection, strLocalFileName, strServerFileName, ftpTransferType, 0)
    If lngReturn <> 1 Then
      strError = ERRNOUPLOAD
      strError = Replace(strError, "%s", strServerFileName)
      On Error GoTo 0
      strError = strError & vbCrLf & GetINETErrorMsg(Err.LastDllError)
      addError errCannotRename, "clsFTP::Upload", strError
    End If
  'Indica que ha enviado el fichero correctamente y sale de la función
    Upload = True
    Exit Function

vbErrorHandler:
  addError Err.Number, "clsFTP::Upload", Err.Description
End Function

Public Function SendCommand(ByVal strCommand As String) As Boolean
'--> Envía un comando al servidor
Dim strError As String
Dim lngResponse As Long
    
  On Error GoTo vbErrorHandler
  'Supone que no puede enviar el comando
    SendCommand = False
  'Si no está conectado, lanza un error
    If lngHndConnection = 0 Then
      On Error GoTo 0
      addError errNotConnectedToSite, "clsFTP::SendCommand", ERRNOCONNECTION
    End If
  'Envía el comando y comprueba el resultado
    If Not FtpCommand(lngHndConnection, False, FTP_TRANSFER_TYPE_BINARY, strCommand, 0, lngResponse) Then
      strError = ERRNOCOMMAND
      On Error GoTo 0
      strError = strError & vbCrLf & GetINETErrorMsg(Err.LastDllError)
      addError Err.Number, "clsFTP::SendCommand", strError
      Exit Function
    End If
  'Indica que se ha enviado correctamente el comando y sale de la función
    SendCommand = True
    Exit Function

vbErrorHandler:
  addError Err.Number, "clsFTP::SendCommand", Err.Description
End Function

Public Function RenameFile(ByVal strExistingName As String, ByVal strNewName As String) As Boolean
'--> Cambia el nombre a un fichero del servidor
Dim strError As String

  On Error GoTo vbErrorHandler
  'Supone que no puede cambiar el nombre al fichero
    RenameFile = False
  'Si no está conectado lanza un error
    If lngHndConnection = 0 Then
      On Error GoTo 0
      addError errNotConnectedToSite, "clsFTP::RenameFile", ERRNOCONNECTION
    End If
  'Intenta cambiar el nombre del fichero y comprueba los errores
    If Not FtpRenameFile(lngHndConnection, strExistingName, strNewName) Then
      strError = ERRNORENAME
      strError = Replace(strError, "%s", strExistingName)
      On Error GoTo 0
      strError = strError & vbCrLf & GetINETErrorMsg(Err.LastDllError)
      addError errCannotRename, "clsFTP::RenameFile", strError
    End If
  'SIndica que ha cambiado el nombre al fichero
    RenameFile = True
    Exit Function

vbErrorHandler:
  addError Err.Number, "clsFTP::RenameFile", Err.Description
End Function

Public Function DeleteFile(ByVal strExistingName As String) As Boolean
'--> Borra un fichero del servidor
Dim strError As String

  On Error GoTo vbErrorHandler
  'Supone que no puede borrar el fichero
    DeleteFile = False
  'Comprueba si está conectado
    If lngHndConnection = 0 Then
      On Error GoTo 0
      addError errNotConnectedToSite, "clsFTP::DeleteFile", ERRNOCONNECTION
    End If
  'Intenta borrar el fichero y comprueba los errores
    If Not FtpDeleteFile(lngHndConnection, strExistingName) Then
        strError = ERRNODELETE
        strError = Replace(strError, "%s", strExistingName)
        On Error GoTo 0
        addError errCannotDelete, "clsFTP::DeleteFile", strError
    End If
  'Indica que se pudo borrar el fichero
    DeleteFile = True
    Exit Function

vbErrorHandler:
  addError Err.Number, "clsFTP::DeleteFile", Err.Description
End Function

Private Function RemoteChDir(ByVal strDir As String) As Boolean
'--> Cambia el directorio remoto
Dim strPathFromRoot As String, strError As String

  On Error GoTo vbErrorHandler
  'Supone que no puede cambiar el directorio
    RemoteChDir = False
  'Convierte las \ del directorio al estándar de Unix (/)
    strDir = Replace(strDir, "\", "/")
  'Comprueba si existe la conexión
    If lngHndConnection = 0 Then
      On Error GoTo 0
      addError errNotConnectedToSite, "clsFTP::RemoteChDir", ERRNOCONNECTION
      Exit Function
    End If
  'Comprueba los cambios del directorio raíz
    If Len(strDir) <> 0 Then
        strPathFromRoot = strDir
        If Len(strDir) = 0 Then
            strDir = "/"
        End If
        Directory = strDir 'Cambia el directorio actual
        If Not FtpSetCurrentDirectory(lngHndConnection, strDir) Then
          strError = ERRCHANGEDIRSTR
          strError = Replace(strError, "%s", strDir)
          On Error GoTo 0
          addError errNoDirChange, "clsFTP::RemoteChDir", strError
        End If
    End If
  'Indica que ha podido cambiar el directorio
    RemoteChDir = True
    Exit Function

vbErrorHandler:
  addError Err.Number, "clsFTP::RemoteChDir", Err.Description
End Function

Private Function GetINETErrorMsg(ByVal lngErrNumber As Long) As String
'--> Obtiene el mensaje de error de INet
Dim lngLength As Long
Dim strBuffer As String

    If lngErrNumber = ERROR_INTERNET_EXTENDED_ERROR Then
      'Obtiene el tamaño del mensaje y el número
        InternetGetLastResponseInfo lngErrNumber, vbNullString, lngLength
        strBuffer = String(lngLength + 1, vbNullChar)
      'Obtiene el mensaje de error
        InternetGetLastResponseInfo lngErrNumber, strBuffer, lngLength
        GetINETErrorMsg = vbCrLf & strBuffer
    End If
End Function

Private Sub Class_Initialize()
'--> Crea la conexión con Internet
  lngHndINet = InternetOpen(SESSION, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
  If lngHndINet = 0 Then
    lngHndConnection = 0
    addError errFatal, "clsFTP::Initialize", ERRFATALERROR
  End If
  lngHndConnection = 0
  blnPassiveConnection = True
End Sub

Private Sub Class_Terminate()
'--> Elimina las conexiones
  'Elimina el handle de la conexión
    If lngHndConnection <> 0 Then
      InternetCloseHandle lngHndConnection
      lngHndConnection = 0
    End If
  'Elimina el handle de la API
    If lngHndINet <> 0 Then
      InternetCloseHandle lngHndINet
      lngHndINet = 0
    End If
  'Libera la colección de errores
    colErrors.Clear
    Set colErrors = Nothing
End Sub

