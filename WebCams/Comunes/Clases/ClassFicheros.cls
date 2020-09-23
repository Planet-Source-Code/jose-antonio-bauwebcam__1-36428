VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFiles"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase con rutinas útiles de control de ficheros
Option Explicit

Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

Public Function existFile(ByVal strFileName As String) As Boolean
'--> Comprueba si existe el fichero
Dim lngFile As Long

  On Error GoTo ErrorExistFile
  existFile = True
  lngFile = FileLen(strFileName)
  Exit Function
  
ErrorExistFile:
  existFile = False
End Function

Public Function getPath(ByVal strFileName As String) As String
'--> A partir de un nombre de fichero consigue el directorio al que pertenece
'--> @param strFileName Nombre completo del fichero (pe: C:\TEMP\Documento.doc)
'--> @return Directorio (pe: c:\TEMP)
Dim strPath As String

  strPath = strFileName
  While Right(strPath, 1) <> "\" And Len(strPath) > 1
    strPath = Left(strPath, Len(strPath) - 1)
  Wend
  getPath = strPath
End Function

Public Function getFileName(ByVal strFileName As String) As String
'--> Obtiene el nombre del fichero del path total introducido
'--> @param strFileName Nombre completo del fichero (pe: C:\TEMP\Documento.doc
'--> @return Nombre del fichero (pe: Documento.doc)
Dim intIndex As Integer
Dim blnFound As Boolean

  intIndex = Len(strFileName)
  blnFound = False
  While Not blnFound And intIndex > 0
    If Mid(strFileName, intIndex, 1) = "\" Then blnFound = True
    intIndex = intIndex - 1
  Wend
  If blnFound Then
    getFileName = Mid(strFileName, intIndex + 2, Len(strFileName))
  Else
    getFileName = strFileName
  End If
End Function

Public Sub deleteFile(ByVal strFileName As String)
'--> Borra un fichero
  On Error Resume Next
  Kill strFileName
End Sub

Public Function dlgGetFileName(ByRef dlgFiles As Object, _
                               Optional ByVal blnLoadSave As Boolean = True, _
                               Optional ByVal strFileName As String = "C:\", _
                               Optional ByVal strFilter As String = "Todos los ficheros (*.*) | *.*") As String
'--> Abre un diálogo para conseguir un nombre de fichero
  On Error GoTo ErrorApertura
  With dlgFiles
    .CancelError = True
    If strFileName <> "" Then
      .InitDir = getPath(strFileName)
      .FileName = getFileName(strFileName)
    Else
      .InitDir = "C:\"
      .FileName = ""
    End If
    .Filter = strFilter
    .FilterIndex = 0
    If blnLoadSave Then
      .Flags = cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNLongNames Or _
                           cdlOFNAllowMultiselect Or cdlOFNPathMustExist Or cdlOFNNoReadOnlyReturn
      .DialogTitle = "Abrir Archivos"
      .ShowOpen
    Else
      .Flags = cdlOFNExplorer Or cdlOFNLongNames Or _
                           cdlOFNAllowMultiselect Or cdlOFNPathMustExist
      .DialogTitle = "Guardar Archivo"
      .ShowSave
    End If
    dlgGetFileName = dlgFiles.FileName
  End With
  Exit Function
  
ErrorApertura:
  dlgGetFileName = ""
End Function

Public Function getFileNameWithoutExtension(ByVal strFileName As String)
'--> Establece el nombre del fichero (le quita el Path y la extensión)
Dim intIndex As Integer
Dim blnEncontrado As Boolean

  intIndex = Len(strFileName)
  blnEncontrado = False
  While Not blnEncontrado And intIndex > 0
    If Mid(strFileName, intIndex, 1) = "\" Then
      blnEncontrado = True
    Else
      intIndex = intIndex - 1
    End If
  Wend
  If blnEncontrado Then strFileName = Mid(strFileName, intIndex + 1, Len(strFileName))
  getFileNameWithoutExtension = getParameter(strFileName, ".")
End Function

Public Function getFileExtension(ByVal strFileName As String) As String
'--> Obtiene la extensión del fichero
Dim intIndex As Integer

  getFileExtension = ""
  intIndex = Len(strFileName)
  While Mid(strFileName, intIndex, 1) <> "." And intIndex > 1
    intIndex = intIndex - 1
  Wend
  If intIndex > 0 Then
    getFileExtension = Mid(strFileName, intIndex + 1)
  End If
End Function

Public Function getFullPath(ByVal strPath As String) As String
'--> Obtiene el nombre de directorio completo (c:\misdoc~1 --> c:\mis documentos)
Dim strBuffer As String
Dim lngReturn As Long


'Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" _
'    (ByVal lpszShortPath As String, ByVal lpszLongPath As String, _
'    ByVal cchBuffer As Long) As Long
'
'    Dim sBuf As String * 260
'    Dim i As Long
'    i = GetLongPathName(Text2, sBuf, Len(sBuf))
'    Text1 = Left$(sBuf, i)

  'Crea el buffer
    strBuffer = Space(255)
  'Copia el directorio actual al buffer y le añade 'a'
    lngReturn = GetFullPathName("a", 255, strBuffer, "")
  'Elimina los nullChar innecesarios
    getFullPath = Left(strBuffer, lngReturn)
End Function

Public Function getShortPath(ByVal strFileName As String) As String
'--> Obtiene el nombre corto del directorio
Dim strPath As String
Dim lngRes As Long

  'Create a buffer
    strPath = String$(165, 0)
  'retrieve the short pathname
    lngRes = GetShortPathName(strFileName, strPath, 164)
  'remove all unnecessary chr$(0)'s
    getShortPath = Left(strPath, lngRes)
End Function
