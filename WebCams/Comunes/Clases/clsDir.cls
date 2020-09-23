VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase para manejo de directorios
Option Explicit

Private Const ATTR_DIRECTORY As Integer = 16

Public Function removeDir(ByVal strPath As String) As Boolean
'--> Elimina un directorio y todos sus subdirectorios
Dim strFileName As String

  On Error GoTo errorRemoveDir
  removeDir = True
  If Right(strPath, 1) <> "\" Then
    strPath = strPath & "\"
  End If
  removeFile strPath & "*.*"
  strFileName = Dir(strPath & "*.*", ATTR_DIRECTORY)
  While strFileName <> ""
    If strFileName <> "." And strFileName <> ".." Then
      If (GetAttr(strPath & strFileName) And ATTR_DIRECTORY) = ATTR_DIRECTORY Then
        'Es un directorio y lo eliminamos
          removeDir strPath & strFileName
          RmDir strPath & strFileName
          strFileName = Dir(strPath & "*.*", ATTR_DIRECTORY)
      Else
        strFileName = Dir
      End If
    Else
      strFileName = Dir
    End If
  Wend
  Exit Function
  
errorRemoveDir:
  Debug.Print Err.Description
  removeDir = False
End Function

Public Function removeFile(ByVal strFileName As String) As Boolean
'--> Elimina un fichero
  On Error Resume Next
  Kill strFileName
  removeFile = True
End Function

Public Function copyDir(ByVal strSourcePath As String, ByVal strTargetPath As String) As Boolean
'--> Copia un directorio en otro
Dim strDirReturn As String
Dim arrStrPath() As String
Dim intIndex As Integer

   On Error GoTo errorCopyDir
  Debug.Print strSourcePath
  Debug.Print strTargetPath
  'Crea el directorio destino
    makeDir strTargetPath
  'Pone la \ final
    If Right(strSourcePath, 1) <> "\" Then
      strSourcePath = strSourcePath & "\"
    End If
    strSourcePath = UCase(strSourcePath)
    If Right(strTargetPath, 1) <> "\" Then
      strTargetPath = strTargetPath & "\"
    End If
    strTargetPath = UCase(strTargetPath)
  'Recoge los nombres de los sub directorios
    ReDim arrStrPath(0)
    strDirReturn = Dir(strSourcePath & "*.*", ATTR_DIRECTORY)
    While strDirReturn <> ""
      If strDirReturn <> "." And strDirReturn <> ".." Then 'Se salta . y ..
        If (GetAttr(strSourcePath & strDirReturn) And ATTR_DIRECTORY) = ATTR_DIRECTORY Then
          'Añade el directorio al vector
            arrStrPath(UBound(arrStrPath)) = strSourcePath & strDirReturn
            ReDim Preserve arrStrPath(UBound(arrStrPath) + 1)
        End If
      End If
      strDirReturn = Dir
    Wend
  'Copia todos los ficheros
    strDirReturn = Dir(strSourcePath & "*.*", 7)
    While strDirReturn <> ""
      If (GetAttr(strSourcePath & strDirReturn) And ATTR_DIRECTORY) <> ATTR_DIRECTORY Then
        FileCopy strSourcePath & strDirReturn, strTargetPath & strDirReturn
      End If
      strDirReturn = Dir
    Wend
   'Y copia los subdirectorios
      For intIndex = 0 To UBound(arrStrPath) - 1
        strTargetPath = strTargetPath & Right(arrStrPath(intIndex), Len(arrStrPath(intIndex)) - Len(strSourcePath))
        copyDir arrStrPath(intIndex), strTargetPath
      Next intIndex
    'Sale de la función
      copyDir = True
      Exit Function
  
errorCopyDir:
  copyDir = False
End Function

Public Function makeDir(ByVal strTargetPath As String) As Boolean
'--> Crea un directorio
  On Error GoTo errorMakeDir
    MkDir strTargetPath
    makeDir = True
  Exit Function
  
errorMakeDir:
  makeDir = False
End Function
