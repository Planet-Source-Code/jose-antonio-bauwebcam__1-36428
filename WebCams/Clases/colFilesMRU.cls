VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colFilesMRU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Colección con los últimos proyectos cargados
Option Explicit

Private arrStrFiles() As String 'array con los nombres de los ficheros

Public Sub Clear()
'--> Elimina los proyectos de la colección
  ReDim arrStrFiles(0)
End Sub

Private Function exist(ByVal strName As String) As Boolean
'--> Comprueba si existe el nombre en el array
Dim intIndex As Integer

  exist = False
  For intIndex = 0 To UBound(arrStrFiles) - 1
    If UCase(Trim(strName)) = UCase(Trim(arrStrFiles(intIndex))) Then
      exist = True
    End If
  Next intIndex
End Function

Public Function Add(ByVal strName As String) As Boolean
'--> Crea una nueva entrada en la tabla de últimos proyectos cargados
Dim intIndex As Integer

  If Not exist(strName) Then
    'Si se ha llegado al tope de cinco elementos se quita el último
      If UBound(arrStrFiles) = 5 Then
        For intIndex = 0 To 3
          arrStrFiles(intIndex) = arrStrFiles(intIndex + 1)
        Next intIndex
        ReDim Preserve arrStrFiles(UBound(arrStrFiles) - 1)
      End If
    'Se añade el elemento
      arrStrFiles(UBound(arrStrFiles)) = strName
      ReDim Preserve arrStrFiles(UBound(arrStrFiles) + 1)
  End If
End Function

Public Sub load()
'--> Carga los datos del registro
Dim objRegistry As New clsRegistry
Dim strValue As String
Dim intIndex As Integer

  'Carga los datos
    If objRegistry.ExistKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry) Then
      For intIndex = 1 To 5
        strValue = Trim(objRegistry.QueryKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "MRUProject" & intIndex))
        If strValue <> "" Then
          Add strValue
        End If
      Next intIndex
    End If
  'Libera la memoria
    Set objRegistry = Nothing
End Sub

Public Sub save()
'--> Graba los datos en el registro
Dim objRegistry As New clsRegistry
Dim intIndex As Integer
  
  'Recorre los nombres de los ficheros grabándolos sobre el registro
    For intIndex = 0 To UBound(arrStrFiles) - 1
      objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "MRUProject" & (intIndex + 1), arrStrFiles(intIndex)
    Next intIndex
  'Libera la memoria
    Set objRegistry = Nothing
End Sub

Public Property Get Item(ByVal intIndex As Integer) As String
'--> Obtiene el fichero intIndex de la colección
  Item = arrStrFiles(intIndex)
End Property

Public Property Get itemName(ByVal intIndex As Integer) As String
'--> Obtiene el nombre de fichero intIndex de la colección cortado
Dim objFile As New clsFiles
  
  itemName = Left(objFile.getPath(arrStrFiles(intIndex)), 10) & "\...\" & _
             objFile.getFileName(arrStrFiles(intIndex))
  Set objFile = Nothing
End Property

Public Property Get Count() As Long
'--> Obtiene el número de elementos en la colección (en este caso el número de formularios)
  Count = UBound(arrStrFiles)
End Property

Private Sub Class_Initialize()
  Clear
End Sub

Private Sub Class_Terminate()
  Clear
  Erase arrStrFiles
End Sub


