VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colFTPErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsFTPError"
Attribute VB_Ext_KEY = "Member0" ,"clsFTPError"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Colección con los errores FTP
Option Explicit

Private colErrors As Collection 'Variable local con la colección

Public Function Add(ByVal errorFTP As errFtpErrors, ByVal strSource As String, _
                    ByVal strDescription As String, Optional strKey As String = "") As clsFTPError
'--> Añade un elemento a la colección de errores
Dim objNewMember As clsFTPError

  'Crea un nuevo objeto
    Set objNewMember = New clsFTPError
  'Establece las propiedades que se transfieren al método
    objNewMember.FTPError = errorFTP
    objNewMember.Source = strSource
    objNewMember.Description = strDescription
    If Len(strKey) = 0 Then
      colErrors.Add objNewMember
    Else
      colErrors.Add objNewMember, strKey
    End If
  'Devuelve el objeto creado y libera la memoria
    Set Add = objNewMember
    Set objNewMember = Nothing
  'Muestra un mensaje de depuración
    Debug.Print "Se ha añadido el siguiente error a la colección de errores FTP:"
    Debug.Print strDescription
End Function

Public Property Get Item(ByVal varIndexKey As Variant) As clsFTPError
'--> Obtiene un elemento de la colección
  Set Item = colErrors(varIndexKey)
End Property

Public Property Get Count() As Long
'Obtiene el número de elementos de la colección
  Count = colErrors.Count
End Property

Public Sub Remove(ByVal vntIndexKey As Variant)
'--> Elimina un elemento de la colección
    colErrors.Remove vntIndexKey
End Sub

Public Sub Clear()
'--> Elimina todos los elementos de la colección
Dim intIndex As Integer

  For intIndex = 1 To colErrors.Count
    colErrors.Remove intIndex
  Next intIndex
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'--> Permite enumerar esta colección con la sintaxis For...Each
  Set NewEnum = colErrors.[_NewEnum]
End Property

Private Sub Class_Initialize()
'--> Crea la colección cuando se crea la clase
  Set colErrors = New Collection
End Sub

Private Sub Class_Terminate()
'--> Destruye la colección cuando se termina de utilizar la clase
  Clear
  Set colErrors = Nothing
End Sub

