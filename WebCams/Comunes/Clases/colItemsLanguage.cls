VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colItemsLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsItemLanguage"
Attribute VB_Ext_KEY = "Member0" ,"clsItemLanguage"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Colección con los distintos textos dependientes del lenguage
Option Explicit

Private Const cnstStrFileName As String = "xmlLanguage.xml"

Private colItemLanguage As Collection
Public Language As String

Public Function Add(ByVal strID As String, ByVal strText As String) As clsItemLanguage
'--> Crea un nuevo objeto
Dim objNewMember As clsItemLanguage

  'Crea el nuevo objeto
    Set objNewMember = New clsItemLanguage
  'Establece las propiedades del objeto
    objNewMember.Key = strID
    objNewMember.Caption = strText
    colItemLanguage.Add objNewMember, "K" & strID
  'Devuelve el objeto creado y libera la memoria
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Sub Clear()
'--> Elimina los elementos de la colección
Dim intIndex As Integer

  For intIndex = colItemLanguage.Count To 1 Step -1
    colItemLanguage.Remove intIndex
  Next intIndex
End Sub

Public Function loadXMLLanguages(ByVal strPath As String, ByVal strLanguage As String, _
                                 Optional ByVal strFileName As String = "") As Boolean
'--> Carga las propiedades de los banners
Dim intIndex As Integer
Dim objXML As New MSXML.DOMDocument
Dim objXMLNode As MSXML.IXMLDOMNode, objXMLItem As MSXML.IXMLDOMNode, objXMLLanguage As MSXML.IXMLDOMNode
Dim strCaption As String, strID As String

  On Error GoTo errorLoad
  'Supone que todo es correcto
    loadXMLLanguages = True
  'Limpia la colección
    Clear
  'Abre el fichero XML
    While Right(strPath, 1) = "\"
      strPath = Left(strPath, Len(strPath) - 1)
    Wend
    If strFileName = "" Then
      objXML.load strPath & "\" & cnstStrFileName
    Else
      objXML.load strPath & "\" & strFileName
    End If
  'Guarda el lenguaje en la variable pública
    Language = strLanguage
  'Recorre los elementos del fichero y los almacena en la colección
    strLanguage = UCase(strLanguage)
    For Each objXMLNode In objXML.childNodes
      If objXMLNode.nodeType = NODE_ELEMENT Then
        For Each objXMLItem In objXMLNode.childNodes
          If UCase(objXMLItem.baseName) = "MENUITEM" Then
            'Obtiene el ID
              strID = objXMLItem.Attributes(0).Text
            'Obtiene el lenguage
              strCaption = ""
              For Each objXMLLanguage In objXMLItem.childNodes
                If strCaption = "" Or UCase(objXMLLanguage.Attributes(0).Text) = strLanguage Then
                  strCaption = objXMLLanguage.Text
                End If
              Next objXMLLanguage
              Set objXMLLanguage = Nothing
            'Añade el elemento a la colección
              If strID <> "" And strCaption <> "" Then
                Add strID, strCaption
              End If
          End If
        Next objXMLItem
        Set objXMLItem = Nothing
      End If
    Next objXMLNode
  'Comprueba si ha habido algún error de interpretación
    If objXML.parseError.errorCode <> 0 Then
      loadXMLLanguages = False
    End If
  'Libera el objeto XML
    Set objXML = Nothing
  'Sale de la función
    Exit Function
    
errorLoad:
  loadXMLLanguages = False
End Function

Public Property Get Item(vntIndexKey As Variant) As clsItemLanguage
Attribute Item.VB_UserMemId = 0
'--> Obtiene un elemento
  On Error Resume Next
  Set Item = colItemLanguage(vntIndexKey)
  If Err.Number <> 0 Then
    Set Item = New clsItemLanguage
    Item.Caption = "<Unknown>"
  End If
End Property

Public Property Get Count() As Long
'--> Obtiene el número de elementos
  Count = colItemLanguage.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
'--> Elimina un elemento de la colección
  colItemLanguage.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'--> Permite enumerar los elementos de esta colección mediante For...Each
  Set NewEnum = colItemLanguage.[_NewEnum]
End Property

Private Sub Class_Initialize()
'--> Crea la colección al inicializar la clase
  Set colItemLanguage = New Collection
End Sub

Private Sub Class_Terminate()
'--> Elimina la colección al terminar la clase
  Clear
  Set colItemLanguage = Nothing
End Sub
