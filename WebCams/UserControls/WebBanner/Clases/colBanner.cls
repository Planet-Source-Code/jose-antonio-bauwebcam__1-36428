VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colBanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"clsBanner"
Attribute VB_Ext_KEY = "Member0" ,"clsBanner"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Colección con la infomración sobre los banners
Option Explicit

Private colBanners As Collection

Public Function Add(ByVal strImage As String, ByVal strURL As String, ByVal strToolTip As String, _
                    ByVal intEllapseTime As Integer, Optional ByVal strKey As String) As clsBanner
'--> Añade un objeto a la colección
Dim objNewMember As clsBanner

  'Crea el nuevo objeto
    Set objNewMember = New clsBanner
  'Establece las propiedades
    With objNewMember
      .Image = strImage
      .URL = strURL
      .ToolTip = strToolTip
      .EllapseTime = intEllapseTime
    End With
    If Len(strKey) = 0 Then
      colBanners.Add objNewMember
    Else
      colBanners.Add objNewMember, strKey
    End If
  'Devuelve el objeto creado y libera la memoria
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Function loadXMLBanners(ByVal strFileName As String, ByVal strCaption As String, ByVal strURLPublicity As String) As Boolean
'--> Carga las propiedades de los banners
Dim intIndex As Integer
Dim objXML As New MSXML.DOMDocument
Dim objXMLNode As MSXML.IXMLDOMNode, objXMLBanner As MSXML.IXMLDOMNode, objXMLParameters As MSXML.IXMLDOMNode
Dim strImage As String, strURL As String, strToolTip As String
Dim intEllapseTime As Integer

  On Error GoTo errorLoad
  'Supone que todo es correcto
    loadXMLBanners = True
  'Añade uno vacío
    If strCaption <> "" Then
      Add "", strURLPublicity, strCaption, 5
    End If
  'Abre el fichero XML
    objXML.Load strFileName
  'Recorre los elementos del fichero y los almacena en la colección
    For Each objXMLNode In objXML.childNodes
      If objXMLNode.nodeType = NODE_ELEMENT Then
        For Each objXMLBanner In objXMLNode.childNodes
          intEllapseTime = -1
          For Each objXMLParameters In objXMLBanner.childNodes
            If UCase(objXMLParameters.baseName) = "IMAGE" Then
              strImage = objXMLParameters.Text
            End If
            If UCase(objXMLParameters.baseName) = "URL" Then
              strURL = objXMLParameters.Text
            End If
            If UCase(objXMLParameters.baseName) = "TOOLTIP" Then
              strToolTip = objXMLParameters.Text
            End If
            If UCase(objXMLParameters.baseName) = "TIME" Then
              intEllapseTime = CInt(objXMLParameters.Text)
            End If
            If strImage <> "" And strURL <> "" And strToolTip <> "" And intEllapseTime <> -1 Then
              Add strImage, strURL, strToolTip, intEllapseTime
            End If
          Next objXMLParameters
          Set objXMLParameters = Nothing
        Next objXMLBanner
        Set objXMLBanner = Nothing
      End If
    Next objXMLNode
  'Libera el objeto XML
    Set objXML = Nothing
  'Sale de la función
    Exit Function
    
errorLoad:
  loadXMLBanners = False
End Function

Public Property Get Item(ByVal varIndexKey As Variant) As clsBanner
Attribute Item.VB_UserMemId = 0
'--> Devuelve el elemento X de la colección
  Set Item = colBanners(varIndexKey)
End Property

Public Property Get Count() As Long
'--> Obtiene el número de elementos de la colección
  Count = colBanners.Count
End Property

Public Sub Remove(varIndexKey As Variant)
'--> Elimina un elemento de la colección
  colBanners.Remove varIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'--> Permite las enumeraciones (For...Each)
  Set NewEnum = colBanners.[_NewEnum]
End Property

Private Sub Class_Initialize()
'--> Crea la colección al crear la clase
  Set colBanners = New Collection
End Sub

Private Sub Class_Terminate()
'--> Destruye la colección al destruir la clase
  Set colBanners = Nothing
End Sub
