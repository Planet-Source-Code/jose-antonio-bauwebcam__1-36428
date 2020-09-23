VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colWebCams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Colección de proyectos
Option Explicit

Private colWebCams As Collection 'Colección con objetos de tipo clsWebCam

Public Sub Clear()
'--> Elimina los proyectos de la colección
Dim intIndice As Integer

  For intIndice = colWebCams.Count To 1 Step -1
    colWebCams.Remove intIndice
  Next intIndice
End Sub

Public Function Add(ByVal strName As String, ByVal strDescription As String, _
                    ByVal strURL As String, ByVal strURLWeb As String, _
                    ByVal strEMail As String, ByVal strICQ As String, _
                    ByVal intInterval As Integer, Optional ByVal strKey As String = "") As clsWebCam
'--> Crea un nuevo webCam
Dim objNewWebCam As New clsWebCam

  'Crea el nuevo proyecto
    Set objNewWebCam = New clsWebCam
  'Guarda los datos del proyecto
    With objNewWebCam
      .Name = strName
      .Description = strDescription
      .URL = strURL
      .WebURL = strURLWeb
      .Interval = intInterval
      .eMail = strEMail
      .ICQ = strICQ
    End With
  'Añade el nuevo proyecto a la colección
    If Len(strKey) = 0 Then
      colWebCams.Add objNewWebCam
    Else
      colWebCams.Add objNewWebCam, strKey
    End If
  'Devuelve el proyecto
    Set Add = objNewWebCam
    Set objNewWebCam = Nothing
End Function

Public Property Get Item(ByVal Index As Variant) As clsWebCam
'--> Obtiene el formulario index de la colección
  Set Item = colWebCams(Index)
End Property

Public Property Get Count() As Long
'--> Obtiene el número de elementos en la colección (en este caso el número de formularios)
  Count = colWebCams.Count
End Property

Public Sub Remove(ByVal Index As Variant)
'--> Borra un elemento de la colección
  colWebCams.Remove Index
End Sub

Public Property Get NewEnum() As IUnknown
'--> Propiedad necesaria en VB para las estructuras del tipo for each ...
  Set NewEnum = colWebCams.[_NewEnum]
End Property

Private Sub Class_Initialize()
  Set colWebCams = New Collection
End Sub

Private Sub Class_Terminate()
  Clear
  Set colWebCams = Nothing
End Sub
