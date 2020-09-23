VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWebCam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Colección con los datos del proyecto
Option Explicit

Public Name As String 'Nombre de la web
Public Description As String 'Descripción de la web
Public URL As String 'URL de la imagen
Public Interval As Integer 'Intervalo (en segundos)
Public WebURL As String 'URL de la web
Public eMail As String 'EMail de la webCam
Public ICQ As String 'ICQ de la webCam
