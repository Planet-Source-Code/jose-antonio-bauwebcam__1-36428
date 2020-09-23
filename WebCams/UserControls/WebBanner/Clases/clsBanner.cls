VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase con las propiedades de los banners
Option Explicit

Public Image As String 'Nombre del fichero de la imagen en el servidor FTP
Public URL As String 'URL a la que se dirige el programa al pulsar sobre el banner
Public ToolTip As String 'Tooltip que se muestra al pasar el cursor sobre la imagen
Public EllapseTime As Long 'Tiempo que tarda en pasar a la siguiente imagen
