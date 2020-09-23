VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHyperlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Hyperlink: manejo de hipervínculos y correo
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Constantes del API de Windows
Public Enum EnumExplorerState 'Modos de apertura del navegador
  SW_SHOWNORMAL = 1
  SW_SHOWMINIMIZED = 2
  SW_SHOWMAXIMIZED = 3
End Enum

'Local variables
Private mstrURL As String
Private mintState As EnumExplorerState

Property Let URL(ByVal strURL As String)
'--> Establece la URL
    mstrURL = strURL
End Property

Property Let ExplorerStatus(ByVal intState As EnumExplorerState)
'--> Indica el modo de apertura de la ventana del explorador
  mintState = SW_SHOWMAXIMIZED
End Property

Public Sub OpenURL()
  ShellExecute 0&, vbNullString, mstrURL, _
               vbNullString, "C:\", mintState
End Sub


Public Sub Mail()
  ShellExecute 0&, vbNullString, "mailto:" & mstrURL, vbNullString, _
               "C:\", mintState
End Sub
