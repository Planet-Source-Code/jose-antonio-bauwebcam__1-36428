Attribute VB_Name = "modGeneral"
'--> Módulo de uso general
Option Explicit

Public Const cnstStrRootRegistry As String = "Software\Bau\bauWebCams" 'Path inicial del registro
Public Const cnstStrQuotes As String = """" 'Constante con las comillas

Public Enum IconButtons 'Enumerado con los índices de la lista de imágenes de los botones del programa
  IconBtOk = 1
  IconBtOkOver
  IconBtOkClick
  IconBtCancel
  IconBtCancelOver
  IconBtCancelClick
  IconBtHelp
  IconBtHelpOver
  IconBtHelpClick
  IconBtAbout
  IconBtAboutOver
  IconBtAboutClick
  IconBtClose
  IconBtCloseOver
  IconBtCloseClick
End Enum

Public Type typeConnection
  intConnection As Integer 'Tipo de conexión
  strServer As String 'Servidor proxy
  intPort As Integer 'Puerto de proxy
  strUser As String 'Usuario de proxy
  strPassword As String 'Contraseña de usuario proxy
End Type

Public dwfConnection As typeConnection 'Parámetros de la conexión a Internet

Public Function getCData(ByVal strValue As String) As String
'--> Obtiene una construcción CDATA de XML
  getCData = "<![CDATA[" & strValue & "]]>"
End Function

Public Function getHeaderXML() As String
'--> Obtiene la cabecera de un fichero XML
    getHeaderXML = "<?xml version=" & cnstStrQuotes & "1.0" & cnstStrQuotes & _
                   " encoding=" & cnstStrQuotes & "ISO-8859-1" & cnstStrQuotes & _
                   " ?>" & vbCrLf
End Function

Public Function getHeaderXSL(ByVal strStyleSheet As String) As String
'--> Obtiene la cabecera que apunta a un fichero XSL
  getHeaderXSL = getHeaderXML()
  getHeaderXSL = getHeaderXSL & "<?xml:stylesheet" & _
                                " type=" & cnstStrQuotes & "text/xsl" & cnstStrQuotes & _
                                " href=" & cnstStrQuotes & strStyleSheet & cnstStrQuotes & _
                                "?>" & vbCrLf
End Function
