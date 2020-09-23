VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFTPError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'--> Clase con los errores de FTP
Option Explicit

Public Enum errFtpErrors 'Errores de FTP
    errCannotConnect = vbObjectError + 2001
    errNoDirChange = vbObjectError + 2002
    errCannotRename = vbObjectError + 2003
    errCannotDelete = vbObjectError + 2004
    errNotConnectedToSite = vbObjectError + 2005
    errGetFileError = vbObjectError + 2006
    errInvalidProperty = vbObjectError + 2007
    errFatal = vbObjectError + 2008
End Enum

Public FTPError As errFtpErrors 'Código de error
Public Source As String 'Rutina que genera el error
Public Description As String 'Descripción del error


