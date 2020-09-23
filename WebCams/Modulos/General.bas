Attribute VB_Name = "General"
'--> Clase con las rutinas más utilizadas en los programas
Option Explicit

Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
'HWND hwnd;  /* handle of window requesting help */
'LPCSTR lpszHelpFile;  /* address of directory-path string */
'UINT fuCommand; /* type of help */
'DWORD dwData; /* additional data  */
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Const HELP_COMMAND = &H102&
Private Const HELP_CONTENTS = &H3&
Private Const HELP_CONTEXT = &H1
Private Const HELP_FORCEFILE = &H9&
Private Const HELP_HELPONHELP = &H4
Private Const HELP_INDEX = &H3
Private Const HELP_KEY = &H101
Private Const HELP_MULTIKEY = &H201&
Private Const HELP_PARTIALKEY = &H105&
Private Const HELP_QUIT = &H2
Private Const HELP_SETCONTENTS = &H5&
Private Const HELP_SETINDEX = &H5
Private Const HELP_SETWINPOS = &H203&

'fuCommand dwData  Action
'HELP_CONTEXT  An unsigned long integer containing the context number for the topic. Displays Help for a particular topic identified by a context number that has been defined in the [MAP] section of the .HPJ file.
'HELP_CONTENTS Ignored; applications should set to 0L. Displays the Help contents topic as defined by the Contents option in the [OPTIONS] section of the .HPJ file.
'HELP_SETCONTENTS  An unsigned long integer containing the context number for the topic theapplication wants to designate as the Contents topic.   Determines which Contents topic Help should display when a user presses the F1 key.
'HELP_CONTEXTPOPUP An unsigned long integer containing the context number for a topic. Displays in a pop-up window a particular Help topic identified by a context number that has been defined in the [MAP] section of the .HPJ file.
'HELP_KEY  A long pointer to a string that contains a keyword for the desired topic. Displays the topic found in the keyword list that matches the keyword passed in the dwData parameter if there is one exact match. If there is more than one match, displays the Search dialog box with the topics listed in the Go To list box.
'HELP_PARTIALKEY A long pointer to a string that contains a keyword for the desired topic. Displays the topic found in the keyword list that matches the keyword passed in the dwData parameter if there is one exact match. If there is more than one match, displays the Search dialog box with the topics found listed in the Go To list box. If there is no match, displays the Search dialog box. If you just want to bring up the Search dialog box without passing a keyword (the third result), you should use a long pointer to an empty string.
'HELP_MULTIKEY A long pointer to the MULTIKEYHELP structure, as defined inWINDOWS.H. This structure specifies the table footnote character and the keyword.  Displays the Help topic identified by a keyword in an alternate key word table.
'HELP_COMMAND  A long pointer to a string that contains a Help macro to be executed. Executes a Help macro.
'HELP_SETWINPOS  A long pointer to the HELPWININFO structure, as defined in WINDOWS.H. Thisstructure specifies the size and position of the primary Help window or a secondary window to be displayed.   Displays the Help window if it is minimized or in memory, and positions it according to the data passed.
'HELP_FORCEFILE  Ignored; applications should set to 0L. Ensures that WinHelp is displaying the correct Help file. If the correct Help file is currently displayed, there is no action. If the incorrect Help file is displayed, WinHelp opens the correct file.
'HELP_HELPONHELP Ignored; applications should set to 0L. Displays the Contents topic of the designated Using Help file.
'HELP_QUIT Ignored; applications should set to 0L. Informs the Help application that Help is no longer needed. If no other applications have asked for Help, Windows closes the Help application.

Public Enum etMensajeError 'Tipos de mensajes de error
  mErrInformacion = 0
  mErrInterrogacion
  mErrError
  mErrComunicacion
  mErrErrorBaseDatos
End Enum

Private TituloAplicacion As String

'Public Sub EstablecerBarraEstado(FormBarraEstado As Object)
''--> Establece la Barra de Estado
''--> Cosas de VB5: da errores al intentar establecer directamente FR_Principal.BR_Estado, por eso se le pasa directamente
''--> el formulario donde está la barra de estado y se busca ésta
'Dim Indice As Integer
'
'  If TypeOf FormBarraEstado Is Form Then
'    For Indice = 0 To FormBarraEstado.Controls.Count - 1
'      If TypeOf FormBarraEstado.Controls(Indice) Is CTL_BarraEstado Then
'        Set BarraEstado = FormBarraEstado.Controls(Indice)
'      End If
'    Next Indice
'  End If
'End Sub

Public Sub AbrirVentanaModal(Ventana As Object)
'--> Abre una ventana de forma modal, es interesante para poder poner después el form utilizado a Nothing.
  DoEvents
  Ventana.Show vbModal
  Set Ventana = Nothing
  DoEvents
End Sub

Public Function ComprobarNulosCadena(Cadena) As String
'--> Si la cadena es nula se devuelve un espacio.
  If IsNull(Cadena) Then
    ComprobarNulosCadena = " "
  Else
    ComprobarNulosCadena = Trim(Cadena)
  End If
End Function

Public Function getToken(ByVal Cadena As String, ByVal Separador As String) As String
'--> Dada una cadena devuelve los primeros caracteres antes de encontrar un carácter separador.
Dim Indice As Integer
Dim CadenaSalida As String
    
  CadenaSalida = ""
  Indice = 1
  While Mid(Cadena, Indice, 1) <> Separador And Indice <= Len(Cadena)
    CadenaSalida = CadenaSalida + Mid(Cadena, Indice, 1)
    Indice = Indice + 1
  Wend
  getToken = Trim(CadenaSalida)
End Function

Public Function getParameter(Cadena As String, ByVal Separador As String, _
                             Optional ByVal GuardarEspacios As Boolean = False) As String
'--> Dada una cadena recoge el token indicado y lo quita
Dim Indice As Integer
Dim CadenaSalida As String

  Cadena = IIf(GuardarEspacios, Cadena, Trim(Cadena))
  CadenaSalida = ""
  Indice = 1
  While Mid(Cadena, Indice, 1) <> Separador And Indice <= Len(Cadena)
    CadenaSalida = CadenaSalida + Mid(Cadena, Indice, 1)
    Indice = Indice + 1
  Wend
  Cadena = Mid(Cadena, Len(CadenaSalida) + 2, Len(Cadena))
  getParameter = IIf(GuardarEspacios, CadenaSalida, Trim(CadenaSalida))
End Function

Public Sub SelCampoEdicion(CampoEdicion As Object)
'--> Selecciona por completo un campo TextEdit, se utiliza para no tener que borrar el contenido
'--> sino que directamente aparezca completamente seleccionado.
  On Error GoTo ErrorSeleccion
  CampoEdicion.SetFocus
  If Len(CampoEdicion.Text) <> 0 And CampoEdicion.SelLength < Len(CampoEdicion.Text) Then
    CampoEdicion.SelStart = 0
    CampoEdicion.SelLength = Len(CampoEdicion.Text)
    If CampoEdicion.Enabled Then CampoEdicion.SetFocus
  End If
  Exit Sub
  
ErrorSeleccion:
  Exit Sub
End Sub

Public Function QuitarApostrofes(ByVal Texto As String) As String
'--> Quita las apóstrofes de un textbox, se utiliza para evitar que en una nombre de elemento _
     puedan aparecer apóstrofes que darían un error a la hora de buscar por el campos DSxxxx _
     por ejemplo: SELECT IDModelo FROM Modelos WHERE DSModelo='Nombre del modelo con ap'ostrofe' _
     Dará un error porque se encuentra con dos cadenas separadas por un único '
Dim Cadena As String, Caracter As String * 1
Dim Indice As Integer

  Texto = Trim(Texto)
  Cadena = ""
  For Indice = 1 To Len(Texto)
    Caracter = Mid(Texto, Indice, 1)
    If Caracter <> "'" Then Cadena = Cadena + Caracter
  Next Indice
  QuitarApostrofes = Cadena
End Function

Public Sub AyudaDe(ByVal FicheroAyuda As String, ByVal hWnd As Long, ByVal ContextId As Integer)
'--> Ayuda acerca de determinado valor
  If ContextId <> 0 Then WinHelp hWnd, FicheroAyuda, HELP_CONTEXT, ContextId
End Sub

Public Function CadenaEspaciada(ByVal PrimerCampo As String, ByVal LongitudPrimerCampo As Byte, SegundoCampo As String) As String
'--> Devuelva una cadena con dos valores, LogitudPrimerCampo indica el número de caracteres que tiene que tener el PrimerCampo.
'--> si PrimerCampo tiene menos de LongitudPrimerCampoCaracteres se añaden tantos espacioes como sea necesario.
  CadenaEspaciada = PrimerCampo + Space(LongitudPrimerCampo - Len(Trim(Right(PrimerCampo, LongitudPrimerCampo))) + 1) + SegundoCampo
End Function

Public Function CargarCadenaIni(ByVal NombreFichero As String, ByVal Tema As String, ByVal Seccion As String, _
                                Optional ByVal ValorDefecto = "@@", _
                                Optional ByVal TamanioCadena As Integer = 255) As String
'--> Recubre a la rutina de la API GetPrivateProfileString
Dim NumBytes As Integer
Dim CadSalida As String

  CadSalida = Space(TamanioCadena)
  NumBytes = GetPrivateProfileString(Tema, Seccion, ValorDefecto, CadSalida, TamanioCadena, NombreFichero)
  CadSalida = Trim(CadSalida)
  If CadSalida <> "" Then CadSalida = Left(CadSalida, Len(CadSalida) - 1)
  CargarCadenaIni = CadSalida
End Function

Public Sub GrabarCadenaIni(ByVal NombreFichero As String, ByVal Tema As String, ByVal Seccion As String, ByVal Valor As String)
'--> Recubre a la rutina de la API WritePrivateProfileString
Dim NumBytes As Integer

  NumBytes = WritePrivateProfileString(Tema, Seccion, Valor, NombreFichero)
End Sub

Public Sub CargarRegistroVentana(Formulario As Object, Optional ByVal AnchuraMinima As Integer = 5000, Optional ByVal AlturaMinima As Integer = 5500)
'--> Carga los parámetros de la última sesión.
  Screen.MousePointer = vbHourglass
  If TypeOf Formulario Is Form Then
    Formulario.WindowState = Val(CargarRegistroW95("Parametros", "EstadoVentana", Str(vbNormal)))
    If Formulario.WindowState = vbNormal Then
      Formulario.Top = Val(CargarRegistroW95("Parametros", "Top", "0"))
      Formulario.Left = Val(CargarRegistroW95("Parametros", "Left", "0"))
      Formulario.Width = Val(CargarRegistroW95("Parametros", "Width", Str(AnchuraMinima)))
      Formulario.Height = Val(CargarRegistroW95("Parametros", "Height", Str(AlturaMinima)))
    End If
  End If
  Screen.MousePointer = vbDefault
End Sub

Public Sub GrabarRegistroVentana(Formulario As Object)
'--> Graba los parámetros para la siguiente vez que el usuario entre en la aplicación.
  Screen.MousePointer = vbHourglass
  If TypeOf Formulario Is Form Then
    GrabarRegistroW95 "Parametros", "EstadoVentana", Str(Formulario.WindowState)
    If Formulario.WindowState = vbNormal Then
      GrabarRegistroW95 "Parametros", "Top", Str(Formulario.Top)
      GrabarRegistroW95 "Parametros", "Left", Str(Formulario.Left)
      GrabarRegistroW95 "Parametros", "Width", Str(Formulario.Width)
      GrabarRegistroW95 "Parametros", "Height", Str(Formulario.Height)
    End If
  End If
  Screen.MousePointer = vbDefault
End Sub

Public Function CargarRegistroW95(ByVal Tema As String, ByVal Capitulo As String, _
                                  Optional ByVal ValorDefecto As String = " ") As String
'--> Carga un parámetro del registro de W95
  CargarRegistroW95 = VBA.GetSetting(TituloAplicacion, Tema, Capitulo, ValorDefecto)
End Function

Public Sub GrabarRegistroW95(ByVal Tema As String, ByVal Capitulo As String, ByVal Cadena As String)
'--> Graba un parámetro en el registro de W95
  SaveSetting TituloAplicacion, Tema, Capitulo, Cadena
End Sub

Public Function ExisteValorEnLista(ByVal Lista As Object, ByVal Valor As String) As Boolean
'--> Comprueba si existe un valor en una lista
Dim Indice As Integer
Dim Encontrado As Boolean

  Indice = 0
  Encontrado = False
  'Dado que los elementos siempre se añaden sin espacios el valor debe estar también sin espacios
  Valor = Trim(Valor)
  While Indice < Lista.ListCount And Not Encontrado
    If Lista.List(Indice) = Valor Then Encontrado = True
    Indice = Indice + 1
  Wend
  ExisteValorEnLista = Encontrado
End Function

