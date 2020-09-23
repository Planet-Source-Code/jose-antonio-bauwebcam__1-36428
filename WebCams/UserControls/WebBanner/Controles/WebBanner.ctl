VERSION 5.00
Object = "*\A..\..\DownloadURL\DownloadFile.vbp"
Begin VB.UserControl WebBanner 
   Alignable       =   -1  'True
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10350
   HitBehavior     =   2  'Use Paint
   ScaleHeight     =   1485
   ScaleWidth      =   10350
   Begin DownloadFile.InternetFile dwnImage 
      Left            =   9270
      Top             =   270
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin webLoadBanner.AnimatedGif imgBanner 
      Height          =   885
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   1561
   End
   Begin VB.Timer tmrBanner 
      Left            =   8280
      Top             =   240
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inserte aquí su publicidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   0
      Left            =   150
      MousePointer    =   10  'Up Arrow
      TabIndex        =   0
      Top             =   30
      Width           =   4515
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inserte aquí su publicidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   180
      TabIndex        =   1
      Top             =   60
      Width           =   4515
   End
End
Attribute VB_Name = "WebBanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--> Control que permite recoger imágenes de Internet
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private strHost As String
Private intPort As Integer
Private strServerFileName As String, strLocalPath As String
Private strUser As String, strPassword As String
Private strCaption As String, strURLPublicity As String
Private colBanners As New colBanner
Private strActualURL As String
Private intActualBanner As Integer
Private dtmFinalTime As Date

Public Property Let Host(ByVal strNewHost As String)
'--> Cambia la URL de la Host del que debe cargar el fichero
  strHost = strNewHost
  PropertyChanged
End Property

Public Property Get Host() As String
'--> Obtiene la URL de la Host del que debe cargar el fichero
  Host = strHost
End Property

Public Property Let Port(ByVal intNewPort As Integer)
'--> Cambia el puerto del Host del que debe cargar el fichero
  intPort = intNewPort
  PropertyChanged
End Property

Public Property Get Port() As Integer
'--> Obtiene el puerto del Host del que debe cargar el fichero
  Port = intPort
End Property

Public Property Let ServerFile(ByVal strNewFileName As String)
'--> Cambia el nombre de fichero que contiene las imágenes
  strServerFileName = strNewFileName
  PropertyChanged
End Property

Public Property Get ServerFile() As String
'--> Obtiene el nombre de fichero que contiene las imágenes
  ServerFile = strServerFileName
End Property

Public Property Let LocalPath(ByVal strNewLocalPath As String)
'--> Cambia el directorio donde se guardan los datos
  strLocalPath = strNewLocalPath
  If Right(strLocalPath, 1) = "\" Then
    strLocalPath = Left(strLocalPath, Len(strLocalPath) - 1)
  End If
  PropertyChanged
End Property

Public Property Get LocalPath() As String
'--> Cambia el directorio donde se guardan los datos
  LocalPath = strLocalPath
End Property

Public Property Let User(ByVal strNewUser As String)
'--> Cambia el usuario del Host
  strUser = strNewUser
  PropertyChanged
End Property

Public Property Get User() As String
'--> Obtiene el usuario del Host
  User = strUser
End Property

Public Property Let Password(ByVal strNewPassword As String)
'--> Cambia la contraseña del usuario
  strPassword = strNewPassword
  PropertyChanged
End Property

Public Property Get Password() As String
'--> Obtiene la contraseña del usuario
  Password = strPassword
End Property

Public Property Let Caption(ByVal strNewCaption As String)
'--> Cambia el título de "Inserte su publicidad"
  strCaption = strNewCaption
  lblCaption(0).Caption = strNewCaption
  lblCaption(1).Caption = strNewCaption
  PropertyChanged
End Property

Public Property Get Caption() As String
'--> Obtiene el título de "Inserte su publicidad"
  Caption = strCaption
End Property

Public Property Let URLPublicity(ByVal strNewURLPublicity As String)
'--> Cambia la URL de "Inserte su publicidad"
  strURLPublicity = strNewURLPublicity
  PropertyChanged
End Property

Public Property Get URLPublicity() As String
'--> Obtiene la URL de "Inserte su publicidad"
  URLPublicity = strURLPublicity
End Property

Private Sub addError(ByVal lngNumber As Long, ByVal strSource As String, ByVal strDescription As String)
'--> Añade el error
End Sub

Private Function downloadFile(ByVal strServerFileName As String, ByVal strLocalFileName As String, ByVal blnAscii As Boolean) As Boolean
'--> Descarga un fichero del servidor
  On Error GoTo errorDownloadFile
    dwnImage.ConnectType = Preconfig
    dwnImage.URL = strServerFileName
    dwnImage.LocalFile = strLocalFileName
    dwnImage.StartDownload
  'Sale de la función
    Exit Function
    
errorDownloadFile:
  addError -1, "downloadFile", "Error en la descarga del fichero " & strServerFileName & " sobre " & _
               strLocalFileName & vbCrLf & Err.Description
End Function

Private Function downloadBanner() As Boolean
'--> Carga / muestra un banner
Dim objBanner As clsBanner
Dim intLastBanner As Integer

  intLastBanner = intActualBanner
  intActualBanner = intActualBanner + 1
  If intActualBanner > colBanners.Count Then
    intActualBanner = 1
  End If
  If intActualBanner <= colBanners.Count And intActualBanner <> intLastBanner Then
    'Recoge un banner
      Set objBanner = colBanners.Item(intActualBanner)
    'Lo descarga
      If objBanner.Image = "" Then
        tmrBanner.Enabled = False
        lblCaption(0).Caption = strCaption
        lblCaption(1).Caption = strCaption
        strActualURL = strURLPublicity
        lblCaption(0).Visible = True
        lblCaption(1).Visible = True
        lblCaption(0).Top = (Height - lblCaption(0).Height) / 2
        lblCaption(0).Left = (Width - lblCaption(0).Width) / 2
        lblCaption(1).Top = lblCaption(0).Top + 30
        lblCaption(1).Left = lblCaption(0).Left + 30
        imgBanner.StopGif
        imgBanner.Visible = False
        tmrBanner.Interval = 30000
        tmrBanner.Enabled = True
      Else
        If downloadFile(objBanner.Image, strLocalPath & "\banner.gif", False) Then
          'Lo muestra
            tmrBanner.Enabled = False
            tmrBanner.Interval = 30000
            imgBanner.StopGif
            imgBanner.GifPath = strLocalPath & "\banner.gif"
            imgBanner.StartGif
            dtmFinalTime = DateAdd("n", objBanner.EllapseTime, Now)
            strActualURL = objBanner.URL
            imgBanner.ToolTipText = objBanner.ToolTip
            tmrBanner.Enabled = True
          'Quita los label
            lblCaption(0).Visible = False
            lblCaption(1).Visible = False
            imgBanner.Visible = True
        End If
      End If
  End If
End Function

Public Function readBanners() As Boolean
'--> Lee el fichero de banners de la Web, carga la imagen adecuada ...
  On Error GoTo errorReadBanners
  'Supone que todo es correcto
    readBanners = True
  'Descarga el fichero y lee los banners
    If downloadFile(strServerFileName, strLocalPath & "\xmlBanner.xml", True) Then
      If colBanners.loadXMLBanners(strLocalPath & "\xmlBanner.xml", strCaption, strURLPublicity) Then
        intActualBanner = 0
        downloadBanner
        readBanners = True
      Else
        addError -1, "readBanners", "Error al leer el fichero de descripción de banners"
        readBanners = False
      End If
    Else
      addError -1, "readBanners", "Error al descargar el fichero de descripción de banners"
      readBanners = False
    End If
  'Sale de la función
    Exit Function
  
errorReadBanners:
  addError -1, "readBanners", Err.Description
  readBanners = False
End Function

Private Sub showURL()
  ShellExecute 0&, vbNullString, strActualURL, vbNullString, "C:\", 1
End Sub

Private Sub imgBanner_onClick()
  showURL
End Sub

Private Sub lblCaption_Click(Index As Integer)
  showURL
End Sub

Private Sub tmrBanner_Timer()
  If Now > dtmFinalTime Then
    downloadBanner
  End If
End Sub

Private Sub UserControl_InitProperties()
  Host = ""
  Port = 21
  ServerFile = ""
  LocalPath = ""
  User = ""
  Password = ""
  tmrBanner.Enabled = False
  intActualBanner = -1
  strCaption = ""
  strURLPublicity = ""
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Host = PropBag.ReadProperty("Host", "")
  Port = PropBag.ReadProperty("Port", 21)
  ServerFile = PropBag.ReadProperty("ServerFile", "")
  LocalPath = PropBag.ReadProperty("LocalPath", "")
  User = PropBag.ReadProperty("User", "")
  Password = PropBag.ReadProperty("Password", "")
  Caption = PropBag.ReadProperty("Caption", "Inserte aquí su publicidad")
  URLPublicity = PropBag.ReadProperty("URLPublicity", "www.galeon.com/bauconsultors")
End Sub

Private Sub UserControl_Resize()
  With imgBanner
    .Left = ScaleLeft
    .Top = ScaleTop
    .Width = ScaleWidth
    .Height = ScaleHeight
  End With
End Sub

Private Sub UserControl_Terminate()
  Set colBanners = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Host", Host, ""
  PropBag.WriteProperty "Port", Port, 21
  PropBag.WriteProperty "ServerFile", ServerFile, ""
  PropBag.WriteProperty "LocalPath", LocalPath, ""
  PropBag.WriteProperty "User", User, ""
  PropBag.WriteProperty "Password", Password, ""
  PropBag.WriteProperty "Caption", Caption, "Inserte aquí su publicidad"
  PropBag.WriteProperty "URLPublicity", URLPublicity, "www.galeon.com/bauconsultors"
End Sub
