VERSION 5.00
Object = "{8D5F548D-FB40-41D9-B8CC-493E031AE985}#3.0#0"; "BAUCommonControls.ocx"
Object = "{D8331088-AA21-4B99-ABD2-5394E2C2A3B4}#1.0#0"; "DownloadFile.ocx"
Begin VB.UserControl WebCam 
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   ScaleHeight     =   4245
   ScaleWidth      =   5250
   Begin BAUCommonControls.progressBar brTime 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   3960
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   503
      ShowPercent     =   -1  'True
      Gradiant        =   1
      BorderStyle     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
   Begin DownloadFile.InternetFile dwnImage 
      Left            =   120
      Top             =   1650
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Timer tmrDownload 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1680
      Top             =   3420
   End
   Begin bauWebCams.ScrollPicture imgWebCam 
      Height          =   3435
      Left            =   90
      TabIndex        =   1
      Top             =   60
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   6059
   End
   Begin VB.Label lblDownload 
      BackColor       =   &H8000000E&
      Caption         =   "Cargando imagen ..."
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   3630
      Visible         =   0   'False
      Width           =   4875
   End
End
Attribute VB_Name = "WebCam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--> Control de usuario que gestiona la descarga de imágenes de la Web
Option Explicit

Private strURL As String, strExtension As String
Private strLocalFileName As String

Public Event imageDownloaded(ByVal strFileName As String)

Private blnStarted As Boolean
'Private WithEvents dwnImage As DownloadFile.InternetFile

Public Property Let TimeDownload(ByVal intTimeDownload As Integer)
'--> Cambia el tiempo de descarga de las imágenes de la WebCam
  initProgressBar intTimeDownload
End Property

Public Property Let URL(ByVal pStrURL As String)
'--> Cambia la URL de donde se descarga la imagen (además recoge la extensión para el momento que realice la descarga)
  strURL = pStrURL
  strExtension = getFileExtension(strURL)
  If strURL <> "" And strLocalFileName <> "" Then
    tmrDownload.Enabled = True
  End If
End Property

Public Property Get URL() As String
'--> Obtiene la URL de donde se descarga la imagen
  URL = strURL
End Property

Public Property Let Stretch(ByVal blnStretch As Boolean)
'--> Indica si se debe redimensionar la imagen
  imgWebCam.Stretch = blnStretch
End Property

Public Property Get Stretch() As Boolean
'--> Obtiene si se redimensiona la imagen
  Stretch = imgWebCam.Stretch
End Property

Public Property Get Stopped() As Boolean
'--> Obtiene si se ha detenido la descarga
  Stopped = Not blnStarted
End Property

Public Property Let LocalFileName(ByVal pStrLocalFileName As String)
'--> Cambia el nombre de fichero donde se graba la imagen
  strLocalFileName = pStrLocalFileName & "." & strExtension
  If strURL <> "" And strLocalFileName <> "" Then
    tmrDownload.Enabled = True
  End If
End Property

Public Property Get LocalFileName() As String
'--> Obtiene el nombre de fichero donde se graba la imagen
  LocalFileName = strLocalFileName
End Property

Public Sub OptionsDownload(ByVal intType As DownloadFile.IFile_ConnectType, _
                           ByVal strServer As String, ByVal intPort As Integer, _
                           ByVal strUser As String, ByVal strPassword As String)
'--> Establece los parámetros del control de descarga
  With dwnImage
    .ConnectType = intType
    .ProxyServer = strServer
    .ProxyPort = intPort
    .ProxyUser = strUser
    .ProxyPassword = strPassword
  End With
End Sub

Private Function getFileExtension(ByVal strFileName As String) As String
'--> Obtiene la extensión del fichero
Dim intIndex As Integer

  getFileExtension = ""
  intIndex = Len(strFileName)
  While Mid(strFileName, intIndex, 1) <> "." And intIndex > 1
    intIndex = intIndex - 1
  Wend
  If intIndex > 0 Then
    getFileExtension = Mid(strFileName, intIndex + 1)
  End If
End Function

Private Sub initProgressBar(ByVal intTimeDownload As Integer)
'--> Inicializa la barra de progreso
  tmrDownload.Enabled = False
  brTime.Maximus = intTimeDownload
  brTime.Value = 0
  If strURL <> "" And strLocalFileName <> "" Then
    tmrDownload.Enabled = True
  End If
End Sub

Private Sub incrementProgress()
'--> Incrementa la barra de progreso
  If strURL <> "" And strLocalFileName <> "" And blnStarted Then
    If brTime.Value = brTime.Maximus Then
      downloadImage
      brTime.Value = 0
    Else
      brTime.Value = brTime.Value + 1
    End If
  End If
End Sub

Private Sub downloadImage()
'--> Descarga la imagen
Dim lngRetVal As Long

  If strURL <> "" And strLocalFileName <> "" And Parent.Visible Then
    On Error Resume Next
    brTime.Visible = False
    lblDownload.Visible = True
    lblDownload.ZOrder 0
    DoEvents
    dwnImage.URL = strURL
    dwnImage.LocalFile = strLocalFileName
    dwnImage.StartDownload
    'Carga la imagen en la pantalla
      imgWebCam.loadImage strLocalFileName
      If Err.Number = 0 Then
        RaiseEvent imageDownloaded(strLocalFileName)
      End If
    lblDownload.Visible = False
    brTime.Visible = True
  End If
End Sub

Public Sub StartDownload()
'--> Comienza la descarga
  downloadImage
  blnStarted = True
End Sub

Public Sub StopDownload()
'--> Para la descarga
  blnStarted = False
End Sub

Private Sub dwnImage_DownloadProgress(lBytesRead As Long)
  lblDownload.Caption = "Cargando imagen " & lBytesRead & " bytes leídos ..."
  DoEvents
End Sub

Private Sub tmrDownload_Timer()
  incrementProgress
End Sub

Private Sub UserControl_Initialize()
  initProgressBar 15
  blnStarted = False
End Sub

Private Sub UserControl_Resize()
  On Error Resume Next
  With imgWebCam
    .Top = ScaleTop
    .Left = ScaleLeft
    .Width = ScaleWidth - .Left
    .Height = ScaleHeight - .Top - brTime.Height
  End With
  With lblDownload
    .Top = brTime.Top
    .Left = brTime.Left
    .Width = brTime.Width
    .Height = brTime.Height
  End With
End Sub

Private Sub UserControl_Terminate()
  On Error Resume Next
  Kill strLocalFileName
End Sub
