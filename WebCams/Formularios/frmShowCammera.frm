VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmShowCammera 
   Caption         =   "Cámara"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   Icon            =   "frmShowCammera.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3540
   ScaleWidth      =   5295
   WindowState     =   2  'Maximized
   Begin BAUCommonControls.Directorio ctlPath 
      Left            =   210
      Top             =   1200
      _ExtentX        =   582
      _ExtentY        =   582
   End
   Begin MSComctlLib.ImageList imlTools 
      Left            =   2070
      Top             =   1500
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowCammera.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowCammera.frx":6B64
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowCammera.frx":CDFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowCammera.frx":D810
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowCammera.frx":13432
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowCammera.frx":18C24
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowCammera.frx":1B3D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowCammera.frx":1DB88
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowCammera.frx":1FE02
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmShowCammera.frx":206DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin bauWebCams.WebCam wbcCammera 
      Height          =   2985
      Left            =   210
      TabIndex        =   1
      Top             =   600
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   5265
   End
   Begin MSComctlLib.Toolbar tlbTools 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlTools"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Tiempo"
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "5"
                  Text            =   "5 segundos"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "10"
                  Text            =   "10 segundos"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "15"
                  Text            =   "15 segundos"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "30"
                  Text            =   "30 segundos"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "60"
                  Text            =   "1 minuto"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "300"
                  Text            =   "5 minutos"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Tag             =   "600"
                  Text            =   "10 minutos"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ajustar"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Detener"
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Web"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Enviar e-mail"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Información"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pantalla completa"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmShowCammera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Formulario que recoge la imagen de la WebCam
Option Explicit

Public objWebCam As clsWebCam
Public strFilePrefix As String

Private Enum eButtons
  ButtonTime = 1
  ButtonStretch = 2
  ButtonStop = 3
  ButtonSave = 4
  ButtonWeb = 6
  ButtonEMail = 7
  ButtonInformation = 8
  ButtonScreen = 10
End Enum

Private blnSave As Boolean
Private intIndexSaved As Integer
Private strPath As String

Public Sub init()
'--> Inicializa los datos de la webCam
Dim objDir As New clsDir

  'Crea el directorio
    objDir.makeDir App.Path & "\Images"
  'Pasa a la webCam el control de descarga
    wbcCammera.OptionsDownload dwfConnection.intConnection, dwfConnection.strServer, _
                               dwfConnection.intPort, dwfConnection.strUser, _
                               dwfConnection.strPassword
  'Pasa los datos al control de manejo de webCams
    wbcCammera.URL = objWebCam.URL
    wbcCammera.LocalFileName = App.Path & "\Images\" & strFilePrefix
    wbcCammera.TimeDownload = objWebCam.Interval
  'Carga la primera imagen
    wbcCammera.StartDownload
  'Inicializa las variables locales
    blnSave = False
    intIndexSaved = 1
End Sub

Private Sub initLanguage()
'--> Inicializa los títulos
  If frmMDIMain.colLanguage.Count > 0 Then
    With tlbTools
      .Buttons(ButtonTime).ToolTipText = frmMDIMain.colLanguage("K400").Caption
      .Buttons(ButtonStretch).ToolTipText = frmMDIMain.colLanguage("K401").Caption
      .Buttons(ButtonStop).ToolTipText = frmMDIMain.colLanguage("K402").Caption
      .Buttons(ButtonSave).ToolTipText = frmMDIMain.colLanguage("K403").Caption
      .Buttons(ButtonScreen).ToolTipText = frmMDIMain.colLanguage("K404").Caption
      .Buttons(ButtonWeb).ToolTipText = frmMDIMain.colLanguage("K409").Caption
      .Buttons(ButtonEMail).ToolTipText = frmMDIMain.colLanguage("K410").Caption
      .Buttons(ButtonInformation).ToolTipText = frmMDIMain.colLanguage("K411").Caption
      .Buttons(ButtonTime).ButtonMenus(1).Text = "5 " & frmMDIMain.colLanguage("K405").Caption
      .Buttons(ButtonTime).ButtonMenus(2).Text = "10 " & frmMDIMain.colLanguage("K405").Caption
      .Buttons(ButtonTime).ButtonMenus(3).Text = "15 " & frmMDIMain.colLanguage("K405").Caption
      .Buttons(ButtonTime).ButtonMenus(4).Text = "30 " & frmMDIMain.colLanguage("K405").Caption
      .Buttons(ButtonTime).ButtonMenus(5).Text = "1 " & frmMDIMain.colLanguage("K406").Caption
      .Buttons(ButtonTime).ButtonMenus(6).Text = "5 " & frmMDIMain.colLanguage("K406").Caption
      .Buttons(ButtonTime).ButtonMenus(7).Text = "10 " & frmMDIMain.colLanguage("K406").Caption
    End With
  End If
End Sub

Private Sub saveImage(ByVal strFileName As String)
'--> Graba la imagen de la webCam en el directorio indicado
Dim strTargetFileName As String
Dim objFile As New clsFiles

  On Error GoTo errorSave
  If blnSave Then
    If Right(strPath, 1) <> "\" Then
      strPath = strPath & "\"
    End If
    strTargetFileName = objFile.getFileNameWithoutExtension(strFileName) & "_" & intIndexSaved & _
                        "." & objFile.getFileExtension(strFileName)
    FileCopy strFileName, strPath & strTargetFileName
    intIndexSaved = intIndexSaved + 1
  End If
  Set objFile = Nothing
  Exit Sub
  
errorSave:
  frmMDIMain.brStatus.Message = "Error al grabar la imagen"
End Sub

Private Sub getSavePath()
'--> Obtiene el directorio de grabación de las imágenes de la cámara
  If strPath = "" Then
    strPath = ctlPath.BrowseForFolder(Me.hwnd, "Directorio de grabación")
    If strPath = "" Then
      blnSave = False
    End If
  End If
End Sub

Private Sub showScreenTotal()
'--> Muestra a pantalla completa
  frmMDIMain.showScreenTotal Me.Caption, objWebCam.URL, strFilePrefix, blnSave, _
                             strPath, objWebCam.Interval, intIndexSaved
End Sub

Private Sub showWeb()
'--> Muestra la web de la cámara
Dim objExplorer As New clsHyperlink

  'Muestra la Web
    If Trim(objWebCam.WebURL) = "" Then
      frmMDIMain.messageBox.showMessage frmMDIMain.colLanguage.Item("K407").Caption, "bauWebCams", MsgInformation
    Else
      With objExplorer
        'Establece las propiedades
          .URL = objWebCam.WebURL
          .ExplorerStatus = SW_SHOWNORMAL
        'Abre el editor de correo
          .OpenURL
      End With
    End If
  'Libera la memoria
    Set objExplorer = Nothing
End Sub

Private Sub showEMail()
'--> Abre el correo para ese destinatario
Dim objExplorer As New clsHyperlink

  'Abre la aplicación de correo
    If Trim(objWebCam.eMail) = "" Then
      frmMDIMain.messageBox.showMessage frmMDIMain.colLanguage.Item("K408").Caption, "bauWebCams", MsgInformation
    Else
      With objExplorer
        'Establece las propiedades
          .URL = objWebCam.eMail
          .ExplorerStatus = SW_SHOWNORMAL
        'Abre el editor de correo
          .Mail
      End With
    End If
  'Libera la memoria
    Set objExplorer = Nothing
End Sub

Private Sub showInformation()
'--> Muestra la pantalla de información
Dim frmInfo As New frmInformation

  'Muestra la pantalla
    Set frmInfo.objWebCam = objWebCam
    frmInfo.Show vbModal
  'Libera la memoria
    Set frmInfo = Nothing
End Sub

Private Sub Form_Load()
  initLanguage
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  With wbcCammera
    .Top = tlbTools.Height
    .Left = Me.ScaleLeft
    .Width = Me.ScaleWidth - .Left
    .Height = Me.ScaleHeight - .Top
  End With
End Sub

Private Sub tlbTools_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index
    Case ButtonTime
      wbcCammera.TimeDownload = 20
    Case ButtonStretch
      wbcCammera.Stretch = Not wbcCammera.Stretch
    Case ButtonStop
      If wbcCammera.Stopped Then
        wbcCammera.StartDownload
        tlbTools.Buttons(3).Image = 4
      Else
        wbcCammera.StopDownload
        tlbTools.Buttons(3).Image = 3
      End If
    Case ButtonSave
      wbcCammera.StopDownload
      blnSave = Not blnSave
      If blnSave = False Then
        'intIndexSaved = 1
      Else
        getSavePath
      End If
      wbcCammera.StartDownload
    Case ButtonWeb
      showWeb
    Case ButtonEMail
      showEMail
    Case ButtonInformation
      showInformation
    Case ButtonScreen
      showScreenTotal
  End Select
End Sub

Private Sub tlbTools_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  wbcCammera.TimeDownload = CInt(ButtonMenu.Tag)
End Sub

Private Sub wbcCammera_imageDownloaded(ByVal strFileName As String)
  saveImage strFileName
End Sub
