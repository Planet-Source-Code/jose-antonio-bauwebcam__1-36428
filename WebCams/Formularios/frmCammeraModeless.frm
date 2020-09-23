VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCammeraModeless 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Cámara"
   ClientHeight    =   6555
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
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
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCammeraModeless.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCammeraModeless.frx":629A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCammeraModeless.frx":C534
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCammeraModeless.frx":CF46
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCammeraModeless.frx":12B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCammeraModeless.frx":1835A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCammeraModeless.frx":1AB0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
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
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCammeraModeless.frx":1AE26
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCammeraModeless.frx":210C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCammeraModeless.frx":2735A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCammeraModeless.frx":27D6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCammeraModeless.frx":2D98E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCammeraModeless.frx":33180
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbTools 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8760
      _ExtentX        =   15452
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
         NumButtons      =   7
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
            Object.ToolTipText     =   "Pantalla completa"
            ImageIndex      =   6
            Style           =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siempre visible"
            ImageIndex      =   7
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin bauWebCams.WebCam wbcCammera 
      Height          =   2985
      Left            =   3330
      TabIndex        =   0
      Top             =   2190
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   5265
   End
   Begin BAUCommonControls.Directorio ctlPath 
      Left            =   0
      Top             =   0
      _ExtentX        =   582
      _ExtentY        =   582
   End
End
Attribute VB_Name = "frmCammeraModeless"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Formulario que recoge la imagen de la WebCam a pantalla completa
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_FLAGS = SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW Or SWP_NOACTIVATE

Private Enum eButtons
  ButtonTime = 1
  ButtonStretch = 2
  ButtonStop = 3
  ButtonSave = 4
  ButtonScreen = 6
  ButtonVisible = 7
End Enum

Public blnSave As Boolean, blnEverVisible As Boolean
Public intIndexSaved As Integer
Public strURL As String
Public strFilePrefix As String
Public intInterval As Integer
Public strPath As String

Private Sub init()
'--> Inicializa los datos de la webCam
Dim objDir As New clsDir

  'Recoloca el control de usuario
    Form_Resize
  'Crea el directorio
    objDir.makeDir App.Path & "\Images"
  'Pasa a la webCam el control de descarga
    wbcCammera.OptionsDownload dwfConnection.intConnection, dwfConnection.strServer, _
                               dwfConnection.intPort, dwfConnection.strUser, _
                               dwfConnection.strPassword
  'Pasa los datos al control de manejo de webCams
    wbcCammera.URL = strURL
    wbcCammera.LocalFileName = App.Path & "\Images\" & strFilePrefix
    wbcCammera.TimeDownload = intInterval
  'Carga la primera imagen
    wbcCammera.StartDownload
  'Inicializa el idioma de la pantalla
    initLanguage
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

Private Sub getSavePath()
'--> Obtiene el directorio de grabación de las imágenes de la cámara
  If strPath = "" Then
    strPath = ctlPath.BrowseForFolder(Me.hWnd, "Directorio de grabación")
    If strPath = "" Then
      blnSave = False
    End If
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

Private Sub showVisible()
'--> Muestra la ventana sobre todas las demás o no
  blnEverVisible = Not blnEverVisible
  If blnEverVisible Then
    'Siempre visible
      SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
  Else
    'No siempre visible
      SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
  End If
End Sub

Private Sub Form_Load()
  'Inicializa los valores
    init
  'Muestra siempre visible
    blnEverVisible = False
    showVisible
End Sub

Private Sub Form_Resize()
  With wbcCammera
    .Top = tlbTools.Height
    .Left = Me.ScaleLeft
    .Width = Me.ScaleWidth - .Left
    .Height = Me.ScaleHeight - .Top
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMDIMain.Visible = True
End Sub

Private Sub tlbTools_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index
    Case ButtonTime
      wbcCammera.TimeDownload = 20
    Case ButtonStretch
      wbcCammera.Stretch = Not wbcCammera.Stretch
      Me.Height = Me.Height + 10
      Form_Resize
      Me.Height = Me.Height - 10
      Form_Resize
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
      If blnSave Then
        getSavePath
      End If
      wbcCammera.StartDownload
    Case ButtonScreen
      Unload Me
    Case ButtonVisible
      showVisible
  End Select
End Sub

Private Sub tlbTools_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  wbcCammera.TimeDownload = CInt(ButtonMenu.Tag)
End Sub

Private Sub wbcCammera_imageDownloaded(ByVal strFileName As String)
  saveImage strFileName
End Sub
