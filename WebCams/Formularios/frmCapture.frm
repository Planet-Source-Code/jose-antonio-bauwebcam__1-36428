VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCapture 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capturar webcam"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4680
   ScaleWidth      =   4245
   Begin MSComctlLib.ImageList imlTools 
      Left            =   210
      Top             =   630
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
            Picture         =   "frmCapture.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCapture.frx":629A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCapture.frx":C534
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCapture.frx":CF46
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCapture.frx":12B68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCapture.frx":1835A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCapture.frx":1AB0C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picWebCamm 
      Height          =   3975
      Left            =   150
      ScaleHeight     =   3915
      ScaleWidth      =   3825
      TabIndex        =   1
      Top             =   540
      Width           =   3885
   End
   Begin VB.PictureBox picClipboard 
      Height          =   3945
      Left            =   120
      ScaleHeight     =   3885
      ScaleWidth      =   3885
      TabIndex        =   0
      Top             =   570
      Width           =   3945
   End
   Begin MSComctlLib.Toolbar tlbTools 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4245
      _ExtentX        =   7488
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
End
Attribute VB_Name = "frmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Formulario de captura de imágnes de la webCam
Option Explicit

Private Const WS_VISIBLE = &H10000000
Private Const WS_CHILD = &H40000000
Private Const WM_USER = 1024
Private Const WM_CAP_EDIT_COPY = WM_USER + 30
Private Const WM_CAP_DRIVER_CONNECT = WM_USER + 10
Private Const WM_CAP_SET_PREVIEW = WM_USER + 50
Private Const WM_CAP_SET_OVERLAY = WM_USER + 51
Private Const WM_CAP_SET_PREVIEWRATE = WM_USER + 52
Private Const WM_CAP_SEQUENCE = WM_USER + 62
Private Const WM_CAP_SINGLE_FRAME_OPEN = WM_USER + 70
Private Const WM_CAP_SINGLE_FRAME_CLOSE = WM_USER + 71
Private Const WM_CAP_SINGLE_FRAME = WM_USER + 72
Private Const DRV_USER = &H4000
Private Const DVM_DIALOG = DRV_USER + 100
Private Const PREVIEWRATE = 30

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
                            (ByVal hWnd As Long, ByVal wMsg As Long, _
                             ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" _
                            (ByVal a As String, ByVal b As Long, ByVal c As Integer, _
                             ByVal d As Integer, ByVal e As Integer, ByVal f As Integer, _
                             ByVal g As Long, ByVal h As Integer) As Long

Private hWndHandle As Long 'Handle a la ventana en que se realiza la captura

Private Sub init()
'--> Inicializa la pantalla
  On Error GoTo initError
    'Inicializa el handle de ventana
      hWndHandle = 0
    'Captura el vídeo en la ventana
      hWndHandle = capCreateCaptureWindow("CaptureWindow", WS_CHILD Or WS_VISIBLE, 0, 0, _
                                          picWebCamm.Width, picWebCamm.Height, picWebCamm.hWnd, 0)
  Exit Sub
  
initError:
  hWndHandle = 0
End Sub

Private Sub captureWebCamm()
'--> Captura la imagen de la webCam
  On Error GoTo errorCapture
    If hWndHandle <> 0 Then
      SendMessage hWndHandle, WM_CAP_DRIVER_CONNECT, 0, 0
      SendMessage hWndHandle, WM_CAP_SET_PREVIEW, 1, 0
      SendMessage hWndHandle, WM_CAP_SET_PREVIEWRATE, PREVIEWRATE, 0
      SendMessage Me.hWnd, WM_CAP_EDIT_COPY, 1, 0
      picClipboard.Picture = Clipboard.GetData '¿¿¿¿¿¿¿¿¿¿¿???????????????????
    End If
  Exit Sub
  
errorCapture:
End Sub

Private Sub Form_Load()
  init
End Sub
