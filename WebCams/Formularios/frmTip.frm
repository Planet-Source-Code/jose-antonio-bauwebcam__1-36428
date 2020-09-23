VERSION 5.00
Begin VB.Form frmTip 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Truco del día"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BAUCommonControls.messageBox messageBox 
      Left            =   270
      Top             =   870
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin BAUCommonControls.ImageButton cmdClose 
      Height          =   465
      Left            =   7200
      TabIndex        =   0
      Top             =   1163
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   820
      Caption         =   "&Cerrar"
      CaptionPosition =   4
      BackColorOver   =   -2147483633
      BackColorDown   =   -2147483633
      ForeColorOver   =   -2147483635
      ForeColorDown   =   128
      ImageTop        =   4
      ImageLeft       =   3
      LabelTop        =   8
      LabelLeft       =   30
      NbLine          =   0
      NbLineOver      =   1
      CaptionPictureSpace=   4
      ColorLineUpLeftOne=   -2147483628
      ColorLineUpLeftTwo=   -2147483626
      ColorLineDownRightOne=   -2147483627
      ColorLineDownRightTwo=   -2147483632
      BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin BAUCommonControls.ImageButton cmdNextTip 
      Height          =   465
      Left            =   7170
      TabIndex        =   1
      Top             =   623
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   820
      Caption         =   "&Siguiente"
      CaptionPosition =   4
      BackColorOver   =   -2147483633
      BackColorDown   =   -2147483633
      ForeColorOver   =   -2147483635
      ForeColorDown   =   128
      ImageTop        =   4
      ImageLeft       =   3
      LabelTop        =   8
      LabelLeft       =   30
      NbLine          =   0
      NbLineOver      =   1
      CaptionPictureSpace=   4
      ColorLineUpLeftOne=   -2147483628
      ColorLineUpLeftTwo=   -2147483626
      ColorLineDownRightOne=   -2147483627
      ColorLineDownRightTwo=   -2147483632
      BeginProperty FontOver {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDown {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   1020
      TabIndex        =   2
      Top             =   218
      Width           =   5985
   End
   Begin VB.Image imgTip 
      Height          =   480
      Left            =   180
      Picture         =   "frmTip.frx":0000
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Cuadro de diálogo que muestra los trucos del día
Option Explicit

Public strLanguage As String
Private colTips As New colItemsLanguage

Private Sub initLanguage()
'--> Cambia los mensajes del cuadro de diálogo
  Me.Caption = frmMDIMain.colLanguage.Item("K500").Caption
  cmdClose.Caption = frmMDIMain.colLanguage.Item("K200").Caption
  cmdNextTip.Caption = frmMDIMain.colLanguage.Item("K501").Caption
End Sub

Private Sub loadTips()
  On Error GoTo errorLoadTips
    colTips.loadXMLLanguages App.Path & "\lib", strLanguage, "Tips.xml"
  Exit Sub
  
errorLoadTips:
  messageBox.showMessage frmMDIMain.colLanguage.Item("K502").Caption, "bauWebCams", MsgExclamation '"Error al cargar el truco del día"
End Sub

Private Sub showNextTip()
'--> Muestra el siguiente truco
  lblMessage.Caption = colTips.Item("K" & (Int(colTips.Count * Rnd) + 1)).Caption
End Sub

Private Sub cmdClose_Click()
'--> Cierra la ventana
  Unload Me
End Sub

Private Sub cmdNextTip_Click()
'--> Pasa al siguiente truco
  showNextTip
End Sub

Private Sub Form_Load()
  initLanguage
  loadTips
  showNextTip
  With cmdClose
    Set .Picture = frmMDIMain.imlButtons.ListImages(IconBtClose).Picture
    Set .PictureOver = frmMDIMain.imlButtons.ListImages(IconBtCloseOver).Picture
    Set .PictureDown = frmMDIMain.imlButtons.ListImages(IconBtCloseClick).Picture
  End With
  With cmdNextTip
    Set .Picture = frmMDIMain.imlButtons.ListImages(IconBtHelp).Picture
    Set .PictureOver = frmMDIMain.imlButtons.ListImages(IconBtHelpOver).Picture
    Set .PictureDown = frmMDIMain.imlButtons.ListImages(IconBtHelpClick).Picture
  End With
End Sub


