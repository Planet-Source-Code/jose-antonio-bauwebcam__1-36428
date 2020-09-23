VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000004&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acerca de ..."
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BAUCommonControls.ImageButton cmdClose 
      Height          =   465
      Left            =   1905
      TabIndex        =   4
      Top             =   2220
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
   Begin VB.Label lblWeb 
      Alignment       =   1  'Right Justify
      Caption         =   "www.bauconsultors.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2970
      TabIndex        =   3
      Top             =   1860
      Width           =   2025
   End
   Begin VB.Label lblMailComment 
      Alignment       =   1  'Right Justify
      Caption         =   "Sugerencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   1590
      Width           =   1635
   End
   Begin VB.Image Image1 
      Height          =   2610
      Left            =   60
      Picture         =   "frmAbout.frx":0000
      Top             =   120
      Width           =   1020
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      Caption         =   "Versión 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1155
      TabIndex        =   1
      Top             =   1140
      Width           =   3675
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "bauWebCams"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   885
      Index           =   0
      Left            =   1185
      TabIndex        =   0
      Top             =   240
      Width           =   3675
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Cuadro de diálogo Acerca de ...
Option Explicit

Private Sub initLanguage()
'--> Cambia los mensajes del cuadro de diálogo
  Me.Caption = frmMDIMain.colLanguage.Item("K44").Caption
  lblMailComment.Caption = frmMDIMain.colLanguage.Item("K100").Caption
  cmdClose.Caption = frmMDIMain.colLanguage.Item("K200").Caption
End Sub

Private Sub cmdClose_Click()
'--> Cierra la ventana
  Unload Me
End Sub

Private Sub Form_Load()
  initLanguage
  With cmdClose
    Set .Picture = frmMDIMain.imlButtons.ListImages(IconBtClose).Picture
    Set .PictureOver = frmMDIMain.imlButtons.ListImages(IconBtCloseOver).Picture
    Set .PictureDown = frmMDIMain.imlButtons.ListImages(IconBtCloseClick).Picture
  End With
End Sub

Private Sub lblMailComment_Click()
'--> Abre una ventana de correo
Dim objExplorer As New clsHyperlink

  With objExplorer
    'Establece las propiedades
      .URL = "bau1970@hotmail.com"
      .ExplorerStatus = SW_SHOWNORMAL
    'Abre el editor de correo
      .Mail
  End With
  'Libera la memoria
    Set objExplorer = Nothing
End Sub

Private Sub lblWeb_Click()
'--> Abre una ventana del explorador
Dim objExplorer As New clsHyperlink

  With objExplorer
    'Establece las propiedades
      .URL = "http://www.galeon.com/bauconsultors/"
      .ExplorerStatus = SW_SHOWNORMAL
    'Abre el editor de correo
      .OpenURL
  End With
  'Libera la memoria
    Set objExplorer = Nothing
End Sub
