VERSION 5.00
Begin VB.Form frmInformation 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Información"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BAUCommonControls.ImageButton cmdClose 
      Height          =   465
      Left            =   3150
      TabIndex        =   12
      Top             =   4290
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
   Begin VB.Label lblICQ 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   285
      Left            =   2250
      TabIndex        =   11
      Top             =   3900
      Width           =   5115
   End
   Begin VB.Label lblEMail 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   255
      Left            =   2250
      TabIndex        =   10
      Top             =   3570
      Width           =   5115
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "ICQ:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   4
      Left            =   510
      TabIndex        =   9
      Top             =   3900
      Width           =   1545
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "eMail:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   3
      Left            =   510
      TabIndex        =   8
      Top             =   3570
      Width           =   1545
   End
   Begin VB.Label lblURLWeb 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   495
      Left            =   2250
      TabIndex        =   7
      Top             =   2940
      Width           =   5115
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Página Principal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   2
      Left            =   510
      TabIndex        =   6
      Top             =   2940
      Width           =   1545
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      Height          =   495
      Left            =   2250
      TabIndex        =   5
      Top             =   2370
      Width           =   5115
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Imagen:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Index           =   1
      Left            =   510
      TabIndex        =   3
      Top             =   2370
      Width           =   1065
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H8000000D&
      Caption         =   "Web"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   465
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   1740
      Width           =   7545
   End
   Begin VB.Label lblDescription 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   1155
      Left            =   360
      TabIndex        =   1
      Top             =   540
      Width           =   7065
   End
   Begin VB.Label lblName 
      BackColor       =   &H8000000D&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   7545
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1965
      Left            =   360
      TabIndex        =   4
      Top             =   2250
      Width           =   7065
   End
End
Attribute VB_Name = "frmInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Ventana que muestra la información sobre la cámara
Option Explicit

Public objWebCam As clsWebCam

Private Sub initLanguage()
'--> Cambia los mensajes del cuadro de diálogo
  Me.Caption = frmMDIMain.colLanguage.Item("K700").Caption
  cmdClose.Caption = frmMDIMain.colLanguage.Item("K200").Caption
  lblCaption(0).Caption = frmMDIMain.colLanguage.Item("K705").Caption
  lblCaption(1).Caption = frmMDIMain.colLanguage.Item("K701").Caption
  lblCaption(2).Caption = frmMDIMain.colLanguage.Item("K702").Caption
  lblCaption(3).Caption = frmMDIMain.colLanguage.Item("K703").Caption
  lblCaption(4).Caption = frmMDIMain.colLanguage.Item("K704").Caption
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  initLanguage
  lblName.Caption = objWebCam.Name
  lblDescription.Caption = objWebCam.Description
  lblURL.Caption = objWebCam.URL
  lblURLWeb.Caption = objWebCam.WebURL
  lblEMail.Caption = objWebCam.eMail
  lblICQ.Caption = objWebCam.ICQ
  With cmdClose
    Set .Picture = frmMDIMain.imlButtons.ListImages(IconBtClose).Picture
    Set .PictureOver = frmMDIMain.imlButtons.ListImages(IconBtCloseOver).Picture
    Set .PictureDown = frmMDIMain.imlButtons.ListImages(IconBtCloseClick).Picture
  End With
End Sub
