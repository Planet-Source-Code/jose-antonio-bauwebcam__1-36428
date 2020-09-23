VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmParameters 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Parámetros"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboInterval 
      Height          =   315
      Left            =   2580
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3900
      Width           =   3495
   End
   Begin BAUCommonControls.messageBox messageBox 
      Left            =   330
      Top             =   2820
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   360
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin BAUCommonControls.Edicion txtName 
      Height          =   315
      Left            =   2565
      TabIndex        =   1
      Top             =   60
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   556
      BackColorCaption=   -2147483643
      Caption         =   "<Nombre de la webCam>"
   End
   Begin BAUCommonControls.Edicion txtURL 
      Height          =   1065
      Left            =   2565
      TabIndex        =   5
      Top             =   1590
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   1879
      Multiline       =   -1  'True
      BackColorCaption=   -2147483643
      Caption         =   "<URL de la WebCam>"
   End
   Begin BAUCommonControls.ImageButton cmdAccept 
      Height          =   435
      Left            =   1110
      TabIndex        =   6
      Top             =   5250
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Caption         =   "&Aceptar"
      Picture         =   "frmParameters.frx":0000
      PictureOver     =   "frmParameters.frx":0457
      PictureDown     =   "frmParameters.frx":08CB
      CaptionPosition =   4
      BackColorOver   =   -2147483633
      BackColorDown   =   -2147483633
      ForeColorOver   =   -2147483635
      ForeColorDown   =   128
      ImageTop        =   3
      ImageLeft       =   8
      LabelTop        =   7
      LabelLeft       =   30
      NbLine          =   0
      NbLineDown      =   1
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
   Begin BAUCommonControls.ImageButton cmdCancel 
      Height          =   435
      Left            =   2760
      TabIndex        =   7
      Top             =   5280
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Caption         =   "&Cancelar"
      Picture         =   "frmParameters.frx":0D3F
      PictureOver     =   "frmParameters.frx":118C
      PictureDown     =   "frmParameters.frx":15EC
      CaptionPosition =   4
      BackColorOver   =   -2147483633
      BackColorDown   =   -2147483633
      ForeColorOver   =   -2147483635
      ForeColorDown   =   128
      ImageTop        =   3
      ImageLeft       =   8
      LabelTop        =   7
      LabelLeft       =   30
      NbLine          =   0
      NbLineDown      =   1
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
   Begin BAUCommonControls.ImageButton cmdHelp 
      Height          =   435
      Left            =   4410
      TabIndex        =   8
      Top             =   5280
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Caption         =   "A&yuda"
      Picture         =   "frmParameters.frx":16D8
      PictureOver     =   "frmParameters.frx":1B4E
      PictureDown     =   "frmParameters.frx":1FCC
      CaptionPosition =   4
      BackColorOver   =   -2147483633
      BackColorDown   =   -2147483633
      ForeColorOver   =   -2147483635
      ForeColorDown   =   128
      ImageTop        =   3
      ImageLeft       =   8
      LabelTop        =   7
      LabelLeft       =   30
      NbLine          =   0
      NbLineDown      =   1
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
   Begin BAUCommonControls.Edicion txtDescription 
      Height          =   1065
      Left            =   2550
      TabIndex        =   3
      Top             =   450
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   1879
      Multiline       =   -1  'True
      BackColorCaption=   -2147483643
      Caption         =   "<Descripción de la WebCam>"
   End
   Begin BAUCommonControls.Edicion txtURLWeb 
      Height          =   1065
      Left            =   2580
      TabIndex        =   9
      Top             =   2760
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   1879
      Multiline       =   -1  'True
      BackColorCaption=   -2147483643
      Caption         =   "<URL de la Web>"
   End
   Begin BAUCommonControls.Edicion txtEMail 
      Height          =   315
      Left            =   2565
      TabIndex        =   13
      Top             =   4290
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   556
      BackColorCaption=   -2147483643
      Caption         =   "<Dirección de correo electrónico>"
   End
   Begin BAUCommonControls.Edicion txtICQ 
      Height          =   315
      Left            =   2565
      TabIndex        =   15
      Top             =   4680
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   556
      BackColorCaption=   -2147483643
      Caption         =   "<Identificador ICQ>"
   End
   Begin VB.Label lblCaption 
      Caption         =   "IC&Q:"
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
      Index           =   6
      Left            =   1110
      TabIndex        =   16
      Top             =   4710
      Width           =   1305
   End
   Begin VB.Label lblCaption 
      Caption         =   "&eMail"
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
      Index           =   5
      Left            =   1110
      TabIndex        =   14
      Top             =   4320
      Width           =   1305
   End
   Begin VB.Label lblCaption 
      Caption         =   "&Intervalo:"
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
      Left            =   1110
      TabIndex        =   11
      Top             =   3930
      Width           =   1305
   End
   Begin VB.Label lblCaption 
      Caption         =   "&URL Web:"
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
      Left            =   1200
      TabIndex        =   10
      Top             =   2790
      Width           =   1305
   End
   Begin VB.Image Image1 
      Height          =   2610
      Left            =   60
      Picture         =   "frmParameters.frx":20B4
      Top             =   -60
      Width           =   1020
   End
   Begin VB.Label lblCaption 
      Caption         =   "&Descripción:"
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
      Left            =   1200
      TabIndex        =   2
      Top             =   420
      Width           =   1305
   End
   Begin VB.Label lblCaption 
      Caption         =   "&Nombre:"
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
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   90
      Width           =   1305
   End
   Begin VB.Label lblCaption 
      Caption         =   "&URL Imagen:"
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
      Left            =   1200
      TabIndex        =   4
      Top             =   1590
      Width           =   1305
   End
End
Attribute VB_Name = "frmParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Formulario que recoge los parámetros de un proyecto VBP
Option Explicit

Public blnCancel As Boolean 'Indica si se ha cancelado la selección de parámetros
Public strName As String 'Nombre del WebCam
Public strDescription As String 'Descripción de la webCam
Public strURL As String 'URL de la imagen WebCam
Public strURLWeb As String 'URL de la web
Public intInterval As Integer 'Intervalo (en segundos)
Public strEMail As String 'Dirección de eMail
Public strICQ As String 'Identificador de ICQ

Private Sub init()
'--> Inicializa la pantalla
Dim intIndex As Integer

  cboInterval.ListIndex = 0
  If strName <> "" Then
    txtName.Text = strName
    txtDescription.Text = strDescription
    txtURL.Text = strURL
    txtURLWeb.Text = strURLWeb
    txtEMail.Text = strEMail
    txtICQ.Text = strICQ
    For intIndex = 0 To cboInterval.ListIndex - 1
      If cboInterval.ItemData(intIndex) = intInterval Then
        cboInterval.ListIndex = intIndex
      End If
    Next intIndex
  End If
End Sub

Private Sub addInterval(ByVal intValue As Integer, ByVal strCaption As String)
'--> Añade un intervalo al combo
  cboInterval.addItem strCaption
  cboInterval.ItemData(cboInterval.ListCount - 1) = intValue
End Sub

Private Sub initLanguage()
'--> Inicializa los títulos
  If frmMDIMain.colLanguage.Count > 0 Then
    Me.Caption = frmMDIMain.colLanguage("K130").Caption
    lblCaption(0).Caption = frmMDIMain.colLanguage("K131").Caption
    txtName.Caption = frmMDIMain.colLanguage("K132").Caption
    lblCaption(1).Caption = frmMDIMain.colLanguage("K133").Caption
    txtDescription.Caption = frmMDIMain.colLanguage("K134").Caption
    lblCaption(2).Caption = frmMDIMain.colLanguage("K135").Caption
    txtURL.Caption = frmMDIMain.colLanguage("K136").Caption
    lblCaption(3).Caption = frmMDIMain.colLanguage("K140").Caption
    txtURLWeb.Caption = frmMDIMain.colLanguage("K141").Caption
    cmdAccept.Caption = frmMDIMain.colLanguage("K201").Caption
    cmdCancel.Caption = frmMDIMain.colLanguage("K202").Caption
    cmdHelp.Caption = frmMDIMain.colLanguage("K203").Caption
    lblCaption(4).Caption = frmMDIMain.colLanguage("K142").Caption
    lblCaption(5).Caption = frmMDIMain.colLanguage("K143").Caption
    txtEMail.Caption = frmMDIMain.colLanguage("K144").Caption
    lblCaption(6).Caption = frmMDIMain.colLanguage("K145").Caption
    txtICQ.Caption = frmMDIMain.colLanguage("K146").Caption
    'Añade los intervalos de tiempo al combo
      cboInterval.Clear
      addInterval 5, "5 " & frmMDIMain.colLanguage("K405").Caption
      addInterval 10, "10 " & frmMDIMain.colLanguage("K405").Caption
      addInterval 15, "15 " & frmMDIMain.colLanguage("K405").Caption
      addInterval 30, "30 " & frmMDIMain.colLanguage("K405").Caption
      addInterval 60, "1 " & frmMDIMain.colLanguage("K406").Caption
      addInterval 300, "5 " & frmMDIMain.colLanguage("K406").Caption
      addInterval 600, "10 " & frmMDIMain.colLanguage("K406").Caption
  End If
End Sub

Private Function getParameters() As Boolean
'--> Obtiene los parámetros a partir de los controles de texto
  getParameters = False
  If txtName.Text = "" Then
    messageBox.showMessage frmMDIMain.colLanguage("K137").Caption, "bauWebCam", MsgInformation
  ElseIf txtDescription.Text = "" Then
    messageBox.showMessage frmMDIMain.colLanguage("K138").Caption, "bauWebCam", MsgInformation
  ElseIf txtURL.Text = "" Then
    messageBox.showMessage frmMDIMain.colLanguage("K139").Caption, "bauWebCam", MsgInformation
  Else
    strName = txtName.Text
    strDescription = txtDescription.Text
    strURL = txtURL.Text
    strURLWeb = txtURLWeb.Text
    intInterval = cboInterval.ItemData(cboInterval.ListIndex)
    strEMail = txtEMail.Text
    strICQ = txtICQ.Text
    getParameters = True
  End If
End Function

Private Sub cmdAccept_Click()
  If getParameters() Then
    blnCancel = False
    Unload Me
  End If
End Sub

Private Sub cmdCancel_Click()
  blnCancel = True
  Unload Me
End Sub

Private Sub Form_Load()
  'Inicializa las variables
    blnCancel = True
    initLanguage 'Debe ir antes del init
    init
End Sub
