VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Opciones"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboConnection 
      Height          =   315
      Left            =   1725
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   120
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proxy"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1875
      Left            =   135
      TabIndex        =   3
      Top             =   570
      Width           =   5295
      Begin BAUCommonControls.Edicion txtServer 
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   556
         BackColorCaption=   -2147483643
         Caption         =   "<Dirección del servidor proxy>"
      End
      Begin BAUCommonControls.Edicion txtUser 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   630
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         BackColorCaption=   -2147483643
         Caption         =   "<Código de usuario>"
      End
      Begin BAUCommonControls.Edicion txtPassword 
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   10
         Top             =   1020
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         PasswordChar    =   "*"
         BackColorCaption=   -2147483643
         Caption         =   "<Contraseña>"
      End
      Begin BAUCommonControls.Edicion txtPassword 
         Height          =   315
         Index           =   1
         Left            =   1680
         TabIndex        =   13
         Top             =   1380
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         PasswordChar    =   "*"
         BackColorCaption=   -2147483643
         Caption         =   "<Repita la contraseña para verificación>"
      End
      Begin BAUCommonControls.Edicion txtServer 
         Height          =   315
         Index           =   1
         Left            =   4110
         TabIndex        =   14
         Top             =   240
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         BackColorCaption=   -2147483643
         Caption         =   "<Puerto>"
         MaxLength       =   3
      End
      Begin VB.Label lblCaption 
         Caption         =   "Con&traseña:"
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
         Left            =   150
         TabIndex        =   12
         Top             =   1440
         Width           =   1305
      End
      Begin VB.Label lblCaption 
         Caption         =   "Con&traseña:"
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
         Left            =   150
         TabIndex        =   11
         Top             =   1050
         Width           =   1305
      End
      Begin VB.Label lblCaption 
         Caption         =   "&Servidor:"
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
         Index           =   8
         Left            =   150
         TabIndex        =   7
         Top             =   300
         Width           =   1305
      End
      Begin VB.Label lblCaption 
         Caption         =   "&Usuario:"
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
         Index           =   7
         Left            =   150
         TabIndex        =   6
         Top             =   660
         Width           =   1305
      End
   End
   Begin BAUCommonControls.ImageButton cmdAccept 
      Height          =   435
      Left            =   405
      TabIndex        =   0
      Top             =   2610
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Caption         =   "&Aceptar"
      Picture         =   "frmOptions.frx":0000
      PictureOver     =   "frmOptions.frx":0457
      PictureDown     =   "frmOptions.frx":08CB
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
      Left            =   2055
      TabIndex        =   1
      Top             =   2610
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Caption         =   "&Cancelar"
      Picture         =   "frmOptions.frx":0D3F
      PictureOver     =   "frmOptions.frx":118C
      PictureDown     =   "frmOptions.frx":15EC
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
      Left            =   3630
      TabIndex        =   2
      Top             =   2610
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Caption         =   "A&yuda"
      Picture         =   "frmOptions.frx":16D8
      PictureOver     =   "frmOptions.frx":1B4E
      PictureDown     =   "frmOptions.frx":1FCC
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
   Begin VB.Label lblCaption 
      Caption         =   "Co&nexión:"
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
      Left            =   255
      TabIndex        =   9
      Top             =   150
      Width           =   1305
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Formulario para modificar las opciones: configuración internet ...
Option Explicit

Public intConnection As Integer 'Conexión
Public strServer As String 'Servidor proxy
Public intPort As Integer 'Puerto proxy
Public strUser As String 'Usuario
Public strPassword As String 'Contaseña
Public blnCancel As Boolean 'Indica si se han cancelado los cambios

Private Sub init()
'--> Inicializa los controles
  'Inicializa los lenguages
    initLanguage
  'Supone que se cancela el formulario
    blnCancel = True
  'Mete los datos en los controles
    cboConnection.ListIndex = intConnection
    txtServer(0).Text = strServer
    txtServer(1).Text = "" & intPort
    txtUser.Text = strUser
    txtPassword(0).Text = strPassword
    txtPassword(1).Text = strPassword
End Sub

Private Sub initLanguage()
'--> Cambia los mensajes del cuadro de diálogo
'  Me.Caption = frmMDIMain.colLanguage.Item("K700").Caption
  'Mete los datos en el combo
    With cboConnection
      .Clear
      .addItem "Configuración de Internet Explorer"
      .addItem "Conexión a través de Proxy"
      .addItem "Conexión directa"
    End With
End Sub

Private Sub acceptChanges()
'--> Comprueba los datos y descarga el formulario
  MsgBox "Comprobar los datos"
  'Recoge los datos
    intConnection = cboConnection.ListIndex
    strServer = txtServer(0).Text
    intPort = txtServer(1).Text
    strUser = txtUser.Text
    strPassword = txtPassword(0).Text
  'Indica que se han aceptado los cambios
    blnCancel = False
  'Cierra la ventana
    Unload Me
End Sub

Private Sub cmdAccept_Click()
  acceptChanges
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  init
End Sub
