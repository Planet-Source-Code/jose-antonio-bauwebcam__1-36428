VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl statusBar 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10365
   ScaleHeight     =   405
   ScaleWidth      =   10365
   Begin BAUCommonControls.progressBar brProgreso 
      Height          =   255
      Left            =   1470
      TabIndex        =   2
      Top             =   90
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   450
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
   Begin MSComctlLib.StatusBar stBar 
      Height          =   255
      Left            =   6990
      TabIndex        =   0
      Top             =   60
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   450
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   706
            TextSave        =   "15:39"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "28/06/2002"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMessage 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preparado ..."
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
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.Image imgStatus 
      Height          =   255
      Index           =   0
      Left            =   8730
      Picture         =   "statusBar.ctx":0000
      Stretch         =   -1  'True
      Top             =   30
      Width           =   240
   End
   Begin VB.Image imgStatus 
      Height          =   255
      Index           =   1
      Left            =   9000
      Picture         =   "statusBar.ctx":0411
      Stretch         =   -1  'True
      Top             =   30
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "statusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--> Control de la barra de estado de la ventana principal
Option Explicit

Private Const cnstStrDefaultMessage As String = "Preparado ... "

Private Enum eIcons
  iconLibre = 0
  iconOcupado
End Enum

Private strReadyMessage As String 'strMessage que muestra en la barra de estado

Private Sub Resize()
'--> Cambia los tamaños de la barra de progreso y la etiqueta cuando está o no visible
Dim intImages As eIcons

  On Error Resume Next 'Pueden darse errores si la ventana es demasiado pequeña
  lblMessage.Top = (ScaleHeight - lblMessage.Height) / 2
  brProgreso.Top = (ScaleHeight - brProgreso.Height) / 2
  With stBar
    .Top = brProgreso.Top
    .Height = brProgreso.Height
    .Left = ScaleWidth - .Width - imgStatus(iconLibre).Width - 70
  End With
  If brProgreso.Visible Then
    brProgreso.Left = lblMessage.Width + lblMessage.Left + 25
    brProgreso.Width = ScaleWidth - brProgreso.Left - stBar.Width - imgStatus(iconLibre).Width - 150
  End If
  For intImages = iconLibre To iconOcupado
    With imgStatus(intImages)
      .Left = ScaleWidth - .Width
      .Top = (ScaleHeight - .Height) / 2
    End With
  Next intImages
  stBar.Refresh
  lblMessage.Refresh
End Sub

Public Sub writeMessage(Optional ByVal strMessage As String = "")
'--> Escribe un Message sobre la barra de estado
  lblMessage = IIf(strMessage <> "", strMessage + " ", strReadyMessage)
End Sub

Public Sub changeProgressBar(Optional ByVal intIncrement As Integer = 1)
'--> Incrementa en uno la barra de progreso
  brProgreso.Value = brProgreso.Value + intIncrement
End Sub

Public Sub closeProgressBar()
'--> Cierra la barra de progreso, deja como Message el establecido en <B> strReadyMessage </B>
  brProgreso.Visible = False
  imgStatus(iconLibre).Visible = True
  imgStatus(iconOcupado).Visible = False
  Resize
  writeMessage
End Sub

Public Sub initProgressBar(ByVal strMessage As String, ByVal lngMin As Long, ByVal lngMax As Long)
'--> Inicializa la barra de progreso.
  On Error Resume Next
  writeMessage strMessage
  With brProgreso
    .Maximus = lngMax
    .Value = 0
    brProgreso.Visible = True
  End With
  imgStatus(iconLibre).Visible = False
  imgStatus(iconOcupado).Visible = True
  Resize
End Sub

Property Let Message(ByVal strMessage As String)
'--> Define un Message como Message por defecto sobre la barra de estado.
  strReadyMessage = strMessage
  writeMessage strReadyMessage
  PropertyChanged
End Property

Property Get Message() As String
'--> Obtiene un Message por defecto
  Message = strReadyMessage
End Property

Private Sub UserControl_Initialize()
  writeMessage strReadyMessage
  brProgreso.Visible = False
End Sub

Private Sub UserControl_InitProperties()
  Message = cnstStrDefaultMessage
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Message = PropBag.ReadProperty("Message", cnstStrDefaultMessage)
End Sub

Private Sub UserControl_Resize()
  If Height < 264 Then
    Height = 264
  Else
    Resize
  End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "setHelpID", lblMessage.WhatsThisHelpID, 0
  PropBag.WriteProperty "Message", strReadyMessage, cnstStrDefaultMessage
End Sub
