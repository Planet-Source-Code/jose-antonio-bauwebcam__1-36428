VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMsgBox 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Administración de Metadatos"
   ClientHeight    =   2370
   ClientLeft      =   2070
   ClientTop       =   2370
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "fmsgbox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2370
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlImage 
      Left            =   5340
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmsgbox.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmsgbox.frx":2E96
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmsgbox.frx":5110
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin BAUCommonControls.ImageButton cmdNo 
      Height          =   465
      Left            =   3300
      TabIndex        =   2
      Top             =   1830
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   820
      Caption         =   "&Cancelar"
      Picture         =   "fmsgbox.frx":542A
      PictureOver     =   "fmsgbox.frx":5877
      PictureDown     =   "fmsgbox.frx":5CD7
      CaptionPosition =   4
      BackColorOver   =   -2147483633
      BackColorDown   =   -2147483633
      ForeColorOver   =   -2147483635
      ForeColorDown   =   128
      ImageTop        =   3
      ImageLeft       =   3
      LabelTop        =   7
      LabelLeft       =   27
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
   Begin VB.Frame Frame1 
      Height          =   1668
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   5724
      Begin VB.Label lblMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "LLLL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1305
         Left            =   930
         TabIndex        =   1
         Top             =   210
         Width           =   4620
      End
      Begin VB.Image imgError 
         Appearance      =   0  'Flat
         Height          =   600
         Left            =   180
         Top             =   180
         Width           =   675
      End
   End
   Begin BAUCommonControls.ImageButton cmdYes 
      Height          =   465
      Left            =   1260
      TabIndex        =   3
      Top             =   1830
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   820
      Caption         =   "&Aceptar"
      Picture         =   "fmsgbox.frx":5DC3
      PictureOver     =   "fmsgbox.frx":621A
      PictureDown     =   "fmsgbox.frx":668E
      CaptionPosition =   4
      BackColorOver   =   -2147483633
      BackColorDown   =   -2147483633
      ForeColorOver   =   -2147483635
      ForeColorDown   =   128
      ImageTop        =   3
      ImageLeft       =   3
      LabelTop        =   7
      LabelLeft       =   27
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
      Height          =   465
      Left            =   4470
      TabIndex        =   4
      Top             =   1860
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   820
      Caption         =   "&Cancelar"
      Picture         =   "fmsgbox.frx":678F
      PictureOver     =   "fmsgbox.frx":6C26
      PictureDown     =   "fmsgbox.frx":70C5
      CaptionPosition =   4
      BackColorOver   =   -2147483633
      BackColorDown   =   -2147483633
      ForeColorOver   =   -2147483635
      ForeColorDown   =   128
      ImageTop        =   3
      ImageLeft       =   3
      LabelTop        =   7
      LabelLeft       =   27
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
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Ventana para la presentación de mensajes de error.
Option Explicit

Private Const cnstIntBetweenButtons As Integer = 500 'Separación de los botones de aceptar y cancelar

Public strMessage As String
Public strTitle As String
Public intType As Integer
Public intResult As Integer

Private Sub cmdYes_Click()
  intResult = 0
  Unload Me
End Sub

Private Sub cmdCancel_Click()
  intResult = 2
  Unload Me
End Sub

Private Sub cmdNo_Click()
  intResult = 1
  Unload Me
End Sub

Private Sub Form_Load()
'--> Inicializa la ventana
  lblMessage.Caption = strMessage
  Me.Caption = strTitle
  Select Case intType
    Case 0 'msgInformation
      Set imgError.Picture = imlImage.ListImages(2).Picture
    Case 1 'msgQuestion
      Set imgError.Picture = imlImage.ListImages(1).Picture
    Case 2 'msgExclamation
      Set imgError.Picture = imlImage.ListImages(3).Picture
    Case 3 'msgQuestionCancel
      Set imgError.Picture = imlImage.ListImages(1).Picture
  End Select
  If intType = 0 Or intType = 2 Then 'Information, exclamation
    With cmdYes
      .Left = (Me.ScaleWidth - .Width) / 2
    End With
    cmdNo.Visible = False
    intResult = 0 'Yes
  ElseIf intType = 1 Then 'Question
    intResult = 1 'No
  ElseIf intType = 3 Then 'Question - Cancel
    intResult = 2 'Cancel
    cmdYes.Caption = "&Sí"
    cmdYes.Left = 210
    cmdNo.Caption = "&No"
    cmdNo.Left = 2250
    cmdCancel.Visible = True
    cmdCancel.Left = 4470
  End If
End Sub
