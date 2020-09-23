VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3600
   ClientLeft      =   4500
   ClientTop       =   3840
   ClientWidth     =   4725
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   ScaleHeight     =   3600
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "@ 2002"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2370
      TabIndex        =   5
      Top             =   2970
      Width           =   2145
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Bau Consultors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2370
      TabIndex        =   4
      Top             =   2760
      Width           =   2145
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Cámaras web"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   450
      TabIndex        =   3
      Top             =   2550
      Width           =   2265
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Bau WebCam"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   2070
      Width           =   3705
   End
   Begin VB.Label lblMessage 
      BackColor       =   &H00808080&
      Caption         =   "Cargando ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   3300
      Width           =   4485
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Bau WebCam"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   465
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   2130
      Width           =   3705
   End
   Begin VB.Image Image1 
      Height          =   3600
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      Top             =   0
      Width           =   4725
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Pantalla de presentación
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Function waitTime(ByVal sngSecToDelay As Single)
Dim sngMarkTime As Single

  sngMarkTime = GetTickCount
  While GetTickCount() < sngMarkTime + sngSecToDelay * 1000 'Lo convierte en segundos
    DoEvents
  Wend
End Function

Public Sub showMessage(ByVal strMessage As String)
'--> Cambia el mensaje que se muestra en pantalla
  lblMessage.Caption = strMessage
  lblMessage.Refresh
  DoEvents
  'DelayTime 2
End Sub

Private Sub Form_Load()
  lblTitle(0).Caption = App.ProductName
  lblTitle(1).Caption = App.ProductName
  lblDescription.Caption = App.Comments
End Sub
