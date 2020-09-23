VERSION 5.00
Begin VB.Form frmSeeAvi 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Visualización de ficheros AVI"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraVideo 
      Caption         =   "Vídeo"
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
      Height          =   5325
      Left            =   60
      TabIndex        =   4
      Top             =   1020
      Width           =   8805
      Begin AVIFile.AVIShow aviShow 
         Height          =   4905
         Left            =   150
         TabIndex        =   5
         Top             =   270
         Width           =   8445
         _ExtentX        =   14896
         _ExtentY        =   8652
         BorderStyle     =   3
      End
   End
   Begin VB.Frame fraParameters 
      Caption         =   "Parámetros del vídeo"
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
      Height          =   795
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   8715
      Begin BAUCommonControls.ImageButton cmdFile 
         Height          =   375
         Left            =   8160
         TabIndex        =   1
         Top             =   270
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         Caption         =   "..."
         BackColorOver   =   -2147483633
         BackColorDown   =   -2147483633
         ForeColorOver   =   -2147483635
         ForeColorDown   =   128
         ImageTop        =   5
         ImageLeft       =   14
         LabelTop        =   6
         LabelLeft       =   10
         NbLine          =   0
         NbLineOver      =   1
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
      Begin BAUCommonControls.Edicion txtFileName 
         Height          =   315
         Left            =   2280
         TabIndex        =   2
         Top             =   300
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   503
         BackColorCaption=   -2147483643
         Caption         =   "<Nombre del fichero AVI>"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontCaption {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
         Caption         =   "Nombre del fichero:"
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
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1875
      End
   End
   Begin BAUCommonControls.ImageButton cmdAccept 
      Height          =   435
      Left            =   2730
      TabIndex        =   6
      Top             =   6450
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Caption         =   "&Ver vídeo"
      Picture         =   "frmSeeAvi.frx":0000
      PictureOver     =   "frmSeeAvi.frx":0457
      PictureDown     =   "frmSeeAvi.frx":08CB
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
      Left            =   4740
      TabIndex        =   7
      Top             =   6450
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Caption         =   "&Cerrar"
      Picture         =   "frmSeeAvi.frx":0D3F
      PictureOver     =   "frmSeeAvi.frx":118C
      PictureDown     =   "frmSeeAvi.frx":15EC
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
End
Attribute VB_Name = "frmSeeAvi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Ventana para visualizar ficheros AVI
Option Explicit

Private Sub getFileNameAVI()
'--> Obtiene el nombre del fichero AVI
Dim objFile As New clsFiles
Dim strPath As String, strFileName As String

  'Obtiene el nombre de fichero
    strPath = Trim(objFile.getPath(txtFileName.Text))
    If strPath = "" Then
      strPath = "C:\"
    End If
    strFileName = Trim(objFile.dlgGetFileName(frmMDIMain.dlgCommon, True, strPath, _
                                              "Ficheros de vídeo (*.avi)|*.avi"))
    If strFileName <> "" Then
      txtFileName.Text = strFileName
    End If
  'Libera la memoria
    Set objFile = Nothing
End Sub

Private Sub showAVI()
'--> Muestra el AVI
  If txtFileName.Text = "" Then
    frmMDIMain.messageBox.showMessage "Seleccione el nombre de fichero", App.Title, MsgExclamation
  Else
    With aviShow
      .ResourceID = 0
      .FileName = txtFileName.Text
      '.Autoplay = True
      .AutoSize = True
      .StartPlay
    End With
  End If
End Sub

Private Sub cmdAccept_Click()
  showAVI
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdFile_Click()
  getFileNameAVI
End Sub
