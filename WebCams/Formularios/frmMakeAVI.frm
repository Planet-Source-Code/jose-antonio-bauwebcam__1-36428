VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{8D5F548D-FB40-41D9-B8CC-493E031AE985}#2.0#0"; "BAUCommonControls.ocx"
Object = "*\A..\..\UserControls\AVIFiles\prjAVIFiles.vbp"
Begin VB.Form frmMakeAVI 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Creación de vídeo"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin AVIFile.AVIFiles ctlAVIFile 
      Left            =   180
      Top             =   5340
      _ExtentX        =   1032
      _ExtentY        =   1032
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
      Left            =   30
      TabIndex        =   8
      Top             =   90
      Width           =   8715
      Begin BAUCommonControls.ImageButton cmdPath 
         Height          =   375
         Index           =   0
         Left            =   8160
         TabIndex        =   9
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
         TabIndex        =   10
         Top             =   300
         Width           =   5895
         _ExtentX        =   10398
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
         Caption         =   "Fichero destino:"
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
         TabIndex        =   11
         Top             =   360
         Width           =   1875
      End
   End
   Begin VB.Frame fraImage 
      Caption         =   "Imágenes"
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
      Height          =   4545
      Left            =   30
      TabIndex        =   0
      Top             =   900
      Width           =   8745
      Begin MSComctlLib.ListView lswImages 
         Height          =   3195
         Left            =   3660
         TabIndex        =   12
         Top             =   1230
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5636
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin BAUCommonControls.Directorio dirPath 
         Left            =   3270
         Top             =   4080
         _ExtentX        =   582
         _ExtentY        =   582
      End
      Begin VB.FileListBox lstFiles 
         Height          =   3210
         Left            =   240
         MultiSelect     =   2  'Extended
         TabIndex        =   1
         Top             =   1230
         Width           =   2925
      End
      Begin BAUCommonControls.ImageButton cmdAddImage 
         Height          =   375
         Left            =   3180
         TabIndex        =   2
         Top             =   2430
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         Caption         =   ">>"
         BackColorOver   =   -2147483633
         BackColorDown   =   -2147483633
         ForeColorOver   =   -2147483635
         ForeColorDown   =   128
         ImageTop        =   5
         ImageLeft       =   14
         LabelTop        =   6
         LabelLeft       =   8
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
      Begin BAUCommonControls.ImageButton cmdPath 
         Height          =   375
         Index           =   1
         Left            =   8250
         TabIndex        =   3
         Top             =   240
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
      Begin VB.Label lblPath 
         Caption         =   "lblPath"
         Height          =   315
         Left            =   1290
         TabIndex        =   5
         Top             =   360
         Width           =   6855
      End
      Begin VB.Label lblImage 
         Caption         =   "Incluir en el fichero:"
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
         Left            =   3510
         TabIndex        =   7
         Top             =   930
         Width           =   2955
      End
      Begin VB.Label lblCaption 
         Caption         =   "Directorio:"
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
         Left            =   180
         TabIndex        =   6
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label lblImage 
         Caption         =   "Imágenes del directorio:"
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
         Left            =   180
         TabIndex        =   4
         Top             =   960
         Width           =   2955
      End
   End
   Begin BAUCommonControls.ImageButton cmdAccept 
      Height          =   435
      Left            =   1530
      TabIndex        =   13
      Top             =   5580
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Caption         =   "&Generar vídeo"
      Picture         =   "frmMakeAVI.frx":0000
      PictureOver     =   "frmMakeAVI.frx":0457
      PictureDown     =   "frmMakeAVI.frx":08CB
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
      Left            =   5550
      TabIndex        =   14
      Top             =   5550
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Caption         =   "&Cerrar"
      Picture         =   "frmMakeAVI.frx":0D3F
      PictureOver     =   "frmMakeAVI.frx":118C
      PictureDown     =   "frmMakeAVI.frx":15EC
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
   Begin BAUCommonControls.ImageButton cmdShowAVI 
      Height          =   435
      Left            =   3480
      TabIndex        =   15
      Top             =   5580
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Caption         =   "&Ver vídeo"
      Picture         =   "frmMakeAVI.frx":16D8
      PictureOver     =   "frmMakeAVI.frx":1B2F
      PictureDown     =   "frmMakeAVI.frx":1FA3
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
Attribute VB_Name = "frmMakeAVI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Formulario para la creación de un fichero AVI
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
         ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub init()
'--> Inicializa el formulario
  'Pone un scroll horizontal en las listas
    SendMessage lstFiles.hWnd, &H194, 200, ByVal 0&
  'Inicializa el filtro de la lista de ficheros
    lstFiles.Pattern = "*.jpg;*.gif;*.bmp"
   'Cambia los parámetros del listView
    With lswImages
      .View = lvwReport
      .ColumnHeaders.Add , , "Directorio"
      .ColumnHeaders.Add , , "Archivo"
    End With
  'Inicializa los directorios
    lstFiles.Path = App.Path
    lblPath.Caption = lstFiles.Path
End Sub

Private Sub changePath(ByVal intIndex As Integer)
Dim strPath As String

  'Recoge el directorio
    strPath = dirPath.BrowseForFolder(Me.hWnd, "Seleccione el directorio")
  'Comprueba si hay algo que modificar
    If strPath <> "" Then
      Select Case intIndex
        Case 0 'Cambia el fichero de salida
          txtFileName.Text = Replace(strPath & "\aviCam.avi", "\\", "\")
        Case 1 'Cambia la ruta de los ficheros
          lblPath.Caption = strPath
          lstFiles.Path = strPath
      End Select
    End If
End Sub

Private Sub addImage()
'--> Añade las imágenes seleccionadas
Dim intIndex As Integer
Dim itmNode As ListItem

  For intIndex = 0 To lstFiles.ListCount - 1
    If lstFiles.Selected(intIndex) Then
      Set itmNode = lswImages.ListItems.Add(, , lstFiles.Path)
      itmNode.ListSubItems.Add , , lstFiles.List(intIndex)
    End If
  Next intIndex
End Sub

Private Sub createAVI()
'--> Crea el fichero AVI
Dim intIndex As Integer

  If txtFileName.Text = "" Then
    frmMDIMain.messageBox.showMessage "Seleccione el nombre del fichero", App.Title, MsgExclamation
  ElseIf lswImages.ListItems.Count = 0 Then
    frmMDIMain.messageBox.showMessage "Seleccione los ficheros que desea pasar a vídeo", App.Title, MsgExclamation
  Else
    With ctlAVIFile
      'Añade los ficheros de imagen al AVI
        For intIndex = 1 To lswImages.ListItems.Count
          .addFile Replace(lswImages.ListItems(intIndex).Text & "\" & lswImages.ListItems(intIndex).ListSubItems(1).Text, "\\", "\")
        Next intIndex
      'Crea el AVI
        If Not .createFile(Me.hWnd, txtFileName.Text) Then
          frmMDIMain.messageBox.showMessage "Error al generar el fichero" & vbCrLf & .ErrorMessage, _
                                            App.Title, MsgExclamation
        Else
          frmMDIMain.messageBox.showMessage "Se ha generado el fichero AVI", App.Title, MsgInformation
        End If
    End With
  End If
End Sub

Private Sub showAVI()
'--> Abre el reproductor de medios para mostrar el vídeo
  If txtFileName.Text = "" Then
    frmMDIMain.messageBox.showMessage "Seleccione el nombre del fichero", App.Title, MsgExclamation
  Else
    On Error Resume Next
    ShellExecute 0&, vbNullString, txtFileName.Text, vbNullString, "C:\", 1
  End If
End Sub

Private Sub cmdAccept_Click()
  createAVI
End Sub

Private Sub cmdAddImage_Click()
  addImage
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdPath_Click(Index As Integer)
  changePath Index
End Sub

Private Sub cmdShowAVI_Click()
  showAVI
End Sub

Private Sub Form_Load()
  init
End Sub
