VERSION 5.00
Begin VB.UserControl ScrollPicture 
   Alignable       =   -1  'True
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7875
   LockControls    =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   7875
   Begin VB.PictureBox imgFondo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      ForeColor       =   &H00808080&
      Height          =   4335
      Left            =   0
      ScaleHeight     =   4275
      ScaleWidth      =   7455
      TabIndex        =   2
      Top             =   0
      Width           =   7515
      Begin VB.PictureBox imgImage 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4500
         Left            =   0
         Picture         =   "ScrollPicture.ctx":0000
         ScaleHeight     =   4500
         ScaleWidth      =   6000
         TabIndex        =   3
         Top             =   0
         Width           =   6000
      End
   End
   Begin VB.VScrollBar VScroll 
      Height          =   4335
      Left            =   7530
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4380
      Visible         =   0   'False
      Width           =   7500
   End
End
Attribute VB_Name = "ScrollPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--> Control para realizar el scroll de una imagen
Option Explicit

Public Event ImageClick()
Public Event ImageDblClick()

Private strFileImage As String

Private Sub AdjustScrollBars()
'--> Ajusta las barras de scroll de acuerdo al tamaño de imgImage
Dim objActiveControl As Control

  On Error Resume Next
  'Scroll vertical
    With VScroll
      .Value = 0
      .Min = 0
      .Max = IIf(imgFondo.Height > imgImage.Height, 0, imgImage.Height - imgFondo.Height)
      .SmallChange = (.Max / 20) + 1
      .LargeChange = (.Max / 5) + 1
      .Visible = (.Max > 0)
      .Refresh
    End With
  'Scroll horizontal
    With HScroll
      .Value = 0
      .Min = 0
      .Max = IIf(imgFondo.Width > imgImage.Width, 0, imgImage.Width - imgFondo.Width)
      .SmallChange = (.Max / 20) + 1
      .LargeChange = (.Max / 5) + 1
      .Visible = (.Max > 0)
      .Refresh
    End With
  'Evita el parpadeo en las barras de scroll
    If imgFondo.Visible Then
      Set objActiveControl = UserControl.ActiveControl
      imgFondo.SetFocus
      Select Case objActiveControl.Name
        Case "HScroll1", "VScroll1"
          'If objActiveControl.Max <> 0 Then
          '  objActiveControl.SetFocus
          'End If
        Case Else
          objActiveControl.SetFocus
      End Select
      Set objActiveControl = Nothing
    End If
End Sub

Public Function loadImage(ByVal strFileName As String) As Boolean
'--> Carga una imagen
  On Error Resume Next
  'Supone que todo es correcto
    loadImage = True
  'Carga la imagen
    imgImage.Picture = LoadPicture(strFileName)
    strFileImage = strFileName
  'Comprueba si todo es correcto
    If Err.Number <> 0 Then
      loadImage = False
    Else
      AdjustScrollBars
    End If
End Function

Public Property Let Stretch(ByVal blnStretch As Boolean)
'--> Indica si se debe redimensionar la imagen
  'imgImage.Stretch = blnStretch
  AdjustScrollBars
End Property

Public Property Get Stretch() As Boolean
'--> Obtiene si se redimensiona la imagen
  'Stretch = imgImage.Picture.Stretch
End Property

Public Sub PrintImage()
'--> Imprime la imagen
  If imgImage.Picture <> 0 Then
    'frmPrintScreen.PrintBitmap imgImage.Picture
  Else
    MsgBox "No existe ninguna imagen", vbInformation, "Scrool picture"
  End If
End Sub

Public Sub Save(ByVal strFileName As String)
'--> Graba la imagen en disco
  If imgImage.Picture = 0 Then
    MsgBox "No existe ninguna imagen", vbInformation, "Scroll picture"
  Else
    SavePicture imgImage.Image, strFileName
  End If
End Sub

Public Sub Copy()
'--> Copia la imagen
  If imgImage.Picture <> 0 Then
    Clipboard.Clear
    DoEvents
    Clipboard.SetData imgImage.Picture
    DoEvents
  Else
    MsgBox "No existe niguna imagen.", vbInformation, "Scroll picture"
  End If
End Sub

Public Property Get FileName() As String
'--> Devuelve la imagen cargada
  FileName = strFileImage
End Property

Private Sub imgFondo_DblClick()
  RaiseEvent ImageDblClick
End Sub

Private Sub imgImage_DblClick()
  RaiseEvent ImageDblClick
End Sub

Private Sub imgImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    RaiseEvent ImageClick
  End If
End Sub

Private Sub UserControl_DblClick()
  RaiseEvent ImageDblClick
End Sub

Private Sub UserControl_Initialize()
'--> Inicializa los controles
  'Coloca imgImage en la parte superior izquierda
    With imgImage
      .Left = 0
      .Top = 0
    End With
End Sub

Private Sub VScroll_Change()
'--> Controla el scroll vertical de la imagen
  If imgImage.Height > imgFondo.Height Then
    imgImage.Top = 0 - (VScroll.Value + 60)
  End If
End Sub

Private Sub HScroll_Change()
'--> Controla el scroll horizontal de la imagen
  If imgImage.Width > imgFondo.Width Then
    imgImage.Left = 0 - (HScroll.Value + 60)
  End If
End Sub

Private Sub UserControl_Resize()
'--> Controla la redimensión del control de usuario
  On Error Resume Next
  'Tamaño mínimo del control de usuario
    If Width < 400 Then
      Width = 400
    ElseIf Height < 400 Then
      Height = 400
    Else
      'Modifica el tamaño de la imagen de fondo
        With imgFondo
          .Width = Width - VScroll.Width
          .Height = Height - .Top - HScroll.Height
        End With
      'Modifica el scroll vertical
        With VScroll
          .Left = Width - VScroll.Width
          .Height = imgFondo.Height
          .Top = imgFondo.Top
          .Value = 0
        End With
      'Modifica el scroll horizontal
        With HScroll
          .Top = Height - HScroll.Height
          .Width = imgFondo.Width
          .Value = 0
          .Left = imgFondo.Left
        End With
      'Ajusta las barras de scroll
        AdjustScrollBars
    End If
End Sub
