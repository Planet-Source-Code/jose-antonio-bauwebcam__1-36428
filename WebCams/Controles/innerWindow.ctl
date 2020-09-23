VERSION 5.00
Begin VB.UserControl innerWindow 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5115
   ControlContainer=   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3330
   ScaleWidth      =   5115
   Begin bauWebCams.GradientLabel lblCaption 
      Height          =   225
      Left            =   120
      Top             =   90
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   397
      Caption         =   "innerWindow"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdVisible 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   4824
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   48
      Width           =   228
   End
End
Attribute VB_Name = "innerWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--> <H3> Control para permitir alineaciones una simulación del tipo de ventanas hija de VB 5.0</H3>
'--> Existen dos lblCaption para que no se vea el corte de palabra cuando lel ancho es demasiado pequeño
Option Explicit

Private Const AnchuraMinima = 700
Private Const AlturaMinima = 700

Public Event CloseWindow()
Public Event Click()
Public Event Resize()

Public Sub Redimensionar()
Dim IDControl As Control

  'On Error Resume Next
  If Width < AnchuraMinima Then
    Width = AnchuraMinima
  ElseIf Height < AlturaMinima Then
    Height = AlturaMinima
  Else
    With lblCaption
      .Left = ScaleLeft
      .Top = ScaleTop
      .Width = ScaleWidth
    End With
    With cmdVisible
      .Left = ScaleWidth - .Width
      .Top = ScaleTop
      .Height = lblCaption.Height
    End With
    RaiseEvent Resize
    cmdVisible.ZOrder 0
  End If
End Sub

Public Property Let Enabled(ByVal pEnable As Boolean)
  UserControl.Enabled = pEnable
  PropertyChanged "Enabled"
End Property

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Caption(ByVal Cadena As String)
  lblCaption.Caption = Cadena
  PropertyChanged
End Property

Public Property Get Caption() As String
  Caption = lblCaption.Caption
End Property

Private Sub cmdVisible_Click()
  RaiseEvent Click
  RaiseEvent CloseWindow
End Sub

Private Sub UserControl_Initialize()
  lblCaption.Caption = ""
End Sub

Private Sub UserControl_InitProperties()
  Caption = "VentanaInterna"
  Enabled = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Caption = PropBag.ReadProperty("Caption", "VentanaInterna")
  Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
  Redimensionar
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Caption", Caption, "Caption"
  PropBag.WriteProperty "Enabled", Enabled, True
End Sub
