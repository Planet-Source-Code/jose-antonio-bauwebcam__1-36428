VERSION 5.00
Begin VB.UserControl progressBar 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4170
   ScaleHeight     =   225
   ScaleWidth      =   4170
End
Attribute VB_Name = "progressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--> Barra de progreso
Option Explicit

Public Enum typeGradiant
  gradNone = 0
  gradRed
  gradGreen
  gradBlue
End Enum

Public Enum typeBorder
  borderNone = 0
  borderSingle
End Enum

Private intMax As Integer, intIncrement As Integer, intValue As Integer
Private blnShowPercent As Boolean
Private intType As Integer
Private bytRed As Byte, bytGreen As Byte, bytBlue As Byte
Private sngStep As Single

Public Sub initProgress()
  Value = 0
End Sub

Public Sub changeProgress()
  Value = Value + Increment
End Sub

Private Sub updateProgress()
'--> Repinta la barra de progreso
Dim sngPercent As Single, sngValc As Single
Dim strData As String

  If intMax <> 0 Then
    'Calcula el porcentaje
      sngPercent = intValue * 100 / intMax
    'Parámetros de la imagen
      If Not UserControl.AutoRedraw Then
        UserControl.AutoRedraw = True
      End If
      UserControl.Cls
      UserControl.ScaleWidth = 100
    'Rellena la imagen
      If sngStep = 0 Then
        sngStep = 0.2
      End If
      For sngValc = 0 To sngPercent Step sngStep
        Select Case intType
          Case gradNone
            UserControl.Line (sngValc, 0)-(sngValc, UserControl.ScaleHeight), _
                             RGB(bytRed, bytGreen, bytBlue)
          Case gradRed
            UserControl.Line (sngValc, 0)-(sngValc, UserControl.ScaleHeight), _
                             RGB(sngValc * 2.55, bytGreen, bytBlue)
          Case gradGreen
            UserControl.Line (sngValc, 0)-(sngValc, UserControl.ScaleHeight), _
                             RGB(bytRed, sngValc * 2.55, bytBlue)
          Case gradBlue
            UserControl.Line (sngValc, 0)-(sngValc, UserControl.ScaleHeight), _
                             RGB(bytRed, bytGreen, sngValc * 2.55)
        End Select
      Next sngValc
    'Muestra el porcentaje
      If blnShowPercent = True Then
        strData = Format(sngPercent, "0") + "%"
        UserControl.CurrentX = (UserControl.ScaleWidth - UserControl.TextWidth(strData)) / 2
        UserControl.CurrentY = (UserControl.ScaleHeight - UserControl.TextHeight(strData)) / 2
        Print strData
      End If
    'Actualiza la pantalla
      UserControl.Refresh
  End If
End Sub

Public Property Let Maximus(ByVal intNewMax As Integer)
  intMax = intNewMax
  updateProgress
  PropertyChanged
End Property

Public Property Get Maximus() As Integer
  Maximus = intMax
End Property

Public Property Let Increment(ByVal intNewIncrement As Integer)
  intIncrement = intNewIncrement
  updateProgress
  PropertyChanged
End Property

Public Property Get Increment() As Integer
  Increment = intIncrement
End Property

Public Property Let Value(ByVal intNewValue As Integer)
  intValue = intNewValue
  If intValue > intMax Then
    intValue = intMax
  End If
  updateProgress
  PropertyChanged
End Property

Public Property Get Value() As Integer
  Value = intValue
End Property

Public Property Let SingleStep(ByVal sngNewStep As Single)
  sngStep = sngNewStep
  If sngNewStep < 0.2 Then
    sngStep = 0.2
  End If
  updateProgress
  PropertyChanged
End Property

Public Property Get SingleStep() As Single
  SingleStep = sngStep
End Property

Public Property Let ColorRed(ByVal bytNewRed As Byte)
  bytRed = bytNewRed
  updateProgress
  PropertyChanged
End Property

Public Property Get ColorRed() As Byte
  ColorRed = bytRed
End Property

Public Property Let ColorBlue(ByVal bytNewBlue As Byte)
  bytBlue = bytNewBlue
  updateProgress
  PropertyChanged
End Property

Public Property Get ColorBlue() As Byte
  ColorBlue = bytBlue
End Property

Public Property Let ColorGreen(ByVal bytNewGreen As Byte)
  bytGreen = bytNewGreen
  updateProgress
  PropertyChanged
End Property

Public Property Get ColorGreen() As Byte
  ColorGreen = bytGreen
End Property

Public Property Let ShowPercent(ByVal blnNewShowPercent As Boolean)
  blnShowPercent = blnNewShowPercent
  updateProgress
  PropertyChanged
End Property

Public Property Get ShowPercent() As Boolean
  ShowPercent = blnShowPercent
End Property

Public Property Let Gradiant(ByVal intNewType As typeGradiant)
  intType = intNewType
  updateProgress
  PropertyChanged
End Property

Public Property Get Gradiant() As typeGradiant
  Gradiant = intType
End Property

Public Property Let BorderStyle(ByVal intNewBorder As typeBorder)
  UserControl.BorderStyle = intNewBorder
End Property

Public Property Get BorderStyle() As typeBorder
  BorderStyle = UserControl.BorderStyle
End Property

Public Property Set Font(ByVal fntNewFont As Font)
   Set UserControl.Font = fntNewFont
   updateProgress
   PropertyChanged "Font"
End Property

Public Property Get Font() As Font
   Set Font = UserControl.Font
End Property

Public Property Let ForeColor(ByVal colNewForeColor As OLE_COLOR)
  UserControl.ForeColor = colNewForeColor
  updateProgress
  PropertyChanged
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = UserControl.ForeColor
End Property

Private Sub UserControl_InitProperties()
  Increment = 1
  Maximus = 100
  Value = 0
  SingleStep = 0.2
  ColorRed = 0
  ColorGreen = 255
  ColorBlue = 0
  ShowPercent = False
  Gradiant = gradRed
  BorderStyle = borderNone
  Set Font = Ambient.Font
  ForeColor = vbWhite
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Maximus = PropBag.ReadProperty("Maximus", 100)
  Increment = PropBag.ReadProperty("Increment", 1)
  Value = PropBag.ReadProperty("Value", 0)
  SingleStep = PropBag.ReadProperty("SingleStep", 0.2)
  ColorRed = PropBag.ReadProperty("ColorRed", 0)
  ColorGreen = PropBag.ReadProperty("ColorGreen", 255)
  ColorBlue = PropBag.ReadProperty("ColorBlue", 0)
  ShowPercent = PropBag.ReadProperty("ShowPercent", False)
  Gradiant = PropBag.ReadProperty("Gradiant", gradRed)
  BorderStyle = PropBag.ReadProperty("BorderStyle", borderNone)
  Set Font = PropBag.ReadProperty("Font", Ambient.Font)
  ForeColor = PropBag.ReadProperty("ForeColor", Ambient.ForeColor)
End Sub

Private Sub UserControl_Resize()
  If Width < 100 Then
    Width = 100
  ElseIf Height < 225 Then
    Height = 225
  Else
    updateProgress
  End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Maximus", Maximus, 100
  PropBag.WriteProperty "Increment", Increment, 1
  PropBag.WriteProperty "Value", Value, 0
  PropBag.WriteProperty "SingleStep", SingleStep, 0.2
  PropBag.WriteProperty "ColorRed", ColorRed, 0
  PropBag.WriteProperty "ColorGreen", ColorGreen, 255
  PropBag.WriteProperty "ColorBlue", ColorBlue, 0
  PropBag.WriteProperty "ShowPercent", ShowPercent, False
  PropBag.WriteProperty "Gradiant", Gradiant, gradGreen
  PropBag.WriteProperty "BorderStyle", BorderStyle, borderNone
  PropBag.WriteProperty "Font", Font, Ambient.Font
  PropBag.WriteProperty "ForeColor", ForeColor, Ambient.ForeColor
End Sub
