VERSION 5.00
Begin VB.UserControl Edicion 
   BackStyle       =   0  'Transparent
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6810
   LockControls    =   -1  'True
   ScaleHeight     =   2865
   ScaleWidth      =   6810
   ToolboxBitmap   =   "Edicion.ctx":0000
   Begin VB.TextBox txtEdicion 
      Height          =   1395
      Index           =   1
      Left            =   420
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1140
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.TextBox txtEdicion 
      Height          =   285
      Index           =   0
      Left            =   420
      TabIndex        =   1
      Top             =   510
      Width           =   4755
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "<Mensaje Inicial>"
      Height          =   285
      Left            =   30
      TabIndex        =   0
      Top             =   180
      Width           =   6735
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Edicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--> Control de texto con colores de fondo cuando se selecciona y control de caracteres numéricos
Option Explicit

Private cSelectedBackColor As OLE_COLOR 'Color de fondo cuando se selecciona el control
Private cSelectedForeColor As OLE_COLOR 'Color de texto cuando se selecciona el control
Private cUnselectedBackColor As OLE_COLOR 'Color de fondo cuando se deselecciona el control
Private cUnselectedForeColor As OLE_COLOR 'Color de texto cuando se deselecciona el control
Private blnMultiline As Boolean 'Indica si es multilínea o no
Private blnNumerico As Boolean 'Indica si es o no numérico
Private intNumDecimales As Integer 'Número de decimales

Public Event Change() 'Evento lanzado al cambiar el texto del control

Private strCharDecimal As String 'Caracter decimal
Private intIndexEdicion As Integer
'Default Property Values:
Const m_def_EnterTab = True
'Property Variables:
Dim m_EnterTab As Boolean
'Event Declarations:
Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event EnterPress()



Private Function CountNumeroDecimales(ByVal intNumeroControl As Integer) As Integer
'--> Cuenta el número de decimales que hay hasta ahora en el campo
Dim intPosicionPunto As Integer

  intPosicionPunto = InStr(txtEdicion(intNumeroControl).Text, strCharDecimal)
  If intPosicionPunto = 0 Then
    CountNumeroDecimales = 0
  Else
    CountNumeroDecimales = Len(txtEdicion(intNumeroControl).Text) - intPosicionPunto
  End If
End Function

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal pEnabled As Boolean)
  UserControl.Enabled = pEnabled
  txtEdicion(0).Enabled = pEnabled
  txtEdicion(1).Enabled = pEnabled
  PropertyChanged
End Property

Public Property Get Caption() As String
  Caption = lblMensaje
End Property

Public Property Let Caption(ByVal strCaption As String)
  lblMensaje = strCaption
  UserControl_Resize
  PropertyChanged
End Property


Public Property Get Value() As Double
  Value = ValorNumerico(Text)
End Property

Public Function ValorNumerico(ByVal Cadena As String) As Double
'--> Devuelve el valor numérico de una cadena de tipo "#,##.00"
  Dim CadSalida As String
  Dim Indice    As Integer

  CadSalida = ""
  For Indice = 1 To Len(Cadena)
    If Mid(Cadena, Indice, 1) = "." Then
      CadSalida = CadSalida 'No se añade nada
    ElseIf Mid(Cadena, Indice, 1) = "," Then
      CadSalida = CadSalida + "."
    Else
      CadSalida = CadSalida + Mid(Cadena, Indice, 1)
    End If
  Next Indice
  ValorNumerico = Val(CadSalida)
  
End Function
Public Property Let Value(ByVal dblNewValue As Double)
  Text = Format(dblNewValue, "0" + IIf(DecimalNumber <> 0, "." + String(DecimalNumber, "0"), ""))
End Property

Public Property Get PasswordChar() As String
  PasswordChar = txtEdicion(intIndexEdicion).PasswordChar
End Property

Public Property Let PasswordChar(ByVal strNewPasswordChar As String)
  txtEdicion(intIndexEdicion).PasswordChar = strNewPasswordChar
  PropertyChanged
End Property

Public Property Get MaxLength() As Integer
  MaxLength = txtEdicion(intIndexEdicion).MaxLength
End Property

Public Property Let MaxLength(ByVal intMaxLength As Integer)
  txtEdicion(intIndexEdicion).MaxLength = intMaxLength
  PropertyChanged
End Property

Public Property Get CaptionAlignment() As AlignmentConstants
  CaptionAlignment = lblMensaje.Alignment
End Property

Public Property Let CaptionAlignment(ByVal intNewAlignment As AlignmentConstants)
  lblMensaje.Alignment = intNewAlignment
  PropertyChanged
End Property

Public Property Get Multiline() As Boolean
  Multiline = blnMultiline
End Property

Public Property Let Multiline(ByVal blnNewMultiline As Boolean)
  blnMultiline = blnNewMultiline
  If blnMultiline Then
    intIndexEdicion = 1
  Else
    intIndexEdicion = 0
  End If
  'Y cambia los valores que ya hubiera
  With txtEdicion(intIndexEdicion)
    .BackColor = txtEdicion(1 - intIndexEdicion).BackColor
    .ForeColor = txtEdicion(1 - intIndexEdicion).ForeColor
    .PasswordChar = txtEdicion(1 - intIndexEdicion).PasswordChar
    .MaxLength = txtEdicion(1 - intIndexEdicion).MaxLength
    .Text = txtEdicion(1 - intIndexEdicion).Text
  End With
  'Redibuja el control
  UserControl_Resize
  'Y avisa de que se han cambiado las propiedades
  PropertyChanged
End Property

Public Property Get Numeric() As Boolean
  Numeric = blnNumerico
End Property

Public Property Let Numeric(ByVal blnNewNumeric As Boolean)
  blnNumerico = blnNewNumeric
  PropertyChanged
End Property

Public Property Get DecimalNumber() As Integer
  DecimalNumber = intNumDecimales
End Property

Public Property Let DecimalNumber(ByVal intNewDecimalNumber As Integer)
  intNumDecimales = intNewDecimalNumber
  PropertyChanged
End Property

Public Property Let SelectedBackColor(ByVal Color As OLE_COLOR)
  cSelectedBackColor = Color
  PropertyChanged
End Property

Public Property Get SelectedBackColor() As OLE_COLOR
  SelectedBackColor = cSelectedBackColor
End Property

Public Property Let UnselectedBackColor(ByVal Color As OLE_COLOR)
  cUnselectedBackColor = Color
  txtEdicion(intIndexEdicion).BackColor = cUnselectedBackColor
  PropertyChanged
End Property

Public Property Get UnselectedBackColor() As OLE_COLOR
  UnselectedBackColor = cUnselectedBackColor
End Property

Public Property Let SelectedForeColor(ByVal Color As OLE_COLOR)
  cSelectedForeColor = Color
  PropertyChanged
End Property

Public Property Get SelectedForeColor() As OLE_COLOR
  SelectedForeColor = cSelectedForeColor
End Property

Public Property Let UnselectedForeColor(ByVal Color As OLE_COLOR)
  cUnselectedForeColor = Color
  PropertyChanged
End Property

Public Property Get UnselectedForeColor() As OLE_COLOR
  UnselectedForeColor = cUnselectedForeColor
End Property

Public Property Let ForeColorCaption(ByVal Color As OLE_COLOR)
  lblMensaje.ForeColor = Color
  PropertyChanged
End Property

Public Property Get ForeColorCaption() As OLE_COLOR
  ForeColorCaption = lblMensaje.ForeColor
End Property

Public Property Let ForeColor(ByVal Color As OLE_COLOR)
  txtEdicion(0).ForeColor = Color
  txtEdicion(1).ForeColor = Color
  PropertyChanged
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = txtEdicion(0).ForeColor
End Property

Public Property Let BackColorCaption(ByVal Color As OLE_COLOR)
  lblMensaje.BackColor = Color
  PropertyChanged
End Property

Public Property Get BackColorCaption() As OLE_COLOR
  BackColorCaption = lblMensaje.BackColor
End Property

Public Property Let Locked(ByVal blnNewLock As Boolean)
  txtEdicion(intIndexEdicion).Locked = blnNewLock
  PropertyChanged
End Property

Public Property Get Locked() As Boolean
  Locked = txtEdicion(intIndexEdicion).Locked
End Property

Public Function Version() As String
  Version = App.Major & "." & App.Minor & "." & App.Revision
End Function

Private Sub txtEdicion_Change(Index As Integer)
  RaiseEvent Change
End Sub

Private Sub txtEdicion_Click(Index As Integer)
  RaiseEvent Click
End Sub


Private Sub txtEdicion_DblClick(Index As Integer)
  RaiseEvent DblClick
End Sub


Private Sub txtEdicion_GotFocus(Index As Integer)
  txtEdicion(Index).BackColor = cSelectedBackColor
  txtEdicion(Index).ForeColor = cSelectedForeColor
  lblMensaje_Click
  'DoEvents
  SelCampoEdicion txtEdicion(Index)
End Sub

Public Sub SelCampoEdicion(ByRef CampoEdicion As Object)
'--> Selecciona por completo un campo TextEdit, se utiliza para no tener que borrar el contenido _
     sino que directamente aparezca completamente seleccionado.
  On Error GoTo ErrorSeleccion
  
  If Len(CampoEdicion.Text) <> 0 And CampoEdicion.SelLength < Len(CampoEdicion.Text) Then
    CampoEdicion.SelStart = 0
    CampoEdicion.SelLength = Len(CampoEdicion.Text)
  End If
  
ErrorSeleccion:
  Exit Sub
  
End Sub

Private Sub txtEdicion_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
  
  'Exportamos el evento
  RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub txtEdicion_KeyPress(Index As Integer, KeyAscii As Integer)
  'Validamos si se ha pulsado enter
  If KeyAscii = 13 Then
    'Simulamos la pulsación de un tab, si procede.
    If EnterTab Then
      SendKeys "{TAB}", True
      'Anulamos esta presión de tecla para que no se ponga el enter en modo multilínea
      KeyAscii = 0
    End If
    'Enviamos el evento de pulsación de enter
    RaiseEvent EnterPress
  End If
  'Validamos resto de teclas
  If KeyAscii <> 8 Then
    If Numeric Then
      If (Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9") And Chr(KeyAscii) <> strCharDecimal Then
        KeyAscii = 0
      ElseIf Chr(KeyAscii) = strCharDecimal Then
        If InStr(txtEdicion(Index).Text, strCharDecimal) <> 0 Or Len(txtEdicion(Index).Text) = 0 Or DecimalNumber = 0 Then
          KeyAscii = 0
        End If
      ElseIf (CountNumeroDecimales(Index) >= DecimalNumber And DecimalNumber <> 0) And txtEdicion(Index).SelLength <> Len(txtEdicion(Index).Text) Then
        KeyAscii = 0
      End If
    End If
  End If
  'Exportamos el evento
  RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub txtEdicion_LostFocus(Index As Integer)
  txtEdicion(Index).BackColor = cUnselectedBackColor
  txtEdicion(Index).ForeColor = cUnselectedForeColor
  If Left(txtEdicion(Index).Text, 1) = strCharDecimal Then
    txtEdicion(Index).Text = "0," + Right(txtEdicion(Index).Text, Len(txtEdicion(Index).Text) - 1)
  End If
  UserControl_Resize
End Sub

Private Sub lblMensaje_Click()
  lblMensaje.Visible = False
  If blnMultiline Then
    txtEdicion(1).Visible = True
    txtEdicion(1).SetFocus
  Else
    txtEdicion(0).Visible = True
    txtEdicion(0).SetFocus
  End If
End Sub

Private Sub UserControl_GotFocus()
  lblMensaje_Click
End Sub

Private Sub UserControl_Initialize()
  strCharDecimal = Mid(Format(1.5, "0.0"), 2, 1)
  intIndexEdicion = 0
End Sub

Private Sub UserControl_InitProperties()
  Enabled = True
  Locked = False
  SelectedBackColor = &HC0FFFF
  SelectedForeColor = vbBlack
  UnselectedBackColor = vbWhite
  UnselectedForeColor = vbBlack
  ForeColorCaption = &H80000011
  BackColorCaption = vbWhite
  Multiline = False
  PasswordChar = ""
  Caption = ""
  MaxLength = 0
  Numeric = False
  DecimalNumber = 0
  m_EnterTab = m_def_EnterTab
End Sub

Private Sub UserControl_LostFocus()
'  changeCaption
End Sub

Private Sub UserControl_Paint()
  txtEdicion(0).ToolTipText = Extender.ToolTipText
  txtEdicion(1).ToolTipText = Extender.ToolTipText
  lblMensaje.ToolTipText = Extender.ToolTipText
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Enabled = PropBag.ReadProperty("Enabled", True)
  Locked = PropBag.ReadProperty("Locked", False)
  PasswordChar = PropBag.ReadProperty("PasswordChar", "")
  Multiline = PropBag.ReadProperty("Multiline", False)
  SelectedBackColor = PropBag.ReadProperty("SelectedBackColor", &HC0FFFF)
  UnselectedBackColor = PropBag.ReadProperty("UnselectedBackColor", vbWhite)
  SelectedForeColor = PropBag.ReadProperty("SelectedForeColor", vbBlack)
  UnselectedForeColor = PropBag.ReadProperty("UnselectedForeColor", vbBlack)
  ForeColorCaption = PropBag.ReadProperty("ForeColorCaption", &H80000011)
  ForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
  BackColorCaption = PropBag.ReadProperty("BackColorCaption", vbWhite)
  Caption = PropBag.ReadProperty("Caption", "")
  CaptionAlignment = PropBag.ReadProperty("CaptionAlignment", vbCenter)
  MaxLength = PropBag.ReadProperty("MaxLength", 0)
  Numeric = PropBag.ReadProperty("Numeric", False)
  DecimalNumber = PropBag.ReadProperty("DecimalNumber", 0)
  Set Font = PropBag.ReadProperty("Font", Ambient.Font)
  FontCaption = PropBag.ReadProperty("FontCaption", Ambient.Font)
  ToolTipText = PropBag.ReadProperty("ToolTipText", "")
  m_EnterTab = PropBag.ReadProperty("EnterTab", m_def_EnterTab)
End Sub

Private Sub UserControl_Resize()
  'On Error Resume Next
  lblMensaje.Visible = (lblMensaje.Caption <> "" And Trim(txtEdicion(intIndexEdicion).Text) = "")
  txtEdicion(intIndexEdicion).Visible = Not lblMensaje.Visible
  txtEdicion(1 - intIndexEdicion).Visible = False
  If Not blnMultiline Then
    Height = txtEdicion(intIndexEdicion).Height
  End If
  'Cambia las dimensiones del control label
    With lblMensaje
      .Top = ScaleTop
      .Left = ScaleLeft
      .Width = ScaleWidth
      .Height = ScaleHeight
    End With
  'Cambia las dimensiones del primer control de edición
    With txtEdicion(intIndexEdicion)
      .Top = ScaleTop
      .Left = ScaleLeft
      .Width = ScaleWidth
      .Height = ScaleHeight
    End With
'  changeCaption
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Enabled", Enabled, True
  PropBag.WriteProperty "Locked", Locked, False
  PropBag.WriteProperty "PasswordChar", PasswordChar, ""
  PropBag.WriteProperty "Multiline", Multiline, False
  PropBag.WriteProperty "SelectedBackColor", SelectedBackColor, &HC0FFFF
  PropBag.WriteProperty "UnselectedBackColor", UnselectedBackColor, vbWhite
  PropBag.WriteProperty "SelectedForeColor", SelectedForeColor, vbBlack
  PropBag.WriteProperty "UnselectedForeColor", UnselectedForeColor, vbBlack
  PropBag.WriteProperty "ForeColorCaption", ForeColorCaption, &H80000011
  PropBag.WriteProperty "ForeColor", ForeColor, vbBlack
  PropBag.WriteProperty "BackColorCaption", BackColorCaption, vbWhite
  PropBag.WriteProperty "Caption", Caption, ""
  PropBag.WriteProperty "CaptionAlignment", CaptionAlignment, vbCenter
  PropBag.WriteProperty "MaxLength", MaxLength, 0
  PropBag.WriteProperty "Numeric", Numeric, False
  PropBag.WriteProperty "DecimalNumber", DecimalNumber, 0
  PropBag.WriteProperty "Font", Font, Ambient.Font
  PropBag.WriteProperty "FontCaption", FontCaption, Ambient.Font
  PropBag.WriteProperty "ToolTipText", ToolTipText, ""
'  Call PropBag.WriteProperty("EnterTab", m_EnterTab, m_def_EnterTab)
  Call PropBag.WriteProperty("EnterTab", m_EnterTab, m_def_EnterTab)
End Sub

Public Property Get Font() As Font
  Set Font = txtEdicion(0).Font
End Property

Public Property Set Font(ByVal fntNewFont As Font)
  Set txtEdicion(0).Font = fntNewFont
  Set txtEdicion(1).Font = txtEdicion(0).Font
  UserControl_Resize
  PropertyChanged
End Property

Public Property Get FontCaption() As Font
  Set FontCaption = lblMensaje.Font
End Property

Public Property Set FontCaption(ByVal fntNewFont As Font)
  Set lblMensaje.Font = fntNewFont
  PropertyChanged
End Property

Public Property Get ToolTipText() As String
  ToolTipText = txtEdicion(0).ToolTipText
End Property

Public Property Let ToolTipText(ByVal strNewToolTipText As String)
  txtEdicion(0).ToolTipText() = strNewToolTipText
  txtEdicion(1).ToolTipText() = strNewToolTipText
  lblMensaje.ToolTipText() = strNewToolTipText
  PropertyChanged ToolTipText
End Property
'
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,true
Public Property Get EnterTab() As Boolean
  EnterTab = m_EnterTab
End Property

Public Property Let EnterTab(ByVal New_EnterTab As Boolean)
  m_EnterTab = New_EnterTab
  PropertyChanged "EnterTab"
End Property


Public Property Let Text(ByVal strNewText As String)
Attribute Text.VB_UserMemId = 0
  txtEdicion(intIndexEdicion).Text = Trim(strNewText)
  UserControl_Resize
End Property

Public Property Get Text() As String
  Dim strValor As String

  strValor = Trim(txtEdicion(intIndexEdicion))
  Text = strValor
End Property

