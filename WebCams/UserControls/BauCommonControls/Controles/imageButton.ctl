VERSION 5.00
Begin VB.UserControl ImageButton 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5595
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   373
   ToolboxBitmap   =   "imageButton.ctx":0000
   Begin VB.PictureBox PicMain 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   240
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   300
      Width           =   480
   End
   Begin VB.PictureBox PicDown 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1620
      Picture         =   "imageButton.ctx":00FA
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   3
      Tag             =   "2"
      Top             =   1980
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicOver 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   840
      Picture         =   "imageButton.ctx":01FC
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Tag             =   "2"
      Top             =   1980
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox PicNormal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   60
      Picture         =   "imageButton.ctx":02FE
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      Tag             =   "2"
      Top             =   1980
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Line LineRight2 
      BorderColor     =   &H80000010&
      X1              =   128
      X2              =   128
      Y1              =   1
      Y2              =   88
   End
   Begin VB.Line LineRight1 
      BorderColor     =   &H80000015&
      X1              =   156
      X2              =   156
      Y1              =   0
      Y2              =   60
   End
   Begin VB.Line LineTop1 
      BorderColor     =   &H80000016&
      X1              =   0
      X2              =   104
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line LineTop2 
      BorderColor     =   &H80000014&
      X1              =   1
      X2              =   88
      Y1              =   1
      Y2              =   1
   End
   Begin VB.Line LineDown1 
      BorderColor     =   &H80000015&
      X1              =   0
      X2              =   68
      Y1              =   116
      Y2              =   116
   End
   Begin VB.Line LineDown2 
      BorderColor     =   &H80000010&
      X1              =   1
      X2              =   112
      Y1              =   88
      Y2              =   88
   End
   Begin VB.Line LineLeft2 
      BorderColor     =   &H80000014&
      X1              =   1
      X2              =   1
      Y1              =   1
      Y2              =   100
   End
   Begin VB.Line LineLeft1 
      BorderColor     =   &H80000016&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   80
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Image Button"
      Height          =   195
      Left            =   780
      TabIndex        =   0
      Top             =   900
      Width           =   945
   End
End
Attribute VB_Name = "ImageButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==================================================================================
'   Permet de faire énormement de bouton differents
'   grace aux proprietes differentes pour les 3 etats du boutons
'   NormalPosition
'   UpPosition
'   DownPosition
'
'   si vous trouvez un BUG ou que vous enrichisser cet OCX
'   merci de me faire parvenir vos améliorations et/ou commentaires
'
'   Si un effet vous manque encore avec cet OCX
'   écrivez moi on verra ca que l'on peut faire ;-)
'
'   adresse en cours    : fred.just@free.fr
'   site actuel         : http://fred.just.free.fr/
'   adresse de secours   : fredjust@hotmail.com
'
'==================================================================================

Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

' differente position du bouton
Enum EnumPosition
    down
    Normal
    up
End Enum

' differente epaiseur de l'ombre
Enum EnumStyleLine
    zero
    one
    Two
End Enum

' diferente position du caption par rapport à l'image
Enum EnumLabelPosition
    lpBottom
    lpLeft
    lpRight
    lpTop
    lpFree
End Enum

' style du bouton
Public Enum EnumStyle
    Standard
    Switch
End Enum

'Valeurs de propriétés par défaut:
Const m_def_CaptionPictureSpace = 0
Const m_def_NbLine = 2
Const m_def_NbLineOver = 2
Const m_def_NbLineDown = 2
Const m_def_ButtonStyle = 0
Const m_def_CaptionPosition = 0
Const m_def_ImageTop = 0
Const m_def_ImageLeft = 0
Const m_def_LabelTop = 0
Const m_def_LabelLeft = 0
Const m_def_Position = 0
'Variables de propriétés:
Dim m_CaptionPictureSpace As Long
Dim m_NbLine As EnumStyleLine
Dim m_NbLineOver As EnumStyleLine
Dim m_NbLineDown As EnumStyleLine
'Dim m_ButtonStyle As EnumStyleLine
Dim m_CaptionPosition As EnumLabelPosition
Dim m_ImageTop As Long
Dim m_ImageLeft As Long
Dim m_LabelTop As Long
Dim m_LabelLeft As Long
Dim m_position As EnumPosition
Dim posActu As EnumPosition
Dim m_Style As EnumStyle
Dim CaptureIcon As Boolean
Dim MousePointerDefault As Long


'Déclarations d'événements:
Event Click()
Attribute Click.VB_Description = "Se produit lorsque l'utilisateur appuie sur un bouton de la souris puis le relâche au-dessus d'un objet."
Event RightMouseUP()
Attribute RightMouseUP.VB_Description = "Sans Commentaire"
Event EnterButton()
Attribute EnterButton.VB_Description = "Se produit lorsque le curseur est sur le bouton"
Event ExitButton()
Attribute ExitButton.VB_Description = "Se produit lorsque le curseur quitte le bouton"

'==================================================================================
'
'==================================================================================
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCapture PicNormal.hWnd
End Sub

Private Sub PicMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCapture PicNormal.hWnd
End Sub

Private Sub PicNormal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then DOWNposition
    If Button = 2 And Shift = 7 Then About
End Sub

'==================================================================================
'
'==================================================================================
Private Sub PicNormal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Xi, Yi As Single
Dim xyCursor As POINTAPI

    RaiseEvent EnterButton
    
    If Not CaptureIcon Then
        MousePointerDefault = Screen.MousePointer
        CaptureIcon = True
    End If
    
    GetCursorPos xyCursor
    ScreenToClient UserControl.hWnd, xyCursor
    
    Xi = xyCursor.X
    Yi = xyCursor.Y
    
    If Xi < 0 Or Yi < 0 Or Xi > UserControl.ScaleWidth Or Yi > UserControl.ScaleHeight Then     'si on sort du contrôle
        
          RaiseEvent ExitButton
          CaptureIcon = False
          Screen.MousePointer = MousePointerDefault
          
        ReleaseCapture
        If m_Style = Standard Then
            If posActu <> Normal Then
                NormalPosition
                m_position = Normal
            End If
        Else
            Select Case m_position
                Case Normal
                    If posActu <> Normal Then
                        NormalPosition
                        m_position = Normal
                    End If
                Case up
                    If posActu <> up Then
                        UPposition
                        m_position = up
                    End If
                Case down
                    If posActu <> down Then
                        DOWNposition
                        m_position = down
                    End If
            End Select
        End If
    Else
        Screen.MousePointer = 99
        Set Screen.MouseIcon = PicNormal.MouseIcon
        If m_Style = Standard Then
            If Button <> 1 Then
                If posActu <> up Then UPposition
            Else
                If posActu <> down Then DOWNposition
            End If
        Else
            If Button <> 1 Then
                If m_position <> down Then
                  If posActu <> up Then UPposition
                Else
                    If posActu <> down Then DOWNposition
                End If
            End If
        End If
    
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub PicNormal_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If m_Style = Standard Then
            UPposition
            m_position = up
            SetCapture PicNormal.hWnd
            RaiseEvent Click
        Else
            If m_position = down Then
                NormalPosition
                m_position = Normal
            Else
                DOWNposition
                m_position = down
            End If
            
        End If
    Else
        RaiseEvent RightMouseUP
    End If
End Sub

'==================================================================================
'
'==================================================================================
Private Sub UserControl_Initialize()
    PicDown.FillColor = vb3DDKShadow
    PicOver.FillColor = vb3DLight
    PicNormal.FillColor = vb3DHighlight
    PicMain.FillColor = vb3DShadow
    CalculPosition
    NormalPosition
End Sub

'==================================================================================
'
'==================================================================================
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SetCapture PicNormal.hWnd
End Sub

'==================================================================================
'
'==================================================================================
Private Sub UserControl_Resize()
    CalculPosition
    
    LineTop1.X2 = UserControl.ScaleWidth - 1
    LineTop2.X2 = UserControl.ScaleWidth - 2
    
    LineLeft1.Y2 = UserControl.ScaleHeight - 1
    LineLeft2.Y2 = UserControl.ScaleHeight - 2
    
    LineDown1.Y1 = UserControl.ScaleHeight - 1
    LineDown1.Y2 = UserControl.ScaleHeight - 1
    LineDown1.X2 = UserControl.ScaleWidth
    
    LineDown2.Y1 = UserControl.ScaleHeight - 2
    LineDown2.Y2 = UserControl.ScaleHeight - 2
    LineDown2.X2 = UserControl.ScaleWidth - 2
    
    LineRight1.X1 = UserControl.ScaleWidth - 1
    LineRight1.X2 = UserControl.ScaleWidth - 1
    LineRight1.Y2 = UserControl.ScaleHeight
    
    LineRight2.X1 = UserControl.ScaleWidth - 2
    LineRight2.X2 = UserControl.ScaleWidth - 2
    LineRight2.Y2 = UserControl.ScaleHeight - 1
    
End Sub

'==================================================================================
'
'==================================================================================
Private Sub CalculPosition()
    Select Case m_CaptionPosition
        Case lpBottom
            m_ImageTop = (UserControl.ScaleHeight - (PicMain.Height + Label1.Height + m_CaptionPictureSpace)) \ 2
            PicMain.Top = m_ImageTop
            m_ImageLeft = (UserControl.ScaleWidth - PicMain.Width) \ 2
            PicMain.Left = m_ImageLeft
            m_LabelTop = PicMain.Top + PicMain.Height + m_CaptionPictureSpace
            Label1.Top = m_LabelTop
            m_LabelLeft = (UserControl.ScaleWidth - Label1.Width) \ 2
            Label1.Left = m_LabelLeft
        Case lpTop
            m_ImageTop = (UserControl.ScaleHeight - (PicMain.Height + Label1.Height + m_CaptionPictureSpace)) \ 2 + m_CaptionPictureSpace + Label1.Height
            PicMain.Top = m_ImageTop
            m_ImageLeft = (UserControl.ScaleWidth - PicMain.Width) \ 2
            PicMain.Left = m_ImageLeft
            m_LabelTop = PicMain.Top - Label1.Height - m_CaptionPictureSpace
            Label1.Top = m_LabelTop
            m_LabelLeft = (UserControl.ScaleWidth - Label1.Width) \ 2
            Label1.Left = m_LabelLeft
        Case lpLeft
            m_ImageTop = (UserControl.ScaleHeight - PicMain.Height) \ 2
            PicMain.Top = m_ImageTop
            m_ImageLeft = (UserControl.ScaleWidth - (PicMain.Width + Label1.Width + m_CaptionPictureSpace)) \ 2
            PicMain.Left = m_ImageLeft
            m_LabelTop = (UserControl.ScaleHeight - Label1.Height) \ 2
            Label1.Top = m_LabelTop
            m_LabelLeft = PicMain.Left + PicMain.Width + m_CaptionPictureSpace
            Label1.Left = m_LabelLeft
        Case lpRight
            m_ImageTop = (UserControl.ScaleHeight - PicMain.Height) \ 2
            PicMain.Top = m_ImageTop
            m_ImageLeft = (UserControl.ScaleWidth - (PicMain.Width + Label1.Width + m_CaptionPictureSpace)) \ 2 + m_CaptionPictureSpace + Label1.Width
            PicMain.Left = m_ImageLeft
            m_LabelTop = (UserControl.ScaleHeight - Label1.Height) \ 2
            Label1.Top = m_LabelTop
            m_LabelLeft = PicMain.Left - Label1.Width - m_CaptionPictureSpace
            Label1.Left = m_LabelLeft
        Case lpFree
            PicMain.Top = m_ImageTop
            PicMain.Left = m_ImageLeft
            Label1.Top = m_LabelTop
            Label1.Left = m_LabelLeft
    End Select

End Sub

'==================================================================================
'
'==================================================================================
Public Sub ForceDownPosition()
    m_position = down
    DOWNposition
End Sub

Public Sub ForceNormalPosition()
    m_position = Normal
    NormalPosition
End Sub

Public Sub ForceUpPosition()
    m_position = up
    UPposition
End Sub
'==================================================================================
'
'==================================================================================
Public Sub DOWNposition()
Attribute DOWNposition.VB_Description = "Place le bouton en position basse"

    PicMain.Picture = PicDown.Image
    
    Set Label1.Font = PicDown.Font
    
    If PicDown.Picture = 0 Then
        PicMain.Height = 0
        PicMain.Width = 0
    End If
    
    CalculPosition
    
    PicMain.Top = m_ImageTop + 1
    Label1.Top = m_LabelTop + 1
    
    Label1.Left = m_LabelLeft + 1
    PicMain.Left = m_ImageLeft + 1
    
    
    
    UserControl.BackColor = PicDown.BackColor
    Label1.BackColor = PicDown.BackColor
    Label1.ForeColor = PicDown.ForeColor
 
    LineDown1.Visible = m_NbLineDown > one
    LineDown2.Visible = m_NbLineDown > zero
    
    LineLeft1.Visible = m_NbLineDown > one
    LineLeft2.Visible = m_NbLineDown > zero
    
    LineRight1.Visible = m_NbLineDown > one
    LineRight2.Visible = m_NbLineDown > zero
    
    LineTop1.Visible = m_NbLineDown > one
    LineTop2.Visible = m_NbLineDown > zero
        
    
    LineLeft1.BorderColor = PicDown.FillColor
    LineTop1.BorderColor = PicDown.FillColor

    LineLeft2.BorderColor = PicMain.FillColor
    LineTop2.BorderColor = PicMain.FillColor

    LineDown1.BorderColor = PicNormal.FillColor
    LineRight1.BorderColor = PicNormal.FillColor

    LineDown2.BorderColor = PicOver.FillColor
    LineRight2.BorderColor = PicOver.FillColor
    posActu = down
End Sub

'==================================================================================
'
'==================================================================================
Public Sub NormalPosition()
Attribute NormalPosition.VB_Description = "Place le bouton en position Normal"
    PicMain.Picture = PicNormal.Image
    
    If PicNormal.Picture = 0 Then
        PicMain.Height = 0
        PicMain.Width = 0
    End If
    
    CalculPosition
    
    PicMain.Top = m_ImageTop
    Label1.Top = m_LabelTop
    
    Label1.Left = m_LabelLeft
    PicMain.Left = m_ImageLeft
    
    Set Label1.Font = PicNormal.Font
    
    
    
    UserControl.BackColor = PicNormal.BackColor
    Label1.BackColor = PicNormal.BackColor
    Label1.ForeColor = PicNormal.ForeColor

    LineDown1.Visible = m_NbLine > one
    LineDown2.Visible = m_NbLine > zero
    
    LineLeft1.Visible = m_NbLine > one
    LineLeft2.Visible = m_NbLine > zero
    
    LineRight1.Visible = m_NbLine > one
    LineRight2.Visible = m_NbLine > zero
    
    LineTop1.Visible = m_NbLine > one
    LineTop2.Visible = m_NbLine > zero

    LineLeft1.BorderColor = PicNormal.FillColor
    LineTop1.BorderColor = PicNormal.FillColor

    LineLeft2.BorderColor = PicOver.FillColor
    LineTop2.BorderColor = PicOver.FillColor

    LineDown1.BorderColor = PicDown.FillColor
    LineRight1.BorderColor = PicDown.FillColor

    LineDown2.BorderColor = PicMain.FillColor
    LineRight2.BorderColor = PicMain.FillColor
    posActu = Normal

End Sub

'==================================================================================
'
'==================================================================================
Public Sub UPposition()
Attribute UPposition.VB_Description = "Place le bouton en position haute"

    PicMain.Picture = PicOver.Image
    
    If PicOver.Picture = 0 Then
        PicMain.Height = 0
        PicMain.Width = 0
    End If
    
    CalculPosition
    
    PicMain.Top = m_ImageTop
    Label1.Top = m_LabelTop
    
    Label1.Left = m_LabelLeft
    PicMain.Left = m_ImageLeft
    
    Set Label1.Font = PicOver.Font
    
    
    
    UserControl.BackColor = PicOver.BackColor
    Label1.BackColor = PicOver.BackColor
    Label1.ForeColor = PicOver.ForeColor
    
    LineDown1.Visible = m_NbLineOver > one
    LineDown2.Visible = m_NbLineOver > zero
    
    LineLeft1.Visible = m_NbLineOver > one
    LineLeft2.Visible = m_NbLineOver > zero
    
    LineRight1.Visible = m_NbLineOver > one
    LineRight2.Visible = m_NbLineOver > zero
    
    LineTop1.Visible = m_NbLineOver > one
    LineTop2.Visible = m_NbLineOver > zero
    
    LineLeft1.BorderColor = PicNormal.FillColor
    LineTop1.BorderColor = PicNormal.FillColor

    LineLeft2.BorderColor = PicOver.FillColor
    LineTop2.BorderColor = PicOver.FillColor

    LineDown1.BorderColor = PicDown.FillColor
    LineRight1.BorderColor = PicDown.FillColor

    LineDown2.BorderColor = PicMain.FillColor
    LineRight2.BorderColor = PicMain.FillColor
    posActu = up
    
End Sub

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=PicNormal,PicNormal,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Renvoie ou définit la couleur d'arrière-plan "
    BackColor = PicNormal.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    PicNormal.BackColor() = New_BackColor
    ActualisePosition
    PropertyChanged "BackColor"
End Property

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=Label1,Label1,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Texte du bouton"
    Caption = Label1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Label1.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=PicNormal,PicNormal,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Couleur du texte en position normal"
    ForeColor = PicNormal.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    PicNormal.ForeColor() = New_ForeColor
    ActualisePosition
    PropertyChanged "ForeColor"
End Property

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=PicNormal,PicNormal,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Image du bouton"
    Set Picture = PicNormal.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set PicNormal.Picture = New_Picture
    If New_Picture Is Nothing Then
        PicNormal.Width = 0
        PicNormal.Height = 0
    Else
        PicMain.Picture = PicNormal.Image
        PicMain.Width = PicNormal.Width
        PicMain.Height = PicNormal.Height
    End If
    ActualisePosition
    PropertyChanged "Picture"
End Property

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=PicOver,PicOver,-1,Picture
Public Property Get PictureOver() As Picture
Attribute PictureOver.VB_Description = "Image du bouton lorsque la souris est au dessus"
    Set PictureOver = PicOver.Picture
End Property

'==================================================================================
'
'==================================================================================
Public Property Set PictureOver(ByVal New_PictureOver As Picture)
    Set PicOver.Picture = New_PictureOver
    ActualisePosition
    PropertyChanged "PictureOver"
End Property

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=PicDown,PicDown,-1,Picture
Public Property Get PictureDown() As Picture
Attribute PictureDown.VB_Description = "Image du bouton lorsque le bouton est enfoncé"
    Set PictureDown = PicDown.Picture
End Property

'==================================================================================
'
'==================================================================================
Public Property Set PictureDown(ByVal New_PictureDown As Picture)
    Set PicDown.Picture = New_PictureDown
    ActualisePosition
    PropertyChanged "PictureDown"
End Property

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=14,0,0,0
Public Property Get CaptionPosition() As EnumLabelPosition
Attribute CaptionPosition.VB_Description = "Position du texte par rapport à l' image"
    CaptionPosition = m_CaptionPosition
End Property

Public Property Let CaptionPosition(ByVal New_CaptionPosition As EnumLabelPosition)
    m_CaptionPosition = New_CaptionPosition
    ActualisePosition
    PropertyChanged "CaptionPosition"
End Property

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=PicOver,PicOver,-1,BackColor
Public Property Get BackColorOver() As OLE_COLOR
Attribute BackColorOver.VB_Description = "Renvoie ou définit la couleur d'arrière-plan en position haute"
    BackColorOver = PicOver.BackColor
End Property

Public Property Let BackColorOver(ByVal New_BackColorOver As OLE_COLOR)
    PicOver.BackColor() = New_BackColorOver
    ActualisePosition
    PropertyChanged "BackColorOver"
End Property

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=PicDown,PicDown,-1,BackColor
Public Property Get BackColorDown() As OLE_COLOR
Attribute BackColorDown.VB_Description = "Renvoie ou définit la couleur d'arrière-plan en position basse"
    BackColorDown = PicDown.BackColor
End Property

Public Property Let BackColorDown(ByVal New_BackColorDown As OLE_COLOR)
    PicDown.BackColor() = New_BackColorDown
    ActualisePosition
    PropertyChanged "BackColorDown"
End Property

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=PicOver,PicOver,-1,ForeColor
Public Property Get ForeColorOver() As OLE_COLOR
Attribute ForeColorOver.VB_Description = "Couleur du texte en position haute"
    ForeColorOver = PicOver.ForeColor
End Property

Public Property Let ForeColorOver(ByVal New_ForeColorOver As OLE_COLOR)
    PicOver.ForeColor() = New_ForeColorOver
    ActualisePosition
    PropertyChanged "ForeColorOver"
End Property

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=PicDown,PicDown,-1,ForeColor
Public Property Get ForeColorDown() As OLE_COLOR
Attribute ForeColorDown.VB_Description = "Couleur du texte en position basse"
    ForeColorDown = PicDown.ForeColor
End Property

Public Property Let ForeColorDown(ByVal New_ForeColorDown As OLE_COLOR)
    PicDown.ForeColor() = New_ForeColorDown
    ActualisePosition
    PropertyChanged "ForeColorDown"
End Property

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=8,0,0,0
Public Property Get ImageTop() As Long
Attribute ImageTop.VB_Description = "Position de l'image ( CaptionPosition doit etre Free )"
    ImageTop = m_ImageTop
End Property

Public Property Let ImageTop(ByVal New_ImageTop As Long)
    m_ImageTop = New_ImageTop
    ActualisePosition
    PropertyChanged "ImageTop"
End Property

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=8,0,0,0
Public Property Get ImageLeft() As Long
Attribute ImageLeft.VB_Description = "Position de l'image ( CaptionPosition doit etre Free )"
    ImageLeft = m_ImageLeft
End Property

Public Property Let ImageLeft(ByVal New_ImageLeft As Long)
    m_ImageLeft = New_ImageLeft
    ActualisePosition
    PropertyChanged "ImageLeft"
End Property

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=8,0,0,0
Public Property Get LabelTop() As Long
Attribute LabelTop.VB_Description = "Position du texte ( CaptionPosition doit etre Free )"
    LabelTop = m_LabelTop
End Property

Public Property Let LabelTop(ByVal New_LabelTop As Long)
    m_LabelTop = New_LabelTop
    ActualisePosition
    PropertyChanged "LabelTop"
End Property

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=8,0,0,0
Public Property Get LabelLeft() As Long
Attribute LabelLeft.VB_Description = "Position du texte ( CaptionPosition doit etre Free )"
    LabelLeft = m_LabelLeft
End Property

Public Property Let LabelLeft(ByVal New_LabelLeft As Long)
    m_LabelLeft = New_LabelLeft
    ActualisePosition
    PropertyChanged "LabelLeft"
End Property

'==================================================================================
'
'==================================================================================
'Initialiser les propriétés pour le contrôle utilisateur
Private Sub UserControl_InitProperties()
    'Debug.Print "UserControl_InitProperties " & Time
    m_CaptionPosition = m_def_CaptionPosition
    m_ImageTop = m_def_ImageTop
    m_ImageLeft = m_def_ImageLeft
    m_LabelTop = m_def_LabelTop
    m_LabelLeft = m_def_LabelLeft
    PicMain.Picture = PicNormal.Picture
    m_NbLine = m_def_NbLine
    m_NbLineOver = m_def_NbLineOver
    m_NbLineDown = m_def_NbLineDown
    m_CaptionPictureSpace = m_def_CaptionPictureSpace
    m_position = Normal
    ActualisePosition
End Sub

'==================================================================================
'
'==================================================================================
'Charger les valeurs des propriétés à partir du stockage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    PicNormal.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Label1.Caption = PropBag.ReadProperty("Caption", "MyButton")
    PicNormal.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
    
    Set PicNormal.Picture = PropBag.ReadProperty("Picture", Nothing)
    Set PicOver.Picture = PropBag.ReadProperty("PictureOver", Nothing)
    Set PicDown.Picture = PropBag.ReadProperty("PictureDown", Nothing)
    
    m_CaptionPosition = PropBag.ReadProperty("CaptionPosition", m_def_CaptionPosition)
    
    PicOver.BackColor = PropBag.ReadProperty("BackColorOver", &H80000005)
    PicDown.BackColor = PropBag.ReadProperty("BackColorDown", &H80000005)
    PicOver.ForeColor = PropBag.ReadProperty("ForeColorOver", &H80000008)
    PicDown.ForeColor = PropBag.ReadProperty("ForeColorDown", &H80000008)
    
    m_ImageTop = PropBag.ReadProperty("ImageTop", m_def_ImageTop)
    m_ImageLeft = PropBag.ReadProperty("ImageLeft", m_def_ImageLeft)
    m_LabelTop = PropBag.ReadProperty("LabelTop", m_def_LabelTop)
    m_LabelLeft = PropBag.ReadProperty("LabelLeft", m_def_LabelLeft)
    
    m_NbLine = PropBag.ReadProperty("NbLine", m_def_NbLine)
    m_NbLineOver = PropBag.ReadProperty("NbLineOver", m_def_NbLineOver)
    m_NbLineDown = PropBag.ReadProperty("NbLineDown", m_def_NbLineDown)
    
    m_CaptionPictureSpace = PropBag.ReadProperty("CaptionPictureSpace", m_def_CaptionPictureSpace)
    
    PicNormal.FillColor = PropBag.ReadProperty("ColorLineUpLeftOne", &H0&)
    PicOver.FillColor = PropBag.ReadProperty("ColorLineUpLeftTwo", &H0&)
    PicDown.FillColor = PropBag.ReadProperty("ColorLineDownRightOne", &H0&)
    PicMain.FillColor = PropBag.ReadProperty("ColorLineDownRightTwo", &H0&)
    
    
    Set PicOver.Font = PropBag.ReadProperty("FontOver", Ambient.Font)
    Set PicDown.Font = PropBag.ReadProperty("FontDown", Ambient.Font)
    Set PicNormal.Font = PropBag.ReadProperty("Font", Ambient.Font)
    
    m_Style = PropBag.ReadProperty("Style", 0)
    m_position = PropBag.ReadProperty("Position", 1)
    
    ActualisePosition
    Set PicNormal.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
End Sub

'==================================================================================
'
'==================================================================================
'Écrire les valeurs des propriétés dans le stockage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    Call PropBag.WriteProperty("BackColor", PicNormal.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Caption", Label1.Caption, "MyButton")
    Call PropBag.WriteProperty("ForeColor", PicNormal.ForeColor, &H80000008)
    
    Call PropBag.WriteProperty("Picture", PicNormal.Picture, Nothing)
    Call PropBag.WriteProperty("PictureOver", PicOver.Picture, Nothing)
    Call PropBag.WriteProperty("PictureDown", PicDown.Picture, Nothing)
    
    
    Call PropBag.WriteProperty("CaptionPosition", m_CaptionPosition, m_def_CaptionPosition)
    
    Call PropBag.WriteProperty("BackColorOver", PicOver.BackColor, &H80000005)
    Call PropBag.WriteProperty("BackColorDown", PicDown.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColorOver", PicOver.ForeColor, &H80000008)
    Call PropBag.WriteProperty("ForeColorDown", PicDown.ForeColor, &H80000008)
    
    Call PropBag.WriteProperty("ImageTop", m_ImageTop, m_def_ImageTop)
    Call PropBag.WriteProperty("ImageLeft", m_ImageLeft, m_def_ImageLeft)
    Call PropBag.WriteProperty("LabelTop", m_LabelTop, m_def_LabelTop)
    Call PropBag.WriteProperty("LabelLeft", m_LabelLeft, m_def_LabelLeft)
    
    Call PropBag.WriteProperty("NbLine", m_NbLine, m_def_NbLine)
    Call PropBag.WriteProperty("NbLineOver", m_NbLineOver, m_def_NbLineOver)
    Call PropBag.WriteProperty("NbLineDown", m_NbLineDown, m_def_NbLineDown)

    Call PropBag.WriteProperty("CaptionPictureSpace", m_CaptionPictureSpace, m_def_CaptionPictureSpace)

    Call PropBag.WriteProperty("ColorLineUpLeftOne", PicNormal.FillColor, &H0&)
    Call PropBag.WriteProperty("ColorLineUpLeftTwo", PicOver.FillColor, &H0&)
    Call PropBag.WriteProperty("ColorLineDownRightOne", PicDown.FillColor, &H0&)
    Call PropBag.WriteProperty("ColorLineDownRightTwo", PicMain.FillColor, &H0&)
    
    Call PropBag.WriteProperty("FontOver", PicOver.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontDown", PicDown.Font, Ambient.Font)
    Call PropBag.WriteProperty("Font", PicNormal.Font, Ambient.Font)
    
    Call PropBag.WriteProperty("Style", m_Style, 0)
    Call PropBag.WriteProperty("Position", m_position, 1)
    ActualisePosition
    Call PropBag.WriteProperty("MouseIcon", PicNormal.MouseIcon, Nothing)
End Sub

'==================================================================================
'
'==================================================================================
Public Property Get Position() As EnumPosition
Attribute Position.VB_Description = "Renvoie la position du bouton"
    Position = m_position
End Property


Public Property Let Position(ByVal New_position As EnumPosition)
    m_position = New_position
    ActualisePosition
    PropertyChanged "NbLine"
End Property

'==================================================================================
'
'==================================================================================
Private Sub ActualisePosition()
    Select Case m_position
        Case up
            UPposition
        Case down
            DOWNposition
        Case Normal
            NormalPosition
    End Select
End Sub

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=23,0,0,1
Public Property Get NbLine() As EnumStyleLine
Attribute NbLine.VB_Description = "Epaisseur du bord en position standard"
    NbLine = m_NbLine
End Property

Public Property Let NbLine(ByVal New_NbLine As EnumStyleLine)
    m_NbLine = New_NbLine
    ActualisePosition
    PropertyChanged "NbLine"
End Property

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=23,0,0,2
Public Property Get NbLineOver() As EnumStyleLine
Attribute NbLineOver.VB_Description = "Epaisseur du bord en position relevée"
    NbLineOver = m_NbLineOver
End Property

Public Property Let NbLineOver(ByVal New_NbLineOver As EnumStyleLine)
    m_NbLineOver = New_NbLineOver
    ActualisePosition
    PropertyChanged "NbLineOver"
End Property

'==================================================================================
'
'==================================================================================
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MemberInfo=23,0,0,1
Public Property Get NbLineDown() As EnumStyleLine
Attribute NbLineDown.VB_Description = "Epaisseur du bord en position  Basse"
    NbLineDown = m_NbLineDown
End Property

Public Property Let NbLineDown(ByVal New_NbLineDown As EnumStyleLine)
    m_NbLineDown = New_NbLineDown
    ActualisePosition
    PropertyChanged "NbLineDown"
End Property


'==================================================================================
'
'==================================================================================
Private Function GreyScale(ByVal Colr As Long) As Integer
' Takes a long integer color value and converts it'
' to an equivalent grayscale value between 0 and 255

Dim R As Long, G As Long, B As Long

    R = Colr Mod 256
    Colr = Colr \ 256
    G = Colr Mod 256
    Colr = Colr \ 256
    B = Colr Mod 256
    
    GreyScale = 76 * R / 255 + 150 * G / 255 + 28 * B / 255

End Function
'==================================================================================
'
'==================================================================================
Public Sub ImageGrayScale(Image As EnumPosition)
Attribute ImageGrayScale.VB_Description = "transforme une image en noir et blanc"
Dim PIC As PictureBox
Dim X As Long
Dim Y As Long
Dim tempo As Long

    Select Case Image
        Case up
            Set PIC = PicOver
        Case down
            Set PIC = PicDown
        Case Normal
            Set PIC = PicNormal
    End Select
    
    For X = 0 To PIC.Width
        For Y = 0 To PIC.Height
            tempo = GreyScale(PIC.Point(X, Y))
            PIC.PSet (X, Y), RGB(tempo, tempo, tempo)
        Next Y
    Next X
    ActualisePosition

End Sub

'
'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=PicNormal,PicNormal,-1,FillColor
Public Property Get ColorLineUpLeftOne() As OLE_COLOR
Attribute ColorLineUpLeftOne.VB_Description = "Couleur des  Lignes exterieures du haut et de gauche"
    ColorLineUpLeftOne = PicNormal.FillColor
End Property

Public Property Let ColorLineUpLeftOne(ByVal New_ColorLineUpLeftOne As OLE_COLOR)
    PicNormal.FillColor() = New_ColorLineUpLeftOne
    ActualisePosition
    PropertyChanged "ColorLineUpLeftOne"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=PicOver,PicOver,-1,FillColor
Public Property Get ColorLineUpLeftTwo() As OLE_COLOR
Attribute ColorLineUpLeftTwo.VB_Description = "Couleur des  Lignes interieures du haut et de gauche"
    ColorLineUpLeftTwo = PicOver.FillColor
End Property

Public Property Let ColorLineUpLeftTwo(ByVal New_ColorLineUpLeftTwo As OLE_COLOR)
    PicOver.FillColor() = New_ColorLineUpLeftTwo
    ActualisePosition
    PropertyChanged "ColorLineUpLeftTwo"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=PicDown,PicDown,-1,FillColor
Public Property Get ColorLineDownRightOne() As OLE_COLOR
Attribute ColorLineDownRightOne.VB_Description = "Couleur des  Lignes exterieures du bas et de droite"
    ColorLineDownRightOne = PicDown.FillColor
End Property

Public Property Let ColorLineDownRightOne(ByVal New_ColorLineDownRightOne As OLE_COLOR)
    PicDown.FillColor() = New_ColorLineDownRightOne
    ActualisePosition
    PropertyChanged "ColorLineDownRightOne"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=PicMain,PicMain,-1,FillColor
Public Property Get ColorLineDownRightTwo() As OLE_COLOR
Attribute ColorLineDownRightTwo.VB_Description = "Couleur des  Lignes interieures du bas et de droite"
    ColorLineDownRightTwo = PicMain.FillColor
End Property

Public Property Let ColorLineDownRightTwo(ByVal New_ColorLineDownRightTwo As OLE_COLOR)
    PicMain.FillColor() = New_ColorLineDownRightTwo
    ActualisePosition
    PropertyChanged "ColorLineDownRightTwo"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=PicNormal,PicNormal,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Police utilisée pour le bouton"
Attribute Font.VB_UserMemId = -512
    Set Font = PicNormal.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set PicNormal.Font = New_Font
    CalculPosition
    ActualisePosition
    PropertyChanged "Font"
End Property


'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=PicOver,PicOver,-1,Font
Public Property Get FontOver() As Font
Attribute FontOver.VB_Description = "Renvoie un objet Font."
    Set FontOver = PicOver.Font
End Property

Public Property Set FontOver(ByVal New_FontOver As Font)
    Set PicOver.Font = New_FontOver
    ActualisePosition
    PropertyChanged "FontOver"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=PicDown,PicDown,-1,Font
Public Property Get FontDown() As Font
Attribute FontDown.VB_Description = "Renvoie un objet Font."
    Set FontDown = PicDown.Font
End Property

Public Property Set FontDown(ByVal New_FontDown As Font)
    Set PicDown.Font = New_FontDown
    ActualisePosition
    PropertyChanged "FontDown"
End Property


Public Property Get CaptionPictureSpace() As Long
    CaptionPictureSpace = m_CaptionPictureSpace
End Property

Public Property Let CaptionPictureSpace(ByVal New_Space As Long)
    m_CaptionPictureSpace = New_Space
    ActualisePosition
End Property

Public Property Get Style() As EnumStyle
    Style = m_Style
End Property

Public Property Let Style(ByVal New_Style As EnumStyle)
    m_Style = New_Style
    PropertyChanged "Style"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENTÉES SUIVANTES!
'MappingInfo=PicNormal,PicNormal,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Définit une icône de souris personnalisée."
    Set MouseIcon = PicNormal.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set PicNormal.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property


Public Sub About()
    MsgBox ";-) Fred Just Any Button OCX" & Chr(13) & _
        "It's a FreeWare !" & Chr(13) & _
        "You can use this OCX in your project" & Chr(13) & _
        "Check my site or mail me" & Chr(13) & _
        "fredjust@hotmail.com", vbInformation, "About AnyButton.ocx"
End Sub

