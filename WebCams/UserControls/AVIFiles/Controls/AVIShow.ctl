VERSION 5.00
Begin VB.UserControl AVIShow 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3135
   PropertyPages   =   "AVIShow.ctx":0000
   ScaleHeight     =   85
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   209
   ToolboxBitmap   =   "AVIShow.ctx":0011
End
Attribute VB_Name = "AVIShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' =================================================
' Name:       Visual Coders Animation Control
' Author:      Kim Pedersen, codemagician@get2net.dk
' Date:         17. September 1999
' Updated:    25. November 1999
' Version:    Beta2
'
' Requires:    None
'
' Description:
' A standard implementation of the Animation control in Common
' Controls. Allows you to play an AVI (without sound) on a form
' instead of a window. Lots of extended features.
'
' --------------------------------------------------------------------------------------------------
' Visit vbCode Magician, free source for VB programmers
' http://hjem.get2net.dk/vcoders/cm
' =================================================
'
' VCoders Ctrl Specific
Private anim_hWnd As Long
Private anim_Autoplay As Boolean
Private anim_Autosize As Boolean
Private anim_BorderStyle As Long
Private anim_FileName As String
Private anim_ResID As Long
Private anim_Transparent As Boolean

Private anim_SubClassed As Boolean
Private anim_WndProcNext As Long
Private anim_hInst As Long

Public Enum anim_Border
       animBorderNone = 0
       animBorderRaised = 1
       animBorderRaisedSingle = 2
       animBorderSunken = 3
       animBorderSunkenSingle = 4
End Enum

' Declares (Non specific)
Private Declare Function CreateWindowEx Lib "User32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function MoveWindow Lib "User32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShowWindow Lib "User32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function UpdateWindow Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function CallWindowProc Lib "User32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

' GetWindowLong's
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)
' Windows Messages
Private Const WM_COMMAND = &H111
' Window Styles
Private Const WS_CHILD = &H40000000
Private Const WS_TABSTOP = &H10000
Private Const WS_VISIBLE = &H10000000

' Animation Control Specific
' Thanks to Klaus H. Probst [kprobst@altavista.net]
Private Type tagInitCommonControlsEx
        animSize As Long
        animICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Const WM_USER = &H400
Private Const ICC_ANIMATE_CLASS = &H80
Private Const ANIMATE_CLASSA = "SysAnimate32"
'// begin_r_commctrl
Private Const ACS_CENTER = &H1
Private Const ACS_TRANSPARENT = &H2
Private Const ACS_AUTOPLAY = &H4
Private Const ACS_TIMER = &H8                  '// don't use threads... use timers
'// end_r_commctrl
'// Standard messages
Private Const ACM_OPEN = (WM_USER + 100)
Private Const ACM_PLAY = (WM_USER + 101)
Private Const ACM_STOP = (WM_USER + 102)
'// Notification messages... if you want them
Private Const ACN_START = 1
Private Const ACN_STOP = 2

' Control events
Event Click()
Attribute Click.VB_Description = "Raised when the user clicks the control."
Event DblClick()
Attribute DblClick.VB_Description = "Raised when the user doubleclicks the control."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Raised when the user releases the mousebutton."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Raised when the user press and hold a mouse button."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Raised when the cursor is move over the control."
Event StartPlaying()
Attribute StartPlaying.VB_Description = "Raised when the animation starts playing."
Event StopPlaying()
Attribute StopPlaying.VB_Description = "Raised when the animation stops playing."

Public Sub CloseFile()
Attribute CloseFile.VB_Description = "Closes the open animation and removes it from the control."
        s_Destroy
        Filename = ""
        ResourceID = 0
End Sub

Private Sub s_Create()
        
        ' This Sub will create
        ' a new Anim control
        ' and set all properties
        
        Dim animStyle As Long
        
        ' Destroy the previous Anim control if any
        If anim_hWnd <> 0 Then s_Destroy
        
        ' Set styles
        animStyle = WS_CHILD Or WS_VISIBLE
        animStyle = animStyle Or IIf(anim_Autoplay = True, ACS_AUTOPLAY, &H0)
        animStyle = animStyle Or IIf(anim_Transparent = True, ACS_TRANSPARENT, &H0)
        
        ' Create the Animation Control
        
        If IsNumeric(anim_FileName) = False And anim_ResID = 0 Then
                ' anim_filename is an AVI file
                anim_hWnd = CreateWindowEx(0, ANIMATE_CLASSA, vbNullString, animStyle, 0, 0, 0, 0, UserControl.hwnd, 0&, App.hInstance, ByVal 0&)
                If SendMessage(anim_hWnd, ACM_OPEN, 0, ByVal anim_FileName) = 0 Then
                        ' An error occured
                End If
        Else
                ' could be a DLL
                anim_hInst = LoadLibraryEx(anim_FileName, 0, 1)
                If anim_hInst <> 0 Then
                        ' Succes... The DLL could be opened
                        ' Create the animation control.
                        anim_hWnd = CreateWindowEx(0, ANIMATE_CLASSA, vbNullString, animStyle, 0, 0, 0, 0, UserControl.hwnd, 0&, anim_hInst, ByVal 0&)
                        If anim_ResID > 0 Then
                                If SendMessage(anim_hWnd, ACM_OPEN, 0, ByVal anim_ResID) = 0 And anim_ResID <> 0 Then
                                        ' An error.. No resource in that DLL
                                End If
                        End If
                End If
        End If
        
        s_SizeWindow
        s_DrawBorders
        
End Sub

Private Sub s_Destroy()
        
        ' Destroys the animation control
        
        If anim_hWnd <> 0 Then
                If anim_hInst <> 0 Then
                        ' Release the hInst of an open Library
                        FreeLibrary anim_hInst
                        anim_hInst = 0
                End If
                DestroyWindow anim_hWnd
                anim_hWnd = 0
        End If
End Sub


Private Sub s_DrawBorders()
       
       Dim bStyle As Long
       Dim bSize As RECT
       
       Select Case anim_BorderStyle
       Case 1 ' Raised
              bStyle = EDGE_RAISED
       Case 2 ' RaisedSingle
              bStyle = BDR_RAISEDINNER
       Case 3 ' Sunken
              bStyle = EDGE_SUNKEN
       Case 4 ' SunkenSingle
              bStyle = BDR_SUNKENOUTER
       Case Else ' none
              bStyle = &H0
       End Select
       
       UserControl.Cls
       
       If bStyle > &H0 Then
              ' Draw the border
              With bSize
                     .Left = 0
                     .Top = 0
                     .Bottom = ScaleHeight
                     .Right = ScaleWidth
              End With
              Call DrawEdge(hdc, bSize, bStyle, BF_RECT)
       End If
       
End Sub

Public Property Get BorderStyle() As anim_Border
Attribute BorderStyle.VB_Description = "Gets/Sets the type of border to be displayed when the control is visible."
Attribute BorderStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
       BorderStyle = anim_BorderStyle
End Property

Public Property Let BorderStyle(ByVal vNewValue As anim_Border)
       anim_BorderStyle = vNewValue
       s_SizeWindow
       s_DrawBorders
       PropertyChanged "BorderStyle"
End Property

Private Sub s_SizeWindow()
        If anim_hWnd <> 0 Then
                
                Dim mPix As Byte
                Dim winRect As RECT
                
                Select Case anim_BorderStyle
                Case 0: mPix = 0
                Case 1, 3: mPix = 3
                Case 2, 4: mPix = 2
                End Select
                
                If anim_Autosize And anim_hWnd <> 0 Then
                        ' Autosize the control to fit the animation only
                        ' (And borders :)
                        
                        GetWindowRect anim_hWnd, winRect
                        
                        With UserControl
                                
                                .Width = (winRect.Right - winRect.Left + (2 * mPix)) * Screen.TwipsPerPixelX
                                .Height = (winRect.Bottom - winRect.Top + (2 * mPix)) * Screen.TwipsPerPixelY
                                
                                Call MoveWindow(anim_hWnd, mPix, mPix, winRect.Right - winRect.Left, winRect.Bottom - winRect.Top, True)
                                
                        End With
                Else
                        ' Center the animation in the control
                        
                        GetWindowRect anim_hWnd, winRect
                        
                        With winRect
                                Call MoveWindow(anim_hWnd, (ScaleWidth - (.Right - .Left)) / 2, (ScaleHeight - (.Bottom - .Top)) / 2, .Right - .Left, .Bottom - .Top, True)
                        End With
                        
                End If
                
        End If
        
End Sub

Private Sub s_SubClass()
        
        s_UnSubClass
        
        anim_WndProcNext = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SubWndProc)
        
        If anim_WndProcNext Then
                SetWindowLong hwnd, GWL_USERDATA, ObjPtr(Me)
                anim_SubClassed = True
        End If
        
End Sub

Private Sub s_UnSubClass()
        If anim_SubClassed Then
                SetWindowLong hwnd, GWL_WNDPROC, anim_WndProcNext
                anim_SubClassed = False
        End If
End Sub

Public Sub StopPlay()
Attribute StopPlay.VB_Description = "Stop playing the animation."
        SendMessage anim_hWnd, ACM_STOP, 0, 0
End Sub

Friend Function s_WindowProc(ByVal Wnd As Long, ByVal uMsg As Long, ByVal wParam As Long, lParam As Long) As Long
        
        Select Case uMsg
        Case WM_COMMAND
                Select Case wParam \ &H10000
                Case ACN_START
                        RaiseEvent StartPlaying
                Case ACN_STOP
                        RaiseEvent StopPlaying
                End Select
        End Select
        
        s_WindowProc = CallWindowProc(anim_WndProcNext, Wnd, uMsg, wParam, ByVal lParam)
        
End Function

Private Sub UserControl_Click()
        RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
        RaiseEvent DblClick
End Sub


Private Sub UserControl_Initialize()
        ' Load common control 32 bit library
        Dim animInstanceLib As Long
        animInstanceLib = LoadLibrary("comctl32.dll")
        ' If the handle is valid, try to get the function address.
        If (animInstanceLib <> 0) Then
                Dim animProcAddress As Long
                Dim iccex As tagInitCommonControlsEx
                animProcAddress = GetProcAddress(animInstanceLib, "InitCommonControlsEx")
                If animProcAddress <> 0 Then
                        With iccex
                                .animSize = LenB(iccex)
                                .animICC = ICC_ANIMATE_CLASS
                        End With
                        Call InitCommonControlsEx(iccex)
                        s_Create
                        FreeLibrary animInstanceLib
                End If
        Else
                MsgBox "Common Controls Library can not be initialized.", 16, "Error"
                Err.Raise vbObjectError
        End If
        
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub


Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
        
        With PropBag
                Autoplay = .ReadProperty("Autoplay", False)
                Autosize = .ReadProperty("Autosize", False)
                BorderStyle = .ReadProperty("BorderStyle", 0)
                Filename = .ReadProperty("FileName", "")
                ResourceID = .ReadProperty("ResourceID", 0)
                Transparent = .ReadProperty("Transparent", False)
        End With
        
        If Ambient.UserMode Then s_SubClass
        
End Sub


Private Sub UserControl_Resize()
        s_SizeWindow
        s_DrawBorders
End Sub

Private Sub UserControl_Terminate()
  StopPlay
  CloseFile
        s_UnSubClass
        s_Destroy
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
        With PropBag
                Call .WriteProperty("Autoplay", anim_Autoplay, False)
                Call .WriteProperty("Autosize", anim_Autosize, False)
                Call .WriteProperty("BorderStyle", anim_BorderStyle, 0)
                Call .WriteProperty("FileName", anim_FileName, "")
                Call .WriteProperty("ResourceID", anim_ResID, 0)
                Call .WriteProperty("Transparent", anim_Transparent, False)
        End With
End Sub

Public Property Get Autoplay() As Boolean
Attribute Autoplay.VB_Description = "Start Playing the animation automatically."
Attribute Autoplay.VB_ProcData.VB_Invoke_Property = ";Behavior"
        Autoplay = anim_Autoplay
End Property

Public Property Let Autoplay(ByVal vNewValue As Boolean)
        anim_Autoplay = vNewValue
        If anim_hWnd <> 0 Then
                s_Create
                PropertyChanged "Autoplay"
        End If
End Property




Public Property Get Transparent() As Boolean
Attribute Transparent.VB_Description = "Gets/Sets whether the animation should be displayed as transparent. The color at 0, 0 in the animation will be transparent."
Attribute Transparent.VB_ProcData.VB_Invoke_Property = ";Appearance"
        Transparent = anim_Transparent
End Property

Public Property Let Transparent(ByVal vNewValue As Boolean)
        anim_Transparent = vNewValue
        s_Create
        PropertyChanged "Transparent"
End Property


Public Property Get Filename() As Variant
Attribute Filename.VB_Description = "Get/Sets the name of the AVI or Resource AVI to be opened."
Attribute Filename.VB_ProcData.VB_Invoke_Property = "ppageFiles;Files"
Attribute Filename.VB_MemberFlags = "200"
        Filename = anim_FileName
End Property

Public Property Let Filename(ByVal vNewValue As Variant)
        anim_FileName = vNewValue
        s_Create
        PropertyChanged "FileName"
End Property

Public Property Get ResourceID() As Long
Attribute ResourceID.VB_Description = "Gets/Sets the ID number of the animation within a dll. Set to 0 if AVI file is loaded."
Attribute ResourceID.VB_ProcData.VB_Invoke_Property = "ppageFiles;Files"
        ResourceID = anim_ResID
End Property

Public Property Let ResourceID(ByVal vNewValue As Long)
        anim_ResID = vNewValue
        s_Create
        PropertyChanged "ResourceID"
End Property

Public Property Get Autosize() As Boolean
Attribute Autosize.VB_Description = "Get/Sets if the control should autosize to fit the animation loaded."
Attribute Autosize.VB_ProcData.VB_Invoke_Property = ";Behavior"
        Autosize = anim_Autosize
End Property

Public Property Let Autosize(ByVal vNewValue As Boolean)
        anim_Autosize = vNewValue
        s_Create
        PropertyChanged "AutoSize"
End Property

Public Sub StartPlay(Optional ByVal From As Integer, Optional ByVal To_ As Integer = -1, Optional ByVal Repeat As Long = -1)
Attribute StartPlay.VB_Description = "Start playing the animation."
        If Len(anim_FileName) > 0 And anim_hWnd > 0 Then
                SendMessage anim_hWnd, ACM_PLAY, Repeat, ByVal MakeLong(From, To_)
        End If
End Sub

