VERSION 5.00
Begin VB.UserControl PopMenu 
   ClientHeight    =   1005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1005
   ScaleWidth      =   1950
   ToolboxBitmap   =   "cspPMenu.ctx":0000
   Begin VB.PictureBox picTest 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1980
      Left            =   840
      ScaleHeight     =   1920
      ScaleWidth      =   1920
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "cspPMenu.ctx":0182
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "PopMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' The sub classing control.  We need to use this because it
' ensures consistent sub-classing even if the form is being
' sub-classed by another control:
Implements ISubclass

' Need this to extract DRAWITEM and MEASUREITEM information
' from the owner draw messages:
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
    
' The messages we will intercept:
Private Const WM_MENUSELECT = &H11F
Private Const WM_MEASUREITEM = &H2C
Private Const WM_DRAWITEM = &H2B
Private Const WM_COMMAND = &H111
Private Const WM_MENUCHAR = &H120
Private Const WM_SYSCOMMAND = &H112
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_WININICHANGE = &H1A
Private Const WM_MDISETMENU = &H230

Private Const m_hColorDef = &HD6BEB5
Private Const m_bColorDef = &H6B2408
Private Const m_fColorDef = 0
Private Const m_fHColorDef = 0
Private Const m_ShadowDef = True
Private Const m_ShadowTopDef = True
Private Const m_LineColorDef = 0
Private Const m_ShadowColorDef = 0

' Enumerations:
Public Enum CSPPMENUSysCommandConstants
  SC_RESTORE = &HF120&
  SC_MOVE = &HF010&
  SC_SIZE = &HF000&
  SC_MAXIMIZE = &HF030&
  SC_MINIMIZE = &HF020&
  SC_CLOSE = &HF060&
  
  SC_ARRANGE = &HF110&
  SC_HOTKEY = &HF150&
  SC_HSCROLL = &HF080&
  SC_KEYMENU = &HF100&
  SC_MANAGER_CONNECT = &H1&
  SC_MANAGER_CREATE_SERVICE = &H2&
  SC_MANAGER_ENUMERATE_SERVICE = &H4&
  SC_MANAGER_LOCK = &H8&
  SC_MANAGER_MODIFY_BOOT_CONFIG = &H20&
  SC_MANAGER_QUERY_LOCK_STATUS = &H10&
  SC_MOUSEMENU = &HF090&
  SC_NEXTWINDOW = &HF040&
  SC_PREVWINDOW = &HF050&
  SC_SCREENSAVE = &HF140&
  SC_TASKLIST = &HF130&
  SC_VSCROLL = &HF070&
  SC_ZOOM = SC_MAXIMIZE
  SC_ICON = SC_MINIMIZE
End Enum

Public Enum CSPShowPopupMenuConstants   ' Track popup menu constants:
  TPM_CENTERALIGN = &H4&
  TPM_LEFTALIGN = &H0&
  TPM_LEFTBUTTON = &H0&
  TPM_RIGHTALIGN = &H8&
  TPM_RIGHTBUTTON = &H2&
  
  TPM_HORIZONTAL = &H0          '/* Horz alignment matters more */
  TPM_VERTICAL = &H40           '/* Vert alignment matters more */
  
  ' Win98/2000 menu animation and menu within menu options:
  TPM_RECURSE = &H1&
  TPM_HORPOSANIMATION = &H400&
  TPM_HORNEGANIMATION = &H800&
  TPM_VERPOSANIMATION = &H1000&
  TPM_VERNEGANIMATION = &H2000&
  ' Win2000 only:
  TPM_NOANIMATION = &H4000&
End Enum

Public Enum CSPHighlightStyleConstants
  cspHighlightStandard
  cspHighlightButton
  cspHighlightXP    '/* This is part of the mods put on the code by */
                    '/* ackbar. This setting will highlight it in a */
                    '/* xp style coloring sort of.                  */
End Enum

Private Type tSubMenuItem
  hMenu As Long
  hSysMenuOwner As Long
End Type

Private Type tVBMenuInfo
  sCaption As String
  sName As String
  sTag As String
  bHasIndex As Boolean
  iIndex As Long
  bUsed As Boolean
End Type

' Array of menu items
Private m_tMI() As tMenuItem
Private m_iMenuCount As Long
' Next id to choose for a menu item:
Private m_lLastMaxId As Long
' Height of a menu item:
Private m_lMenuItemHeight As Long
Private m_lIconSize As Long

' Hwnd of parent form:
Private m_hWndParent As Long
' If MDI form, then the hwnd of the MDI client area:
Private m_hWndMDIClient As Long
Private m_hLastMDIMenu As Long
Private m_cStoreMenus() As cStoreMenu
Private m_iStoreMenuCount As Long

' Handle to image list for drawing icons:
Private m_hIml As Long
' Where to get a tick icon for checked stuff:
Private m_lTickIconIndex As Long

' Sub menus we have created:
Private m_hSubMenus() As tSubMenuItem
Private m_lSubMenuCount As Long

' Subclassing response
Private m_emr As EMsgResponse

' When adding system menu items, set their id to this:
Private m_lNextSysMenuID As Long
Private Const WM_MENUBASE = &H2000&

' Whether to make top level menu items owner-draw:
Private m_bLeaveTopLevel As Boolean

' Display stuff, used to draw the control and also
' to evaluate menu font item sizes:
Private m_HDC As Long
Private m_hBMPOLd As Long
Private m_hBMPDither As Long
Private m_bUseDither As Boolean
Private m_hBmp As Long
Private m_hFntOld As Long
Private m_cNCM As cNCMetrics
Private m_bGotFont As Boolean
Private m_hDCBack As Long
Private m_lBitmapW As Long
Private m_lBitmapH As Long
Private m_tP As POINTAPI
Private m_lButton As Long
Private m_eStyle As CSPHighlightStyleConstants
Private m_hColor As OLE_COLOR
Private m_bColor As OLE_COLOR
Private m_fColor As OLE_COLOR
Private m_fHColor As OLE_COLOR
Private m_LineColor As OLE_COLOR
Private m_ShadowTop As Boolean
Private m_ShadowColor As OLE_COLOR

Private m_ShadowXPHighlight As Boolean

' Events:
Public Event Click(ItemNumber As Long)
Public Event SystemMenuClick(ItemNumber As Long)
Public Event ItemHighlight(ItemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)
Public Event SystemMenuItemHighlight(ItemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)
Public Event MenuExit()
Public Event InitPopupMenu(ParentItemNumber As Long)
Public Event WinIniChange()
Public Event NewMDIMenu()
Public Event RequestNewMenuDetails(ByRef sCaption As String, ByRef sKey As String, ByRef iIcon As Long, ByRef lItemData As Long, ByRef sHelptext As String, ByRef sTag As String)

Public Property Get HighlightStyle() As CSPHighlightStyleConstants
  HighlightStyle = m_eStyle
End Property

Public Property Let HighlightStyle(ByVal eStyle As CSPHighlightStyleConstants)
  m_eStyle = eStyle
  PropertyChanged "HighlightStyle"
End Property

Public Property Get ShadowXPHighlight() As Boolean
  ShadowXPHighlight = m_ShadowXPHighlight
End Property

Public Property Let ShadowXPHighlight(ByVal eSHighlight As Boolean)
  m_ShadowXPHighlight = eSHighlight
  PropertyChanged "ShadowXPHighlight"
End Property

Public Property Get ShadowXPHighlightTopMenu() As Boolean
  ShadowXPHighlightTopMenu = m_ShadowTop
End Property

Public Property Let ShadowXPHighlightTopMenu(ByVal eSHighlightTop As Boolean)
  m_ShadowTop = eSHighlightTop
  PropertyChanged "ShadowXPHighlightTopMenu"
End Property

Public Property Get HighlightColor() As OLE_COLOR
  HighlightColor = m_hColor
End Property

Public Property Let HighlightColor(ByVal hColor As OLE_COLOR)
  m_hColor = hColor
  PropertyChanged "HighlightColor"
End Property

Public Property Get LineColor() As OLE_COLOR
  LineColor = m_LineColor
End Property

Public Property Let LineColor(ByVal lColor As OLE_COLOR)
  m_LineColor = lColor
  PropertyChanged "LineColor"
End Property

Public Property Get ShadowColor() As OLE_COLOR
  ShadowColor = m_ShadowColor
End Property

Public Property Let ShadowColor(ByVal sColor As OLE_COLOR)
  m_ShadowColor = sColor
  PropertyChanged "ShadowColor"
End Property

Public Property Get ForeColor() As OLE_COLOR
  ForeColor = m_fColor
End Property

Public Property Let ForeColor(ByVal fColor As OLE_COLOR)
  m_fColor = fColor
  PropertyChanged "ForeColor"
End Property

Public Property Get HighlightForeColor() As OLE_COLOR
  HighlightForeColor = m_fHColor
End Property

Public Property Let HighlightForeColor(ByVal fHColor As OLE_COLOR)
  m_fHColor = fHColor
  PropertyChanged "HForeColor"
End Property

Public Property Get BorderColor() As OLE_COLOR
  BorderColor = m_bColor
End Property

Public Property Let BorderColor(ByVal bColor As OLE_COLOR)
  m_bColor = bColor
  PropertyChanged "BorderColor"
End Property

Public Function ShowPopupMenu(ByRef objTo As Object, ByVal vKeyParent As Variant, ByVal x As Single, ByVal y As Single, _
                              Optional ByVal eOptions As CSPShowPopupMenuConstants = TPM_LEFTALIGN Or TPM_HORIZONTAL) As Long
Dim lIndex As Long
Dim tP As POINTAPI
Dim tR As RECT
Dim eMode As VBRUN.ScaleModeConstants
Dim hMenu As Long
Dim lID As Long
Dim i As Long

   hMenu = hPopupMenu(vKeyParent)
   If (hMenu <> 0) Then
      eOptions = eOptions Or TPM_RETURNCMD
      With objTo
         On Error Resume Next
         eMode = .ScaleMode
         If (Err.Number = 0) Then
            ' Object has scalemode
            tP.x = .ScaleX(x, eMode, vbPixels)
            tP.y = .ScaleY(y, eMode, vbPixels)
         Else
            ' Object is scaled in twips
            tP.x = x \ Screen.TwipsPerPixelX
            tP.y = y \ Screen.TwipsPerPixelY
         End If
      End With
      ClientToScreen objTo.hwnd, tP
      lID = TrackPopupMenu(hMenu, eOptions, tP.x, tP.y, 0, m_hWndParent, tR)
      ' Find the ID:
      If (lID <> 0) Then
         For i = 1 To m_iMenuCount
            If (m_tMI(i).lActualID = lID) Then
               RaiseEvent Click(i)
               Exit For
            End If
         Next i
      End If
   End If
End Function

Public Property Set BackgroundPicture(ByRef sPic As StdPicture)
Dim tBm As BITMAP
   On Error Resume Next
   Set picTest.Picture = sPic
   If (Err.Number = 0) Then
      m_hDCBack = picTest.hdc
      GetObjectAPI picTest.Picture.Handle, Len(tBm), tBm
      m_lBitmapW = tBm.bmWidth
      m_lBitmapH = tBm.bmHeight
   Else
      m_hDCBack = 0
      m_lBitmapW = 0
      m_lBitmapH = 0
   End If
End Property

Public Property Get BackgroundPicture() As StdPicture
   Set BackgroundPicture = picTest.Picture
End Property

Public Sub ClearBackgroundPicture()
   Set picTest.Picture = Nothing
   m_hDCBack = 0
   m_lBitmapW = 0
   m_lBitmapH = 0
End Sub

Public Sub GetVersion(ByRef lMajor As Long, ByRef lMinor As Long, ByRef lRevision As Long)
   lMajor = App.Major
   lMinor = App.Minor
   lRevision = App.Revision
End Sub

Public Property Get HighlightCheckedItems() As Boolean
    HighlightCheckedItems = m_bUseDither
End Property

Public Property Let HighlightCheckedItems(ByVal bState As Boolean)
    m_bUseDither = bState
    PropertyChanged "HighlightCheckedItems"
End Property

Public Property Get MenuExists(ByVal vKey As Variant) As Boolean
   MenuExists = (plMenuIndex(vKey) > 0)
End Property

Public Property Get MenuIndex(ByVal vKey As Variant) As Long
Dim i As Long

   i = plMenuIndex(vKey)
   MenuIndex = i
   If (i < 0) Then
      'Err.Raise 9, App.EXEName & ".cPopMenu"
   End If
End Property

Private Function plMenuIndex(ByVal vKey As Variant) As Long
Dim i As Long

   ' Signal default
   plMenuIndex = -1
   ' Check for numeric key (i.e. index):
   If (IsNumeric(vKey)) Then
      i = CLng(vKey)
      If (i > 0) And (i <= m_iMenuCount) Then
         plMenuIndex = i
      End If
   Else
      ' Check for string key:
      For i = 1 To m_iMenuCount
          If (m_tMI(i).sKey = vKey) Then
              plMenuIndex = i
              Exit Function
          End If
      Next i
   End If
End Function

Public Property Get MenuKey(ByVal lIndex As Long) As String
    MenuKey = m_tMI(lIndex).sKey
End Property

Public Property Let MenuKey(ByVal lIndex As Long, ByVal sKey As String)
    If (pbIsValidKey(sKey)) Then
        m_tMI(lIndex).sKey = sKey
    End If
End Property

Public Property Get MenuTag(ByVal vKey As Variant) As String
Dim lIndex As Long
   lIndex = MenuIndex(vKey)
   If (lIndex > 0) Then
      MenuTag = m_tMI(lIndex).sTag
   End If
End Property

Public Property Let MenuTag(ByVal vKey As Variant, ByVal sTag As String)
Dim lIndex As Long
   lIndex = MenuIndex(vKey)
   If (lIndex > 0) Then
      m_tMI(lIndex).sTag = sTag
   End If
End Property

Public Property Get MenuDefault(ByVal vKey As Variant) As Boolean
Dim lIndex As Long
   lIndex = MenuIndex(vKey)
   If (lIndex > 0) Then
      MenuDefault = m_tMI(lIndex).bDefault
   End If
End Property

Public Property Let MenuDefault(ByVal vKey As Variant, ByVal bState As Boolean)
Dim lIndex As Long

   lIndex = MenuIndex(vKey)
   If (lIndex > 0) Then
      m_tMI(lIndex).bDefault = bState
   End If
End Property

Private Function pbIsValidKey(ByRef sKey As String) As Boolean
Dim i As Long
Dim bInvalid As Boolean

    If (Trim$(sKey) = "") Then
        sKey = Trim$(sKey)
        ' you're allowed to have a null key:
        pbIsValidKey = True
    Else
        For i = 1 To m_iMenuCount
            If (m_tMI(i).sKey = sKey) Then
                bInvalid = True
                Exit For
            End If
        Next i
        If (bInvalid) Then
            Err.Raise 457, App.EXEName & ".cPopMenu"
        Else
            pbIsValidKey = True
        End If
    End If
End Function

Public Property Let TickIconIndex(ByVal lIconIndex As Long)
    m_lTickIconIndex = lIconIndex
    PropertyChanged "TickIconIndex"
End Property

Public Property Get TickIconIndex() As Long
    TickIconIndex = m_lTickIconIndex
End Property

Public Property Get SystemMenuCaption(ByVal lPosition As Long) As String
Dim tMII As MENUITEMINFO
Dim hSysMenu As Long

    hSysMenu = GetSystemMenu(m_hWndParent, 0)
    If (hSysMenu <> 0) Then
        tMII.fMask = MIIM_DATA
        tMII.cch = 127
        tMII.dwTypeData = String$(128, 0)
        tMII.cbSize = LenB(tMII)
        GetMenuItemInfo hSysMenu, (lPosition - 1), 1, tMII
        SystemMenuCaption = left$(tMII.dwTypeData, tMII.cch)
    End If
End Property

Public Property Get SystemMenuCount() As Long
Dim hSysMenu As Long

    hSysMenu = GetSystemMenu(m_hWndParent, 0)
    If (hSysMenu <> 0) Then
        SystemMenuCount = GetMenuItemCount(hSysMenu)
    End If
End Property

Public Sub SystemMenuRemoveItem(ByVal lPosition As Long)
Dim hSysMenu As Long

    hSysMenu = GetSystemMenu(m_hWndParent, 0)
    If (hSysMenu <> 0) Then
        RemoveMenu hSysMenu, (lPosition - 1), MF_BYPOSITION
    End If
End Sub

Public Function SystemMenuAppendItem(ByVal sCaption As String) As Long
Dim hSysMenu As Long

    hSysMenu = GetSystemMenu(m_hWndParent, 0)
    If (hSysMenu <> 0) Then
        If (sCaption = "-") Then
            AppendMenuBylong hSysMenu, MF_SEPARATOR, m_lNextSysMenuID, m_lNextSysMenuID
        Else
            AppendMenuByString hSysMenu, MF_STRING, m_lNextSysMenuID, sCaption
        End If
        SystemMenuAppendItem = m_lNextSysMenuID
        m_lNextSysMenuID = m_lNextSysMenuID + 1
    End If
End Function

Public Sub SystemMenuRestore()
    GetSystemMenu m_hWndParent, 1
    m_lNextSysMenuID = WM_MENUBASE
End Sub

Private Function plParseMenuChar(ByVal hMenu As Long, ByVal iChar As Integer) As Long
Dim sChar As String
Dim l As Long
Dim lH() As Long
Dim sItems() As String
    
    sChar = UCase$(Chr$(iChar))
    For l = 1 To m_iMenuCount
        If (m_tMI(l).hMenu = hMenu) Then
            If (m_tMI(l).sAccelerator = sChar) Then
               pHierarchyForIndex l, lH(), sItems()
               plParseMenuChar = &H20000 Or lH(UBound(lH)) - 1
               ' Debug.Print "Found Menu Char"
               Exit Function
            End If
        End If
    Next l
End Function

Public Property Let ImageList(ByRef vImageList As Variant)
    m_hIml = 0
    If (VarType(vImageList) = vbLong) Then
        ' Assume a handle to an image list:
        m_hIml = vImageList
    ElseIf (VarType(vImageList) = vbObject) Then
        ' Assume a VB image list:
        On Error Resume Next
        ' Get the image list initialised..
        vImageList.ListImages(1).Draw 0, 0, 0, 1
        m_hIml = vImageList.hImageList
        If (Err.Number = 0) Then
            ' OK
        Else
            Debug.Print "Failed to Get Image list Handle", "cVGrid.ImageList"
        End If
        On Error GoTo 0
    End If
    If (m_hIml <> 0) Then
        Dim cX As Long, cY As Long
        If (ImageList_GetIconSize(m_hIml, cX, cY) <> 0) Then
            m_lIconSize = cY
            ' Set the menu item height accordingly:
            pSelectMenuFont
        End If
    End If
End Property

Public Property Get Caption(ByVal vKey As Variant) As String
Dim lIndex As Long
    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        Caption = m_tMI(lIndex).sCaption
    End If
End Property

Public Property Let Caption(ByVal vKey As Variant, ByVal sCaption As String)
Dim lIndex As Long
Dim i As Long
    
   ' Fixed bug where the menu item did not change size to accomodate the
   ' text.
   lIndex = MenuIndex(vKey)
   ReplaceItem lIndex, sCaption
      
   ' If other items with accelerators on this menu, then we need to ensure
   ' that all the other items at this menu level are also replaced.  This
   ' is the only way to ensure that the menus resize correctly.
   ' Do this each time because there isn't really a performance hit here.
   For i = 1 To m_iMenuCount
      If (i <> lIndex) And (m_tMI(i).hMenu = m_tMI(lIndex).hMenu) Then
         ReplaceItem i
      End If
   Next i
   
End Property

Public Property Get Enabled(ByVal vKey As Variant) As Boolean
Dim tMII As MENUITEMINFO
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        tMII.fMask = MIIM_STATE
        tMII.cbSize = LenB(tMII)
        GetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, False, tMII
        m_tMI(lIndex).bEnabled = Not ((tMII.fState And MFS_DISABLED) = MFS_DISABLED)
        Enabled = m_tMI(lIndex).bEnabled
    End If
End Property

Public Property Let Enabled(ByVal vKey As Variant, ByVal bEnabled As Boolean)
Dim lFlag As Long
Dim lFlagNot As Long
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        m_tMI(lIndex).bEnabled = bEnabled
        If (bEnabled) Then
            lFlag = MF_ENABLED
            lFlagNot = MF_GRAYED
        Else
            lFlag = MF_DISABLED
            lFlagNot = MF_GRAYED
        End If
        pSetMenuFlag lIndex, lFlag, lFlagNot
    End If
End Property

Public Property Get Checked(ByVal vKey As Variant) As Boolean
Dim tMII As MENUITEMINFO
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        tMII.fMask = MIIM_STATE
        tMII.cbSize = LenB(tMII)
        GetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, False, tMII
        m_tMI(lIndex).bChecked = ((tMII.fState And MFS_CHECKED) = MFS_CHECKED)
        Checked = m_tMI(lIndex).bChecked
    End If
End Property

Public Property Let Checked(ByVal vKey As Variant, ByVal bChecked As Boolean)
Dim lFlag As Long
Dim lFlagNot As Long
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        m_tMI(lIndex).bChecked = bChecked
        If (bChecked) Then
            lFlag = MF_CHECKED
            lFlagNot = 0
        Else
            lFlag = 0
            lFlagNot = MF_CHECKED
        End If
        pSetMenuFlag lIndex, lFlag, lFlagNot
    End If
End Property

Public Property Get HelpText(ByVal vKey As Variant) As String
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        HelpText = m_tMI(lIndex).sHelptext
    End If
End Property

Public Property Let HelpText(ByVal vKey As Variant, ByVal sHelptext As String)
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        m_tMI(lIndex).sHelptext = sHelptext
    End If
End Property

Public Property Let ItemIcon(ByVal vKey As Variant, ByVal lIconIndex As Long)
Dim lPrevIconIndex As Long
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        lPrevIconIndex = m_tMI(lIndex).lIconIndex
        m_tMI(lIndex).lIconIndex = lIconIndex
        If (((lPrevIconIndex = -1) Or (lIconIndex = -1)) And (lPrevIconIndex <> lIconIndex)) Then
            If (pbIsTopLevelmenu(lIndex)) Then
                ' Somehow we need to re-measure all the top menu items.
                ' How do we do this?
            End If
        End If
    End If
End Property

Public Property Get ItemIcon(ByVal vKey As Variant) As Long
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        ItemIcon = m_tMI(lIndex).lIconIndex
    End If
End Property

Public Property Get ItemData(ByVal vKey As Variant) As Long
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        ItemData = m_tMI(lIndex).lItemData
    End If
End Property

Public Property Let ItemData(ByVal vKey As Variant, ByVal lItemData As Long)
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        m_tMI(lIndex).lItemData = lItemData
    End If
End Property

Public Property Get hPopupMenu(ByVal vKey As Variant) As Long
Dim tMII As MENUITEMINFO
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        tMII.fMask = MIIM_SUBMENU
        tMII.cbSize = LenB(tMII)
        GetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, False, tMII
        hPopupMenu = tMII.hSubMenu
    End If
End Property

Public Property Get PositionInMenu(ByVal vKey As Variant) As Long
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
    End If
End Property

Public Property Get NextSibling(ByVal vKey As Variant) As Long
Dim lIndex As Long
Dim lParentId As Long
Dim l As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        ' The next sibling is the next item with the same parent id:
        lParentId = m_tMI(lIndex).lParentId
        For l = lIndex + 1 To m_iMenuCount
            If (m_tMI(l).lParentId = lParentId) Then
                NextSibling = l
                Exit For
            End If
        Next l
    End If
End Property

Public Property Get SiblingCount(ByVal vKey As Variant) As Long
Dim lIndex As Long
Dim lParentId As Long
Dim l As Long
Dim iCount As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        ' The sibling count is the total of items with the same parent id:
        lParentId = m_tMI(lIndex).lParentId
        For l = 1 To m_iMenuCount
            If (m_tMI(l).lParentId = lParentId) Then
                iCount = iCount + 1
            End If
        Next l
        SiblingCount = iCount
    End If
End Property

Public Property Get HasChildren(ByVal vKey As Variant)
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        ' An item has children if there are any menu items with the
        ' the parent id set to the index of this item:
        HasChildren = (FirstChild(lIndex) <> 0)
    End If
End Property

Public Property Get FirstChild(ByVal vKey As Variant) As Long
Dim lIndex As Long
Dim lID As Long
Dim l As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        ' Return the first item which has a ParentId = lIndex:
        lID = m_tMI(lIndex).lID
        For l = 1 To m_iMenuCount
            If (m_tMI(l).lParentId = lID) Then
                FirstChild = l
                Exit For
            End If
        Next l
    End If
End Property

Public Property Get Parent(ByVal vKey As Variant) As Long
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        Parent = plGetIndexForId(m_tMI(lIndex).lParentId)
    End If
End Property

Public Property Get UltimateParent(ByVal vKey As Variant) As Long
Dim lIndex As Long
Dim lR As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        lR = plGetIndexForId(m_tMI(lIndex).lParentId)
        If (lR <> 0) Then
            Do While m_tMI(lR).lParentId <> 0
                lR = plGetIndexForId(m_tMI(lR).lParentId)
            Loop
        End If
        If (lR = 0) Then
            UltimateParent = lIndex
        Else
            UltimateParent = lR
        End If
    End If
End Property

Private Sub pSetMenuFlag(ByVal lIndex As Long, ByVal lFlag As Long, ByVal lFlagNot As Long)
Dim tMII As MENUITEMINFO
Dim lFlags As Long

    lFlags = plMenuFlags(lIndex)
    lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING
    tMII.fMask = MIIM_SUBMENU
    tMII.cbSize = LenB(tMII)
    GetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, False, tMII
    If (tMII.hSubMenu <> 0) Then
        lFlags = lFlags Or MF_POPUP
    End If
    lFlags = lFlags And Not MF_BYPOSITION Or MF_BYCOMMAND
    
    lFlags = lFlags Or lFlag
    lFlags = lFlags And Not lFlagNot
    
    ModifyMenuByLong m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, lFlags, m_tMI(lIndex).lActualID, m_tMI(lIndex).lActualID
End Sub

Property Get IndexForMenuHierarchyParamArray(ParamArray vHierarchy() As Variant) As Long
Dim lH() As Long
Dim l As Long

    ReDim lH(LBound(vHierarchy) To UBound(vHierarchy)) As Long
    For l = LBound(vHierarchy) To UBound(vHierarchy)
        lH(l) = vHierarchy(l)
    Next l
    IndexForMenuHierarchyParamArray = IndexForMenuHierarchy(lH())
End Property

Property Get IndexForMenuHierarchy(ByRef lHierarchy() As Long) As Long
Dim l As Long
Dim lEnd As Long
Dim hMenuSeek As Long
Dim lRet As Long
Dim lFindIndex As Long

    hMenuSeek = GetMenu(m_hWndParent)
    lEnd = UBound(lHierarchy, 1)
    For l = LBound(lHierarchy, 1) To lEnd
        lFindIndex = plFindItemInMenu(hMenuSeek, lHierarchy(l))
        If (lFindIndex <> 0) Then
            If (l = lEnd) Then
                lRet = lFindIndex
            Else
                hMenuSeek = hPopupMenu(lFindIndex)
                If (hMenuSeek = 0) Then
                    Exit For
                End If
            End If
        Else
            Exit For
        End If
    Next l
    IndexForMenuHierarchy = lRet
End Property

Public Sub GetHierarchyForIndexPosition(ByVal vKey As Variant, ByRef lHierarchy() As Long)
Dim sItems() As String
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        pHierarchyForIndex lIndex, lHierarchy(), sItems()
    End If
End Sub

Property Get HierarchyPath(ByVal vKey As Variant, ByVal lStartLevel As Long, ByVal sSeparator As String) As String
Dim sItems() As String
Dim lH() As Long
Dim lItem As Long
Dim sOut As String
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        pHierarchyForIndex lIndex, lH(), sItems()
        For lItem = lStartLevel To UBound(sItems)
            sOut = sOut & sItems(lItem) & sSeparator
        Next lItem
        If Len(sOut) > 0 Then
            sOut = left$(sOut, Len(sOut) - 1)
            HierarchyPath = sOut
        End If
    End If
End Property

Private Function pHierarchyForIndex(ByVal lIndex As Long, ByRef lHierarchy() As Long, ByRef sItems() As String) As String
Dim lH() As Long
Dim sI() As String
Dim lItems As Long
Dim hMenuSeek As Long
Dim lPid As Long
Dim bComplete As Boolean
Dim l As Long
Dim lNewIndex As Long
Dim sOut As String

    Erase lHierarchy
    Erase sItems
    ' Now determine the hierarchy for this item:
    hMenuSeek = m_tMI(lIndex).hMenu
    Do
        lItems = lItems + 1
        ReDim Preserve lH(1 To lItems) As Long
        ReDim Preserve sI(1 To lItems) As String
        lH(lItems) = plMenuPositionFOrIndex(hMenuSeek, lIndex)
        sI(lItems) = m_tMI(lIndex).sCaption
        lPid = m_tMI(lIndex).lParentId
        If (lPid <> 0) Then
            lNewIndex = plGetIndexForId(m_tMI(lIndex).lParentId)
            ' Debug.Print lNewIndex
            lIndex = lNewIndex
            hMenuSeek = m_tMI(lIndex).hMenu
        Else
            bComplete = True
        End If
    Loop While Not (bComplete)
    
    ReDim lHierarchy(1 To lItems) As Long
    ReDim sItems(1 To lItems) As String
    For l = lItems To 1 Step -1
        lHierarchy(l) = lH(lItems - l + 1)
        sItems(l) = sI(lItems - l + 1)
    Next l

End Function

Private Function IndexForId(ByVal lID As Long)
Dim lItem As Long

    For lItem = 1 To m_iMenuCount
        If (m_tMI(lItem).lActualID = lID) Then
            IndexForId = lItem
            Exit For
        End If
    Next lItem
End Function

Private Function plMenuPositionFOrIndex(ByVal hMenuSeek As Long, ByVal lIndex As Long) As Long
Dim l As Long
Dim lPos As Long
Dim tMII As MENUITEMINFO
Dim lCount As Long

   ' fixed bug where this returned the wrong menu item...
   lCount = GetMenuItemCount(hMenuSeek)
   If (lCount > 0) Then
      For l = 0 To lCount - 1
         tMII.cbSize = Len(tMII)
         tMII.fMask = MIIM_ID
         GetMenuItemInfo hMenuSeek, l, True, tMII
         If (tMII.wID = m_tMI(lIndex).lActualID) And (m_tMI(lIndex).hMenu = hMenuSeek) Then
            plMenuPositionFOrIndex = l + 1
         End If
      Next l
   End If
   

    'For l = 1 To lIndex
    '    If (m_tMI(l).hMenu = hMenuSeek) Then
    '        lPos = lPos + 1
    '    End If
    'Next l
    'plMenuPositionFOrIndex = lPos
End Function

Private Function plFindItemInMenu(ByVal hMenuSeek As Long, ByVal lPosition As Long) As Long
Dim lPos As Long
Dim l As Long, i As Long
Dim lID As Long
Dim lCount As Long
Dim tMII As MENUITEMINFO

   ' fixed bug where this returned the wrong menu item...
   tMII.cbSize = Len(tMII)
   tMII.fMask = MIIM_ID
   GetMenuItemInfo hMenuSeek, lPosition - 1, True, tMII
      
   For i = 1 To m_iMenuCount
      If m_tMI(i).lActualID = tMII.wID And m_tMI(i).hMenu = hMenuSeek Then
         plFindItemInMenu = i
         Exit Function
      End If
   Next i
   
   'For l = 1 To m_iMenuCount
   '   If (m_tMI(l).hMenu = hMenuSeek) Then
   '         lPos = lPos + 1
   '         If (lPos = lPosition) Then
   '             plFindItemInMenu = l
   '             Exit For
   '         End If
   '   End If
   'Next l
End Function

Public Function ClearSubMenusOfItem(ByVal vKey As Variant) As Long
Dim hMenu As Long
Dim iMenu As Long
Dim lIndex As Long
    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        ' The idea is to leave just the submenu
        ' but with nothing in it:
        
        ' The ActualID of a sub-menu will be the
        ' handle to the submenu:
        hMenu = m_tMI(lIndex).lActualID
            
        ' Now remove all the items in the sub-menu,
        ' mark them for destruction and also do
        ' any sub-menus they may have:
        For iMenu = m_iMenuCount To 1 Step -1
            If (iMenu <= m_iMenuCount) Then
                If (m_tMI(iMenu).hMenu = hMenu) Then
                    pRemoveItem iMenu
                End If
            End If
        Next iMenu
        
        For iMenu = 1 To m_iMenuCount
            If (m_tMI(iMenu).lActualID = hMenu) Then
                ClearSubMenusOfItem = iMenu
                Exit For
            End If
        Next iMenu
    End If
End Function

Public Sub RemoveItem(ByVal vKey As Variant)
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        pRemoveItem lIndex
    End If
End Sub

Private Sub pRemoveItem(ByVal lIndex As Long)
Dim hMenusToDestroy() As Long
Dim lCount As Long
Dim lDestroy As Long
Dim lRealCount As Long
Dim lR As Long
Dim lMaxID As Long

    ' Remove the Item:
    lR = RemoveMenu(m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, MF_BYCOMMAND)
    m_tMI(lIndex).bMarkToDestroy = True
    ' Loop though all the children of the item at Index and determine
    ' what there is to remove:
    pRemoveSubMenus m_tMI(lIndex).lActualID, 1, hMenusToDestroy(), lCount
            
    ' Destroy the menus:
    For lDestroy = 1 To lCount
        DestroyMenu hMenusToDestroy(lDestroy)
        Debug.Print "Destroyed sub-menu:" & hMenusToDestroy(lDestroy)
    Next lDestroy
    
    ' Now repopulate the array & sort out the indexes to remove
    ' the indexes marked for deletion:
    If (lCount > 0) Or (lR <> 0) Then
        lRealCount = 0
        For lIndex = 1 To m_iMenuCount
            If Not (m_tMI(lIndex).bMarkToDestroy) Then
                If (GetMenuItemCount(m_tMI(lIndex).lActualID) = -1) Then
                    If (m_tMI(lIndex).lActualID > lMaxID) Then
                        lMaxID = m_tMI(lIndex).lActualID
                    End If
                End If
                lRealCount = lRealCount + 1
                If (lRealCount <> lIndex) Then
                     ' A much neater way than previously (set all the items independently!
                     ' what was I thinking of)
                     LSet m_tMI(lRealCount) = m_tMI(lIndex)
                End If
            End If
        Next lIndex
        ReDim Preserve m_tMI(1 To lRealCount) As tMenuItem
        m_iMenuCount = lRealCount
        If (lMaxID > m_iMenuCount) Then
            m_lLastMaxId = lMaxID
        Else
            m_lLastMaxId = m_iMenuCount
        End If
    End If
End Sub

Private Sub pRemoveSubMenus(ByVal lParentId As Long, ByVal lStartIndex As Long, ByRef hMenusToDestroy() As Long, ByRef lMenuToDestroyCount As Long)
Dim lIndex As Long
    
    For lIndex = 1 To m_iMenuCount
        If (m_tMI(lIndex).lParentId = lParentId) Then
            m_tMI(lIndex).bMarkToDestroy = True
            pAddToDestroyArray m_tMI(lIndex).hMenu, hMenusToDestroy(), lMenuToDestroyCount
            pRemoveSubMenus m_tMI(lIndex).lActualID, lIndex, hMenusToDestroy(), lMenuToDestroyCount
        End If
    Next lIndex
End Sub

Private Sub pAddToDestroyArray(ByVal hMenu As Long, ByRef hMenusToDestroy() As Long, ByRef lMenuToDestroyCount As Long)
Dim lIndex As Long
Dim bFound As Boolean

    For lIndex = 1 To lMenuToDestroyCount
        If (hMenusToDestroy(lIndex) = hMenu) Then
            bFound = True
            Exit For
        End If
    Next lIndex
    If Not (bFound) Then
        lMenuToDestroyCount = lMenuToDestroyCount + 1
        ReDim Preserve hMenusToDestroy(1 To lMenuToDestroyCount) As Long
        hMenusToDestroy(lMenuToDestroyCount) = hMenu
    End If
End Sub

Private Sub RaiseMenuExitEvent()
    RaiseEvent MenuExit
End Sub

Private Function RaiseClickEvent(lID As Long) As Boolean
' Return true from this if we have completely handled the
' click on our own:
Dim lIndex As Long

   ' Check whether this is an MDI child system menu command:
   If (lID = SC_SIZE) Or (lID = SC_MOVE) Or (lID = SC_MINIMIZE) Or (lID = SC_MAXIMIZE) Or (lID = SC_RESTORE) Or (lID = SC_CLOSE) Then
      RaiseClickEvent = False
   Else

      ' Find the Index of this menu id within our own array:
      lIndex = plGetIndexForId(lID)
      
      ' If we find it, then raise a click event for it:
      If (lIndex > 0) Then
      
          ' Send a click event with the index:
          RaiseEvent Click(lIndex)
          
          ' If this was one of the VB menu entries we have
          ' subclassed, we want to return false.  Then the
          ' click will filter through to the original Click
          ' event so your code should work as normal:
          If Not (m_tMI(lIndex).bIsAVBMenu) Then
              RaiseClickEvent = True
          End If
          
      Else
          ' This is a problem.  We've got a click on
          ' a menu id which doesn't seem to be any
          ' of the menu items of the form.  It shouldn't
          ' happen, but return true anyway so we eat
          ' the message.  This should prevent unwanted
          ' interference with any other controls on the
          ' form which seem to think the message is their
          ' own...
          Debug.Print "Failed to find index"
          RaiseClickEvent = True
      End If
   End If
End Function

Private Function pbIdIsSysMenuId(ByVal lID As Long) As Boolean
' Determine whether the menu id lId is actually one of the standard
' system menu items, or if it is an item which has been added to the
' system menu by this control:

    Select Case lID
    Case SC_RESTORE, SC_MOVE, SC_SIZE, SC_MAXIMIZE, SC_MINIMIZE, SC_CLOSE
        ' This is a standard menu id item:
        pbIdIsSysMenuId = True
    Case WM_MENUBASE To m_lNextSysMenuID - 1
        ' This is a new item added to the system menu by this control:
        pbIdIsSysMenuId = True
    End Select
    
End Function

Private Sub RaiseHighlightEvent(lID As Long)
Dim lIndex As Long

    ' Debug.Print lItem
    lIndex = plGetIndexForId(lID)
    If (lIndex > 0) Then
        RaiseEvent ItemHighlight(lIndex, m_tMI(lIndex).bEnabled, (Trim$(m_tMI(lIndex).sCaption = "-")))
    Else
        ' It may be a sys menu item:
        If (pbIdIsSysMenuId(lID)) Then
            Dim tMII As MENUITEMINFO
            Dim hSysMenu As Long
            hSysMenu = GetSystemMenu(m_hWndParent, 0)
            If (hSysMenu <> 0) Then
                tMII.fMask = MIIM_STATE
                tMII.cbSize = LenB(tMII)
                GetMenuItemInfo hSysMenu, lID, False, tMII
            End If
            RaiseEvent SystemMenuItemHighlight(lID, ((tMII.fState And MFS_DISABLED) <> MFS_DISABLED), False)
        Else
            Debug.Print "Failed to find Index for Highlight Id:", lID, lIndex
        End If
    End If
End Sub

Private Sub DrawItem(ByRef lParam As Long, ByRef wParam As Long)
Dim tDI As DRAWITEMSTRUCT
Dim lHDC As Long
Dim lIndex As Long
Dim lColour As Long
Dim lDiff As Long
Dim lFillColour As Long
Dim bDisabled As Boolean
Dim bSelected As Boolean
Dim bChecked As Boolean
Dim bIsTopLevel As Boolean
Dim hBrush As Long
Dim tP As POINTAPI
Dim tB As RECT
Dim tC As RECT
Dim tOB As RECT
Dim tS As RECT
Dim tFR As RECT
Dim tIR As RECT
Dim tFC As RECT
Dim sText As String
Dim x As Long
Dim hFntBold As Long, hFntOld As Long

   CopyMemory tDI, ByVal lParam, Len(tDI)
   'Debug.Print "CtlID:", tDI.CtlID, "CtlType:", tDI.CtlType, "HwndItem:", tDI.hwndItem
   If (tDI.CtlType = 1) Then ' Menu
      lIndex = (plGetIndexForId(tDI.itemID))
      If (lIndex <> 0) Then
         ' Debug.Print "Found item to draw."
         lHDC = tDI.hdc
         SetBkMode lHDC, TRANSPARENT
         
         ' 'Default' (bolded) item?
         If (m_tMI(lIndex).bDefault) Then
            hFntBold = m_cNCM.BoldenedFontHandle(MenuFOnt)
            If (hFntBold <> 0) Then
               hFntOld = SelectObject(lHDC, hFntBold)
            End If
         End If
         
         bDisabled = ((tDI.itemState And ODS_DISABLED) = ODS_DISABLED) Or ((tDI.itemState And ODS_GRAYED) = ODS_GRAYED)
         bSelected = ((tDI.itemState And ODS_SELECTED) = ODS_SELECTED)
         bChecked = ((tDI.itemState And ODS_CHECKED) = ODS_CHECKED)
         bIsTopLevel = pbIsTopLevelmenu(lIndex)
         
         tP.x = tDI.rcItem.left
         tP.y = tDI.rcItem.tOp + 1
         
         ' Store the rectangle into tOB so we can draw an
         ' outer border if this is a top level item:
         CopyMemory tOB, tDI.rcItem, LenB(tB)
          
         If Not (bIsTopLevel) Then
            If (m_hDCBack <> 0) Then
               TileArea lHDC, tDI.rcItem.left, tDI.rcItem.tOp, tDI.rcItem.Right - tDI.rcItem.left, tDI.rcItem.Bottom - tDI.rcItem.tOp, m_hDCBack, m_lBitmapW, m_lBitmapH
            End If
         End If
          
         If (bDisabled) Then
            If (Trim$(m_tMI(lIndex).sCaption = "-")) Then
               ' If this is a separator, then draw the separator line:
               If (m_eStyle <> cspHighlightXP) Then
                 tS.left = tP.x
                 tS.Right = tDI.rcItem.Right
                 tS.tOp = tP.y + 1
                 tS.Bottom = tS.tOp + 2
                 DrawEdge lHDC, tS, EDGE_ETCHED, BF_TOP
               Else
                 tS.tOp = tP.y + 1
                 tS.Bottom = tS.tOp + 1
                 tS.left = m_lMenuItemHeight + 5
                 tS.Right = tDI.rcItem.Right - 2
                 hBrush = CreateSolidBrush(m_LineColor)
                 FillRect lHDC, tS, hBrush
               End If
            Else
               ' Get the position to output the text:
               pGetTextPosition lHDC, lIndex, bIsTopLevel, tDI.rcItem
               
               ' Draw the text grayed:
               With tDI.rcItem
                   .left = .left + 1
                   .tOp = .tOp + 1
               End With
               SetTextColor lHDC, GetSysColor(COLOR_BTNHIGHLIGHT)
               pDrawMenuCaption lHDC, lIndex, tDI.rcItem
               With tDI.rcItem
                   .left = .left - 1
                   .tOp = .tOp - 1
               End With
               SetTextColor lHDC, GetSysColor(COLOR_BTNSHADOW)
               pDrawMenuCaption lHDC, lIndex, tDI.rcItem
               
               SetTextColor lHDC, GetSysColor(COLOR_MENUTEXT)
            End If
         Else
            If (Trim$(m_tMI(lIndex).sCaption = "-")) Then
               ' Draw a separator line:
               If (bIsTopLevel) Then
                  ' We draw nothing - a separator at
                  ' the top level just leaves a space.
                  Debug.Print "Separator at top level"
               Else
                  ' We draw the separator line:
                  tS.left = tP.x
                  tS.tOp = tP.y + 1
                  tS.Bottom = tS.tOp + 2
                  tS.Right = tDI.rcItem.Right
                  DrawEdge lHDC, tS, EDGE_ETCHED, BF_TOP
               End If
            Else
              
               ' Set the back colour and text colour for the menu item:
               If bSelected And Not (bIsTopLevel) And (m_eStyle = cspHighlightStandard) Then
                  lFillColour = COLOR_HIGHLIGHT
                  lColour = m_fHColor 'GetSysColor(COLOR_HIGHLIGHTTEXT)
               Else
                  lFillColour = COLOR_MENU
                  lColour = m_fColor 'GetSysColor(COLOR_MENUTEXT)
               End If
               CopyMemory tFR, tDI.rcItem, LenB(tFR)
               If (m_tMI(lIndex).lIconIndex > -1) Or (bChecked) Then
                  ' Erase the icon background:
                  If (m_hDCBack = 0) Or (bIsTopLevel) Then
                     CopyMemory tIR, tDI.rcItem, LenB(tIR)
                     tIR.Right = tIR.left + m_lMenuItemHeight - 1
                     hBrush = GetSysColorBrush(COLOR_MENU)
                     FillRect lHDC, tIR, hBrush
                     DeleteObject hBrush
                     tFR.left = tFR.left + m_lMenuItemHeight
                  End If
               End If
               If (m_eStyle <> cspHighlightXP And m_eStyle <> cspHighlightButton And (bSelected) Or (bIsTopLevel)) Then
                    hBrush = GetSysColorBrush(lFillColour)
                    FillRect lHDC, tFR, hBrush
                    DeleteObject hBrush
               Else
                  If m_hDCBack = 0 Then
                    hBrush = GetSysColorBrush(lFillColour)
                    FillRect lHDC, tOB, hBrush
                    DeleteObject hBrush
                  End If
               End If
               SetTextColor lHDC, lColour
                  
               ' Get the position to output the text:
               If (m_eStyle = cspHighlightXP) And Not (bIsTopLevel) Then tDI.rcItem.left = tDI.rcItem.left + 3
               pGetTextPosition lHDC, lIndex, bIsTopLevel, tDI.rcItem
               pDrawMenuCaption lHDC, lIndex, tDI.rcItem
                                  
               If bSelected And (m_eStyle = cspHighlightXP) Then
                 Dim lBorderColor As Long, lBrush As Long
                 lColour = m_fHColor
                 lFillColour = m_hColor 'RGB(181, 190, 214)
                 lBorderColor = m_bColor 'RGB(8, 36, 107)
                 If (m_ShadowXPHighlight) And (Not (bIsTopLevel) Or (m_ShadowTop)) Or ((bIsTopLevel And (m_ShadowTop))) Then
                   hBrush = CreateSolidBrush(m_ShadowColor)
                   tOB.left = tOB.left + 3
                   tOB.tOp = tOB.tOp + 3
                   FillRect lHDC, tOB, hBrush
                   tOB.left = tOB.left - 2
                   tOB.tOp = tOB.tOp - 2
                 Else
                   tOB.left = tOB.left + 1
                   tOB.tOp = tOB.tOp + 1
                 End If
                 tOB.Right = tOB.Right - 1
                 tOB.Bottom = tOB.Bottom - 1
                 hBrush = CreateSolidBrush(lBorderColor)
                 FillRect lHDC, tOB, hBrush
                 DeleteObject hBrush
                 tOB.left = tOB.left + 1
                 tOB.tOp = tOB.tOp + 1
                 tOB.Right = tOB.Right - 1
                 tOB.Bottom = tOB.Bottom - 1
                 lBrush = CreateSolidBrush(lFillColour)
                 FillRect lHDC, tOB, lBrush
                 DeleteObject lBrush
                 SetTextColor lHDC, lColour
                 ' Get the position to output the text:
                 pDrawMenuCaption lHDC, lIndex, tDI.rcItem
               End If
               If ((bIsTopLevel) Or (m_eStyle = cspHighlightButton)) And (bSelected) And (m_eStyle <> cspHighlightXP) Then
                  ' We draw a sunken box around selected
                  ' top level menu items:
                  tOB.Right = tOB.Right - 1
                  tOB.Bottom = tOB.Bottom
                  If (bIsTopLevel) Then
                     DrawEdge lHDC, tOB, BDR_SUNKENOUTER, BF_RECT
                  Else
                     If (GetAsyncKeyState(vbKeyLButton) <> 0) And (m_tMI(lIndex).lActualID = m_tMI(lIndex).lID) Then
                        DrawEdge lHDC, tOB, BDR_SUNKENOUTER, BF_RECT
                     Else
                        DrawEdge lHDC, tOB, BDR_RAISEDINNER, BF_RECT
                     End If
                  End If
               End If
            End If
         End If
          
         If (bChecked And m_eStyle <> cspHighlightXP) Then
          ' We draw a sunken box around the checked item:
          'With the xp style menu's this would look really bad
          'if we drew in this method so we gotta check if it's
          'xp or not and handle it.
          tB.left = tP.x
          tB.tOp = tP.y + 1
          tB.Right = tB.left + m_lMenuItemHeight - 1
          tB.Bottom = tB.tOp + m_lMenuItemHeight - 2
          DrawEdge lHDC, tB, BDR_SUNKENOUTER, BF_RECT
            ' If we're not disabled, we should fill the background.
            ' If we're selected, we use BTNFACE, otherwise we use
            ' BTNHIGHLIGHT (although really we should be dithering
            ' it... but its too hard to do right now!)>
            ' SPM: 29/07/98 - now dithering is enabled.  Only works in
            ' the compiled OCX though!
         
         'End If
            If Not (bDisabled) Then
               If (bSelected) Then
                   hBrush = GetSysColorBrush(COLOR_BTNFACE)
               Else
               
                   If m_bUseDither Then
                       hBrush = CreatePatternBrush(m_hBMPDither)
                   Else
                       hBrush = CreateSolidBrush(GetSysColor(COLOR_BTNFACE))
                   End If
               End If
               CopyMemory tFC, tB, LenB(tFC)
               With tFC
                   .left = .left + 1
                   .tOp = .tOp + 1
                   .Right = .Right - 2
                   .Bottom = .Bottom - 2
               End With
               FillRect lHDC, tFC, hBrush
               DeleteObject hBrush
            End If
              
            ' If we don't have an icon, we ought to draw the tick icon
            ' inside it:
            If (m_tMI(lIndex).lIconIndex = -1) Then
               If (bIsTopLevel) Then
                  x = tP.x
               Else
                  x = tP.x + (m_lMenuItemHeight - 18) \ 2
               End If
               If (bDisabled) Then
                  ImageListDrawIconDisabled lHDC, m_hIml, m_lTickIconIndex, x + 1, tP.y + 3, 16
               Else
                  ImageListDrawIcon lHDC, m_hIml, m_lTickIconIndex, x + 1, tP.y + 3, , (bDisabled)
               End If
             End If
         
         ElseIf (bSelected And Not (bIsTopLevel)) And Not (bDisabled) Then
            ' We should draw a raised box around the icon:
            If (m_tMI(lIndex).lIconIndex > -1) And (m_eStyle = cspHighlightStandard) Then
               tB.left = tP.x
               tB.tOp = tP.y - 1
               tB.Right = tB.left + m_lMenuItemHeight - 1
               tB.Bottom = tB.tOp + m_lMenuItemHeight
               DrawEdge lHDC, tB, BDR_RAISEDINNER, BF_RECT
            End If

         End If
            If ((m_eStyle = cspHighlightXP) And (bChecked)) Then
              tB.left = tP.x + 1
              tB.tOp = tP.y
              tB.Right = tB.left + m_lMenuItemHeight - 1
              tB.Bottom = tB.tOp + m_lMenuItemHeight - 1
              hBrush = CreateSolidBrush(m_bColor)
              FillRect lHDC, tB, hBrush
             
              DeleteObject hBrush
              tB.left = tB.left + 1
              tB.tOp = tB.tOp + 1
              tB.Right = tB.Right - 1
              tB.Bottom = tB.Bottom - 1
             
              hBrush = CreateSolidBrush(m_hColor)
              FillRect lHDC, tB, hBrush
              DeleteObject hBrush
            End If
   
         If (m_tMI(lIndex).lIconIndex > -1) Then
            If (bIsTopLevel) Then
               x = tP.x
            Else
               x = tP.x + (m_lMenuItemHeight - 18) \ 2
            End If
            If (bDisabled) Then
               ImageListDrawIconDisabled lHDC, m_hIml, m_tMI(lIndex).lIconIndex, x + 1, tP.y + 3, 16
            Else
               ImageListDrawIcon lHDC, m_hIml, m_tMI(lIndex).lIconIndex, x + 1, tP.y + 3, , (bDisabled)
            End If
         End If
                    
         If (hFntBold <> 0) Then
            If (hFntOld <> 0) Then
               SelectObject lHDC, hFntOld
            Else
               SelectObject lHDC, m_cNCM.FontHandle(MenuFOnt)
            End If
            DeleteObject hFntBold
         End If
          
          
      Else
         Debug.Print "Failed to find item to draw.", tDI.itemID, lIndex
      End If
   End If
End Sub

Private Sub pDrawMenuCaption(ByVal lHDC As Long, ByVal lIndex As Long, ByRef tR As RECT)
Dim sText As String
Dim tSR As RECT

   sText = Trim$(m_tMI(lIndex).sCaption)
   DrawText lHDC, sText, Len(sText), tR, DT_LEFT
   sText = Trim$(m_tMI(lIndex).sShortCutDisplay)
   If (sText <> "") Then
      CopyMemory tSR, tR, LenB(tR)
      tSR.left = m_tMI(lIndex).lShortCutStartPos
      DrawText lHDC, sText, Len(sText), tSR, DT_LEFT
   End If
End Sub

Private Function pbIsTopLevelmenu(ByVal lIndex As Long) As Boolean
   pbIsTopLevelmenu = (m_tMI(lIndex).hMenu = GetMenu(m_hWndParent))
End Function

Private Function pGetTextPosition(ByVal lHDC As Long, ByVal lIndex As Long, ByVal bIsTopLevel As Boolean, _
                                  ByRef rcItem As RECT)
Dim tC As RECT
Dim lDiff As Long
Dim lMenuHeight As Long

    If (bIsTopLevel) Then
        lMenuHeight = GetSystemMetrics(SM_CYMENU) - 2 ' Allow for border
    Else
        lMenuHeight = m_lMenuItemHeight
    End If
    
    ' Determine the size of the text to draw:
    DrawText lHDC, m_tMI(lIndex).sCaption, Len(m_tMI(lIndex).sCaption), tC, DT_CALCRECT
    
    ' We want to centre the text vertically:
    lDiff = lMenuHeight - (tC.Bottom - tC.tOp)
    If (lDiff > 0) Then
        rcItem.tOp = rcItem.tOp + lDiff \ 2
    End If
    
    ' Now move the left position of the text
    ' across to accommodate icon/selection rectangle:
    If (bIsTopLevel) Then
        ' If its a top level item, then move across
        ' to accommodate the border.  Additionally, if
        ' there is an icon, move across to accomodate
        ' the icon:
        If (m_tMI(lIndex).lIconIndex > -1) Then
            rcItem.left = rcItem.left + 18
        Else
            rcItem.left = rcItem.left + 4
        End If
    Else
        ' All normal menu items are indented by 26 to
        ' accomodate icon & checked surround for icon:
        rcItem.left = rcItem.left + 26
    End If
End Function

Private Function plGetIndexForId(ByVal lItemId As Long) As Long
Dim l As Long
Dim lIndex As Long
    'Debug.Print "Finding Index:"
    'Debug.Print lItemId
    lIndex = 0
    For l = 1 To m_iMenuCount
        'Debug.Print "    Index at l = " & m_tMI(l).lId
        If (m_tMI(l).lActualID = lItemId) Then
            lIndex = l
            Exit For
        End If
    Next l
    plGetIndexForId = lIndex
End Function

Private Sub pGetTheMenuFont()

End Sub

Private Sub MeasureItem(ByVal lItemId As Long, ByRef lWidth As Long, ByRef lHeight As Long)
Dim lIndex As Long
Dim tR As RECT
Dim bDontEvalWidth As Long
Dim bIsTopLevel As Long
Dim sLongestCaption As String
Dim sLongestShortCut As String
Dim l As Long
Dim lItemsOnMenu() As Long
Dim lItemCount As Long
Dim hMenu As Long
Dim lOrigWidth As Long
Dim hFnt As Long, hFntOld As Long

    pGetTheMenuFont

    lIndex = plGetIndexForId(lItemId)
    If (lIndex <> 0) Then
        bIsTopLevel = pbIsTopLevelmenu(lIndex)
        If (bIsTopLevel) Then
            lHeight = GetSystemMetrics(SM_CYMENU)
        Else
            lHeight = m_lMenuItemHeight
        End If
        ' Determine the width of the item:
        If bIsTopLevel Then
            'Debug.Print "Top Level"
            If Trim$(m_tMI(lIndex).sCaption) = "-" Or Trim$(m_tMI(lIndex).sCaption) = "" Then
                lWidth = 4
                bDontEvalWidth = True
            Else
                If (m_tMI(lIndex).lIconIndex > -1) Then
                    lWidth = 18
                Else
                    lWidth = 0
                End If
            End If
            'Debug.Print lWidth
        Else
            If Trim$(m_tMI(lIndex).sCaption = "-") Then
                lHeight = 6
                lWidth = 32
                bDontEvalWidth = True
            Else
               If (m_tMI(lIndex).sCaption = "") Then
                  lWidth = m_lIconSize - 6
               Else
                  lWidth = 32
               End If
            End If
        End If
        
        If Not (bDontEvalWidth) Then
            If (m_tMI(lIndex).bDefault) Then
               hFnt = m_cNCM.BoldenedFontHandle(MenuFOnt)
               If (hFnt <> 0) Then
                  hFntOld = SelectObject(m_HDC, hFnt)
               End If
            End If
        
            If bIsTopLevel Then
                ' For top level items we evaluate the width of
                ' the actual text item only:
                DrawText m_HDC, m_tMI(lIndex).sCaption, Len(m_tMI(lIndex).sCaption), tR, DT_CALCRECT
                lWidth = lWidth + tR.Right
            Else
               ' Return the total width.  If CTRL accelerators on this menu level,
               ' we need to evaluate the maximum size as well to make sure
               ' these work too.
               lOrigWidth = lWidth
               hMenu = m_tMI(lIndex).hMenu
               For l = 1 To m_iMenuCount
                  If (m_tMI(l).hMenu = hMenu) Then
                     If Len(m_tMI(l).sCaption) > Len(sLongestCaption) Then
                        sLongestCaption = m_tMI(l).sCaption
                     End If
                     If (Len(m_tMI(l).sShortCutDisplay) > Len(sLongestShortCut)) Then
                        sLongestShortCut = m_tMI(l).sShortCutDisplay
                     End If
                     lItemCount = lItemCount + 1
                     ReDim Preserve lItemsOnMenu(1 To lItemCount) As Long
                     lItemsOnMenu(lItemCount) = l
                  End If
               Next l
                
               DrawText m_HDC, sLongestCaption, Len(sLongestCaption), tR, DT_CALCRECT
               lWidth = lWidth + tR.Right
               If (sLongestShortCut <> "") Then
                  DrawText m_HDC, sLongestShortCut, Len(sLongestShortCut), tR, DT_CALCRECT
                  lWidth = lWidth + 8
                  For l = 1 To lItemCount
                     m_tMI(lItemsOnMenu(l)).lShortCutStartPos = lWidth
                  Next l
                  lWidth = lWidth + tR.Right
               End If
                
            End If
            
            If (hFnt <> 0) Then
               If (hFntOld <> 0) Then
                  SelectObject m_HDC, hFntOld
               End If
               DeleteObject hFnt
            End If
        End If
        'Debug.Print "Width " & lWidth
    End If
End Sub

Property Get IDForIndex(ByVal vKey As Variant) As Long
Dim lIndex As Long

    lIndex = MenuIndex(vKey)
    If (lIndex > 0) Then
        IDForIndex = m_tMI(lIndex).lActualID
    End If
End Property

Public Function AddItem(ByVal sCaption As String, Optional ByVal sKey As String = "", _
                        Optional ByVal sHelptext As String = "", Optional ByVal lItemData As Long = 0, _
                        Optional ByVal lParentIndex As Long = 0, Optional ByVal lIconIndex As Long = -1, _
                        Optional ByVal bChecked As Boolean = False, Optional ByVal bEnabled As Boolean = True) As Long
Dim lID As Long

   ' Appends a new item to the end of a menu:
   If (pbIsValidKey(sKey)) Then
      m_iMenuCount = m_iMenuCount + 1
      ReDim Preserve m_tMI(1 To m_iMenuCount) As tMenuItem
      lID = plGetNewID()
      With m_tMI(m_iMenuCount)
         .lID = lID
         .lActualID = lID
         pSetMenuCaption m_iMenuCount, sCaption, (sCaption = "-")
         .sAccelerator = psExtractAccelerator(sCaption)
         .sHelptext = sHelptext
         .lIconIndex = lIconIndex
         .lParentId = m_tMI(lParentIndex).lActualID
         .lItemData = lItemData
         .bChecked = bChecked
         .bEnabled = bEnabled
         .bCreated = True
         ' Bug fix:
         .sKey = sKey
      End With
      pAddNewMenuItem m_tMI(m_iMenuCount)
      AddItem = m_iMenuCount
   End If
End Function

Public Function ReplaceItem(ByVal vKey As Variant, Optional ByVal sCaption As Variant, _
                            Optional ByVal sHelptext As Variant, Optional ByVal lItemData As Variant, _
                            Optional ByVal lIconIndex As Variant, Optional ByVal bChecked As Variant, _
                            Optional ByVal bEnabled As Variant) As Long
Dim lIndex As Long
Dim sItems() As String
Dim lH() As Long
Dim lR As Long
Dim lFlags As Long
Dim lPosition As Long
Dim tMI As MENUITEMINFO
Dim hSubMenu As Long

   ' Replaces a menu item with a new one.  Works
   ' around a bug with the caption property where if
   ' you changed the size of the caption the menu did
   ' not resize.  Also allows you to change the help
   ' text, item data, icon, check and enable at the
   ' same time.
   
   ' Check valid index:
   lIndex = MenuIndex(vKey)
   If (lIndex > 0) Then
      If Not IsMissing(sCaption) Then
         pSetMenuCaption lIndex, sCaption, (sCaption = "-")
      End If
      If Not IsMissing(sHelptext) Then
         m_tMI(lIndex).sHelptext = sHelptext
      End If
      If Not IsMissing(lItemData) Then
         m_tMI(lIndex).lItemData = lItemData
      End If
      If Not IsMissing(lIconIndex) Then
         m_tMI(lIndex).lIconIndex = lIconIndex
      End If
      If Not IsMissing(bChecked) Then
         m_tMI(lIndex).bChecked = bChecked
      End If
      If Not IsMissing(bEnabled) Then
         m_tMI(lIndex).bEnabled = bEnabled
      End If
      
      pHierarchyForIndex lIndex, lH(), sItems()
      lPosition = lH(UBound(lH)) - 1
      ' Check if there is a sub menu:
      tMI.cbSize = Len(tMI)
      tMI.fMask = MIIM_SUBMENU
      GetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, 0, tMI
      hSubMenu = tMI.hSubMenu
      ' Remove the menu item:
      lR = RemoveMenu(m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, MF_BYCOMMAND)
      ' Insert it back again at the corect position with the same ID etc:
      lFlags = plMenuFlags(lIndex)
      lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING Or MF_BYPOSITION
      lR = InsertMenuByLong(m_tMI(lIndex).hMenu, lPosition, lFlags, m_tMI(lIndex).lID, m_tMI(lIndex).lID)
      If (hSubMenu <> 0) Then
         ' If we had a submenu then put that back again:
         lFlags = lFlags And Not MF_BYPOSITION Or MF_BYCOMMAND
         lFlags = lFlags Or MF_POPUP
         lR = ModifyMenuByLong(m_tMI(lIndex).hMenu, m_tMI(lIndex).lID, lFlags, hSubMenu, m_tMI(lIndex).lActualID)
      End If
      If (lR = 0) Then
         Debug.Print "Failed to insert new menu item."
      End If
      If (m_tMI(lIndex).hMenu = GetMenu(m_hWndParent)) Then
         DrawMenuBar m_hWndParent
      End If
   End If
End Function

Public Function InsertItem(ByVal sCaption As String, ByVal vKeyBefore As Variant, _
                           Optional ByVal sKey As String = "", Optional ByVal sHelptext As String = "", _
                           Optional ByVal lItemData As Long = 0, Optional ByVal lIconIndex As Long = -1, _
                           Optional ByVal bChecked As Boolean = False, Optional ByVal bEnabled As Boolean = True) As Long
Dim lIndexBefore As Long
Dim lID As Long
   'Inserts an item into a menu:
   If (pbIsValidKey(sKey)) Then
      lIndexBefore = MenuIndex(vKeyBefore)
      If (lIndexBefore > 0) Then
         m_iMenuCount = m_iMenuCount + 1
         ReDim Preserve m_tMI(1 To m_iMenuCount) As tMenuItem
         lID = plGetNewID()
         With m_tMI(m_iMenuCount)
            .lID = lID
            .lActualID = lID
             pSetMenuCaption m_iMenuCount, sCaption, (sCaption = "-")
            .sAccelerator = psExtractAccelerator(sCaption)
            .sHelptext = sHelptext
            .lIconIndex = lIconIndex
            .lItemData = lItemData
            .bChecked = bChecked
            .bEnabled = bEnabled
            .bCreated = True
            .sKey = sKey
         End With
         pInsertNewMenuitem m_tMI(m_iMenuCount), lIndexBefore
         InsertItem = m_iMenuCount
      End If
   End If
End Function

Private Sub pInsertNewMenuitem(ByRef tMI As tMenuItem, ByVal lIndexBefore As Long)
Dim lPIndex As Long
Dim hMenu As Long
Dim lFlags As Long
Dim lPosition As Long
Dim lR As Long
Dim lH() As Long
Dim sItems() As String

   ' Find out where we're inserting:
   ' 1) is this inserted in the top level menu item?
   If (m_tMI(lIndexBefore).lParentId = 0) Then
      ' inserting into the top level:
      hMenu = GetMenu(m_hWndParent)
   Else
      ' inserting into an existing sub menu:
      lPIndex = plGetIndexForId(m_tMI(lIndexBefore).lParentId)
      If (lPIndex = 0) Then
         Debug.Print " **** Couldn't find parent... *** "
         Err.Raise 9, App.EXEName & ".cPopMenu"
         Exit Sub
      Else
         hMenu = m_tMI(lIndexBefore).hMenu
      End If
   End If
   If (hMenu <> 0) Then
      pHierarchyForIndex lIndexBefore, lH(), sItems()
      lPosition = lH(UBound(lH)) - 1
      
      lFlags = plMenuFlags(m_iMenuCount)
      lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING Or MF_BYPOSITION
      lR = InsertMenuByLong(hMenu, lPosition, lFlags, tMI.lID, tMI.lID)
      If (lR = 0) Then
          Debug.Print "Failed to insert new Menu item"
      Else
         ' Store the hMenu for this item:
         tMI.hMenu = hMenu
      End If
   End If
End Sub

Public Sub EnsureMenuSeparators(ByVal hMenu As Long)
Dim i As Long
Dim lCount As Long
   
   For i = 1 To m_iMenuCount
      If (m_tMI(i).hMenu = hMenu) Then
         lCount = lCount + 1
      End If
   Next i
End Sub

Private Function plGetNewID() As Long
Dim lID As Long

    If (m_lLastMaxId < m_iMenuCount) Then
        m_lLastMaxId = m_iMenuCount
    Else
        m_lLastMaxId = m_lLastMaxId + 1
    End If
    lID = m_lLastMaxId
    Do Until (pbIDIsUnique(lID))
        lID = lID + 1
        m_lLastMaxId = lID
    Loop
    plGetNewID = lID
End Function

Private Function pbIDIsUnique(ByVal lID As Long) As Boolean
Dim bFound As Boolean
Dim lMenu As Long

    For lMenu = 1 To m_iMenuCount
        If (m_tMI(lMenu).lActualID = lID) Then
            bFound = True
            Exit For
        End If
    Next lMenu
    pbIDIsUnique = Not (bFound)
End Function

Private Function psExtractAccelerator(ByVal sCaption As String)
Dim i As Long

   For i = 1 To Len(sCaption)
      If (Mid$(sCaption, i, 1) = "&") Then
         If (i < Len(sCaption)) Then
             psExtractAccelerator = UCase$(Mid$(sCaption, (i + 1), 1))
         End If
         Exit For
      End If
   Next i
End Function

Private Sub pAddNewMenuItem(ByRef tMI As tMenuItem)
Dim tMII As MENUITEMINFO
Dim hMenu As Long
Dim lPIndex As Long
Dim lFlags As Long
Dim lR As Long
Dim bTopLevel As Boolean

    ' Find out where we're adding this item:
    With tMI
      If (.lParentId = 0) Then
          ' This is a new top level menu item:
          hMenu = GetMenu(m_hWndParent)
          bTopLevel = True
      Else
         ' We are adding to an existing menu:
         ' First we need to determine if there is already a sub menu for the parent item:
         lPIndex = plGetIndexForId(tMI.lParentId)
         If (lPIndex = 0) Then
            Debug.Print " *** Couldn't find parent... *** "
         Else
            ' Determine if the parent menu has a sub-menu:
            tMII.fMask = MIIM_SUBMENU
            tMII.cbSize = LenB(tMII)
            GetMenuItemInfo m_tMI(lPIndex).hMenu, m_tMI(lPIndex).lActualID, False, tMII
            hMenu = tMII.hSubMenu
            If (hMenu = 0) Then
               ' We don't have a sub menu for this item so we're
               ' going to have to add one:
               ' Debug.Print "Adding new sub-menu:"
               
               ' Create the new menu item and store it's handle so we can clear up
               ' again later:
               hMenu = CreatePopupMenu()
               If (hMenu = 0) Then
                   Debug.Print " *** Failed to create sub menu *** "
               Else
                  m_lSubMenuCount = m_lSubMenuCount + 1
                  ReDim Preserve m_hSubMenus(1 To m_lSubMenuCount) As tSubMenuItem
                  m_hSubMenus(m_lSubMenuCount).hMenu = hMenu
                  m_hSubMenus(m_lSubMenuCount).hSysMenuOwner = m_hLastMDIMenu
                  
                  ' Now set the parent item so it has a popup menu:
                  lFlags = plMenuFlags(lPIndex)
                  lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING
                  lFlags = lFlags Or MF_POPUP
                  lFlags = lFlags And Not MF_BYPOSITION Or MF_BYCOMMAND
                  
                  lR = ModifyMenuByLong(m_tMI(lPIndex).hMenu, m_tMI(lPIndex).lActualID, lFlags, hMenu, m_tMI(lPIndex).lActualID)
                  If (lR = 0) Then
                     Debug.Print "Failed to modify menu to add the sub menu " & GetLastError()
                  End If
                  
                  ' WHen you add a sub menu to an item, its id becomes the sub menu handle:
                  m_tMI(lPIndex).lActualID = hMenu
                  tMI.lParentId = hMenu
               End If
             End If
         End If
      End If
      
      If (hMenu <> 0) Then
         lFlags = plMenuFlags(m_iMenuCount)
         lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING Or MF_BYPOSITION
         lR = AppendMenuBylong(hMenu, lFlags, tMI.lID, tMI.lID)
         If (lR = 0) Then
            Debug.Print "Failed to add new Menu item"
         Else
            ' Store the hMenu for this item:
            .hMenu = hMenu
         End If
         If (bTopLevel) Then
            DrawMenuBar m_hWndParent
         End If
      End If
        
   End With
End Sub

Public Sub Clear()
   m_iMenuCount = 0
   Erase m_tMI
End Sub

Property Get Count() As Integer
   Count = m_iMenuCount
End Property

Private Sub pRemoveMenuItems(ByVal hMenuOwner As Long)
Dim lMenu As Long
Dim i As Long

   For lMenu = m_lSubMenuCount To 1 Step -1
      If (m_hSubMenus(lMenu).hSysMenuOwner = hMenuOwner) Or hMenuOwner = 0 Then
         DestroyMenu m_hSubMenus(lMenu).hMenu
         For i = lMenu + 1 To m_lSubMenuCount
            LSet m_hSubMenus(i - 1) = m_hSubMenus(i)
         Next i
         m_lSubMenuCount = m_lSubMenuCount - 1
      End If
   Next lMenu
End Sub

Private Function plMenuFlags(ByVal lIndex As Long)
Dim lFlags As Long

   With m_tMI(lIndex)
      If (.bChecked) Then
         lFlags = lFlags Or MF_CHECKED
      Else
         lFlags = lFlags Or MF_UNCHECKED
      End If
      If (.bEnabled) Then
         lFlags = lFlags Or MF_ENABLED
      Else
         lFlags = lFlags Or MF_GRAYED
      End If
      If (Trim$(.sCaption) = "-") Then
         lFlags = lFlags Or MF_SEPARATOR
      End If
      If (m_tMI(lIndex).bMenuBarBreak) Then
         lFlags = lFlags Or MF_MENUBARBREAK
      End If
      If (m_tMI(lIndex).bMenuBreak) Then
         lFlags = lFlags Or MF_MENUBREAK
      End If
   End With
   plMenuFlags = lFlags
End Function

Public Sub SubClassMenu(Optional ByVal oForm As Object = Nothing, Optional ByVal bLeaveTopLevelMenus As Boolean = False)
Dim hMenu As Long
Dim tVBInfo() As tVBMenuInfo
Dim iVBMenuCount As Long
Dim i As Long
Dim lIndex As Long
Dim ctl As Control

   Clear
   m_bLeaveTopLevel = bLeaveTopLevelMenus
   If (m_hWndParent <> 0) Then
   
      If Not (oForm Is Nothing) Then
         ' Loop through the form object to find the menus.
         ' Store their caption and name.  We use this to
         ' set the key and tag in the internal array
         ' based on their name:
         For Each ctl In oForm.Controls
            If (TypeOf ctl Is Menu) Then
               iVBMenuCount = iVBMenuCount + 1
               ReDim Preserve tVBInfo(1 To iVBMenuCount) As tVBMenuInfo
               tVBInfo(iVBMenuCount).sName = ctl.Name
               tVBInfo(iVBMenuCount).sCaption = ctl.Caption
               tVBInfo(iVBMenuCount).sTag = ctl.Tag
               On Error Resume Next
               lIndex = ctl.Index
               If (Err.Number = 0) Then
                   tVBInfo(iVBMenuCount).bHasIndex = True
                   tVBInfo(iVBMenuCount).iIndex = lIndex
               End If
               Err.Clear
               tVBInfo(iVBMenuCount).bUsed = ctl.Visible
            End If
         Next ctl
      End If
      
      hMenu = GetMenu(m_hWndParent)
      pUpdateMenuItems hMenu, 0, False, bLeaveTopLevelMenus
      
      ' Now try to associate VB menus with the ones we've just updated:
      If (iVBMenuCount > 0) Then
         i = 0
         For lIndex = 1 To m_iMenuCount
            i = i + 1
            Do While Not (tVBInfo(i).bUsed)
               i = i + 1
               If (i > iVBMenuCount) Then
                  Exit Do
               End If
            Loop
            If (i > iVBMenuCount) Then
               Exit For
            End If
            ' These should match!
            ' Debug.Print tVBInfo(i).sCaption, m_tMI(lIndex).sCaption
            m_tMI(lIndex).sKey = tVBInfo(i).sName
            If (tVBInfo(i).bHasIndex) Then
               m_tMI(lIndex).sKey = m_tMI(lIndex).sKey & "(" & tVBInfo(i).iIndex & ")"
            End If
            m_tMI(lIndex).sTag = tVBInfo(i).sTag
         Next lIndex
      End If
        
      ' Cache the handle to the menu we've just subclassed
      On Error Resume Next
      If (m_hWndMDIClient = 0) Then
         If (TypeOf UserControl.Parent Is MDIForm) Then
            If (Err.Number = 0) Then
               m_hWndMDIClient = (GetWindow(m_hWndParent, GW_CHILD))
            End If
         End If
         On Error GoTo 0
         If (m_hWndMDIClient <> 0) Then
            m_hLastMDIMenu = GetMenu(m_hWndParent)
            AttachMessage Me, m_hWndMDIClient, WM_MDISETMENU
         End If
      End If
         
      ' Draw the menu:
      DrawMenuBar m_hWndParent
   End If
End Sub

Public Sub CheckForNewItems()
Dim i As Long
Dim iActualIndex As Long
   
   ' Initialise check for all items relevant:
   For i = 1 To m_iMenuCount
      m_tMI(i).bIsPresent = False
   Next i
   
   ' Recursively check through the menus
   ' for new items, ticking off all those
   ' items that are in the menu:
   pCheckForNew GetMenu(m_hWndParent), 0
   
   ' Strip out unused items:
   For i = 1 To m_iMenuCount
      If (m_tMI(i).bIsPresent) Then
         iActualIndex = iActualIndex + 1
         If (iActualIndex <> i) Then
            LSet m_tMI(iActualIndex) = m_tMI(i)
         End If
      End If
   Next i
   If (iActualIndex <> m_iMenuCount) Then
      m_iMenuCount = iActualIndex
      ReDim Preserve m_tMI(1 To m_iMenuCount) As tMenuItem
   End If
   
End Sub

Private Function pCheckForNew(ByVal hMenu As Long, ByVal lParentId As Long)
Dim lCount As Long
Dim lMenu As Long
Dim hSubMenu As Long
Dim tMI As MENUITEMINFO
Dim lIndex As Long
Dim sCaption As String
Dim sKey As String
Dim iIcon As Long
Dim lItemData As Long
Dim lFlags As Long
Dim lR As Long
Dim bDontSubClass As Boolean
Dim sTag As String
Dim sHelptext As String

   
   lCount = GetMenuItemCount(hMenu)
   
   For lMenu = 1 To lCount
      tMI.fMask = MIIM_ID Or MIIM_SUBMENU
      tMI.cbSize = Len(tMI)
      GetMenuItemInfo hMenu, lMenu - 1, 1, tMI
      lIndex = IndexForId(tMI.wID)
      
      If (lIndex = 0) Then
         ' We have a new menu - get all the details:
         tMI.fMask = MIIM_DATA Or MIIM_ID Or MIIM_STATE Or MIIM_SUBMENU Or MIIM_TYPE
         tMI.cch = 127
         tMI.dwTypeData = String$(128, 0)
         tMI.cbSize = LenB(tMI)
         GetMenuItemInfo hMenu, lMenu - 1, 1, tMI
         sCaption = left$(tMI.dwTypeData, tMI.cch)
         If ((tMI.fType And MF_SEPARATOR) = MF_SEPARATOR) Then
            sCaption = "-"
         End If
         ' Now we want to add it to the internal array:
         sKey = ""
         iIcon = -1
         RaiseEvent RequestNewMenuDetails(sCaption, sKey, iIcon, lItemData, sHelptext, sTag)
         ' Add or insert the item as required:
         If (pbIsValidKey(sKey)) Then
            m_iMenuCount = m_iMenuCount + 1
            ReDim Preserve m_tMI(1 To m_iMenuCount) As tMenuItem
            With m_tMI(m_iMenuCount)
               .lID = tMI.wID
               .lActualID = tMI.wID
               pSetMenuCaption m_iMenuCount, sCaption, (sCaption = "-")
               .sAccelerator = psExtractAccelerator(sCaption)
               .sTag = sTag
               .sHelptext = sHelptext
               .lIconIndex = iIcon
               ' TODO
               .lParentId = lParentId
               .lItemData = lItemData
               .bChecked = ((tMI.fState And MFS_CHECKED) = MFS_CHECKED)
               .bEnabled = Not ((tMI.fState And MFS_DISABLED) = MFS_DISABLED)
               .bIsPresent = True
               .hMenu = hMenu
            End With
            
            ' Get the flag equivalent for this menu item:
            lFlags = plMenuFlags(m_iMenuCount)
            ' Set it to owner draw:
            lFlags = lFlags Or MF_OWNERDRAW
            ' Ensure the string flag is removed:
            lFlags = lFlags And Not MF_STRING
            ' If there is a popup menu, make sure we keep it:
            If (tMI.hSubMenu <> 0) Then
                lFlags = lFlags Or MF_POPUP
            End If
            ' Modifying by position:
            lFlags = lFlags Or MF_BYPOSITION
            
            bDontSubClass = False
            ' Now set the menu to the owner draw version:
            If (hMenu = GetMenu(m_hWndParent)) Then
                If m_bLeaveTopLevel Then
                  bDontSubClass = True
                End If
            End If
            If Not (bDontSubClass) Then
               lR = ModifyMenuByLong(hMenu, (lMenu - 1), lFlags, m_tMI(m_iMenuCount).lActualID, m_tMI(m_iMenuCount).lActualID)
            End If
         End If
         lIndex = m_iMenuCount
      Else
         ' Mark as present:
         m_tMI(lIndex).bIsPresent = True
      End If
      ' Recurse sub-menus:
      If (tMI.hSubMenu <> 0) Then
         pCheckForNew tMI.hSubMenu, lIndex
      End If
   Next lMenu
End Function

Public Property Get MenuItemsPerScreen() As Long
Dim tWR As RECT
Dim lR As Long

   ' Get the available screen height
   lR = SystemParametersInfo(SPI_GETWORKAREA, 0, tWR, 0)
   If (lR = 0) Then
      ' Call failed - just use standard screen:
      tWR.tOp = 0
      tWR.Bottom = Screen.Height \ Screen.TwipsPerPixelY
   End If
   MenuItemsPerScreen = (tWR.Bottom - tWR.tOp) \ m_lMenuItemHeight
End Property

Public Sub UnsubclassMenu()
Dim i As Long

   Debug.Print "Unsubclass Menu"
   pRemoveMenuItems m_hLastMDIMenu
End Sub

Private Sub pUpdateMenuItems(ByVal hMenu As Long, ByVal lParentId As Long, ByVal bUpdate As Boolean, _
                             ByVal bLeaveTopLevelMenus As Boolean)
Dim lCount As Long
Dim lMenu As Long
Dim hSubMenu As Long
        
    lCount = GetMenuItemCount(hMenu)
    For lMenu = 1 To lCount
        pAddMenuItem hMenu, lMenu, lParentId, bUpdate, bLeaveTopLevelMenus
        hSubMenu = GetSubMenu(hMenu, (lMenu - 1))
            'Debug.Print hSubMenu
        If (hSubMenu <> 0) Then
            ' Recurse for the sub menus:
            pUpdateMenuItems hSubMenu, hSubMenu, bUpdate, bLeaveTopLevelMenus
        End If
    Next lMenu
End Sub
Private Sub pAddMenuItem(ByVal hMenu As Long, ByVal lPosition As Long, ByVal lParentId As Long, _
                         ByVal bUpdate As Boolean, ByVal bLeaveTopLevelMenus As Boolean)
Dim tMI As MENUITEMINFO
Dim lFlags As Long
Dim lR As Long
Dim bTopMenu As Boolean
Dim lIndex As Long
Dim lID As Long
Dim bAlreadyHave As Boolean
    
    bTopMenu = (lParentId = 0)
    
    ' Get information about the current menu item:
    ' Do we already have this menu?
    If (bUpdate) Then
        lID = GetMenuItemID(hMenu, (lPosition - 1))
        lIndex = IndexForId(lID)
        If (lIndex > 0) Then
            bAlreadyHave = True
        End If
    End If
    
    If Not (bAlreadyHave) Then
        tMI.fMask = MIIM_DATA Or MIIM_ID Or MIIM_STATE Or MIIM_SUBMENU Or MIIM_TYPE
        tMI.cch = 127
        tMI.dwTypeData = String$(128, 0)
        tMI.cbSize = LenB(tMI)
        GetMenuItemInfo hMenu, (lPosition - 1), 1, tMI
            
        ' Add a this item to the internal menu item array:
        m_iMenuCount = m_iMenuCount + 1
        ReDim Preserve m_tMI(1 To m_iMenuCount) As tMenuItem
        lIndex = m_iMenuCount
    End If
    
    With m_tMI(lIndex)
        .bIsAVBMenu = True
        .lIconIndex = -1     ' Start off with no icon
        .bChecked = ((tMI.fState And MFS_CHECKED) = MFS_CHECKED)
        .bEnabled = Not ((tMI.fState And MFS_DISABLED) = MFS_DISABLED)
        .hMenu = hMenu
        .lID = tMI.wID
        .lActualID = tMI.wID
        .lParentId = lParentId
    End With
    pSetMenuCaption lIndex, left$(tMI.dwTypeData, tMI.cch), ((tMI.fType And MF_SEPARATOR) = MF_SEPARATOR)
        
    ' Get the flag equivalent for this menu item:
    lFlags = plMenuFlags(lIndex)
    ' Set it to owner draw:
    lFlags = lFlags Or MF_OWNERDRAW
    ' Ensure the string flag is removed:
    lFlags = lFlags And Not MF_STRING
    ' If there is a popup menu, make sure we keep it:
    If (tMI.hSubMenu <> 0) Then
        lFlags = lFlags Or MF_POPUP
    End If
    ' Modifying by position:
    lFlags = lFlags Or MF_BYPOSITION
    
    ' Now set the menu to the owner draw version:
    If (bTopMenu) Then
        If bLeaveTopLevelMenus Then
            Exit Sub
        End If
    End If
    lR = ModifyMenuByLong(hMenu, (lPosition - 1), lFlags, m_tMI(lIndex).lActualID, m_tMI(lIndex).lActualID)
        ' This really shouldn't happen:
        If (lR = 0) Then
            Debug.Print "ModifyMenu failed:" & GetLastError()
        End If
End Sub

Private Sub pSetMenuCaption(ByVal iItem As Long, ByVal sCaption As String, ByVal bSeparator As Boolean)
Dim sCap As String
Dim sShortCut As String
Dim iPos As Long

   If (bSeparator) Then
       m_tMI(iItem).sCaption = "-"
   Else
      ' Check if this menu item will have a menu bar break:
      pParseCaption sCaption, "|", m_tMI(iItem).bMenuBarBreak
      ' Check if this menu item will be on the same line as
      ' the last one:
      pParseCaption sCaption, "^", m_tMI(iItem).bMenuBreak
      
      ' Check if we have a shortcut to the menu item:
      iPos = InStr(sCaption, vbTab)
      If (iPos <> 0) Then
          sCap = left$(sCaption, (iPos - 1))
          ' Extract the ctrl key item:
          sShortCut = Mid$(sCaption, (iPos + 1))
          pParseMenuShortcut iItem, sShortCut
      Else
          sCap = sCaption
      End If
      m_tMI(iItem).sAccelerator = psExtractAccelerator(sCap)
      m_tMI(iItem).sCaption = sCap
    End If
End Sub

Private Sub pParseCaption(ByRef sCaption As String, ByVal sToken As String, ByRef bFlag As Boolean)
Dim iPos As Long
Dim iPos2 As Long
Dim sCap As String

   iPos = InStr(sCaption, sToken)
   If (iPos <> 0) Then
      ' Check for double token (i.e. interpret as untokenised character):
      iPos2 = InStr(sCaption, sToken & sToken)
      If (iPos2 <> 0) Then
         bFlag = False
         If (iPos2 > 1) Then
            sCap = left$(sCaption, iPos - 1)
         End If
         If (iPos2 + 1 < Len(sCaption)) Then
            sCap = sCap & Mid$(sCaption, iPos2 + 1)
         End If
      Else
         bFlag = True
         If (iPos > 1) Then
            sCap = left$(sCaption, iPos - 1)
         End If
         If (iPos < Len(sCaption)) Then
            sCap = sCap & Mid$(sCaption, iPos + 1)
         End If
         sCaption = sCap
      End If
   Else
      bFlag = False
   End If
End Sub

Private Sub pParseMenuShortcut(ByVal iItem As Long, ByVal sShortCut As String)
Dim bNotFKey As Boolean
Dim iPos As Integer
Dim iLen As Integer
Dim sKey As String
Dim SkeyNum As String

    m_tMI(iItem).iShortCutShiftMask = 0
    m_tMI(iItem).iShortCutShiftKey = 0
    m_tMI(iItem).sShortCutDisplay = sShortCut
    
    If (sShortCut <> "") Then
        If (InStr(sShortCut, "Ctrl")) Then
            m_tMI(iItem).iShortCutShiftMask = vbCtrlMask
            bNotFKey = True
        End If
        If (InStr(sShortCut, "Shift")) Then
            m_tMI(iItem).iShortCutShiftMask = m_tMI(iItem).iShortCutShiftMask Or vbShiftMask
            bNotFKey = True
        End If
        
        If (bNotFKey) Then
            ' Find the last + and get the key:
            iLen = Len(sShortCut)
            iPos = iLen
            Do While Mid$(sShortCut, iPos, 1) <> "+" And iPos > 1
                iPos = iPos - 1
            Loop
            sKey = Mid$(sShortCut, iPos)
            If (Len(sKey) = 1) Then
                m_tMI(iItem).iShortCutShiftKey = Asc(sKey)
            Else
                ' Check for F key, Space, Backspace, Del
            End If
        Else
            ' Parse the Fkey:
            iPos = InStr(sShortCut, "F")
            If (iPos <> 0) Then
                SkeyNum = Mid$(sShortCut, (iPos + 1))
                m_tMI(iItem).iShortCutShiftKey = vbKeyF1 + Val(SkeyNum) - 1
            End If
        End If
    End If
End Sub
    
Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    m_emr = emrPreprocess
End Property

Private Property Get ISubclass_MsgResponse() As EMsgResponse
   ISubclass_MsgResponse = emrPreprocess
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lMenuId As Long, hMenu As Long, lItem As Long
Dim lMenuCount As Long
Dim lHiWord As Long
Dim bEnabled As Boolean, bSeparator As Boolean
Dim bFound As Boolean
Dim bNoDefault As Boolean
Dim iChar As Integer
Dim lFlag As Long
Dim i As Long, iIndex As Long, iNewIndex As Long

   Select Case iMsg
   
   ' Handle Menu Select events:
   Case WM_MENUSELECT
      ' Extract the menu id and flags for the selected
      ' menu item:
      lHiWord = wParam \ &H10000
      lMenuId = wParam And &HFFFF&
      
      Debug.Print lHiWord, lMenuId
      
      ' MenuId 0 corresponds to a separator on the system
      ' menu:
      'If (lMenuId <> 0) Then
          
          ' Extract separator & enabled/disabled from the flags
          ' stored in the High Word of wParam:
          bSeparator = ((lHiWord And MF_SEPARATOR) = MF_SEPARATOR)
          bEnabled = ((lHiWord And MF_DISABLED) = MF_DISABLED) Or ((lHiWord And MF_GRAYED) = MF_GRAYED)
          
          ' Menu handle is passed in as lParam:
          hMenu = lParam
          
          ' Now check if the message is a menu item higlight,
          ' or whether it is indicating exit from the menu:
          lMenuCount = GetMenuItemCount(hMenu)
          For lItem = 0 To lMenuCount - 1
              If (lMenuId = GetMenuItemID(hMenu, lItem)) Then
                  bFound = True
                  Exit For
              End If
          Next lItem
          
          ' Raise a highlight or menu exit as required:
          If (bFound) Then
              RaiseHighlightEvent lMenuId
          Else
              RaiseMenuExitEvent
          End If
          
      'End If
      
      ' Let the MENU_SELECT event filter through wherever
      ' else it is going:
      m_emr = emrPostProcess
       
        
   ' Handle menu click events:
   Case WM_COMMAND
      Debug.Print "Got a WM_COMMAND"
      
      ' Commands from menus are identified by an lParam of 0
      ' (otherwise it is set the hWnd of the control):
      If (lParam = 0) Then
          ' Low order word of the wParam item is the menu item id:
          lMenuId = (wParam And &HFFFF&)
          
          Debug.Print "ID: " & lMenuId
          If (RaiseClickEvent(lMenuId)) Then
              ' Don't send on the WM_COMMAND if the item
              ' wasn't a VB menu, it might interfere
              ' with some other control items!
              m_emr = emrConsume
          Else
              ' Otherwise allow the message to parse through
              ' to the click event on the VB menu so your old
              ' code continues to work:
              m_emr = emrPostProcess
          End If
      Else
          m_emr = emrPostProcess
      End If
        
    
   ' Handle system menu click events:
   Case WM_SYSCOMMAND
      'Debug.Print "Got a SYSCOMMAND item"
      
      ' Check if the item is a system menu command:
      If (pbIdIsSysMenuId(wParam)) Then
          ' If it is, send the event:
          RaiseEvent SystemMenuClick(wParam)
      End If
      
      ' Always let the message do its normal work:
      m_emr = emrPostProcess
    
   ' Draw Menu items:
   Case WM_DRAWITEM
      'Debug.Print "Got a draw item",lParam, wParam
      DrawItem lParam, wParam
        
    
   ' Measure Menu items prior to drawing them:
   Case WM_MEASUREITEM
      ' Debug.Print "Measure item"
      Dim tMis As MEASUREITEMSTRUCT
      CopyMemory tMis, ByVal lParam, Len(tMis)
      If tMis.CtlType = 1 Then
          ' Get the required width & height:
          MeasureItem tMis.itemID, tMis.itemWidth, tMis.itemHeight
          ' Put the new items back into the structure:
          CopyMemory ByVal lParam, tMis, Len(tMis)
          ISubclass_WindowProc = 1
          m_emr = emrConsume
      Else
          m_emr = emrPreprocess
      End If
    
   ' Handle accelerator (&key) messages in the menu:
   Case WM_MENUCHAR
      ' Check that this is my menu:
      lFlag = wParam \ &H10000
      If ((lFlag And MF_SYSMENU) <> MF_SYSMENU) Then
         hMenu = lParam
         iChar = (wParam And &HFFFF&)
         ' Debug.Print hMenu, Chr$(iChar)
         ' See if this corresponds to an accelerator on the menu:
         ISubclass_WindowProc = plParseMenuChar(hMenu, iChar)
         m_emr = emrConsume
      End If
        
   Case WM_INITMENUPOPUP
      ' Check the sys menu flag:
      If (lParam \ &H10000) > 0 Then
          ' System menu.
      Else
          hMenu = wParam
          ' Find the item which is the parent
          ' of this popup menu:
          RaiseInitMenuEvent hMenu
      End If
      m_emr = emrPostProcess
        
   Case WM_MDISETMENU
      If (wParam <> 0) Then
         If (wParam <> m_hLastMDIMenu) Then
            Debug.Print "New MDI Menu!"

            ' Store this menu:
            For i = 1 To m_iStoreMenuCount
               If (m_cStoreMenus(i).hMenu = m_hLastMDIMenu) Then
                  iIndex = i
               ElseIf (m_cStoreMenus(i).hMenu = wParam) Then
                  iNewIndex = i
               End If
            Next i
            If (iIndex = 0) Then
               m_iStoreMenuCount = m_iStoreMenuCount + 1
               ReDim Preserve m_cStoreMenus(1 To m_iStoreMenuCount) As cStoreMenu
               Set m_cStoreMenus(m_iStoreMenuCount) = New cStoreMenu
               iIndex = m_iStoreMenuCount
               m_cStoreMenus(iIndex).hMenu = m_hLastMDIMenu
            End If
            Debug.Print "Storing menu in index ", iIndex
            m_cStoreMenus(iIndex).Store m_tMI(), m_iMenuCount
            
            m_hLastMDIMenu = wParam
               
            ' If we have the new menu stored, then restore
            ' that information, otherwise raise an event
            ' saying we have got the changed menu for the
            ' first time:
            If (iNewIndex > 0) Then
               Debug.Print "Restoring menu from index ", iIndex
               m_cStoreMenus(iNewIndex).Restore m_tMI(), m_iMenuCount
            Else
               Debug.Print "Requesting new menu"
               Erase m_tMI
               m_iMenuCount = 0
               RaiseEvent NewMDIMenu
            End If
         End If
      End If
    
   Case WM_WININICHANGE
      ' First ensure we have the correct font:
      m_cNCM.ClearUp
      pSelectMenuFont
      ' Now replace every menu item so the new sizes of the
      ' the menu items are correctly displayed...
      For i = 1 To m_iMenuCount
         ReplaceItem i
      Next i
      
      ' Now allow the event to be responded
      ' to in the form
      RaiseEvent WinIniChange
      ' Make sure we pass the message on for
      ' default processing!
   
   End Select
End Function

Private Sub RaiseInitMenuEvent(ByVal hMenu As Long)
Dim lIndex As Long
Dim lParentId As Long
Dim bFound As Boolean

    ' Firstly, we need to find the index of an item
    ' in hMenu:
    For lIndex = m_iMenuCount To 1 Step -1
        If (m_tMI(lIndex).hMenu = hMenu) Then
            lParentId = m_tMI(lIndex).lParentId
            bFound = True
        End If
        If (bFound) Then
            If (m_tMI(lIndex).lActualID = lParentId) Then
                RaiseEvent InitPopupMenu(lIndex)
                Exit For
            End If
        End If
    Next lIndex
End Sub

Private Sub pCreateSubClass(hWndA As Long)
   AttachMessage Me, hWndA, WM_MENUSELECT
   AttachMessage Me, hWndA, WM_MEASUREITEM
   AttachMessage Me, hWndA, WM_DRAWITEM
   AttachMessage Me, hWndA, WM_COMMAND
   AttachMessage Me, hWndA, WM_MENUCHAR
   AttachMessage Me, hWndA, WM_SYSCOMMAND
   AttachMessage Me, hWndA, WM_INITMENUPOPUP
   AttachMessage Me, hWndA, WM_WININICHANGE
End Sub

Private Sub pDestroySubClass()
   If (m_hWndParent <> 0) Then
      DetachMessage Me, m_hWndParent, WM_MENUSELECT
      DetachMessage Me, m_hWndParent, WM_MEASUREITEM
      DetachMessage Me, m_hWndParent, WM_DRAWITEM
      DetachMessage Me, m_hWndParent, WM_COMMAND
      DetachMessage Me, m_hWndParent, WM_MENUCHAR
      DetachMessage Me, m_hWndParent, WM_SYSCOMMAND
      DetachMessage Me, m_hWndParent, WM_INITMENUPOPUP
      DetachMessage Me, m_hWndParent, WM_WININICHANGE
      If (m_hWndMDIClient <> 0) Then
         DetachMessage Me, m_hWndMDIClient, WM_MDISETMENU
      End If
   End If
   m_hWndParent = 0
End Sub

Private Sub UserControl_Initialize()
    Debug.Print "Initialise"
    m_lNextSysMenuID = WM_MENUBASE
    m_lMenuItemHeight = 22
    m_lIconSize = 16
    Set m_cNCM = New cNCMetrics
    m_lBitmapW = picTest.ScaleWidth \ Screen.TwipsPerPixelX
    m_lBitmapH = picTest.ScaleHeight \ Screen.TwipsPerPixelY - 1
End Sub

Private Sub UserControl_InitProperties()
  m_hColor = m_hColorDef
  m_bColor = m_bColorDef
  m_fColor = m_fColorDef
  m_fHColor = m_fHColorDef
  m_ShadowXPHighlight = m_ShadowDef
  m_ShadowTop = m_ShadowTopDef
  m_LineColor = m_LineColorDef
  m_ShadowColor = m_ShadowColorDef
End Sub

Private Sub UserControl_Paint()
Dim lHDC As Long
Dim tR As RECT
    tR.Right = 40
    tR.Bottom = 40
    lHDC = UserControl.hdc
    DrawEdge lHDC, tR, EDGE_RAISED, BF_RECT
    BitBlt lHDC, 4, 4, 32, 32, m_HDC, 0, 0, SRCCOPY
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ' At the ReadProperties we are now sited
    ' and have a fully usable UserControl
    ' object.
    m_hColor = PropBag.ReadProperty("HighlightColor", m_hColorDef)
    m_bColor = PropBag.ReadProperty("BorderColor", m_bColorDef)
    HighlightCheckedItems = PropBag.ReadProperty("HighlightCheckedItems", True)
    TickIconIndex = PropBag.ReadProperty("TickIconIndex", -1)
    HighlightStyle = PropBag.ReadProperty("HighlightStyle", cspHighlightStandard)
    m_ShadowXPHighlight = PropBag.ReadProperty("ShadowXPHighlight", m_ShadowDef)
    m_ShadowTop = PropBag.ReadProperty("ShadowXPHighlightTopMenu", m_ShadowTopDef)
    m_fColor = PropBag.ReadProperty("ForeColor", m_fColorDef)
    m_fHColor = PropBag.ReadProperty("HForeColor", m_fHColorDef)
    m_LineColor = PropBag.ReadProperty("LineColor", m_LineColorDef)
    m_ShadowColor = PropBag.ReadProperty("ShadowColor", m_ShadowColorDef)
    If (UserControl.Ambient.UserMode) Then
        ' Only do the subclassing stuff whilst
        ' we are in run mode.  Makes it easier
        ' to debug, if nothing else...
        m_hWndParent = UserControl.Parent.hwnd
        ' Make a HDC to allow us to evaluate the
        ' size of menu items.
        m_HDC = CreateCompatibleDC(UserControl.hdc)
        ' Select the menu font into it:
        pSelectMenuFont
        ' Get the dither bitmap from the resource file:
        m_hBMPDither = LoadImageByNum(App.hInstance, 49, IMAGE_BITMAP, 0, 0, LR_LOADMAP3DCOLORS)
        ' Start subclassing:
        Debug.Print "Start subclassing"
        pCreateSubClass m_hWndParent
        
        ' Background picture...
    Else
        ' We don't draw when we're in run mode so
        ' only do it when not in run mode:
        pMakeDisplay
    End If
End Sub

Private Sub pSelectMenuFont()
Dim tM As RECT
   ' If we have already selected the font,
   ' then remove it from the DC:
   If (m_hFntOld <> 0) Then
       SelectObject m_HDC, m_hFntOld
   End If
   ' Get the metrics.  This will delete
   ' the hFont for menu:
   m_cNCM.GetMetrics
   ' Select the latest version of the menu font
   ' into the DC, storing what was there before:
   m_hFntOld = SelectObject(m_HDC, m_cNCM.FontHandle(MenuFOnt))
   
   ' Determine what height to make the menu items:
   DrawText m_HDC, "yY", 2, tM, DT_CALCRECT
   If (tM.Bottom - tM.tOp) > m_lIconSize + 6 Then
       m_lMenuItemHeight = tM.Bottom - tM.tOp + 6
   Else
       m_lMenuItemHeight = m_lIconSize + 6
   End If
   DrawMenuBar m_hWndParent
End Sub

Private Sub pMakeDisplay()
Dim hInst As Long

    m_HDC = CreateCompatibleDC(UserControl.hdc)
    If (m_HDC <> 0) Then
        hInst = App.hInstance
        m_hBmp = LoadImageByNum(hInst, 48, IMAGE_BITMAP, 0, 0, LR_LOADMAP3DCOLORS)
        If (m_hBmp <> 0) Then
            m_hBMPOLd = SelectObject(m_HDC, m_hBmp)
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 40 * Screen.TwipsPerPixelX
    UserControl.Height = 40 * Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_Terminate()
    Debug.Print "Terminate"
    ' Remove any new menus we have created:
    Clear
    pRemoveMenuItems 0
    ' Destroy the sub class:
    pDestroySubClass
    ' Remove the graphics:
    If (m_HDC <> 0) Then
        If (m_hBmp <> 0) Then
            SelectObject m_HDC, m_hBMPOLd
            DeleteObject m_hBmp
        End If
        If (m_hFntOld <> 0) Then
            SelectObject m_HDC, m_hFntOld
        End If
        DeleteObject m_HDC
    End If
    If (m_hBMPDither <> 0) Then
        DeleteObject m_hBMPDither
    End If
    ' Clear the non-client object, removing any fonts:
    Set m_cNCM = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "HighlightCheckedItems", HighlightCheckedItems, True
    PropBag.WriteProperty "TickIconIndex", TickIconIndex, -1
    PropBag.WriteProperty "HighlightStyle", HighlightStyle, cspHighlightStandard
    PropBag.WriteProperty "HighlightColor", HighlightColor, m_hColorDef
    PropBag.WriteProperty "BorderColor", BorderColor, m_bColorDef
    PropBag.WriteProperty "ForeColor", ForeColor, m_fColorDef
    PropBag.WriteProperty "HForeColor", HighlightForeColor, m_fHColorDef
    PropBag.WriteProperty "ShadowXPHighlight", ShadowXPHighlight, m_ShadowDef
    PropBag.WriteProperty "ShadowXPHighlightTopMenu", ShadowXPHighlightTopMenu, m_ShadowTopDef
    PropBag.WriteProperty "LineColor", LineColor, m_LineColorDef
    PropBag.WriteProperty "ShadowColor", ShadowColor, m_ShadowColorDef
    ' ... background picture ...
End Sub

Private Function GetPicture(imgListIndex As Long)
  
End Function
