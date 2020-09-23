Attribute VB_Name = "kMod"
Option Explicit

Public Declare Function GetWindowRect Lib "User32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function DrawEdge Lib "User32" (ByVal hdc As Long, _
        qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

Public Type RECT
       Left As Long
       Top As Long
       Right As Long
       Bottom As Long
End Type

Public Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Src As Any, ByVal L As Long)
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function EnumResourceNamesA Lib "kernel32" (ByVal hModule As Long, ByVal lpszType As String, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
' Border Constants for the Edge parameter in the DrawEdge API
Public Const BDR_INNER = &HC
Public Const BDR_OUTER = &H3
Public Const BDR_RAISED = &H5
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKEN = &HA
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2

' Alternative EDGE styles (Combines the constants above)
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)

' Constants for the grfFlags parameter in the DrawEdge API
Public Const BF_FLAT = &H4000
Public Const BF_LEFT = &H1
Public Const BF_MONO = &H8000
Public Const BF_MIDDLE = &H800
Public Const BF_RIGHT = &H4
Public Const BF_SOFT = &H1000
Public Const BF_TOP = &H2
Public Const BF_ADJUST = &H2000
Public Const BF_BOTTOM = &H8
Public Const BF_DIAGONAL = &H10
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM _
               Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL _
               Or BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)

Public Function MakeLong(ByVal HIWORD As Integer, ByVal LOWORD As Integer) As Long

    MoveMemory MakeLong, HIWORD, 2
    MoveMemory ByVal VarPtr(MakeLong) + 2, LOWORD, 2
    
End Function

Public Function EnumResNameProc(ByVal hModule As Long, ByVal lpszType As Long, ByVal lpszName As Long, ByVal lParam As Long) As Long
        
        Dim Cbo As ComboBox
        
        On Error Resume Next
        
        MoveMemory Cbo, lParam, 4
        
        Cbo.AddItem CStr(lpszName)
        
        MoveMemory Cbo, 0&, 4
        
        DoEvents
        
        EnumResNameProc = 1
        
End Function
Public Function SubWndProc(ByVal Wnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        
        Dim myCtl As AVIShow, Ptr As Long
        
        Ptr = GetWindowLong(Wnd, GWL_USERDATA)
        MoveMemory myCtl, Ptr, 4
        
        SubWndProc = myCtl.s_WindowProc(Wnd, uMsg, wParam, lParam)
        
        MoveMemory myCtl, 0&, 4
        
End Function


