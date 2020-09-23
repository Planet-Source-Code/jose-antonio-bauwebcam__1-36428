VERSION 5.00
Begin VB.UserControl AVIFiles 
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   585
   InvisibleAtRuntime=   -1  'True
   Picture         =   "AVIFiles.ctx":0000
   ScaleHeight     =   585
   ScaleWidth      =   585
   Begin VB.Image imgAVI 
      Height          =   1545
      Left            =   1410
      Top             =   750
      Width           =   1755
   End
End
Attribute VB_Name = "AVIFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--> Control para creación de ficheros AVI
Option Explicit

Private Declare Function SetRect Lib "user32.dll" _
    (ByRef lprc As AVI_RECT, ByVal xLeft As Long, ByVal yTop As Long, _
     ByVal xRight As Long, ByVal yBottom As Long) As Long 'BOOL

Private colURL As Collection 'Colección de URLs
Private colURLConverted As Collection 'Colección de las URLs de las imágenes convertidas
Private strLastError As String

Public Sub addFile(ByVal strFileName As String)
'--> Añade un nombre de fichero a la colección
  colURL.Add strFileName
End Sub

Public Function createFile(ByVal hndWnd As Long, ByVal strFileName As String) As Boolean
'--> Crea el fichero AVI de salida
Dim InitDir As String
Dim szOutputAVIFile As String
Dim res As Long
Dim pfile As Long
Dim bmp As cDIB
Dim ps As Long
Dim psCompressed As Long
Dim strhdr As AVI_STREAM_INFO
Dim BI As BITMAPINFOHEADER
Dim opts As AVI_COMPRESS_OPTIONS
Dim pOpts As Long
Dim i As Long

  If colURL.Count > 0 Then
    convertFiles
    res = AVIFileOpen(pfile, strFileName, OF_WRITE Or OF_CREATE, 0&)
    If (res <> AVIERR_OK) Then GoTo error
    Set bmp = New cDIB
    If bmp.CreateFromFile(colURLConverted.Item(1)) <> True Then
        MsgBox "No File!", vbExclamation + vbOKOnly, "KO!"
        GoTo error
    End If
    With strhdr
        .fccType = mmioStringToFOURCC("vids", 0&)
        .fccHandler = 0&
        .dwScale = 1
        .dwRate = 10
        .dwSuggestedBufferSize = bmp.SizeImage
        SetRect .rcFrame, 0, 0, 300, 300 'TamanoCuadroW, TamanoCuadroH
    End With
    If strhdr.dwRate < 1 Then strhdr.dwRate = 1
    If strhdr.dwRate > 30 Then strhdr.dwRate = 30
    res = AVIFileCreateStream(pfile, ps, strhdr)
    If (res <> AVIERR_OK) Then GoTo error
    pOpts = VarPtr(opts)
    res = AVISaveOptions(hWnd, ICMF_CHOOSE_KEYFRAME Or ICMF_CHOOSE_DATARATE, 1, ps, pOpts)
    If res <> 1 Then
        Call AVISaveOptionsFree(1, pOpts)
        GoTo error
    End If
    res = AVIMakeCompressedStream(psCompressed, ps, opts, 0&)
    If res <> AVIERR_OK Then GoTo error
    With BI
        .biBitCount = bmp.BitCount
        .biClrImportant = bmp.ClrImportant
        .biClrUsed = bmp.ClrUsed
        .biCompression = bmp.Compression
        .biHeight = bmp.Height
        .biWidth = bmp.Width
        .biPlanes = bmp.Planes
        .biSize = bmp.SizeInfoHeader
        .biSizeImage = bmp.SizeImage
        .biXPelsPerMeter = bmp.XPPM
        .biYPelsPerMeter = bmp.YPPM
    End With
    res = AVIStreamSetFormat(psCompressed, 0, ByVal bmp.PointerToBitmapInfo, bmp.SizeBitmapInfo)
    If (res <> AVIERR_OK) Then GoTo error
    For i = 1 To colURLConverted.Count
        bmp.CreateFromFile colURLConverted.Item(i)
        res = AVIStreamWrite(psCompressed, i, 1, bmp.PointerToBits, bmp.SizeImage, AVIIF_KEYFRAME, ByVal 0&, ByVal 0&)
        If res <> AVIERR_OK Then GoTo error
    Next i
  End If
  
error:
    Set bmp = Nothing
    If (ps <> 0) Then Call AVIStreamClose(ps)
    If (psCompressed <> 0) Then Call AVIStreamClose(psCompressed)
    If (pfile <> 0) Then Call AVIFileClose(pfile)
    Call AVIFileExit
    If (res <> AVIERR_OK) Then
        strLastError = "Error!" + Err.Description
    End If
    dropFiles
End Function

Private Sub convertFiles()
'--> Convierte las imágenes a bmp
Dim intIndex As Integer
Dim strFile As String

  'Inicializa la colección de ficheros convertidos
    Set colURLConverted = Nothing
    Set colURLConverted = New Collection
  'Convierte las imágenes
    For intIndex = 1 To colURL.Count
      On Error Resume Next
      'Carga la imagen
        Set imgAVI.Picture = LoadPicture(colURL.Item(intIndex))
      'Graba la imagen como bmp
        strFile = Replace(colURL.Item(intIndex), ".", "_") & ".bmp"
      'Graba el fichero
        If Err.Number = 0 Then
          SavePicture imgAVI.Picture, strFile
        End If
      'Libera la memoria
        Set imgAVI.Picture = Nothing
      'Añade el nuevo nombre a la colección de convertidos
        If Err.Number = 0 Then
          colURLConverted.Add strFile
        End If
      'Quita los errores
        Err.Clear
    Next intIndex
End Sub

Private Sub dropFiles()
'--> Elimina los ficheros intermedios
Dim intIndex As Integer

  On Error Resume Next
  For intIndex = 1 To colURLConverted.Count
    Kill colURLConverted.Item(intIndex)
  Next intIndex
End Sub

Public Property Get ErrorMessage() As String
'--> Obtiene el último mensaje de error
  ErrorMessage = strLastError
End Property

Private Sub UserControl_Initialize()
  Set colURL = New Collection
  Set colURLConverted = New Collection
End Sub

Private Sub UserControl_Resize()
  Width = 585
  Height = 585
End Sub

Private Sub UserControl_Terminate()
  Set colURL = Nothing
  Set colURLConverted = Nothing
End Sub
