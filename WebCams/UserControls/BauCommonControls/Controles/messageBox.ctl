VERSION 5.00
Begin VB.UserControl messageBox 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   InvisibleAtRuntime=   -1  'True
   Picture         =   "messageBox.ctx":0000
   ScaleHeight     =   495
   ScaleWidth      =   495
End
Attribute VB_Name = "messageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--> Control para mostrar un cuadro de mensaje
Option Explicit

Public Enum typeMessage 'Tipos de mensajes
  MsgInformation = 0
  MsgQuestion
  MsgExclamation
  MsgQuestionCancel
End Enum

Public Enum typeResult 'Valores de retorno
  ResultYes = 0
  ResultNo
  ResultCancel
End Enum

Public Function showMessage(ByVal strMessage As String, ByVal strTitle As String, _
                            ByVal tmType As typeMessage) As typeResult
'--> Muestra un cuadro de mensaje
Dim frmMessage As New frmMsgBox

  With frmMessage
    .strMessage = strMessage
    .strTitle = strTitle
    .intType = tmType
    .Show vbModal
    showMessage = .intResult
  End With
  Set frmMessage = Nothing
End Function

Private Sub UserControl_Resize()
  Width = 495
  Height = 495
End Sub
