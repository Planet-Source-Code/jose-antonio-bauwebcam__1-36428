VERSION 5.00
Object = "{8EF5732E-26C2-4934-9070-EAE6628A143E}#1.0#0"; "O3_VER~1.OCX"
Begin VB.Form vTest 
   Caption         =   "Form1"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin O3_VerticalMenu2.O3_VerticalMenu O3_VerticalMenu1 
      Height          =   8355
      Left            =   1740
      TabIndex        =   0
      Top             =   180
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   5424
      MenusMax        =   11
      MenuCaption1    =   "Menu1"
      MenuItemsMax1   =   16
      MenuItemCaption12=   "Item2"
      MenuItemCaption13=   "Item3"
      MenuItemCaption14=   "Item4"
      MenuItemCaption15=   "Item5"
      MenuItemCaption16=   "Item6"
      MenuItemCaption17=   "Item7"
      MenuItemCaption18=   "Item8"
      MenuItemCaption19=   "Item9"
      MenuItemCaption110=   "Item10"
      MenuItemCaption111=   "Item11"
      MenuItemCaption112=   "Item12"
      MenuItemCaption113=   "Item13"
      MenuItemCaption114=   "Item14"
      MenuItemCaption115=   "Item15"
      MenuItemCaption116=   "Item16"
      MenuCaption2    =   "Menu2"
      MenuCaption3    =   "Menu3"
      MenuCaption4    =   "Menu4"
      MenuCaption5    =   "Menu5"
      MenuCaption6    =   "Menu6"
      MenuCaption7    =   "Menu7"
      MenuCaption8    =   "Menu8"
      MenuCaption9    =   "Menu9"
      MenuCaption10   =   "Menu10"
      MenuCaption11   =   "Menu11"
      BackColor       =   8421504
   End
End
Attribute VB_Name = "vTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_DblClick()
  MsgBox O3_VerticalMenu1.Version(True)
End Sub

