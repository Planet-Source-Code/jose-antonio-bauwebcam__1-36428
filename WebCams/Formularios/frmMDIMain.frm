VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMDIMain 
   BackColor       =   &H8000000C&
   Caption         =   "bauWebCams"
   ClientHeight    =   8250
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10890
   Icon            =   "frmMDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin BAUCommonControls.statusBar brStatus 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   9
      Top             =   7950
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   529
   End
   Begin VB.PictureBox picBackground 
      Align           =   1  'Align Top
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   0
      Picture         =   "frmMDIMain.frx":08CA
      ScaleHeight     =   15
      ScaleWidth      =   10890
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   10890
   End
   Begin ctlMenu.PopMenu ctlPopMenu 
      Left            =   6900
      Top             =   2850
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
   End
   Begin bauWebCams.innerWindow picPannel 
      Align           =   3  'Align Left
      Height          =   7005
      Left            =   1275
      TabIndex        =   3
      Top             =   375
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   12039
      Caption         =   "WebCams"
      Begin MSComctlLib.TreeView trvProject 
         Height          =   3495
         Index           =   0
         Left            =   150
         TabIndex        =   6
         Top             =   840
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   6165
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
         OLEDropMode     =   1
      End
      Begin MSComctlLib.TreeView trvProject 
         Height          =   3555
         Index           =   1
         Left            =   150
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   6271
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   353
         Style           =   7
         FullRowSelect   =   -1  'True
         Appearance      =   1
      End
      Begin MSComctlLib.TabStrip tabTree 
         Height          =   4095
         Left            =   60
         TabIndex        =   4
         Top             =   390
         Width           =   2685
         _ExtentX        =   4736
         _ExtentY        =   7223
         MultiRow        =   -1  'True
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Archivo"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Favoritos"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin BAUCommonControls.SpliterVertical splVertical 
         Height          =   120
         Left            =   1860
         TabIndex        =   7
         Top             =   1470
         Width           =   9.47745e8
         _ExtentX        =   1727049082
         _ExtentY        =   32880300
      End
   End
   Begin DownloadFile.InternetFile dwnImage 
      Left            =   4830
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin BAUCommonControls.messageBox messageBox 
      Left            =   4860
      Top             =   2820
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin MSComctlLib.ImageList imlButtons 
      Left            =   5610
      Top             =   2850
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":093D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":0D94
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":1208
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":1309
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":1756
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":1BB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":1CA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":2118
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":2596
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":267E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":2AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":2EF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":2FD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":346A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":3909
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlImages 
      Index           =   1
      Left            =   6240
      Top             =   2850
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":39FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":3B16
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":3C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":3D46
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":3E5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":3F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":5280
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":5398
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":54B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":55C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":56E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":57F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":5910
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":5A28
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":5E7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":62D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":6BAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":7484
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlImages 
      Index           =   0
      Left            =   6240
      Top             =   2220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":9C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":9D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":9E66
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":9F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":A096
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":A1AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":B4B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":B5D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":B6E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":B800
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":B918
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":BA30
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":BB44
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":BC58
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":C0AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":C500
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":CDDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":D6B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommon 
      Left            =   6960
      Top             =   2190
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin webLoadBanner.WebBanner wbnImage 
      Align           =   2  'Align Bottom
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   7380
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   1005
   End
   Begin BAUCommonControls.VerticalMenu vrtMnuProject 
      Align           =   3  'Align Left
      Height          =   7005
      Left            =   0
      TabIndex        =   2
      Top             =   375
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   12039
      MenusMax        =   3
      MenuCaption1    =   "Archivo"
      MenuItemsMax1   =   4
      MenuItemIcon11  =   "frmMDIMain.frx":FE66
      MenuItemCaption11=   "Nuevo"
      MenuItemKey11   =   "vrtMnuNew"
      MenuItemIcon12  =   "frmMDIMain.frx":10180
      MenuItemCaption12=   "Abrir"
      MenuItemKey12   =   "vrtMnuOpen"
      MenuItemIcon13  =   "frmMDIMain.frx":1049A
      MenuItemCaption13=   "Grabar"
      MenuItemKey13   =   "vrtMnuSave"
      MenuItemIcon14  =   "frmMDIMain.frx":107B4
      MenuItemCaption14=   "Buscar Web"
      MenuItemKey14   =   "vrtMnuSearchWeb"
      MenuCaption2    =   "Proyecto"
      MenuItemsMax2   =   5
      MenuItemIcon21  =   "frmMDIMain.frx":10ACE
      MenuItemCaption21=   "Carpeta"
      MenuItemKey21   =   "vrtMnuAddFolder"
      MenuItemIcon22  =   "frmMDIMain.frx":10DE8
      MenuItemCaption22=   "webCam"
      MenuItemKey22   =   "vrtMnuAddItem"
      MenuItemIcon23  =   "frmMDIMain.frx":11102
      MenuItemCaption23=   "Propiedades"
      MenuItemKey23   =   "vrtMnuProperties"
      MenuItemIcon24  =   "frmMDIMain.frx":1141C
      MenuItemCaption24=   "Eliminar"
      MenuItemKey24   =   "vrtMnuDrop"
      MenuItemIcon25  =   "frmMDIMain.frx":115F6
      MenuItemCaption25=   "Generar"
      MenuItemKey25   =   "vrtMnuGenerate"
      MenuCaption3    =   "Ayuda"
      MenuItemsMax3   =   3
      MenuItemIcon31  =   "frmMDIMain.frx":11910
      MenuItemCaption31=   "Ayuda"
      MenuItemKey31   =   "vrtMnuHelp"
      MenuItemIcon32  =   "frmMDIMain.frx":11C2A
      MenuItemCaption32=   "Acerca de ..."
      MenuItemKey32   =   "vrtMnuAbout"
      MenuItemIcon33  =   "frmMDIMain.frx":11F44
      MenuItemCaption33=   "Web"
      MenuItemKey33   =   "vrtMnuWeb"
   End
   Begin MSComctlLib.Toolbar tlbHerramientas 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "N1"
                  Text            =   "Proyecto"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "N2"
                  Text            =   "Carpeta"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "N3"
                  Text            =   "Elemento"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Abrir"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Guardar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cortar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copiar"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Pegar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Propiedades"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Generar"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Visitar web cámara"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ayuda"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ilsIcons 
      Left            =   5580
      Top             =   2190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":1225E
            Key             =   "NULL"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":12578
            Key             =   "OPEN"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":12892
            Key             =   "SAVE"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":12BAC
            Key             =   "PRINT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":12EC6
            Key             =   "CUT"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":131E0
            Key             =   "COPY"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":134FA
            Key             =   "PASTE"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":13814
            Key             =   "FIND"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":13B2E
            Key             =   "BACK"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":13E48
            Key             =   "NEXT"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":14162
            Key             =   "FAVE"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":1447C
            Key             =   "HELP"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":14796
            Key             =   "NET"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":14AB0
            Key             =   "FOLDER"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":14DCA
            Key             =   "DOCUMENT"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDIMain.frx":150E4
            Key             =   "TICK"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuNewProject 
         Caption         =   "&Nuevo proyecto ..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpenProject 
         Caption         =   "&Abrir proyecto ..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSaveProject 
         Caption         =   "&Guardar proyecto ..."
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSaveProjectAs 
         Caption         =   "Guardar &proyecto como ..."
      End
      Begin VB.Menu mnuFileSeparator7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Propiedades ..."
      End
      Begin VB.Menu mnuSeparatorFile2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpenCammera 
         Caption         =   "&Ver cámara ..."
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuViewWebWebcam 
         Caption         =   "Visitar web ca&mara"
      End
      Begin VB.Menu mnuSeparator5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Imprimir"
         Shortcut        =   ^P
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSeparatorFile3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMRUFile 
         Caption         =   "File1"
         Index           =   0
      End
      Begin VB.Menu mnuMRUFile 
         Caption         =   "File2"
         Index           =   1
      End
      Begin VB.Menu mnuMRUFile 
         Caption         =   "File3"
         Index           =   2
      End
      Begin VB.Menu mnuMRUFile 
         Caption         =   "File4"
         Index           =   3
      End
      Begin VB.Menu mnuMRUFile 
         Caption         =   "File5"
         Index           =   4
      End
      Begin VB.Menu mnuSeparatorFile4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Salir"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edición"
      Begin VB.Menu mnuAddFolder 
         Caption         =   "A&gregar carpeta ..."
      End
      Begin VB.Menu mnuAddItem 
         Caption         =   "Agregar &webCam ..."
      End
      Begin VB.Menu mnuEditSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Co&rtar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copiar"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Pegar"
      End
      Begin VB.Menu mnuEditSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDrop 
         Caption         =   "&Quitar"
      End
      Begin VB.Menu mnuEditSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCreateAvi 
         Caption         =   "Crear &vídeo"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Herramientas"
      Begin VB.Menu mnuVer 
         Caption         =   "&Ver"
         Begin VB.Menu mnuShowProject 
            Caption         =   "Ventana &Proyecto"
         End
         Begin VB.Menu mnuShowToolBar 
            Caption         =   "&Barra de herramientas"
         End
         Begin VB.Menu mnuShowBigIcons 
            Caption         =   "&Iconos grandes"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuShowVerticalMenu 
            Caption         =   "&Menú vertical"
         End
      End
      Begin VB.Menu mnuIdioma 
         Caption         =   "&Idioma"
         Begin VB.Menu mnuChangeIdioma 
            Caption         =   "&Español"
            Checked         =   -1  'True
            Index           =   1
         End
         Begin VB.Menu mnuChangeIdioma 
            Caption         =   "&Inglés"
            Index           =   2
         End
      End
      Begin VB.Menu mnuSearchWebCam 
         Caption         =   "&Buscar webCam"
      End
      Begin VB.Menu mnuToolsSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "&Configuración ..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Ventana"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "En cascada"
         Index           =   1
      End
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "Mosaico horizontal"
         Index           =   2
      End
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "Mosaico vertical"
         Index           =   3
      End
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "Organizar iconos"
         Index           =   4
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "A&yuda"
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "&Indice"
      End
      Begin VB.Menu mnuHelpTip 
         Caption         =   "&Truco del día ..."
      End
      Begin VB.Menu mnuHelpSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web ..."
      End
      Begin VB.Menu mnuHelpSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&Acerca de ..."
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Popup"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpAddFolder 
         Caption         =   "Agregar carpeta ..."
      End
      Begin VB.Menu mnuPopupAddItem 
         Caption         =   "Agregar elemento ..."
      End
      Begin VB.Menu mnuPopupSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopUpEditCut 
         Caption         =   "Co&rtar"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPopUpEditCopy 
         Caption         =   "&Copiar"
      End
      Begin VB.Menu mnuPopUpEditPaste 
         Caption         =   "&Pegar"
      End
      Begin VB.Menu mnuPopupSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopUpDrop 
         Caption         =   "&Quitar"
      End
      Begin VB.Menu mnuPopupSeparator4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupProperties 
         Caption         =   "Propiedades ..."
      End
      Begin VB.Menu mnuPopupSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupOpenCammera 
         Caption         =   "Abrir cámara ..."
      End
   End
End
Attribute VB_Name = "frmMDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--> Ventana principal del programa de documentación de proyectos
Option Explicit

Private Const cnstStrTreeMask As String = "000000" 'Máscara de la clave de los nodos del árbol
Private Const cnstStrTreeRoot As String = "R" 'Carácter clave de la raíz del árbol
Private Const cnstStrFolder As String = "F" 'Carácter clave de la carpeta en el árbol
Private Const cnstStrProject As String = "I" 'Carácter clave de un elemento en el árbol
Private Const cnstStrQuotes As String = """" 'Constante con las comillas

Private Enum eIcons 'Iconos de los botones de la barra de herramientas
  IconCopy = 1
  IconCut
  IconDelete
  IconHelp
  IconHelpThis
  IconDiagram
  IconNew
  IconOpen
  IconPast
  IconPrint
  IconProperties
  IconSave
  IconUp
  IconOpenFolder
  IconCloseFolder
  IconCammera
  IconView
  IconWeb
End Enum

Private Enum eButtons 'Enumerado con los botones de la barra de herramientas del generador
  ButtonNew = 1
  ButtonOpen = 2
  ButtonSave = 3
  ButtonCut = 5
  ButtonCopy = 6
  ButtonPaste = 7
  ButtonProperties = 9
  ButtonOpenCammera = 11
  ButtonOpenWebCammera = 12
  ButtonHelp = 14
End Enum

Private Enum eButtonsExplorer 'Enumerado con los botones de la parra de herramientas del explorador
  ButtonBack = 1
  ButtonForward = 2
  ButtonStop = 3
  ButtonRefresh = 4
  ButtonHome = 5
  ButtonSearch = 6
End Enum

Private Enum eIconsType 'Tipo de las imágenes de los iconos
  BigIcons = 0
  SmallIcons = 1
End Enum

Private Enum eProjectType 'Enumerado con losíndices de las webcam y favoritos
  prjFile = 0
  prjFavourites
End Enum

Private intIndexWebCam As Integer 'Indice con el número de la última WebCam abierta
Private imgIconType As eIconsType 'Variable con el tipo de icono (grande o pequeño) a presentar
Private colWebCam(prjFile To prjFavourites) As New colWebCams 'Colección de proyectos
Private strActualProject As String 'Nombre del proyecto actual
Private strLastPath As String 'Nombre del último directorio al que se ha accedido
Private blnSaved As Boolean 'Indica si se han realizado cambios en el proyecto
Private colMRUFiles As New colFilesMRU 'Recoge los ficheros de proyectos cargados
Public colLanguage As New colItemsLanguage 'Recoge los títulos dependientes del idioma
Private frmSplashScreen As frmSplash
Private intActualProject As eProjectType 'Indice con el fichero de webCam actual

Private Sub init()
'--> Inicializa los valores de la aplicación
Dim objDir As New clsDir

  On Error Resume Next
  'Inicializa los datos del registro
    frmSplashScreen.showMessage "Cargando los datos del registro ..." 'Esta es independiente del idioma, aún no se ha cargado nada
    loadRegistry
  'Inicializa los datos de últimos ficheros cargados
    colMRUFiles.load
    initMenuMRUFiles
    initMenu
  'Muestra/oculta las partes de la pantalla de acuerdo con los datos del registro
    frmSplashScreen.showMessage colLanguage.Item("K61").Caption
    mnuShowBigIcons.Checked = Not mnuShowBigIcons.Checked
    mnuShowProject.Checked = Not mnuShowProject.Checked
    mnuShowToolBar.Checked = Not mnuShowToolBar.Checked
    mnuShowVerticalMenu.Checked = Not mnuShowVerticalMenu.Checked
    showBigIcons
    showPanelProyecto
    showToolBar
    showVerticalMenu
  'Inicializa los controles
    With splVertical
      .Top = 0
      .Left = 0
      .Resize Me.Width, picPannel.Height
      .ZOrder 0
    End With
    brStatus.ZOrder 0
  'Inicializa el webBanner
    frmSplashScreen.showMessage colLanguage.Item("K62").Caption
    initWebBanner
  'Inicializa el árbol
    initProjectTree prjFile
  'Inicializa las barras de herramientas
    initToolBars
  'Inicializa las variables
    strActualProject = ""
    intActualProject = prjFile
    blnSaved = True
  'Crea el directorio donde se guardan las webCams y los favoritos
    strLastPath = App.Path & "\My webcams"
    objDir.makeDir strLastPath
    Set objDir = Nothing
  'Asocia los ficheros al programa
    shellPrograms
  'Carga el fichero de favoritos
    loadProject prjFavourites, App.Path & "\My webCams\Favourites.wxml"
End Sub

Private Sub loadRegistry()
'--> Carga los datos del registro
Dim objRegistry As New clsRegistry

  'Carga los datos
    If objRegistry.ExistKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry) Then
      'Configuración de pantalla
        Me.WindowState = objRegistry.QueryKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "Window_State")
        If Me.WindowState <> vbMaximized And Me.WindowState <> vbMinimized Then
          Me.Left = objRegistry.QueryKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "Window_Left")
          Me.Top = objRegistry.QueryKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "Window_Top")
          Me.Width = objRegistry.QueryKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "Window_Width")
          Me.Height = objRegistry.QueryKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "Window_Height")
        End If
        splVertical.SpliterLeft = objRegistry.QueryKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "Spl_Left")
      'Menú ver
        mnuShowToolBar.Checked = objRegistry.QueryKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "ShowToolbar")
        mnuShowProject.Checked = objRegistry.QueryKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "ShowProject")
        mnuShowBigIcons.Checked = objRegistry.QueryKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "BigIcons")
        mnuShowVerticalMenu.Checked = objRegistry.QueryKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "ShowVerticalMenu")
      'Configuración de lenguage
        changeLanguage objRegistry.QueryKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "Language")
      'Configuración de Internet
        dwfConnection.intConnection = objRegistry.QueryKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "ConnInternet")
        dwfConnection.strServer = objRegistry.QueryKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "ConnProxyServer")
        dwfConnection.intPort = objRegistry.QueryKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "ConnProxyPort")
        dwfConnection.strUser = objRegistry.QueryKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "ConnProxyUser")
        dwfConnection.strPassword = objRegistry.QueryKey(HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "ConnProxyPassword")
    Else
      'Menú ver
        mnuShowToolBar.Checked = True
        mnuShowProject.Checked = True
        mnuShowBigIcons.Checked = False
        mnuShowVerticalMenu.Checked = True
      'Configuración de lenguage
        changeLanguage 1
      'Configuración de Internet
        With dwfConnection
          .intConnection = DownloadFile.IFile_ConnectType.Preconfig
          .strServer = ""
          .intPort = 8080
          .strUser = ""
          .strServer = ""
        End With
    End If
  'Libera la memoria
    Set objRegistry = Nothing
End Sub

Private Sub saveRegistry()
'--> Graba los datos del registro
Dim objRegistry As New clsRegistry

  'Graba los datos
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "Window_State", Me.WindowState
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "Window_Left", Me.Left
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "Window_Top", Me.Top
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "Window_Width", Me.Width
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "Window_Height", Me.Height
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "Spl_Left", splVertical.SpliterLeft
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "ShowToolbar", mnuShowToolBar.Checked
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "ShowProject", mnuShowProject.Checked
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "BigIcons", mnuShowBigIcons.Checked
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "ShowVerticalMenu", mnuShowVerticalMenu.Checked
    If mnuChangeIdioma(1).Checked Then
      objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "Language", 1
    ElseIf mnuChangeIdioma(2).Checked Then
      objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "Language", 2
    End If
  'Configuración de Internet
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "ConnInternet", dwfConnection.intConnection
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "ConnProxyServer", dwfConnection.strServer
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "ConnProxyPort", dwfConnection.intPort
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "ConnProxyUser", dwfConnection.strUser
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, cnstStrRootRegistry, "ConnProxyPassword", dwfConnection.strPassword
  'Libera la memoria
    Set objRegistry = Nothing
End Sub

Private Sub shellPrograms()
'--> Actualiza el registro para asociar el programa
Dim objRegistry As New clsRegistry

  'Crea la clase del programa
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\bauWebCams\shell\open", "", "Abrir con bauWebCam"
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\bauWebCams\shell\open\command", "", _
                               App.Path & "\" & App.EXEName & " " & cnstStrQuotes & "%1" & cnstStrQuotes
  'Indica el icono del programa
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\bauWebCams\DefaultIcon", "", _
                               App.Path & "\" & App.EXEName & ", 0"
  'Asocia la extensión del archivo
    objRegistry.CreateKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Classes\.wxml", "", _
                               "bauWebCams"
  'Libera la memoria
    Set objRegistry = Nothing
End Sub

Private Sub initMenuMRUFiles()
'--> Inicializa el menú con los últimos ficheros cargados / grabados
Dim intIndex As Integer
Dim blnFound As Boolean

  For intIndex = 0 To 4
    mnuMRUFile(intIndex).Visible = False
    If intIndex < colMRUFiles.Count Then
      mnuMRUFile(intIndex).Visible = True
      mnuMRUFile(intIndex).Caption = "&" & (intIndex + 1) & " " & colMRUFiles.itemName(intIndex)
      blnFound = True
    End If
  Next intIndex
  mnuSeparatorFile4.Visible = blnFound
End Sub

Private Sub setIconMenu(ByVal strIconKey As String, ByVal strMenuKey As String)
'--> Asigna un icono a una opción de menu
  ctlPopMenu.ItemIcon(strMenuKey) = getIconIndex(strIconKey)
End Sub

Private Function getIconIndex(ByVal strKey As String) As Long
'--> Obtiene el índice de un icono
    getIconIndex = ilsIcons.ListImages.Item(strKey).Index - 1
End Function

Private Sub initMenu()
'--> Inicializa el menú
    With ctlPopMenu
      'Asocia la lista de imágenes
        .ImageList = ilsIcons
      'Recorre el menú diseñado en VB y lo subclasifica
        .SubClassMenu Me
      'Asigna las propiedades al menú
        .HighlightStyle = cspHighlightXP
        .ShadowXPHighlightTopMenu = True
        .ShadowXPHighlight = True
        Set .BackgroundPicture = picBackground.Picture
      'Inicializa los iconos
        setIconMenu "OPEN", "mnuOpenProject"
        setIconMenu "SAVE", "mnuSaveProject"
        'setIconMenu "PRINT", "mnuPrint"
        
        setIconMenu "CUT", "mnuEditCut"
        setIconMenu "COPY", "mnuEditCopy"
        setIconMenu "PASTE", "mnuEditPaste"
        
        setIconMenu "HELP", "mnuHelpIndex"
        setIconMenu "NET", "mnuHelpWeb"

        setIconMenu "OPEN", "mnuPopUpAddFolder"
        .TickIconIndex = getIconIndex("TICK")
    End With
End Sub

Private Sub highlightOptionMenu(ByVal lngItemNumber As Long, ByVal blnEnabled As Boolean, ByVal blnSeparator As Boolean)
'--> Muestra la ayuda de una opción de menú
Dim strText As String

  strText = ""
  If Not blnSeparator And blnEnabled Then
    Select Case lngItemNumber
      Case 2
        strText = "Crear nuevo archivo"
      Case 3
        strText = "Abrir archivo"
      Case 4
        strText = "Guardar archivo"
      Case 5
        strText = "Guardar archivo como"
      Case 7
        strText = "Propiedades de la webCamm"
      Case 9
        strText = "Ver imagen de la webCam"
      Case 10
        strText = "Visitar web de la webCam"
      Case 12, 13, 14, 15
        strText = "Abrir fichero " & Mid(ctlPopMenu.Caption(lngItemNumber), 3)
      Case 17
        strText = "Salir del programa"
      Case 19
        strText = "Agregar carpeta"
      Case 20
        strText = "Agregar webCamm"
      Case 22
        strText = "Cortar webCamm/carpeta"
      Case 23
        strText = "Copiar webCamm/carpeta en el portapapeles"
      Case 24
        strText = "Pegar webCamm/carpeta del portapapeles"
      Case 25
        strText = "Eliminar webCamm/carpeta"
      Case 26
        strText = "Crear un vídeo a partir de las imágenes grabadas"
      Case 31
        strText = "Mostrar/ocultar la ventana de proyecto"
      Case 32
        strText = "Mostrar/ocultar la barra de herramientas"
      Case 33
        strText = "Mostrar/ocultar el menú vertical"
      Case 35
        strText = "Cambiar el idioma de la aplicación a español"
      Case 36
        strText = "Cambiar el idioma de la aplicación a inglés"
      Case 37
        strText = "Ver los ficheros de la web"
      Case 39
        strText = "Opciones de la aplicación"
      Case Else
        strText = "Item: " & lngItemNumber
    End Select
  End If
  brStatus.Message = strText
End Sub

Private Sub changeLanguage(ByVal intIndex As Integer)
'--> Cambia el lenguaje a partir de una opción del menú
Dim intIndexCheck As Integer

  For intIndexCheck = 1 To mnuChangeIdioma.Count
    mnuChangeIdioma.Item(intIndexCheck).Checked = False
  Next intIndexCheck
  If intIndex < 1 Or intIndex > 2 Then
    intIndex = 0
  End If
  mnuChangeIdioma.Item(intIndex).Checked = True
  Select Case intIndex
    Case 1 'Español
      initLanguage "Spanish"
    Case 2 'Inglés
      initLanguage "English"
  End Select
End Sub

Private Sub initLanguage(ByVal strLanguage As String)
'--> Inicializa las características de texto dependientes del idioma
  'Carga la colección
    If colLanguage.loadXMLLanguages(App.Path & "\lib", strLanguage) Then
      'Inicializa los menús
        'Menú de archivos
          mnuFile.Caption = colLanguage.Item("K1").Caption
          mnuNewProject.Caption = colLanguage.Item("K2").Caption
          mnuOpenProject.Caption = colLanguage.Item("K3").Caption
          mnuSaveProject.Caption = colLanguage.Item("K4").Caption
          mnuSaveProjectAs.Caption = colLanguage.Item("K5").Caption
          mnuAddFolder.Caption = colLanguage.Item("K6").Caption
          mnuAddItem.Caption = colLanguage.Item("K7").Caption
          mnuDrop.Caption = colLanguage.Item("K300").Caption
          mnuProperties.Caption = colLanguage.Item("K8").Caption
          mnuOpenCammera.Caption = colLanguage.Item("K9").Caption
          mnuViewWebWebcam.Caption = colLanguage.Item("K301").Caption
          mnuPrint.Caption = colLanguage.Item("K10").Caption
          mnuExit.Caption = colLanguage.Item("K11").Caption
        'Menú edición
          mnuEdit.Caption = colLanguage.Item("K800").Caption
          mnuEditCut.Caption = colLanguage.Item("K801").Caption
          mnuEditCopy.Caption = colLanguage.Item("K802").Caption
          mnuEditPaste.Caption = colLanguage.Item("K803").Caption
        'Menú de herramientas
          mnuTools.Caption = colLanguage.Item("K12").Caption
          mnuVer.Caption = colLanguage.Item("K13").Caption
          mnuShowProject.Caption = colLanguage.Item("K14").Caption
          mnuShowToolBar.Caption = colLanguage.Item("K15").Caption
          mnuShowVerticalMenu.Caption = colLanguage.Item("K16").Caption
          mnuIdioma.Caption = colLanguage.Item("K17").Caption
          mnuChangeIdioma(1).Caption = colLanguage.Item("K18").Caption
          mnuChangeIdioma(2).Caption = colLanguage.Item("K19").Caption
          mnuSearchWebCam.Caption = colLanguage.Item("K20").Caption
        'Menú ventana
          mnuWindow.Caption = colLanguage.Item("K600").Caption
          mnuWindowArrange(1).Caption = colLanguage.Item("K601").Caption
          mnuWindowArrange(2).Caption = colLanguage.Item("K602").Caption
          mnuWindowArrange(3).Caption = colLanguage.Item("K603").Caption
          mnuWindowArrange(4).Caption = colLanguage.Item("K604").Caption
        'Menú ayuda
          mnuHelp.Caption = colLanguage.Item("K21").Caption
          mnuHelpIndex.Caption = colLanguage.Item("K22").Caption
          mnuHelpTip.Caption = colLanguage.Item("K23").Caption
          mnuHelpWeb.Caption = colLanguage.Item("K24").Caption
          mnuHelpAbout.Caption = colLanguage.Item("K25").Caption
        'Menú popup
          mnuPopUpAddFolder.Caption = colLanguage.Item("K6").Caption
          mnuPopupAddItem.Caption = colLanguage.Item("K7").Caption
          mnuPopUpDrop.Caption = colLanguage.Item("K300").Caption
          mnuPopUpEditCut.Caption = colLanguage.Item("K801").Caption
          mnuPopUpEditCopy.Caption = colLanguage.Item("K802").Caption
          mnuPopUpEditPaste.Caption = colLanguage.Item("K803").Caption
          mnuPopupProperties.Caption = colLanguage.Item("K8").Caption
          mnuPopupOpenCammera.Caption = colLanguage.Item("K9").Caption
        'Barra de herramientas
          With tlbHerramientas
            .Buttons(ButtonNew).ToolTipText = colLanguage.Item("K26").Caption
            .Buttons(ButtonNew).ButtonMenus(1) = colLanguage.Item("K27").Caption
            .Buttons(ButtonNew).ButtonMenus(2) = colLanguage.Item("K28").Caption
            .Buttons(ButtonNew).ButtonMenus(3) = colLanguage.Item("K29").Caption
            .Buttons(ButtonOpen).ToolTipText = colLanguage.Item("K30").Caption
            .Buttons(ButtonSave).ToolTipText = colLanguage.Item("K31").Caption
            .Buttons(ButtonCut).ToolTipText = colLanguage.Item("K900").Caption
            .Buttons(ButtonCopy).ToolTipText = colLanguage.Item("K901").Caption
            .Buttons(ButtonPaste).ToolTipText = colLanguage.Item("K902").Caption
            .Buttons(ButtonProperties).ToolTipText = colLanguage.Item("K32").Caption
            .Buttons(ButtonOpenWebCammera).ToolTipText = colLanguage.Item("K301").Caption
            .Buttons(ButtonOpenCammera).ToolTipText = colLanguage.Item("K33").Caption
            .Buttons(ButtonHelp).ToolTipText = colLanguage.Item("K34").Caption
          End With
        'Ventana de proyecto
          picPannel.Caption = colLanguage.Item("K41").Caption
        'Fichas de proyecto
          tabTree.Tabs.Item(1).Caption = colLanguage.Item("K1000").Caption
          tabTree.Tabs.Item(2).Caption = colLanguage.Item("K1001").Caption
        'Menú vertical
          With vrtMnuProject
            'Menú archivo
              .MenuCur = 1
              .MenuCaption = colLanguage.Item("K42").Caption
              .MenuItemCur = 1
              .MenuItemCaption = colLanguage.Item("K26").Caption
              .MenuItemCur = 2
              .MenuItemCaption = colLanguage.Item("K30").Caption
              .MenuItemCur = 3
              .MenuItemCaption = colLanguage.Item("K31").Caption
              .MenuItemCur = 4
              .MenuItemCaption = colLanguage.Item("K20").Caption
            'Menú proyecto
              .MenuCur = 2
              .MenuCaption = colLanguage.Item("K27").Caption
              .MenuItemCur = 1
              .MenuItemCaption = colLanguage.Item("K28").Caption
              .MenuItemCur = 2
              .MenuItemCaption = colLanguage.Item("K29").Caption
              .MenuItemCur = 3
              .MenuItemCaption = colLanguage.Item("K32").Caption
              .MenuItemCur = 4
              .MenuItemCaption = colLanguage.Item("K43").Caption
              .MenuItemCur = 5
              .MenuItemCaption = colLanguage.Item("K33").Caption
            'Menú ayuda
              .MenuCur = 3
              .MenuCaption = colLanguage.Item("K34").Caption
              .MenuItemCur = 1
              .MenuItemCaption = colLanguage.Item("K34").Caption
              .MenuItemCur = 2
              .MenuItemCaption = colLanguage.Item("K44").Caption
              .MenuItemCur = 3
              .MenuItemCaption = colLanguage.Item("K45").Caption
            'Se posiciona de nuevo en el primer menú
              .MenuCur = 1
          End With
        'Barra de estado
          brStatus.Message = colLanguage.Item("K46").Caption
    End If
End Sub

Private Sub initToolBars()
'--> Inicializa las barras de herramientas
  'Barra de herramientas del programa
    With tlbHerramientas
      Set .ImageList = imlImages(imgIconType)
      .Buttons(ButtonNew).Image = IconNew
      .Buttons(ButtonOpen).Image = IconOpen
      .Buttons(ButtonSave).Image = IconSave
      .Buttons(ButtonCut).Image = IconCut
      .Buttons(ButtonCopy).Image = IconCopy
      .Buttons(ButtonPaste).Image = IconPast
      .Buttons(ButtonProperties).Image = IconProperties
      .Buttons(ButtonOpenWebCammera).Image = IconWeb
      .Buttons(ButtonOpenCammera).Image = IconView
      .Buttons(ButtonHelp).Image = IconHelp
      If imgIconType = BigIcons Then
        .Height = 800
      Else
        .Height = 400
      End If
    End With
End Sub

Private Sub initWebBanner()
'--> Inicializa el webBanner
Dim objDir As New clsDir

  'Crea el directorio donde se guardan las imágenes
    objDir.makeDir App.Path & "\Banner"
  'Libera el objeto
    Set objDir = Nothing
  'Inicializa el webBanner
    With wbnImage
      .Host = "ftp.galeon.com"
      .Port = 21
      .User = "bauconsultors"
      .Password = "bau5044"
      .LocalPath = App.Path & "\Banner"
      .ServerFile = "xmlBanner.xml"
      '.readBanners
    End With
End Sub

Private Sub setOptions()
'--> Modifica la configuración de servidor proxy y otras opciones
Dim frmConfig As New frmOptions

  'Abre la ventana y recoge sus datos
    With frmConfig
      'Configura los datos
        .intConnection = dwfConnection.intConnection
        .strServer = dwfConnection.strServer
        .intPort = dwfConnection.intPort
        .strUser = dwfConnection.strUser
        .strPassword = dwfConnection.strPassword
      'Muestra la ventana
        .Show vbModal
      'Recoge los datos
        If .blnCancel Then
          dwfConnection.intConnection = .intConnection
          dwfConnection.strServer = .strServer
          dwfConnection.intPort = .intPort
          dwfConnection.strUser = .strUser
          dwfConnection.strPassword = .strPassword
        End If
    End With
  'Cierra la ventana
    Set frmConfig = Nothing
End Sub

Private Sub searchWeb()
'--> Se dirige a la web para cargar ficheros
'--> Abre una ventana del explorador
Dim objExplorer As New clsHyperlink

  With objExplorer
    'Establece las propiedades
      .URL = "http://www.galeon.com/bauconsultors/webCams/files.htm"
      .ExplorerStatus = SW_SHOWNORMAL
    'Abre el editor de correo
      .OpenURL
  End With
  'Libera la memoria
    Set objExplorer = Nothing
End Sub

Private Sub showPanelProyecto()
'--> Muestra u oculta el panel de proyecto
  mnuShowProject.Checked = Not mnuShowProject.Checked
  picPannel.Visible = mnuShowProject.Checked
  Resize
End Sub

Private Sub showToolBar()
'--> Muestra u oculta la barra de herramientas
  mnuShowToolBar.Checked = Not mnuShowToolBar.Checked
  tlbHerramientas.Visible = mnuShowToolBar.Checked
  Resize
End Sub

Private Sub showVerticalMenu()
'--> Muestra u oculta el menú vertical
  mnuShowVerticalMenu.Checked = Not mnuShowVerticalMenu.Checked
  vrtMnuProject.Visible = mnuShowVerticalMenu.Checked
  Resize
End Sub

Private Sub showBigIcons()
'--> Cambia el tamaño de los iconos
'  mnuShowBigIcons.Checked = Not mnuShowBigIcons.Checked
'  If mnuShowBigIcons.Checked Then
'    imgIconType = BigIcons
'  Else
'    imgIconType = SmallIcons
'  End If
'  initToolBars
'  initProjectTree
'  Resize
End Sub

Private Sub Resize()
'--> Recoloca los controles
Dim intHeightToolBar As Integer, intLastVertMenu As Integer, intIndex As Integer

  On Error Resume Next
  If tlbHerramientas.Visible Then
    intHeightToolBar = tlbHerramientas.Height
  Else
    intHeightToolBar = 0
  End If
  'intLastVertMenu = vrtMnuProject.MenuCur
  'vrtMnuProject.MenuCur = 3
  'vrtMnuProject.Refresh
  'vrtMnuProject.MenuCur = intLastVertMenu
  splVertical.Resize Me.Width, Me.Height
  picPannel.Width = splVertical.SpliterLeft + splVertical.SpliterPictureWidth
  If picPannel.Visible Then
    With tabTree
      .Width = splVertical.SpliterLeft - 100
      .Height = picPannel.Height - .Top - 40
    End With
    For intIndex = 0 To 1
      With trvProject(intIndex)
        .Top = tabTree.Top + tabTree.Tabs(1).Height + 150
        .Width = tabTree.Width - 250
        .Height = tabTree.Height - .Top + tabTree.Tabs(1).Height - 40
      End With
    Next intIndex
  End If
End Sub

Private Sub newProject()
'--> Crea un nuevo proyecto
Dim blnContinuar
Dim strFileName As String, strPath As String
Dim objFile As New clsFiles

  blnContinuar = True
  'Si ya hay un proyecto cargado pregunta si desea cargar otro
    If Not blnSaved Then
      Select Case messageBox.showMessage(colLanguage("K83").Caption & vbCrLf & _
                         colLanguage("K84").Caption, "bauWebCams", MsgQuestionCancel)
        Case ResultYes
          blnContinuar = saveProject()
        Case ResultNo
          blnContinuar = True 'No es necesario, se añade por claridad
        Case ResultCancel
          blnContinuar = False
      End Select
    End If
  'Crea el nuevo proyecto si es necesario
    If blnContinuar Then
      'Obtiene el nombre de fichero
        strPath = Trim(objFile.getPath(strActualProject))
        If strPath = "" Then
          strPath = "C:\"
        End If
        strFileName = Trim(objFile.dlgGetFileName(dlgCommon, False, strPath, _
                                                  colLanguage("K85").Caption))
        If strFileName <> "" Then
          'Elimina el proyecto anterior
            colWebCam(prjFile).Clear
          'Crea el proyecto
            strActualProject = strFileName
          'Inicializa el árbol
            initProjectTree prjFile
          'Inicializa las variables
            blnSaved = True
        End If
    End If
  'Libera los objetos de memoria
    Set objFile = Nothing
End Sub

Private Function openProject() As Boolean
'--> Abre un proyecto
'--> Crea un nuevo proyecto
Dim blnContinuar
Dim strFileName As String, strPath As String
Dim objFile As New clsFiles

  blnContinuar = True
  'Si ya hay un proyecto cargado pregunta si desea cargar otro
    If Not blnSaved Then
      Select Case messageBox.showMessage(colLanguage("K83").Caption & vbCrLf & _
                                         colLanguage("K84").Caption, _
                                         "bauWebCams", MsgQuestionCancel)
        Case ResultYes
          blnContinuar = saveProject()
        Case ResultNo
          blnContinuar = True 'No es necesario, se añade por claridad
        Case ResultCancel
          blnContinuar = False
      End Select
    End If
  'Abre el proyecto si es necesario
    If blnContinuar Then
      'Obtiene el nombre de fichero
        strPath = Trim(objFile.getPath(strActualProject))
        If strPath = "" Then
          strPath = "C:\"
        End If
        strFileName = Trim(objFile.dlgGetFileName(dlgCommon, True, strPath, _
                                                  colLanguage("K85").Caption))
        If strFileName <> "" Then
          'Elimina el proyecto anterior
            colWebCam(prjFile).Clear
          'Carga el árbol de proyectos
            loadProject prjFile, strFileName
          'Carga el fichero sobre la colección de ficheros cargados
            colMRUFiles.Add strFileName
            initMenuMRUFiles
          'Inicializa las variables
            blnSaved = True
        End If
    End If
  'Libera los objetos de memoria
    Set objFile = Nothing
End Function

Private Sub loadXMLProject(ByVal intProject As eProjectType, ByVal objXMLList As MSXML.IXMLDOMNodeList)
'--> Carga los nodos XML del proyecto
Dim objXMLNode As MSXML.IXMLDOMNode
Dim trnNode As Node, trnLast As Node

  'Recorre la lista de nodos
    For Each objXMLNode In objXMLList
      Select Case UCase(Trim(objXMLNode.baseName))
        Case "FOLDER"
          With trvProject(intProject)
            If .SelectedItem Is Nothing Then
              Set .SelectedItem = .Nodes.Item(1)
            End If
            Set trnLast = .SelectedItem
            Set trnNode = .Nodes.Add(.SelectedItem, tvwChild, _
                                     getNextTreeKey(intProject, cnstStrFolder), _
                                     objXMLNode.Attributes.Item(0).nodeValue, IconOpenFolder)
            Set .SelectedItem = trnNode
          End With
          loadXMLProject intProject, objXMLNode.childNodes
          Set trvProject(intProject).SelectedItem = trnLast
        Case "WEBCAM"
          addWebCamFromXML intProject, objXMLNode.xml
      End Select
    Next objXMLNode
End Sub

Private Sub addWebCamFromXML(ByVal intProject As eProjectType, ByVal strXML As String)
Dim objXMLDocument As New MSXML.DOMDocument
Dim objXMLWebCam As MSXML.IXMLDOMNode, objXMLNode As MSXML.IXMLDOMNode
Dim strName As String, strDescription As String, strURL As String, strURLWeb As String
Dim strEMail As String, strICQ As String
Dim intInterval As Integer

  If Not objXMLDocument.loadXML(strXML) Then
    messageBox.showMessage "No se puede cargar la cámara" & vbCrLf & strXML, "bauWebCam", MsgExclamation
  Else
    'Recoge los parámetros
      intInterval = 30
      For Each objXMLWebCam In objXMLDocument.childNodes
        If UCase(objXMLWebCam.baseName) = "WEBCAM" Then
          'Recoge los parámetros de la webCam
            For Each objXMLNode In objXMLWebCam.childNodes
              If UCase(objXMLNode.baseName) = "NAME" Then
                strName = objXMLNode.Text
              End If
              If UCase(objXMLNode.baseName) = "DESCRIPTION" Then
                strDescription = objXMLNode.Text
              End If
              If UCase(objXMLNode.baseName) = "URL" Then
                strURL = objXMLNode.Text
              End If
              If UCase(objXMLNode.baseName) = "URLWEB" Then
                strURLWeb = objXMLNode.Text
              End If
              If UCase(objXMLNode.baseName) = "EMAIL" Then
                strEMail = objXMLNode.Text
              End If
              If UCase(objXMLNode.baseName) = "ICQ" Then
                strICQ = objXMLNode.Text
              End If
              If UCase(objXMLNode.baseName) = "INTERVAL" And IsNumeric(objXMLNode.Text) Then
                intInterval = CInt(objXMLNode.Text)
              End If
            Next objXMLNode
          'Añade la webCam
            If strName <> "" And strURL <> "" Then
              addWebCamCollectionTree intProject, strName, strDescription, strURL, strURLWeb, _
                                      strEMail, strICQ, intInterval
            End If
        End If
      Next objXMLWebCam
  End If
  'Libera los objetos
    Set objXMLDocument = Nothing
End Sub

Private Function loadProject(ByVal intProject As eProjectType, ByVal strFileName As String) As Boolean
'--> Carga el proyecto en el árbol y la colección
Dim objXML As New MSXML.DOMDocument
Dim objXMLNode As MSXML.IXMLDOMNode

  On Error GoTo errorLoad
  'Crea el proyecto
    If intProject = prjFile Then
      strActualProject = strFileName
    End If
  'Inicializa el árbol
    initProjectTree intProject
  'Inicializa la colección de webCams
    Set colWebCam(intProject) = New colWebCams
  'Abre el fichero XML
    objXML.load strFileName
    If objXML.parseError.errorCode <> 0 Then
      messageBox.showMessage colLanguage("K81").Caption & " (" & objXML.parseError.Line & ")" & vbCrLf & _
                             objXML.parseError.reason & vbCrLf & _
                             objXML.parseError.srcText, "bauWebCams", MsgExclamation
    Else
      'Recorre los elementos del fichero y los almacena en la colección de DLLs
        For Each objXMLNode In objXML.childNodes
          If objXMLNode.nodeType = NODE_ELEMENT Then
            If UCase(objXMLNode.baseName) = "WEBCAMS" Then
              loadXMLProject intProject, objXMLNode.childNodes
            End If
          End If
        Next objXMLNode
    End If
  'Libera los objetos
    Set objXMLNode = Nothing
    Set objXML = Nothing
  'Sale de la función
    Exit Function
    
errorLoad:
  messageBox.showMessage colLanguage("K81").Caption & vbCrLf & Err.Description, _
                         "bauWebCams", MsgExclamation
End Function

Private Function getParameterWebCam(ByVal intProject As eProjectType, ByVal strKey As String, ByVal intSpaces As Integer)
'--> Obtiene los parámetros del proyecto
Dim objWebCam As clsWebCam

  getParameterWebCam = ""
  Set objWebCam = colWebCam(intProject).Item(strKey)
  If Not objWebCam Is Nothing Then
    getParameterWebCam = Space(intSpaces) & "<Name>" & getCData(objWebCam.Name) & "</Name>" & vbCrLf
    getParameterWebCam = getParameterWebCam & Space(intSpaces) & "<Description>" & _
                                  getCData(objWebCam.Description) & "</Description>" & vbCrLf
    getParameterWebCam = getParameterWebCam & Space(intSpaces) & "<URL>" & getCData(objWebCam.URL) & "</URL>" & vbCrLf
    getParameterWebCam = getParameterWebCam & Space(intSpaces) & "<URLWeb>" & getCData(objWebCam.WebURL) & "</URLWeb>" & vbCrLf
    getParameterWebCam = getParameterWebCam & Space(intSpaces) & "<Interval>" & getCData(objWebCam.Interval) & "</Interval>" & vbCrLf
    getParameterWebCam = getParameterWebCam & Space(intSpaces) & "<EMail>" & getCData(objWebCam.eMail) & "</EMail>" & vbCrLf
    getParameterWebCam = getParameterWebCam & Space(intSpaces) & "<ICQ>" & getCData(objWebCam.ICQ) & "</ICQ>" & vbCrLf
  End If
  Set objWebCam = Nothing
End Function

Private Function getXMLNodes(ByVal intProject As eProjectType, ByVal trnParent As Node, ByVal intSpaces As Integer) As String
'--> Obtiene el XML de los hijos
Dim trnNode As Node
  
  getXMLNodes = ""
  For Each trnNode In trvProject(intProject).Nodes
    If Not trnNode.Parent Is Nothing Then
      If trnNode.Parent.Key = trnParent.Key Then
        Select Case Left(trnNode.Key, 1)
          Case cnstStrFolder 'Si estamos en una carpeta
            getXMLNodes = getXMLNodes & Space(intSpaces) & "<Folder Name=" & cnstStrQuotes & _
                                                           trnNode.Text & cnstStrQuotes & ">" & _
                                                           vbCrLf
            If Not trnNode.Child Is Nothing Then
              getXMLNodes = getXMLNodes & getXMLNodes(intProject, trnNode, intSpaces + 2)
            End If
            getXMLNodes = getXMLNodes & Space(intSpaces) & "</Folder>" & vbCrLf
          Case cnstStrProject 'Si estamos en un elemento
            getXMLNodes = getXMLNodes & Space(intSpaces) & "<WebCam>" & vbCrLf
            getXMLNodes = getXMLNodes & getParameterWebCam(intProject, trnNode.Key, intSpaces + 2)
            getXMLNodes = getXMLNodes & Space(intSpaces) & "</WebCam>" & vbCrLf
        End Select
      End If
    End If
  Next trnNode
End Function

Private Function getXMLProject(ByVal intProject As eProjectType) As String
'--> Obtiene el XML del proyecto
  getXMLProject = "<WebCams>" & vbCrLf
  getXMLProject = getXMLProject & getXMLNodes(intProject, trvProject(intProject).Nodes.Item(1), 2)
  getXMLProject = getXMLProject & "</WebCams>" & vbCrLf
End Function

Private Sub copyWebCam()
'--> Copia los datos de una webCam en el portapapeles
  If trvProject(intActualProject).SelectedItem Is Nothing Then
    messageBox.showMessage "Seleccione una webCam", "bauWebCams", MsgInformation
  ElseIf Left(trvProject(intActualProject).SelectedItem.Key, 1) <> cnstStrProject Then
    messageBox.showMessage "Seleccione una webCam", "bauWebCams", MsgInformation
  Else
    Clipboard.SetText "<WebCam>" & vbCrLf & getParameterWebCam(intActualProject, trvProject(intActualProject).SelectedItem.Key, 2) & "</WebCam>" & vbCrLf
  End If
End Sub

Private Sub pasteWebCam()
'--> Pega una webCam
  If trvProject(intActualProject).SelectedItem Is Nothing Then
    messageBox.showMessage "Seleccione una carpeta", "bauWebCams", MsgInformation
  ElseIf Left(trvProject(intActualProject).SelectedItem.Key, 1) <> cnstStrFolder And Left(trvProject(intActualProject).SelectedItem.Key, 1) <> cnstStrTreeRoot Then
    messageBox.showMessage "Seleccione una carpeta", "bauWebCams", MsgInformation
  Else
    addWebCamFromXML intActualProject, Clipboard.GetText
  End If
End Sub

Private Function saveProject() As Boolean
'--> Graba el proyecto actualmente en memoria
Dim blnContinuar As Boolean
Dim objFile As New clsFiles

  'Obtiene el nombre del proyecto
    If strActualProject = "" Then
      strActualProject = objFile.dlgGetFileName(dlgCommon, False, strLastPath, colLanguage("K85").Caption)
    End If
  'Guarda el proyecto
    If strActualProject <> "" Then
      'Comprueba si ya existe el fichero
        blnContinuar = True
        If objFile.existFile(strActualProject) Then
          If messageBox.showMessage(Replace(colLanguage("K86").Caption, "%s", strActualProject) & vbCrLf & _
                                    colLanguage("K87").Caption, "bauWebCams", MsgQuestion) = ResultNo Then
            blnContinuar = False
          End If
        End If
      'Si todo es correcto graba el proyecto
        If blnContinuar Then
          'Realiza la grabación
            saveWebCamFiles prjFile, strActualProject
          'Recoge las variables
            strLastPath = objFile.getPath(strActualProject)
            blnSaved = True
            trvProject(prjFile).Nodes.Item(1).Text = objFile.getFileName(strActualProject)
          'Añade el nombre del fichero a la colección de ficheros cargados
            colMRUFiles.Add strActualProject
            initMenuMRUFiles
        End If
    End If
  'Libera los objetos
    Set objFile = Nothing
End Function

Private Sub saveWebCamFiles(ByVal intProject As eProjectType, ByVal strFileName As String)
'--> Graba físicamente una colección de webCams
Dim lngFile As Long

  lngFile = FreeFile()
  Open strFileName For Output As #lngFile
  Print #lngFile, getHeaderXML() & getXMLProject(intProject)
  Close #1
End Sub

Private Sub dropItem()
'--> Elimina un elemento / carpeta
Dim blnContinuar As Boolean

  If trvProject(intActualProject).SelectedItem Is Nothing Then
    messageBox.showMessage colLanguage.Item("K98").Caption, "bauWebCams", MsgInformation '"Seleccione un elemento o carpeta"
  Else
    blnContinuar = False
    If Left(trvProject(intActualProject).SelectedItem.Key, 1) = cnstStrTreeRoot Then
      messageBox.showMessage colLanguage.Item("K94").Caption, "bauWebCams", MsgInformation '"No se puede eliminar el elemento raíz"
    ElseIf Left(trvProject(intActualProject).SelectedItem.Key, 1) = cnstStrFolder Then
      If messageBox.showMessage(colLanguage.Item("K95").Caption & vbCrLf & _
                                colLanguage.Item("K96").Caption, "bauWebCams", MsgQuestion) = ResultYes Then '"Al eliminar una carpeta se eliminarán también sus elementos""¿Desea continuar?"
        blnContinuar = True
      End If
    ElseIf messageBox.showMessage(Replace(colLanguage.Item("K97").Caption, "%s", trvProject(intActualProject).SelectedItem.Text), _
                                  "bauWebCams", MsgQuestion) = ResultYes Then '"¿Realmente desea eliminar el elemento "
      blnContinuar = True
    End If
    If blnContinuar Then 'Borra los elementos
      If Left(trvProject(intActualProject).SelectedItem.Key, 1) = cnstStrProject Then
        dropWebcam trvProject(intActualProject).SelectedItem.Key
      Else
        dropFolder trvProject(intActualProject).SelectedItem.Key
      End If
    End If
  End If
End Sub

Private Sub dropWebcam(ByVal strKeyProject As String)
'--> Elimina una webCam de la colección de proyectos y el árbol
  On Error GoTo errorDrop
    trvProject(intActualProject).Nodes.Remove strKeyProject
    colWebCam(intActualProject).Remove strKeyProject
  Exit Sub
  
errorDrop:
  messageBox.showMessage Replace(colLanguage.Item("K99").Caption, "%s", trvProject(intActualProject).SelectedItem.Text), _
                         "bauWebCams", MsgExclamation 'Error al eliminar el elemento %s
End Sub

Private Sub dropFolder(ByVal strKeyItem As String)
'--> Elimina una carpeta y todos sus elementos
Dim objNode As Node
Dim blnDropped As Boolean

  'Elimina los hijos de esa carpeta
    Do
      blnDropped = False
      For Each objNode In trvProject(intActualProject).Nodes
        If Not objNode.Parent Is Nothing Then
          If objNode.Parent.Key = strKeyItem Then
            If Left(objNode.Key, 1) = cnstStrFolder Then
              dropFolder objNode.Key
              blnDropped = True
              Exit For
            ElseIf Left(objNode.Key, 1) = cnstStrProject Then
              dropWebcam objNode.Key
              blnDropped = True
              Exit For
            End If
          End If
        End If
      Next objNode
    Loop While blnDropped = True
  'Elimina la carpeta
    If Left(strKeyItem, 1) <> cnstStrTreeRoot Then
      trvProject(intActualProject).Nodes.Remove strKeyItem
    End If
End Sub

Private Sub openCammera()
'--> Abre la vista de una cámara
Dim frmCammera As New frmShowCammera
Dim objWebCam As clsWebCam

  If trvProject(intActualProject).SelectedItem Is Nothing Then
    messageBox.showMessage "Seleccione una cámara", "bauWebCams", MsgInformation
  ElseIf Left(trvProject(intActualProject).SelectedItem.Key, 1) <> cnstStrProject Then
    messageBox.showMessage "Seleccione una cámara", "bauWebCams", MsgInformation
  Else
    Set objWebCam = colWebCam(intActualProject).Item(trvProject(intActualProject).SelectedItem.Key)
    If Not objWebCam Is Nothing Then
      With frmCammera
        Set .objWebCam = objWebCam
        .Caption = objWebCam.Name
        .strFilePrefix = "webCam" & Format(intIndexWebCam, "00000")
        intIndexWebCam = intIndexWebCam + 1
        .Show
        .init
      End With
    End If
  End If
End Sub

Public Sub showScreenTotal(ByVal strCaption As String, ByVal strURL As String, _
                           ByVal strFilePrefix As String, ByVal blnSave As Boolean, _
                           ByVal strPath As String, ByVal intInterval As Integer, ByVal intIndexSaved)
'--> Muestra la vista de una cámara a pantalla completa
Dim frmCammera As New frmCammeraModeless

  With frmCammera
    .blnSave = blnSave
    .strURL = strURL
    .strFilePrefix = strFilePrefix
    .intInterval = intInterval
    .intIndexSaved = intIndexSaved
    .strPath = strPath
    
    .Caption = Me.Caption 'Esta debe ser el último, al hacer referencia a caption abre la ventana
    .Show vbModeless, Nothing
  End With
  frmMDIMain.Visible = False
End Sub

Private Sub helpIndex()
'--> Muestra la ayuda
  messageBox.showMessage "Ayuda no disponible", "bauWebCams", MsgInformation
End Sub

Private Sub dayTip()
'--> Muestra el truco del día
Dim frmTipDay As New frmTip

  frmTipDay.strLanguage = colLanguage.Language
  frmTipDay.Show vbModal
  Set frmTipDay = Nothing
End Sub

Private Sub exitApplication()
'--> Sale de la aplicación
  'Graba los datos
    If Not blnSaved Then
      If messageBox.showMessage(colLanguage("K83").Caption & vbCrLf & colLanguage("K84").Caption, _
                                "bauWebCams", MsgQuestion) = ResultYes Then
        saveProject
      End If
    End If
  'Graba los datos del registro
    saveRegistry
  'Graba la colección de favoritos
    saveWebCamFiles prjFavourites, App.Path & "\My webCams\Favourites.wxml"
  'Libera los objetos
    colMRUFiles.Save
    Set colMRUFiles = Nothing
    colWebCam(prjFavourites).Clear
    Set colWebCam(prjFavourites) = Nothing
    colWebCam(prjFile).Clear
    Set colWebCam(prjFile) = Nothing
    colLanguage.Clear
    Set colLanguage = Nothing
  'Termina
    Unload Me
End Sub

Private Sub initProjectTree(ByVal intProject As eProjectType)
'--> Inicializa el árbol de proyectos
Dim objFile As New clsFiles

  'Inicializa el elemento raíz del árbol
    With trvProject(intProject)
      Set .ImageList = imlImages(imgIconType)
      .Nodes.Clear
      .Nodes.Add , , cnstStrTreeRoot, colLanguage("K27").Caption, IconCloseFolder
      If intProject = prjFile And strActualProject <> "" Then
        .Nodes.Item(1).Text = objFile.getFileName(strActualProject)
      ElseIf intProject = prjFavourites Then
        .Nodes.Item(1).Text = colLanguage.Item("K1001").Caption
      End If
    End With
  'Libera la memoria intermedia
    Set objFile = Nothing
End Sub

Private Sub addFolder()
'--> Agrega una carpeta al árbol
Dim strFolder As String

  If trvProject(intActualProject).SelectedItem Is Nothing Then
    Set trvProject(intActualProject).SelectedItem = trvProject(intActualProject).Nodes.Item(1)
  End If
  If Left(trvProject(intActualProject).SelectedItem.Key, "1") = cnstStrProject Then '"No es posible agregar carpetas a un elemento"
    messageBox.showMessage colLanguage("K90").Caption, "bauWebCams", MsgExclamation
  Else
    'Selecciona el nombre de la carpeta
      strFolder = Trim(InputBox(colLanguage("K91").Caption, colLanguage("K92").Caption, _
                                colLanguage("K28").Caption & " " & trvProject(intActualProject).Nodes.Count))
      If strFolder <> "" Then
        'Añade el nodo al árbol
          With trvProject(intActualProject)
            .Nodes.Add .SelectedItem, tvwChild, getNextTreeKey(intActualProject, cnstStrFolder), strFolder, IconOpenFolder
            .SelectedItem.Expanded = True
          End With
        'Cambia la variable que indica que se han añadido modificaciones al documento
          blnSaved = False
      End If
  End If
End Sub

Private Sub addItem()
'--> Agrega un elemento al árbol
Dim frmItem As New frmParameters

  'Comprueba si puede agregar el elemento
    If trvProject(intActualProject).SelectedItem Is Nothing Then
      Set trvProject(intActualProject).SelectedItem = trvProject(intActualProject).Nodes.Item(1)
    End If
    If Left(trvProject(intActualProject).SelectedItem.Key, "1") = cnstStrProject Then '"No es posible agregar elementos dentro de otros elementos"
      messageBox.showMessage colLanguage("K93").Caption, "bauWebCams", MsgExclamation
    Else
      'Recoge los parámetros de la webCam
        With frmItem
          .Show vbModal
          If Not .blnCancel Then
            'Recoge los parámetros de la ventana de parámetros de la webcam
              addWebCamCollectionTree intActualProject, .strName, .strDescription, .strURL, .strURLWeb, _
                                      .strEMail, .strICQ, .intInterval
            'Cambia la variable que indica que se ha modificado el documento
              If intActualProject <> prjFavourites Then
                blnSaved = False
              End If
          End If
        End With
    End If
  'Libera el formulario
    Set frmItem = Nothing
End Sub

Private Sub changeParametersWebCam()
'--> Cambia los parámetros de un proyecto a documentar
Dim objWebCam As clsWebCam
Dim frmParam As New frmParameters
Dim strName As String, strURL As String

  If trvProject(intActualProject).SelectedItem Is Nothing Then
    messageBox.showMessage "Seleccione una webCam", "bauWebCams", MsgInformation
  ElseIf Left(trvProject(intActualProject).SelectedItem.Key, 1) <> cnstStrProject Then
    messageBox.showMessage "Seleccione una webCam", "bauWebCams", MsgInformation
  Else
    Set objWebCam = colWebCam(intActualProject).Item(trvProject(intActualProject).SelectedItem.Key)
    With frmParam
      .strName = objWebCam.Name
      .strDescription = objWebCam.Description
      .strURL = objWebCam.URL
      .strURLWeb = objWebCam.WebURL
      .intInterval = objWebCam.Interval
      .strEMail = objWebCam.eMail
      .strICQ = objWebCam.ICQ
      .Show vbModal
      If Not .blnCancel Then
        objWebCam.Name = .strName
        objWebCam.Description = .strDescription
        objWebCam.URL = .strURL
        objWebCam.WebURL = .strURLWeb
        objWebCam.Interval = .intInterval
        objWebCam.eMail = .strEMail
        objWebCam.ICQ = .strICQ
      End If
    End With
    trvProject(intActualProject).SelectedItem.Text = objWebCam.Name
    Set objWebCam = Nothing
    Set frmParam = Nothing
    If intActualProject <> prjFavourites Then
      blnSaved = False
    End If
  End If
End Sub

Private Sub addWebCamCollectionTree(ByVal intProject As eProjectType, ByVal strName As String, ByVal strDescription As String, _
                                    ByVal strURL As String, ByVal strURLWeb As String, _
                                    ByVal strEMail As String, ByVal strICQ As String, _
                                    ByVal intInterval As Integer)
'--> Rutina intermedia que añade un elemento al árbol y a la colección desde el menú y el árbol
Dim strKeyItem As String

  'Añade el nodo al árbol
    With trvProject(intProject)
      strKeyItem = getNextTreeKey(intProject, cnstStrProject)
      If .SelectedItem Is Nothing Then
        Set .SelectedItem = .Nodes.Item(1)
      End If
      .Nodes.Add .SelectedItem, tvwChild, strKeyItem, strName, IconCammera
      .SelectedItem.Expanded = True
    End With
  'Añade los parámetros del proyecto a la colección
    colWebCam(intProject).Add strName, strDescription, strURL, strURLWeb, strEMail, strICQ, intInterval, strKeyItem
End Sub

Private Sub showAbout()
'--> Muestra la ventana Acerca de ...
Dim frmAboutMe As New frmAbout

  'Muestra la ventana
    frmAboutMe.Show vbModal
  'Libera la memoria
    Set frmAboutMe = Nothing
End Sub

Private Sub openWeb()
'--> Abre una ventana del explorador
Dim objExplorer As New clsHyperlink

  With objExplorer
    'Establece las propiedades
      .URL = "http://www.galeon.com/bauconsultors/"
      .ExplorerStatus = SW_SHOWNORMAL
    'Abre el editor de correo
      .OpenURL
  End With
  'Libera la memoria
    Set objExplorer = Nothing
End Sub

Private Sub openWebWebCamm()
'--> Abre una ventana del explorador a la web de la webCam
Dim objWebCam As clsWebCam
Dim objExplorer As New clsHyperlink

  If trvProject(intActualProject).SelectedItem Is Nothing Then
    messageBox.showMessage "Seleccione una webCam", "bauWebCams", MsgInformation
  ElseIf Left(trvProject(intActualProject).SelectedItem.Key, 1) <> cnstStrProject Then
    messageBox.showMessage "Seleccione una webCam", "bauWebCams", MsgInformation
  Else
    Set objWebCam = colWebCam(intActualProject).Item(trvProject(intActualProject).SelectedItem.Key)
    If Trim(objWebCam.WebURL) = "" Then
      messageBox.showMessage "La webCam no tiene página asociada", "bauWebCams", MsgInformation
    Else
      With objExplorer
        'Establece las propiedades
          .URL = objWebCam.WebURL
          .ExplorerStatus = SW_SHOWNORMAL
        'Abre el editor de correo
          .OpenURL
      End With
    End If
    Set objWebCam = Nothing
  End If
  'Libera la memoria
    Set objExplorer = Nothing
End Sub

Private Function getNextTreeKey(ByVal intProject As eProjectType, ByVal strType As String)
'--> Obtiene una clave para el nodo de la carpeta / elemento sobre el árbol
Dim intIndexKey As Integer
Dim objNode As Node

  On Error GoTo errorGetNextTreeKey
    With trvProject(intProject)
      intIndexKey = .Nodes.Count
      While Not .Nodes.Item(strType & Format(intIndexKey, cnstStrTreeMask)) Is Nothing
        intIndexKey = intIndexKey + 1
      Wend
      getNextTreeKey = strType & Format(intIndexKey, cnstStrTreeMask)
    End With
  Exit Function
  
errorGetNextTreeKey:
  getNextTreeKey = strType & Format(intIndexKey, cnstStrTreeMask)
End Function

Private Sub createAVI()
'--> Crea un fichero AVI a partir de las imágenes
Dim frmAVI As New frmMakeAVI

  frmAVI.Show vbModal
  Set frmAVI = Nothing
End Sub

Private Sub ctlPopMenu_ItemHighlight(itemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)
  highlightOptionMenu itemNumber, bEnabled, bSeparator
End Sub

Private Sub ctlPopMenu_MenuExit()
  brStatus.Message = ""
End Sub

Private Sub MDIForm_Load()
  'Crea el formulario de pantalla inicial
    Me.Visible = False
    Set frmSplashScreen = New frmSplash
    frmSplashScreen.Show
    frmSplashScreen.SetFocus
  'Inicializa la pantalla
    intIndexWebCam = 1
    init
  'Libera el formulario de pantalla inicial y muestra el formulario principal
    frmSplashScreen.SetFocus
    frmSplashScreen.waitTime 5
    Unload frmSplashScreen
    Set frmSplash = Nothing
    Me.Visible = True
    Me.SetFocus
  'Comprueba si hay algo para cargar
    If Command <> "" Then
     loadProject prjFile, Replace(Command, cnstStrQuotes, "")
    End If
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Files As Variant

  'Recorre la lista de ficheros
    If Data.Files.Count > 0 Then
      loadProject prjFavourites, Data.Files(1)
    End If
  'No provoca ninguna acción, para más información ver Object Browser
    Effect = vbDropEffectNone
End Sub

Private Sub MDIForm_Resize()
  If Me.WindowState <> vbMinimized Then
    If Me.Width < 9090 Then
      Me.Width = 9090
    ElseIf Me.Height < 6465 Then
      Me.Height = 6465
    Else
      Resize
    End If
  End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  exitApplication
End Sub

Private Sub mnuAddFolder_Click()
  addFolder
End Sub

Private Sub mnuAddItem_Click()
  addItem
End Sub

Private Sub mnuChangeIdioma_Click(Index As Integer)
  changeLanguage Index
End Sub

Private Sub mnuCreateAvi_Click()
  createAVI
End Sub

Private Sub mnuDrop_Click()
  dropItem
End Sub

Private Sub mnuEditCopy_Click()
  copyWebCam
End Sub

Private Sub mnuEditPaste_Click()
  pasteWebCam
End Sub

Private Sub mnuExit_Click()
  exitApplication
End Sub

Private Sub mnuHelpAbout_Click()
  showAbout
End Sub

Private Sub mnuHelpIndex_Click()
  helpIndex
End Sub

Private Sub mnuHelpTip_Click()
  dayTip
End Sub

Private Sub mnuHelpWeb_Click()
  openWeb
End Sub

Private Sub mnuMRUFile_Click(Index As Integer)
  loadProject prjFile, colMRUFiles.Item(Index)
End Sub

Private Sub mnuNewProject_Click()
  newProject
End Sub

Private Sub mnuOpenCammera_Click()
  openCammera
End Sub

Private Sub mnuOpenProject_Click()
  openProject
End Sub

Private Sub mnuOptions_Click()
  setOptions
End Sub

Private Sub mnuPopUpAddFolder_Click()
  addFolder
End Sub

Private Sub mnuPopupAddItem_Click()
  addItem
End Sub

Private Sub mnuPopUpDrop_Click()
  dropItem
End Sub

Private Sub mnuPopUpEditCopy_Click()
  copyWebCam
End Sub

Private Sub mnuPopUpEditPaste_Click()
  pasteWebCam
End Sub

Private Sub mnuPopupOpenCammera_Click()
  openCammera
End Sub

Private Sub mnuPopupProperties_Click()
  changeParametersWebCam
End Sub

Private Sub mnuProperties_Click()
  changeParametersWebCam
End Sub

Private Sub mnuSaveProject_Click()
  saveProject
End Sub

Private Sub mnuSaveProjectAs_Click()
  strActualProject = ""
  saveProject
End Sub

Private Sub mnuSearchWebCam_Click()
  searchWeb
End Sub

Private Sub mnuShowBigIcons_Click()
  showBigIcons
End Sub

Private Sub mnuShowProject_Click()
  showPanelProyecto
End Sub

Private Sub mnuShowToolBar_Click()
  showToolBar
End Sub

Private Sub mnuShowVerticalMenu_Click()
  showVerticalMenu
End Sub

Private Sub mnuViewWebWebcam_Click()
  openWebWebCamm
End Sub

Private Sub mnuWindowArrange_Click(Index As Integer)
  Select Case Index
    Case 1 'Cascada
      Me.Arrange vbCascade
    Case 2 'Horizontal
      Me.Arrange vbTileHorizontal
    Case 3 'Vertical
      Me.Arrange vbTileVertical
    Case 4 'Organizar iconos
      Me.Arrange vbArrangeIcons
  End Select
End Sub

Private Sub picPannel_CloseWindow()
'--> Oculta la ventana del árbol de proyectos
  showPanelProyecto
End Sub

Private Sub picPannel_Resize()
  Resize
End Sub

Private Sub splVertical_Resize(ByVal SpliterLeft As Integer)
'--> Redimensiona los elementos de la ventana principal
  Resize
End Sub

Private Sub tabTree_Click()
  If tabTree.Tabs(1).Selected Then
    intActualProject = prjFile
  Else
    intActualProject = prjFavourites
  End If
  trvProject(prjFile).Visible = tabTree.Tabs(1).Selected
  trvProject(prjFavourites).Visible = tabTree.Tabs(2).Selected
End Sub

Private Sub tlbHerramientas_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index
    Case ButtonNew
      newProject
    Case ButtonOpen
      openProject
    Case ButtonSave
      saveProject
    Case ButtonCut
    Case ButtonCopy
      copyWebCam
    Case ButtonPaste
      pasteWebCam
    Case ButtonProperties
      changeParametersWebCam
    Case ButtonOpenWebCammera
      openWebWebCamm
    Case ButtonOpenCammera
      openCammera
    Case ButtonHelp
      messageBox.showMessage "Help", "bauWebCams", MsgInformation
  End Select
End Sub

Private Sub tlbHerramientas_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
  Select Case UCase(ButtonMenu.Key)
    Case "N1"
      newProject
    Case "N2"
      addFolder
    Case "N3"
      addItem
  End Select
End Sub

Private Sub trvProject_AfterLabelEdit(Index As Integer, Cancel As Integer, NewString As String)
  If Not trvProject(Index).SelectedItem Is Nothing Then
    If Left(trvProject(Index).SelectedItem.Key, 1) = cnstStrProject Then
      colWebCam(intActualProject).Item(trvProject(Index).SelectedItem.Key).Name = NewString
    End If
  End If
End Sub

Private Sub trvProject_BeforeLabelEdit(Index As Integer, Cancel As Integer)
  If Not trvProject(Index).SelectedItem Is Nothing Then
    If Left(trvProject(Index).SelectedItem.Key, 1) = cnstStrTreeRoot Then
      Cancel = True
    End If
  End If
End Sub

Private Sub trvProject_Collapse(Index As Integer, ByVal Node As MSComctlLib.Node)
  If Left(Node.Key, 1) = cnstStrFolder Then
    Node.Image = IconCloseFolder
  End If
End Sub

Private Sub trvProject_DblClick(Index As Integer)
  If Not trvProject(Index).SelectedItem Is Nothing Then
    If Left(trvProject(Index).SelectedItem.Key, 1) = cnstStrProject Then
      openCammera
    End If
  End If
End Sub

Private Sub trvProject_Expand(Index As Integer, ByVal Node As MSComctlLib.Node)
  If Left(Node.Key, 1) = cnstStrFolder Then
    Node.Image = IconOpenFolder
  End If
End Sub

Private Sub trvProject_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbRightButton Then
    Me.PopupMenu mnuPopUp
    'ctlPopMenu.ShowPopupMenu(trvProject, "mnupopup", 0, 0)
  End If
End Sub

Private Sub trvProject_OLEDragDrop(Index As Integer, Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
  On Error Resume Next
  loadProject prjFile, Data.Files(1)
End Sub

Private Sub vrtMnuProject_MenuItemClick(MenuNumber As Long, MenuItem As Long)
'--> Controla la pulsación sobre el menú vertical
  Select Case MenuNumber
    Case 1 'Archivo
      Select Case MenuItem
        Case 1 'Nuevo
          newProject
        Case 2 'Abrir
          openProject
        Case 3 'Guardar
          saveProject
        Case 4 'Buscar en la web
          searchWeb
      End Select
    Case 2 'Elemento
      Select Case MenuItem
        Case 1 'Carpeta
          addFolder
        Case 2 'Elemento
          addItem
        Case 3 'Propiedades
          changeParametersWebCam
        Case 4 'Eliminar
          dropItem
        Case 5 'Abrir cámara
          openCammera
      End Select
    Case 3 'Ayuda
      Select Case MenuItem
        Case 1 'Ayuda
          helpIndex
        Case 2 'Acerca de
          showAbout
        Case 3 'Web
          openWeb
      End Select
  End Select
End Sub
