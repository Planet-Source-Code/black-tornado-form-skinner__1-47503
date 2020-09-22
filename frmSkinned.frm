VERSION 5.00
Begin VB.Form frmSkinner 
   BorderStyle     =   0  'None
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7905
   DrawMode        =   5  'Not Copy Pen
   FillStyle       =   0  'Solid
   Icon            =   "frmSkinned.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSkinned.frx":08CA
   ScaleHeight     =   352
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   527
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "e-mail me"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   4440
      Width           =   1455
   End
   Begin VB.PictureBox PicMainSkin 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmSkinner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Shell "start mailto:btsoft@burntmail.com?Subject=FormSkinner"
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Form_Load()
    Dim WindowRegion As Long
    PicMainSkin.ScaleMode = vbPixels
    PicMainSkin.AutoRedraw = True
    PicMainSkin.AutoSize = True
    PicMainSkin.BorderStyle = vbBSNone
    BorderStyle = vbBSNone
    Set PicMainSkin.Picture = LoadPicture(App.Path & "\skin.bmp")   ' loads the skin
    Width = PicMainSkin.Width * 15
    Height = PicMainSkin.Height * 15
    
    WindowRegion = MakeRegion(PicMainSkin)
    SetWindowRgn hwnd, WindowRegion, True
    Me.Refresh
    Me.Picture = PicMainSkin.Picture
Me.Refresh
Me.Move 0, 0
End Sub

