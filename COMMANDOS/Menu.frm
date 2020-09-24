VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Image Image6 
      Height          =   510
      Left            =   4080
      Picture         =   "Menu.frx":0000
      Top             =   4680
      Width           =   4590
   End
   Begin VB.Image Image5 
      Height          =   510
      Left            =   4080
      Picture         =   "Menu.frx":3CC8
      Top             =   3960
      Width           =   4605
   End
   Begin VB.Image Image4 
      Height          =   990
      Left            =   1200
      Picture         =   "Menu.frx":7D80
      Top             =   1920
      Width           =   8685
   End
   Begin VB.Image Image3 
      Height          =   1170
      Left            =   720
      Picture         =   "Menu.frx":D1F0
      Top             =   480
      Width           =   9525
   End
   Begin VB.Image Image2 
      Height          =   4995
      Left            =   360
      Picture         =   "Menu.frx":10804
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
SoundEngine.Init Me.hWnd
Set Listener = SoundEngine.Get3DListener
SFXMenu(0).Load App.Path + "\Data\Sfx\15.wav"
SFXMenu(0).Loop_ = True
SFXMenu(0).Play

Image1.Picture = LoadPicture(App.Path + "\Data\Gfx\Menu.jpg")
End Sub

Private Sub Form_Resize()
Image1.Width = Me.Width
Image1.Height = Me.Height
Image2.Top = Me.Height - Image2.Height
Image2.Visible = True
End Sub

Private Sub Image5_Click()
On Error Resume Next
Open App.Path + "\Data\Mission.txt" For Output As #1
Write #1, "1"
Close #1
Main.Show
Unload Me
End Sub

Private Sub Image6_Click()
On Error Resume Next
Main.Show
Unload Me
End Sub
