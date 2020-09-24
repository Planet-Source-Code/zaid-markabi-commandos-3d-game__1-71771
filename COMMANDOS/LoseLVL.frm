VERSION 5.00
Begin VB.Form LoseLVL 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Image Image2 
      Height          =   630
      Left            =   720
      Picture         =   "LoseLVL.frx":0000
      Top             =   720
      Width           =   4605
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "LoseLVL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
Image1.Picture = LoadPicture(App.Path + "\Data\Gfx\Menu.jpg")
Image1.Width = Me.Width
Image1.Height = Me.Height
End Sub

Private Sub Image1_Click()
'Shell App.Path + "\" + App.EXEName + ".exe"
End
End Sub
