VERSION 5.00
Begin VB.Form Intro 
   BackColor       =   &H00000000&
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
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2400
      Top             =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
End
Attribute VB_Name = "Intro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Resize()
Label1.Caption = "Zaid Markabi" + vbCrLf + vbCrLf + "Present"
Label1.Width = Me.Width
Label1.Top = (Me.Height - Label1.Height) \ 2
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Menu.Show
Unload Me
End Sub
