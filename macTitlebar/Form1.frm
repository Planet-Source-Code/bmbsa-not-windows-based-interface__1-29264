VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "bmbas@hotmail.com"
   ClientHeight    =   7590
   ClientLeft      =   3840
   ClientTop       =   2745
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   5775
      Left            =   360
      ScaleHeight     =   5715
      ScaleWidth      =   8475
      TabIndex        =   2
      Top             =   720
      Width           =   8535
   End
   Begin VB.PictureBox mainContainer 
      Height          =   7335
      Left            =   240
      ScaleHeight     =   7275
      ScaleWidth      =   8715
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin Project1.ButtonEx ButtonEx1 
         Height          =   600
         Left            =   3600
         TabIndex        =   3
         Top             =   6600
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1058
         Caption         =   "Exit"
         ForeColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TransparentColor=   16777215
         SkinDown        =   "Form1.frx":0000
         SkinFocus       =   "Form1.frx":3B12
         SkinUp          =   "Form1.frx":7624
         TransparentColor=   16777215
      End
      Begin VB.PictureBox titlebar 
         Height          =   375
         Left            =   720
         ScaleHeight     =   315
         ScaleWidth      =   7515
         TabIndex        =   1
         Top             =   120
         Width           =   7575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButtonEx1_Click()
End
End Sub

Private Sub Form_Load()
    Call CreateMacOSTitleBar(titlebar, " " & Caption & " ", Me)
    
    mainContainer.Top = 0: mainContainer.Left = 0
    mainContainer.Width = Me.Width: mainContainer.Height = Me.Height
    Call ColForm(mainContainer, 217, 211, 213, 125)
    
    Call ColForm(Picture1, 217, 211, 213, 125)
    
End Sub

Private Sub titlebar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DragForm(Me)
End Sub
