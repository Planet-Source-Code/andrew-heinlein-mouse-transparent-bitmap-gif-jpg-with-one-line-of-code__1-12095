VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1740
      Left            =   3960
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1740
      ScaleWidth      =   1530
      TabIndex        =   3
      Top             =   1320
      Width           =   1530
   End
   Begin VB.Timer Timer1 
      Left            =   1560
      Top             =   3120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do It!"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   3240
      Width           =   2775
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1740
      Left            =   120
      Picture         =   "Form1.frx":0EDD
      ScaleHeight     =   1740
      ScaleWidth      =   1530
      TabIndex        =   1
      Top             =   1320
      Width           =   1530
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   120
      Picture         =   "Form1.frx":1DBA
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   360
      TabIndex        =   0
      Top             =   360
      Width           =   5400
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Source Picture:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Streched:"
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Regular:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Timer1.Enabled = True
    Timer1.Interval = 255
End Sub

Private Sub Timer1_Timer()
    Static LAST_FRAME As Integer
    
    Const FRAME_WIDTH = 40
    Const FRAME_HEIGHT = 40
    
    Picture2.Cls
    Picture3.Cls
    
    'i set the transparent color to be VbWhite... which is &HFFFFFF
    TransparentBlt Picture2.hDC, 40, 40, FRAME_WIDTH, FRAME_HEIGHT, Picture1.hDC, LAST_FRAME * FRAME_WIDTH, 0, FRAME_WIDTH, FRAME_HEIGHT, vbWhite
    TransparentBlt Picture3.hDC, 0, 0, Picture3.Width / 15, Picture3.Height / 15, Picture1.hDC, LAST_FRAME * FRAME_WIDTH, 0, FRAME_WIDTH, FRAME_HEIGHT, vbWhite
    
    LAST_FRAME = LAST_FRAME + 1
    
    If LAST_FRAME = 9 Then LAST_FRAME = 0
End Sub
