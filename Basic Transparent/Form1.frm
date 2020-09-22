VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   6150
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1740
      Left            =   2760
      ScaleHeight     =   1740
      ScaleWidth      =   1530
      TabIndex        =   1
      Top             =   480
      Width           =   1530
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1740
      Left            =   360
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1740
      ScaleWidth      =   1530
      TabIndex        =   0
      Top             =   480
      Width           =   1530
   End
   Begin VB.Label Label6 
      Caption         =   "Destination:"
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Current Transp:"
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "Mouse over:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Source Pic:"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2280
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TEMP_COLOR As Long
    'get the color of the pixel the mouse is over:
    TEMP_COLOR = GetPixel(Picture1.hdc, X / 15, Y / 15)
    'set the label's back color to show what we picked:
    Label4.BackColor = TEMP_COLOR
    'clear off the picturebox:
    Picture2.Cls
    'use the function accordinly
    'NOTE: when a value is divided by 15 (i.e. Picture2.Width / 15) it is just converting TWIPS to PIXELS
    'the RIGHT way to do this is: Picture2.Width / Screen.TwipsPerPixelX... but im lazy
    TransparentBlt Picture2.hdc, 0, 0, Picture2.Width / 15, Picture2.Height / 15, Picture1.hdc, 0, 0, Picture1.Width / 15, Picture1.Height / 15, TEMP_COLOR
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'see above if you dont get this:
    Label1.BackColor = GetPixel(Picture1.hdc, X / 15, Y / 15)
End Sub
