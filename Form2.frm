VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8610
   LinkTopic       =   "Form2"
   ScaleHeight     =   4995
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000014&
      Height          =   855
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000014&
      Height          =   855
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000014&
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000014&
      Height          =   5295
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Kembali Ke Menu Login"
      BeginProperty Font 
         Name            =   "Unispace"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Melihat       Produk"
      BeginProperty Font 
         Name            =   "Unispace"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   5
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Mencatat     Laporan"
      BeginProperty Font 
         Name            =   "Unispace"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   4
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   8655
   End
   Begin VB.Image Image1 
      Height          =   5625
      Left            =   1320
      Picture         =   "Form2.frx":0000
      Top             =   -120
      Width           =   8445
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Hide
Form3.Show

End Sub

Private Sub Command2_Click()
Form2.Hide
Form4.Show

End Sub

Private Sub Command3_Click()
Form2.Hide
Form1.Show

End Sub

