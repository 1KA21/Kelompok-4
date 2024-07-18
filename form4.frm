VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7545
   LinkTopic       =   "Form4"
   ScaleHeight     =   4545
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000014&
      Caption         =   "Kembali"
      Height          =   495
      Left            =   6000
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000014&
      Caption         =   "Hdd"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000014&
      Caption         =   "Casing"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000014&
      Caption         =   "Fan"
      Height          =   255
      Left            =   3720
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000014&
      Caption         =   "VgaCard"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000014&
      Caption         =   "Ssd"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000014&
      Caption         =   "Ram"
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000014&
      Caption         =   "Keyboard"
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000014&
      Caption         =   "Mouse"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.Image Image9 
      Height          =   1335
      Left            =   1920
      Picture         =   "form4.frx":0000
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Image Image8 
      Height          =   1335
      Left            =   120
      Picture         =   "form4.frx":5ACA
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Image Image7 
      Height          =   1335
      Left            =   3720
      Picture         =   "form4.frx":69AF
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Image Image6 
      Height          =   1335
      Left            =   1920
      Picture         =   "form4.frx":15398
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Image Image5 
      Height          =   1335
      Left            =   120
      Picture         =   "form4.frx":1DD94
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Image Image4 
      Height          =   1335
      Left            =   3720
      Picture         =   "form4.frx":2D289
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image3 
      Height          =   1335
      Left            =   1920
      Picture         =   "form4.frx":3F7BB
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   1335
      Left            =   120
      Picture         =   "form4.frx":4DD37
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000B&
      Height          =   4695
      Left            =   -120
      TabIndex        =   1
      Top             =   -120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   5625
      Left            =   3240
      Picture         =   "form4.frx":51D45
      Top             =   0
      Width           =   8445
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form4.Hide
Form2.Show
End Sub
