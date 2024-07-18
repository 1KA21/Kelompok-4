VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   2520
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Masuk"
      Height          =   495
      Left            =   7080
      TabIndex        =   0
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000018&
      Caption         =   "Kelompok             Adrian Taufiq, Chandra Budianto, Rafidhia Izdihar,       M Ilyas"
      Height          =   1095
      Left            =   5520
      TabIndex        =   7
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000018&
      Caption         =   "Password"
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      Caption         =   "Username"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000018&
      Height          =   5175
      Left            =   5400
      TabIndex        =   1
      Top             =   0
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   5205
      Left            =   -1320
      Picture         =   "Form1.frx":0000
      Top             =   0
      Width           =   8460
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "user1" And Text2.Text = "123" Then
pesan = MsgBox("Login Berhasil", vbinformaion, "success")
Form1.Hide
Form2.Show

ElseIf Text1.Text = "" And Text2.Text = "" Then
pesan = MsgBox("silahkan isi user dan Password", vbInformation, "hint")
ElseIf Text1.Text = "" Then
pesan = MsgBox("user kosong", vbCritical, "eror!")
ElseIf Text1.Text = "" Then
pesan = MsgBox("passsword kosong", vbCritical, "eror!")
Else
pesan = MsgBox("user dan password salah!", vbCritical, "eror!")
End If
End Sub

Private Sub Command2_Click()
End
End Sub
