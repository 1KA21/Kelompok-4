VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   5265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10650
   LinkTopic       =   "Form3"
   ScaleHeight     =   5265
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      BackColor       =   &H80000014&
      Caption         =   "Kembali"
      Height          =   495
      Left            =   8040
      TabIndex        =   19
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   1800
      TabIndex        =   11
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   1800
      TabIndex        =   10
      Top             =   2160
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H80000014&
      Caption         =   "Keluar"
      Height          =   495
      Left            =   8040
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H80000014&
      Caption         =   "Hapus"
      Height          =   495
      Left            =   1800
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000014&
      Caption         =   "Update"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000014&
      Caption         =   "Tambahkan"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1800
      Top             =   3360
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Adrian Taufiq\Documents\DataPembeli.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Adrian Taufiq\Documents\DataPembeli.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "DataPembeli"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form3.frx":0000
      Height          =   2295
      Left            =   3840
      TabIndex        =   2
      Top             =   960
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4048
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483628
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   360
      Width           =   4095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000014&
      Caption         =   "Cari"
      Height          =   495
      Left            =   8040
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   5625
      Left            =   6000
      Picture         =   "Form3.frx":0015
      Top             =   0
      Width           =   8445
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000018&
      Caption         =   "Pencarian"
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000014&
      Caption         =   "TanggalTransaksi"
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000014&
      Caption         =   "NamaBarang"
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000014&
      Caption         =   "Alamat"
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000014&
      Caption         =   "Nama"
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000014&
      Caption         =   "IdPembeli"
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000018&
      Height          =   5655
      Left            =   -1800
      TabIndex        =   12
      Top             =   -120
      Width           =   10815
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub bersih()
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub


Private Sub Command1_Click()
Adodc1.Recordset.Find "IdPembeli='" & Text1.Text & "'", , adSearchForward, 1
If Not Adodc1.Recordset.EOF Then
Text2.Text = Adodc1.Recordset!IdPembeli
Text3.Text = Adodc1.Recordset!Nama
Text4.Text = Adodc1.Recordset!Alamat
Text5.Text = Adodc1.Recordset!NamaBarang
Text6.Text = Adodc1.Recordset!TanggalTransaksi
Else
MsgBox ("Data Tidak Di Temukan")
End If
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.AddNew
Adodc1.Recordset!IdPembeli = Text2.Text
Adodc1.Recordset!Nama = Text3.Text
Adodc1.Recordset!Alamat = Text4.Text
Adodc1.Recordset!NamaBarang = Text5.Text
Adodc1.Recordset!TanggalTransaksi = Text6.Text
Call bersih

End Sub

Private Sub Command3_Click()
Adodc1.Recordset.Update
Adodc1.Recordset!IdPembeli = Text2.Text
Adodc1.Recordset!Nama = Text3.Text
Adodc1.Recordset!Alamat = Text4.Text
Adodc1.Recordset!NamaBarang = Text5.Text
Adodc1.Recordset!TanggalTransaksi = Text6.Text
Call bersih
End Sub

Private Sub Command4_Click()
konfirmasi = MsgBox("Yakin Akan Di Hapus???", vbYesNo + vbInformation, "Konfirmasi")
If konfirmasi = vbYes Then
Adodc1.Recordset.Delete
Else
End If
End Sub

Private Sub Command5_Click()
konfirmasi = MsgBox("Yakin Mau Keluar???", vbYesNo + vbInformation, "Keluar")
If konfirmasi = vbYes Then
End
Else
End If
End Sub


Sub cek_data()
If Text2.Text = "" And Text3.Text = "" And Text4.Text = "" And Text5.Text = "" And Text6.Text = "" Then
MsgBox ("Data Belum Lengkap")
End If

End Sub

Private Sub Command6_Click()
Form3.Hide
Form2.Show

End Sub

