VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form admin 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DATA ADMIN"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdubah 
      BackColor       =   &H00808000&
      Caption         =   "UBAH"
      Height          =   495
      Left            =   2280
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdtambah 
      BackColor       =   &H00808000&
      Caption         =   "TAMBAH"
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   480
      TabIndex        =   12
      Top             =   3840
      Width           =   9975
      Begin VB.CommandButton cmdsimpan 
         BackColor       =   &H00808000&
         Caption         =   "SIMPAN"
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FFFF&
         Caption         =   "CARI"
         Height          =   495
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00808000&
         Caption         =   "HAPUS"
         Height          =   495
         Left            =   3360
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   5040
         TabIndex        =   13
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "CARI"
      Height          =   495
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4080
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   8520
      Top             =   6600
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\APP REKAP TEKNISI\app_rekap.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\APP REKAP TEKNISI\app_rekap.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "login"
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
      Bindings        =   "admin.frx":0000
      Height          =   1455
      Left            =   480
      TabIndex        =   10
      Top             =   5040
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   2566
      _Version        =   393216
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
            LCID            =   1057
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
            LCID            =   1057
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
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   6480
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3000
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   6480
      TabIndex        =   4
      Top             =   1800
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.Label Label5 
         BackColor       =   &H00808000&
         Caption         =   "DATA ADMIN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "USERNAME"
      Height          =   255
      Left            =   6480
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PASSWORD"
      Height          =   255
      Left            =   6480
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NIK"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NAMA ADMIN"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2640
      Width           =   1335
   End
End
Attribute VB_Name = "admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'CARI
Private Sub Command4_Click()
Adodc1.Recordset.Filter = "nama_admin like '%" + Me.Text5.Text + "%' or nik_admin like '%" + Me.Text5.Text + "%'"
End Sub

'MUNCULKAN DATA SAAT PENCARIAN BERAKHIR
Private Sub Text5_Change()
If Text5.Text = "" Then
Adodc1.Refresh
Call table
Else
'wkwk
End If
End Sub

'OPERASI SAAT FORM DIBUKA
Private Sub Form_Load()
cmdtambah.Visible = True
Call clear
Call table
End Sub

'BERSIHKAN FORM
Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

'HIDUPKAN FORM
Sub enabel()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text1.SetFocus
End Sub

'FORMAT TABEL
Sub table()
With DataGrid1
.Columns(0).Width = 2200
.Columns(1).Width = 3500
.Columns(2).Width = 2000
.Columns(3).Width = 2000

.Columns(0).Caption = "NIK ADMIN"
.Columns(1).Caption = "NAMA ADMIN"
.Columns(2).Caption = "USERNAME"
.Columns(3).Caption = "PASSWORD"

.Columns(3).Value = "*****"
End With
End Sub

'PINDAH DATA DARI TABEL KE FORM
Private Sub DataGrid1_Click()
cmdtambah.Visible = True
Text1.Text = Adodc1.Recordset!nik_admin
Text2.Text = Adodc1.Recordset!nama_admin
Text3.Text = Adodc1.Recordset!UserName
Text4.Text = Adodc1.Recordset!Password
End Sub

'TAMBAH
Private Sub cmdtambah_Click()
cmdtambah.Visible = False
Call clear
Call enabel
End Sub

'SIMPAN
Private Sub cmdsimpan_Click()
'jika ada inputan yang kosong, tampilkan pesan peringatan
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "LENGKAPI DAHULU DATA YANG AKAN ANDA INPUTKAN !", vbInformation, "PERHATIAN !"
Else
'jika semua form sudah terisi, lakukan operasi simpan
Adodc1.Recordset.AddNew 'untuk tambah record'
Adodc1.Recordset.Fields("nik_admin") = Text1.Text
Adodc1.Recordset.Fields("nama_admin") = Text2.Text
Adodc1.Recordset.Fields("username") = Text3.Text
Adodc1.Recordset.Fields("password") = Text4.Text
Adodc1.Recordset.Update
MsgBox "DATA BERHASIL DITAMBAHKAN !", vbInformation, "INFORMASI !"
Call clear
End If
End Sub

'UBAH
Private Sub cmdubah_Click()
'jika ada inputan yang kosong, tampilkan pesan peringatan
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "PILIH DAHULU DATA YANG AKAN DIUBAH !", vbInformation, "PERHATIAN !"
Else
'jika semua form sudah terisi, lakukan operasi ubah
Adodc1.Recordset.Fields("nik_admin") = Text1.Text
Adodc1.Recordset.Fields("nama_admin") = Text2.Text
Adodc1.Recordset.Fields("username") = Text3.Text
Adodc1.Recordset.Fields("password") = Text4.Text
Adodc1.Recordset.Update
MsgBox "DATA BERHASIL DIUBAH !", vbInformation, "INFORMASI !"
Call clear
cmdtambah.Visible = False
End If
End Sub

'HAPUS
Private Sub Command2_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Then
MsgBox "PILIH DAHULU DATA YANG AKAN ANDA HAPUS !", vbInformation, "PERHATIAN !"
Else
xx = MsgBox("Apakah Anda yakin akan menghapus data?", vbOKCancel, "Peringatan")
            If xx = vbOK Then
               Adodc1.Recordset.Delete
               Call clear
MsgBox "DATA ANDA BERHASIL DIHAPUS !", vbInformation, "INFORMASI !"
'Adodc1.Refresh
            End If
           
End If
End Sub


