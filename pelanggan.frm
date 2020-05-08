VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form pelanggan 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DATA PELANGGAN"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15195
   LinkTopic       =   "Form1"
   ScaleHeight     =   10500
   ScaleWidth      =   15195
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo4 
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   315
      Left            =   10560
      TabIndex        =   31
      Top             =   5520
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   16335
      Begin VB.Label Label13 
         BackColor       =   &H00808000&
         Caption         =   "DATA PELANGGAN"
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
         Left            =   600
         TabIndex        =   30
         Top             =   120
         Width           =   5175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   600
      TabIndex        =   23
      Top             =   6600
      Width           =   13935
      Begin VB.CommandButton cmdtambah 
         BackColor       =   &H00808000&
         Caption         =   "TAMBAH"
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdhapus 
         BackColor       =   &H00808000&
         Caption         =   "HAPUS"
         Height          =   495
         Left            =   3840
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdubah 
         BackColor       =   &H00808000&
         Caption         =   "UBAH"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FFFF&
         Caption         =   "CARI"
         Height          =   495
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdsimpan 
         BackColor       =   &H00808000&
         Caption         =   "SIMPAN"
         Height          =   495
         Left            =   240
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   9000
         TabIndex        =   24
         Top             =   240
         Width           =   3495
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "pelanggan.frx":0000
      Height          =   2175
      Left            =   480
      TabIndex        =   22
      Top             =   7800
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   3836
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
   Begin VB.ComboBox Combo3 
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   315
      Left            =   10560
      TabIndex        =   20
      Top             =   4320
      Width           =   3975
   End
   Begin VB.ComboBox Combo2 
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   315
      Left            =   10560
      TabIndex        =   18
      Top             =   3000
      Width           =   3975
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   5760
      TabIndex        =   7
      Top             =   5520
      Width           =   3975
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   5760
      TabIndex        =   6
      Top             =   4200
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   615
      Left            =   5760
      TabIndex        =   5
      Top             =   2880
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   228524033
      CurrentDate     =   43702
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   5760
      TabIndex        =   4
      Top             =   1560
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   600
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   5520
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   4200
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      TabIndex        =   1
      Top             =   2880
      Width           =   3975
   End
   Begin VB.TextBox text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   1560
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   615
      Left            =   10560
      TabIndex        =   16
      Top             =   1560
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   228524033
      CurrentDate     =   43702
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   480
      Top             =   10080
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
      RecordSource    =   "pelanggan"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   2640
      Top             =   10080
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
      RecordSource    =   "combo"
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
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BANDWITH"
      Height          =   255
      Left            =   10560
      TabIndex        =   32
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Line Line2 
      X1              =   10200
      X2              =   10200
      Y1              =   1080
      Y2              =   6240
   End
   Begin VB.Line Line1 
      X1              =   5160
      X2              =   5160
      Y1              =   1080
      Y2              =   6240
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TEKNOLOGI"
      Height          =   255
      Left            =   10560
      TabIndex        =   21
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      Caption         =   "WORKZONE"
      Height          =   255
      Left            =   10560
      TabIndex        =   19
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CONTACT PERSON"
      Height          =   255
      Left            =   5760
      TabIndex        =   17
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TANGGAL MULAI BERLANGANAN"
      Height          =   255
      Left            =   10560
      TabIndex        =   15
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ALAMAT"
      Height          =   255
      Left            =   5760
      TabIndex        =   14
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TANGGAL LAHIR"
      Height          =   255
      Left            =   5760
      TabIndex        =   13
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TEMPAT LAHIR"
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "JENIS KELAMIN"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NAMA PELANGGAN"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NO KTP"
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ID PELANGGAN"
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "pelanggan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'MENAMPILKAN DATA PADA DATABASE KE COMBO
Sub tambahcom()
Adodc2.ConnectionString = conn.ConnectionString
Adodc2.RecordSource = "select * from combo"

For Each gosong In Me.Controls
If TypeOf gosong Is ComboBox Then
gosong.Text = ""
With Adodc2.Recordset
    Do While Not .EOF
    On Error Resume Next
    Combo2.AddItem !workzone
    Combo3.AddItem !teknologi
    Combo4.AddItem !bandwith
    .MoveNext
    Loop
End With
End If
Next
End Sub

'kode anggota otomatis
Sub KodeOtomatis()
Call Koneksi
RS.Open ("select * from pelanggan Where id_pelanggan In(Select Max(id_pelanggan)From pelanggan)Order By id_pelanggan Desc"), conn
RS.Requery
    Dim Urutan As String * 6
    Dim Hitung As Long
    With RS
        If .EOF Then
            Urutan = "PG-" + "001"
            text1 = Urutan
        Else
            Hitung = Right(!id_pelanggan, 3) + 1
            Urutan = "PG-" + Right("000" & Hitung, 3)
        End If
        text1 = Urutan
    End With
End Sub

'CARI
Private Sub Command4_Click()
Adodc1.Recordset.Filter = "id_pelanggan like '%" + Me.Text7.Text + "%' or no_ktp like '%" + Me.Text7.Text + "%' or nama_pelanggan like '%" + Me.Text7.Text + "%'"
End Sub

'MUNCULKAN DATA SAAT PENCARIAN BERAKHIR
Private Sub Text7_Change()
If Text7.Text = "" Then
Adodc1.Refresh
Call table
Else
'wkwk
End If
End Sub


'OPERASI SAAT FORM DIBUKA
Private Sub Form_Load()
Call tambahcom
cmdtambah.Visible = True
Call clear
Call table
Call KodeOtomatis

Combo1.AddItem ("LAKI-LAKI")
Combo1.AddItem ("PEREMPUAN")
End Sub

'BERSIHKAN FORM
Sub clear()
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
DTPicker1.Value = Now
DTPicker1.Value = Now
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Combo4.Text = ""
'Text2.SetFocus
End Sub

'HIDUPKAN FORM
Sub enabel()
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
DTPicker1.Enabled = True
DTPicker1.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
Combo3.Enabled = True
Combo4.Enabled = True
cmdtambah.Enabled = True
Text2.SetFocus
End Sub

'FORMAT TABEL
Sub table()
With DataGrid1
.Columns(0).Width = 1000
.Columns(1).Width = 1500
.Columns(2).Width = 1500
.Columns(3).Width = 2000
.Columns(4).Width = 1500
.Columns(5).Width = 2000
.Columns(6).Width = 1500
.Columns(7).Width = 1500
.Columns(8).Width = 1500
.Columns(9).Width = 1500
.Columns(10).Width = 1500
.Columns(11).Width = 1500

.Columns(0).Caption = "ID PELANGGAN"
.Columns(1).Caption = "NO KTP"
.Columns(2).Caption = "NAMA"
.Columns(3).Caption = "JENIS KELAMIN"
.Columns(4).Caption = "TEMPAT LAHIR"
.Columns(5).Caption = "TANGGAL LAHIR"
.Columns(6).Caption = "ALAMAT"
.Columns(7).Caption = "CONTACT PERSON"
.Columns(8).Caption = "TGL MULAI"
.Columns(9).Caption = "WORKZONE"
.Columns(10).Caption = "TEKNOLOGI"
.Columns(11).Caption = "BANDWITH"
End With
'seting waktu pada datagrid
With DataGrid1
.Columns(5).NumberFormat = "dd MMMM yy"
End With
End Sub

'PINDAH DATA DARI TABEL KE FORM
Private Sub DataGrid1_Click()
cmdtambah.Visible = True
text1.Text = Adodc1.Recordset.Fields("id_pelanggan")
Text2.Text = Adodc1.Recordset.Fields("no_ktp")
Text3.Text = Adodc1.Recordset.Fields("nama_pelanggan")
Combo1.Text = Adodc1.Recordset.Fields("jenis_kelamin")
Text4.Text = Adodc1.Recordset.Fields("tempat_lahir")
DTPicker1.Value = Adodc1.Recordset.Fields("tanggal_lahir")
Text5.Text = Adodc1.Recordset.Fields("alamat")
Text6.Text = Adodc1.Recordset.Fields("cp")
DTPicker2.Value = Adodc1.Recordset.Fields("tgl_gabung")
Combo2.Text = Adodc1.Recordset.Fields("workzone")
Combo3.Text = Adodc1.Recordset.Fields("teknologi")
Combo4.Text = Adodc1.Recordset.Fields("bandwith")
End Sub

'TAMBAH
Private Sub cmdtambah_Click()
cmdtambah.Visible = False
Call clear
Call enabel
Call KodeOtomatis

End Sub

'SIMPAN
Private Sub cmdsimpan_Click()
'jika ada inputan yang kosong, tampilkan pesan peringatan
If text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Or Combo3.Text = "" Or Combo4.Text = "" Then
MsgBox "LENGKAPI DAHULU DATA YANG AKAN DIINPUTKAN !", vbInformation, "PERHATIAN !"
Else
'jika semua form sudah terisi, lakukan operasi simpan
Adodc1.Recordset.AddNew 'untuk tambah record'
Adodc1.Recordset.Fields("id_pelanggan") = text1.Text
Adodc1.Recordset.Fields("no_ktp") = Text2.Text
Adodc1.Recordset.Fields("nama_pelanggan") = Text3.Text
Adodc1.Recordset.Fields("jenis_kelamin") = Combo1.Text
Adodc1.Recordset.Fields("tempat_lahir") = Text4.Text
Adodc1.Recordset.Fields("tanggal_lahir") = DTPicker1.Value
Adodc1.Recordset.Fields("alamat") = Text5.Text
Adodc1.Recordset.Fields("cp") = Text6.Text
Adodc1.Recordset.Fields("tgl_gabung") = DTPicker2.Value
Adodc1.Recordset.Fields("workzone") = Combo2.Text
Adodc1.Recordset.Fields("teknologi") = Combo3.Text
Adodc1.Recordset.Fields("bandwith") = Combo4.Text
Adodc1.Recordset.Update
MsgBox "DATA BERHASIL DITAMBAHKAN !", vbInformation, "INFORMASI !"
Call clear
Call KodeOtomatis
End If
End Sub

'UBAH
Private Sub cmdubah_Click()
'jika ada inputan yang kosong, tampilkan pesan peringatan
'jika ada inputan yang kosong, tampilkan pesan peringatan
If text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Or Combo3.Text = "" Or Combo4.Text = "" Then
MsgBox "PILIH DAHULU DATA YANG AKAN DIUBAH !", vbInformation, "PERHATIAN !"
Else
'jika semua form sudah terisi, lakukan operasi simpan
Adodc1.Recordset.Fields("id_pelanggan") = text1.Text
Adodc1.Recordset.Fields("no_ktp") = Text2.Text
Adodc1.Recordset.Fields("nama_pelanggan") = Text3.Text
Adodc1.Recordset.Fields("jenis_kelamin") = Combo1.Text
Adodc1.Recordset.Fields("tempat_lahir") = Text4.Text
Adodc1.Recordset.Fields("tanggal_lahir") = DTPicker1.Value
Adodc1.Recordset.Fields("alamat") = Text5.Text
Adodc1.Recordset.Fields("cp") = Text6.Text
Adodc1.Recordset.Fields("tgl_gabung") = DTPicker2.Value
Adodc1.Recordset.Fields("workzone") = Combo2.Text
Adodc1.Recordset.Fields("teknologi") = Combo3.Text
Adodc1.Recordset.Fields("bandwith") = Combo4.Text
Adodc1.Recordset.Update
MsgBox "DATA BERHASIL DIUBAH !", vbInformation, "INFORMASI !"
Call clear
Call KodeOtomatis
End If
End Sub

'HAPUS
Private Sub cmdhapus_Click()
If text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Or Combo3.Text = "" Or Combo4.Text = "" Then
MsgBox "PILIH DAHULU DATA YANG AKAN DIHAPUS !", vbInformation, "PERHATIAN !"
Else
xx = MsgBox("APAKAH ANDA YAKIN AKAN MENGHAPUS DATA INI ?", vbOKCancel, "Peringatan")
            If xx = vbOK Then
               Adodc1.Recordset.Delete
               Call clear
               Call KodeOtomatis
MsgBox "DATA ANDA BERHASIL DIHAPUS !", vbInformation, "INFORMASI !"
Adodc1.Refresh
Call table
            End If
End If
End Sub








