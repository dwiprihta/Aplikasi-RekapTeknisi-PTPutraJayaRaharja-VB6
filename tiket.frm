VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form tiket 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DATA TIKET OPEN"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   14760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "CARI"
      Height          =   615
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdhapus 
      BackColor       =   &H00808000&
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   3960
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdubah 
      BackColor       =   &H00808000&
      Caption         =   "UBAH"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   360
      TabIndex        =   22
      Top             =   5040
      Width           =   13935
      Begin VB.CommandButton cmdtambah 
         BackColor       =   &H00808000&
         Caption         =   "TAMBAH"
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   9000
         TabIndex        =   25
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton cmdsimpan 
         BackColor       =   &H00808000&
         Caption         =   "SIMPAN"
         Height          =   495
         Left            =   240
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FFFF&
         Caption         =   "CARI"
         Height          =   495
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "CARI"
      Height          =   495
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5160
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   615
      Left            =   10320
      TabIndex        =   20
      Top             =   1800
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
      _Version        =   393216
      Format          =   94633985
      CurrentDate     =   43658
   End
   Begin VB.ComboBox combo2 
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   315
      Left            =   10320
      TabIndex        =   19
      Top             =   4200
      Width           =   3975
   End
   Begin VB.ComboBox combo1 
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      Height          =   315
      Left            =   10320
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   3000
      Width           =   3975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "tiket.frx":0000
      Height          =   1935
      Left            =   360
      TabIndex        =   17
      Top             =   6240
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   3413
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
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   5280
      TabIndex        =   12
      Top             =   4200
      Width           =   3975
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   5280
      TabIndex        =   10
      Top             =   3000
      Width           =   3975
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   5280
      TabIndex        =   8
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   4200
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      TabIndex        =   4
      Top             =   3000
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      Begin VB.Label Label5 
         BackColor       =   &H00808000&
         Caption         =   "DATA TIKET OPEN"
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
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   5175
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   8400
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
      RecordSource    =   "tiket"
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
      Left            =   2520
      Top             =   8400
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
      Caption         =   "WORKZONE"
      Height          =   255
      Left            =   10320
      TabIndex        =   16
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TEKNOLOGI"
      Height          =   255
      Left            =   10320
      TabIndex        =   15
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TGL OPEN"
      Height          =   255
      Left            =   10320
      TabIndex        =   14
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NAMA TEKNISI"
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CP PELANGGAN"
      Height          =   255
      Left            =   5280
      TabIndex        =   11
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NAMA PELANGGAN"
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LAYANAN TERGANGGU"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NO LAYANAN"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NO TIKET"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
End
Attribute VB_Name = "tiket"
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
    Combo1.AddItem !teknologi
    Combo2.AddItem !workzone
    .MoveNext
    Loop
End With
End If
Next
End Sub

'kode anggota otomatis 1
Sub KodeOtomatis()
Call Koneksi
RS.Open ("select * from tiket Where no_tiket In(Select Max(no_tiket)From tiket)Order By no_tiket Desc"), conn
RS.Requery
    Dim Urutan As String * 6
    Dim Hitung As Long
    With RS
        If .EOF Then
            Urutan = "TK-" + "001"
            Text1 = Urutan
        Else
            Hitung = Right(!no_tiket, 3) + 1
            Urutan = "TK-" + Right("000" & Hitung, 3)
        End If
        Text1 = Urutan
    End With
End Sub

'kode anggota otomatis 2
Sub KodeOtomatis2()
Call Koneksi
RS.Open ("select * from tiket Where no_layanan In(Select Max(no_layanan)From tiket)Order By no_layanan Desc"), conn
RS.Requery
    Dim Urutan As String * 6
    Dim Hitung As Long
    With RS
        If .EOF Then
            Urutan = "LY-" + "001"
            Text2 = Urutan
        Else
            Hitung = Right(!no_tiket, 3) + 1
            Urutan = "LY-" + Right("000" & Hitung, 3)
        End If
        Text2 = Urutan
    End With
End Sub


Private Sub Command1_Click()
cari_pelanggan.Show
End Sub

'CARI
Private Sub Command4_Click()
Adodc1.Recordset.Filter = "no_tiket like '%" + Me.Text7.Text + "%' or nama_pelanggan like '%" + Me.Text7.Text + "%' or no_layanan like '%" + Me.Text7.Text + "%' or nama_teknisi like '%" + Me.Text7.Text + "%' or layanan_terganggu like '%" + Me.Text7.Text + "%'"
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
Call tambahcom
cmdtambah.Visible = True
Call clear
Call table
Call KodeOtomatis
End Sub

'BERSIHKAN FORM
Sub clear()
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo1.Text = ""
Combo2.Text = ""
'Text3.SetFocus
End Sub

'HIDUPKAN FORM
Sub enabel()
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
cmdtambah.Enabled = True
Text2.SetFocus
End Sub

'FORMAT TABEL
Sub table()
With DataGrid1
.Columns(0).Width = 1000
.Columns(1).Width = 1500
.Columns(2).Width = 2000
.Columns(3).Width = 2000
.Columns(4).Width = 1500
.Columns(5).Width = 2000
.Columns(6).Width = 1500
.Columns(7).Width = 1500
.Columns(8).Width = 1500

.Columns(0).Caption = "NO TIKET"
.Columns(1).Caption = "NO LAYANAN"
.Columns(2).Caption = "LAYANAN TERGANGGU"
.Columns(3).Caption = "NAMA PELANGGAN"
.Columns(4).Caption = "CP PELANGGAN"
.Columns(5).Caption = "NAMA TEKNISI"
.Columns(6).Caption = "TANGGAL OPEN"
.Columns(7).Caption = "TEKNOLOGI"
.Columns(8).Caption = "WORKZONE"

End With
End Sub

'PINDAH DATA DARI TABEL KE FORM
Private Sub DataGrid1_Click()
cmdtambah.Visible = True
Text1.Text = Adodc1.Recordset!no_tiket
Text2.Text = Adodc1.Recordset!no_layanan
Text3.Text = Adodc1.Recordset!layanan_terganggu
Text4.Text = Adodc1.Recordset!nama_pelanggan
Text5.Text = Adodc1.Recordset!cp_pelanggan
Text6.Text = Adodc1.Recordset!nama_teknisi
DTPicker1.Value = Adodc1.Recordset!tgl_open
Combo1.Text = Adodc1.Recordset!teknologi
Combo2.Text = Adodc1.Recordset!workzone
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
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or Combo1.Text = "" Then
MsgBox "LENGKAPI DAHULU DATA YANG AKAN DIINPUTKAN !", vbInformation, "PERHATIAN !"
Else
'jika semua form sudah terisi, lakukan operasi simpan
Adodc1.Recordset.AddNew 'untuk tambah record'
Adodc1.Recordset.Fields("no_tiket") = Text1.Text
Adodc1.Recordset.Fields("no_layanan") = Text2.Text
Adodc1.Recordset.Fields("layanan_terganggu") = Text3.Text
Adodc1.Recordset.Fields("nama_pelanggan") = Text4.Text
Adodc1.Recordset.Fields("cp_pelanggan") = Text5.Text
Adodc1.Recordset.Fields("nama_teknisi") = Text6.Text
Adodc1.Recordset.Fields("tgl_open") = DTPicker1.Value
Adodc1.Recordset.Fields("teknologi") = Combo1.Text
Adodc1.Recordset.Fields("workzone") = Combo2.Text
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
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or Combo1.Text = "" Then
MsgBox "PILIH DAHULU DATA YANG AKAN DIUBAH !", vbInformation, "PERHATIAN !"
Else
'jika semua form sudah terisi, lakukan operasi simpan
Adodc1.Recordset.Fields("no_tiket") = Text1.Text
Adodc1.Recordset.Fields("no_layanan") = Text2.Text
Adodc1.Recordset.Fields("layanan_terganggu") = Text3.Text
Adodc1.Recordset.Fields("nama_pelanggan") = Text4.Text
Adodc1.Recordset.Fields("cp_pelanggan") = Text5.Text
Adodc1.Recordset.Fields("nama_teknisi") = Text6.Text
Adodc1.Recordset.Fields("tgl_open") = DTPicker1.Value
Adodc1.Recordset.Fields("teknologi") = Combo1.Text
Adodc1.Recordset.Fields("workzone") = Combo2.Text
Adodc1.Recordset.Update
MsgBox "DATA BERHASIL DIUBAH !", vbInformation, "INFORMASI !"
Call clear
Call KodeOtomatis
End If
End Sub

'HAPUS
Private Sub cmdhapus_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or Combo1.Text = "" Then
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


