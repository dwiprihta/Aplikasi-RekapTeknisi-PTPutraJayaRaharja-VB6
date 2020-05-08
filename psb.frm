VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form psb 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DATA PASANG BARU"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   14805
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      DataSource      =   "Adodc2"
      Height          =   315
      Left            =   10320
      TabIndex        =   27
      Top             =   1680
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   13560
      TabIndex        =   24
      Top             =   8160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393216
      Format          =   84869121
      CurrentDate     =   43703
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   10320
      TabIndex        =   22
      Top             =   4080
      Width           =   3975
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   10320
      TabIndex        =   21
      Top             =   2880
      Width           =   3975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   14775
      Begin VB.Label Label5 
         BackColor       =   &H00808000&
         Caption         =   "DATA PASANG BARU"
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
         Left            =   480
         TabIndex        =   19
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      TabIndex        =   10
      Top             =   1680
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   2880
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      TabIndex        =   8
      Top             =   4080
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   5280
      TabIndex        =   7
      Top             =   1680
      Width           =   3975
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   5280
      TabIndex        =   6
      Top             =   2880
      Width           =   3975
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   5280
      TabIndex        =   5
      Top             =   4080
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   360
      TabIndex        =   1
      Top             =   4920
      Width           =   13935
      Begin VB.CommandButton cmdhapus 
         BackColor       =   &H00808000&
         Caption         =   "HAPUS"
         Height          =   495
         Left            =   3480
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdubah 
         BackColor       =   &H00808000&
         Caption         =   "UBAH"
         Height          =   495
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdtambah 
         BackColor       =   &H00808000&
         Caption         =   "TAMBAH"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdsimpan 
         BackColor       =   &H00808000&
         Caption         =   "SIMPAN"
         Height          =   495
         Left            =   120
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FFFF&
         Caption         =   "CARI"
         Height          =   495
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   9000
         TabIndex        =   2
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "CARI"
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "psb.frx":0000
      Height          =   1935
      Left            =   360
      TabIndex        =   4
      Top             =   6120
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   8160
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
      RecordSource    =   "psb"
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
      Top             =   8160
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
      Caption         =   "PAKET"
      Height          =   255
      Left            =   10320
      TabIndex        =   23
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TEKNOLOGI"
      Height          =   255
      Left            =   10320
      TabIndex        =   20
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NO SC"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NO INET"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NAMA PELANGGAN"
      Height          =   255
      Left            =   360
      TabIndex        =   15
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NO TELPON"
      Height          =   255
      Left            =   5280
      TabIndex        =   14
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ALAMAT PELANGGAN"
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NAMA TEKNISI"
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PANJANG KABEL"
      Height          =   255
      Left            =   10320
      TabIndex        =   11
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "psb"
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
    Combo1.AddItem !panjang_kabel
    .MoveNext
    Loop
End With
End If
Next
End Sub

'kode anggota otomatis 1
Sub KodeOtomatis()
Call Koneksi
RS.Open ("select * from psb Where no_sc In(Select Max(no_sc)From psb)Order By no_sc Desc"), conn
RS.Requery
    Dim Urutan As String * 6
    Dim Hitung As Long
    With RS
        If .EOF Then
            Urutan = "SC-" + "001"
            Text1 = Urutan
        Else
            Hitung = Right(!no_sc, 3) + 1
            Urutan = "SC-" + Right("000" & Hitung, 3)
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
cari_pelanggan2.Show
End Sub

'CARI
Private Sub Command4_Click()
Adodc1.Recordset.Filter = "no_sc like '%" + Me.Text7.Text + "%' or nama like '%" + Me.Text7.Text + "%' or no_inet like '%" + Me.Text7.Text + "%' or nama_teknisi like '%" + Me.Text7.Text + "%'"
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

cmdtambah.Visible = True
Call clear
Call table
Call KodeOtomatis
Call tambahcom
End Sub

'BERSIHKAN FORM
Sub clear()
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Combo1.Text = ""
Text9.Text = ""
Text10.Text = ""

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
Text9.Enabled = True
Text10.Enabled = True
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

.Columns(0).Caption = "NO SC"
.Columns(1).Caption = "NO INET"
.Columns(2).Caption = "NAMA PELANGGAN"
.Columns(3).Caption = "ALAMAT PELANGGAN"
.Columns(4).Caption = "NO TELPON"
.Columns(5).Caption = "NAMA TEKNISI"
.Columns(6).Caption = "PANJANG KABEL"
.Columns(7).Caption = "TEKNOLOGI"
.Columns(8).Caption = "PAKET"

End With
End Sub

'PINDAH DATA DARI TABEL KE FORM
Private Sub DataGrid1_Click()
cmdtambah.Visible = True
Text1.Text = Adodc1.Recordset!no_sc
Text2.Text = Adodc1.Recordset!no_inet
Text3.Text = Adodc1.Recordset!nama
Text4.Text = Adodc1.Recordset!alamat
Text5.Text = Adodc1.Recordset!no_telpon
Text6.Text = Adodc1.Recordset!nama_teknisi
Combo1.Text = Adodc1.Recordset!panjang_kabel
Text9.Text = Adodc1.Recordset!teknologi
Text10.Text = Adodc1.Recordset!bandwith
DTPicker1.Value = Adodc1.Recordset!tgl_pasang
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
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or Text9.Text = "" Or Text10.Text = "" Then
MsgBox "LENGKAPI DAHULU DATA YANG AKAN DIINPUTKAN !", vbInformation, "PERHATIAN !"
Else
'jika semua form sudah terisi, lakukan operasi simpan
Adodc1.Recordset.AddNew 'untuk tambah record'
Adodc1.Recordset.Fields("no_sc") = Text1.Text
Adodc1.Recordset.Fields("no_inet") = Text2.Text
Adodc1.Recordset.Fields("nama") = Text3.Text
Adodc1.Recordset.Fields("alamat") = Text4.Text
Adodc1.Recordset.Fields("no_telpon") = Text5.Text
Adodc1.Recordset.Fields("nama_teknisi") = Text6.Text
Adodc1.Recordset.Fields("panjang_kabel") = Combo1.Text
Adodc1.Recordset.Fields("teknologi") = Text9.Text
Adodc1.Recordset.Fields("bandwith") = Text10.Text
Adodc1.Recordset.Fields("tgl_pasang") = DTPicker1.Value
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
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or Text9.Text = "" Or Text10.Text = "" Then
MsgBox "PILIH DAHULU DATA YANG AKAN DIUBAH !", vbInformation, "PERHATIAN !"
Else
'jika semua form sudah terisi, lakukan operasi simpan

Adodc1.Recordset.Fields("no_sc") = Text1.Text
Adodc1.Recordset.Fields("no_inet") = Text2.Text
Adodc1.Recordset.Fields("nama") = Text3.Text
Adodc1.Recordset.Fields("alamat") = Text4.Text
Adodc1.Recordset.Fields("no_telpon") = Text5.Text
Adodc1.Recordset.Fields("nama_teknisi") = Text6.Text
Adodc1.Recordset.Fields("panjang_kabel") = Combo1.Text
Adodc1.Recordset.Fields("teknologi") = Text9.Text
Adodc1.Recordset.Fields("bandwith") = Text10.Text
Adodc1.Recordset.Fields("tgl_pasang") = DTPicker1.Value
Adodc1.Recordset.Update
MsgBox "DATA BERHASIL DIUBAH !", vbInformation, "INFORMASI !"
Call clear
Call KodeOtomatis
End If
End Sub

'HAPUS
Private Sub cmdhapus_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text5.Text = "" Or Text6.Text = "" Or Combo1.Text = "" Or Text9.Text = "" Or Text10.Text = "" Then
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



