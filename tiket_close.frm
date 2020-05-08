VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form tiket_close 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DATA TIKET CLOSE"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14490
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   14490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "CARI"
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   495
      Left            =   13320
      TabIndex        =   23
      Text            =   "cp pelanggan"
      Top             =   7560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   495
      Left            =   12960
      TabIndex        =   22
      Text            =   "nama prlanggan"
      Top             =   7560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   495
      Left            =   14025
      TabIndex        =   21
      Text            =   "teknologi"
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   495
      Left            =   13680
      TabIndex        =   20
      Text            =   "workzone"
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   495
      Left            =   12720
      TabIndex        =   19
      Text            =   "lyanan terganggu"
      Top             =   7560
      Visible         =   0   'False
      Width           =   150
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   7560
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      RecordSource    =   "tiket_close"
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
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   10200
      TabIndex        =   17
      Top             =   3360
      Width           =   3975
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   5280
      TabIndex        =   13
      Top             =   2160
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      TabIndex        =   11
      Top             =   3360
      Width           =   3975
   End
   Begin VB.TextBox text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   360
      TabIndex        =   9
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   -120
      TabIndex        =   7
      Top             =   0
      Width           =   15375
      Begin VB.Label Label5 
         BackColor       =   &H00808000&
         Caption         =   "DATA TIKET CLOSE"
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
         TabIndex        =   8
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   4320
      Width           =   13935
      Begin VB.CommandButton Command4 
         BackColor       =   &H0000FFFF&
         Caption         =   "CARI"
         Height          =   495
         Left            =   12600
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdhapus 
         BackColor       =   &H00808000&
         Caption         =   "HAPUS"
         Height          =   495
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdsimpan 
         BackColor       =   &H00808000&
         Caption         =   "CLOSE TIKET"
         Height          =   495
         Left            =   360
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1575
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "tiket_close.frx":0000
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   5520
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   3413
      _Version        =   393216
      Enabled         =   -1  'True
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   615
      Left            =   5280
      TabIndex        =   6
      Top             =   3360
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
      _Version        =   393216
      Format          =   56819713
      CurrentDate     =   43658
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   615
      Left            =   10200
      TabIndex        =   25
      Top             =   2160
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1085
      _Version        =   393216
      Format          =   56819713
      CurrentDate     =   43658
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TGL OPEN"
      Height          =   255
      Left            =   5280
      TabIndex        =   18
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NO LAYANAN"
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TGL CLOSE"
      Height          =   255
      Left            =   10200
      TabIndex        =   15
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "KETERANGAN PROSES"
      Height          =   255
      Left            =   10200
      TabIndex        =   14
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NAMA TEKNISI"
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NO TIKET"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "tiket_close"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'MUNCULKAN PENCARIAN TIKET
Private Sub Command3_Click()
cari_tiket.Show
End Sub

'CARI
Private Sub Command4_Click()
Adodc1.Recordset.Filter = "no_tiket like '%" + Me.Text7.Text + "%' or nama_pelanggan like '%" + Me.Text7.Text + "%' or no_layanan like '%" + Me.Text7.Text + "%' or nama_teknisi like '%" + Me.Text7.Text + "%' or layanan_terganggu like '%" + Me.Text7.Text + "%'"
End Sub

'MUNCULKAN DATA SAAT PENCARIAN BERAKHIR
Private Sub Text7_Change()
If Text7.Text = "" Then
'Adodc1.Refresh
Call table
Else
'wkwk
End If
End Sub

'OPERASI SAAT FORM DIBUKA
Private Sub Form_Load()
Call clear
End Sub

'BERSIHKAN FORM
Sub clear()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
DTPicker1.Value = Now
DTPicker2.Value = Now
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
Text1.Text = Adodc1.Recordset.Fields("no_tiket")
Text2.Text = Adodc1.Recordset.Fields("no_layanan")
Text3.Text = Adodc1.Recordset.Fields("nama_teknisi")
DTPicker1.Value = Adodc1.Recordset.Fields("tgl_open")
DTPicker2.Value = Adodc1.Recordset.Fields("tgl_close")
Text4.Text = Adodc1.Recordset.Fields("keterangan_proses")
End Sub

'SIMPAN
Private Sub cmdsimpan_Click()
'jika ada inputan yang kosong, tampilkan pesan peringatan
If Text4 = "" Then
MsgBox "LENGAPI KETERANGAN PROSES DARI TEKNISI !", vbInformation, "PERHATIAN !"
Else
'jika semua form sudah terisi, lakukan operasi simpan
Adodc1.Recordset.AddNew 'untuk tambah record'
Adodc1.Recordset.Fields("no_tiket") = Text1.Text
Adodc1.Recordset.Fields("no_layanan") = Text2.Text
Adodc1.Recordset.Fields("nama_teknisi") = Text3.Text
Adodc1.Recordset.Fields("cp_pelanggan") = Text5.Text
Adodc1.Recordset.Fields("layanan_terganggu") = Text6.Text
Adodc1.Recordset.Fields("tgl_open") = DTPicker1.Value
Adodc1.Recordset.Fields("workzone") = Text8.Text
Adodc1.Recordset.Fields("nama_pelanggan") = Text9.Text
Adodc1.Recordset.Fields("teknologi") = Text10.Text
Adodc1.Recordset.Fields("tgl_close") = DTPicker2.Value
Adodc1.Recordset.Fields("keterangan_proses") = Text4.Text
Adodc1.Recordset.Update
MsgBox "TIKET BERHASIL DICLOSE!", vbInformation, "INFORMASI !"
'cari_tiket.Adodc1.Recordset.Delete
Call clear
End If

End Sub

'HAPUS
Private Sub cmdhapus_Click()
If Text4.Text = "" Then
MsgBox "PILIH DAHULU DATA YANG AKAN DIHAPUS !", vbInformation, "PERHATIAN !"
Else
xx = MsgBox("APAKAH ANDA YAKIN AKAN MENGHAPUS DATA INI ?", vbOKCancel, "Peringatan")
            If xx = vbOK Then
               Adodc1.Recordset.Delete
               Call clear
MsgBox "DATA ANDA BERHASIL DIHAPUS !", vbInformation, "INFORMASI !"
Adodc1.Refresh
Call table
            End If
End If
End Sub








