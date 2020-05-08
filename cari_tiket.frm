VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form cari_tiket 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CARI TIKET"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   14805
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   3600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "cari_tiket.frx":0000
      Height          =   2055
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   3625
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
   Begin VB.TextBox Text7 
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   14055
   End
End
Attribute VB_Name = "cari_tiket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text7_Change()
If Text7.Text = "" Then
Adodc1.Refresh
Else
Adodc1.Recordset.Filter = "no_tiket like '%" + Me.Text7.Text + "%' or nama_pelanggan like '%" + Me.Text7.Text + "%' or no_layanan like '%" + Me.Text7.Text + "%' or nama_teknisi like '%" + Me.Text7.Text + "%' or layanan_terganggu like '%" + Me.Text7.Text + "%'"
End If
End Sub

Private Sub DataGrid1_Click()
'pindaahkan data dari datagrid 1 data anggota, kedalam form transaksi pinjam untuk digunakan mengisi form
tiket_close.text1.Text = Adodc1.Recordset!no_tiket
tiket_close.Text2.Text = Adodc1.Recordset!no_layanan
tiket_close.Text3.Text = Adodc1.Recordset!nama_teknisi
tiket_close.Text5.Text = Adodc1.Recordset!cp_pelanggan
tiket_close.Text6.Text = Adodc1.Recordset!layanan_terganggu
tiket_close.DTPicker1.Value = Adodc1.Recordset!tgl_open
tiket_close.Text10.Text = Adodc1.Recordset!teknologi
tiket_close.Text8.Text = Adodc1.Recordset!workzone
tiket_close.Text9.Text = Adodc1.Recordset!nama_pelanggan
tiket_close.Text4.SetFocus
'jika selesai tutup form ini
Unload Me
End Sub


Private Sub Form_Load()
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

