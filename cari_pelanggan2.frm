VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form cari_pelanggan2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "CARI DATA PELANGGAN"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   14865
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   14055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   3480
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "cari_pelanggan2.frx":0000
      Height          =   2055
      Left            =   360
      TabIndex        =   1
      Top             =   1320
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
End
Attribute VB_Name = "cari_pelanggan2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Text7_Change()
If Text7.Text = "" Then
Adodc1.Refresh
Else
Adodc1.Recordset.Filter = "id_pelanggan like '%" + Me.Text7.Text + "%' or nama_pelanggan like '%" + Me.Text7.Text + "%' or no_ktp like '%" + Me.Text7.Text + "%'"
End If
End Sub

Private Sub DataGrid1_Click()
'pindaahkan data dari datagrid 1 data anggota, kedalam form transaksi pinjam untuk digunakan mengisi form
psb.Text3.Text = Adodc1.Recordset.Fields("nama_pelanggan")
psb.Text4.Text = Adodc1.Recordset.Fields("alamat")
psb.Text5.Text = Adodc1.Recordset.Fields("cp")
psb.Text9.Text = Adodc1.Recordset.Fields("teknologi")
psb.Text10.Text = Adodc1.Recordset.Fields("bandwith")
'tiket.Text3.SetFocus

'jika selesai tutup form ini
Unload Me
End Sub


Private Sub Form_Load()
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

End Sub



