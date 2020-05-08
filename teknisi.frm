VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form teknisi 
   BackColor       =   &H00FFFFFF&
   Caption         =   "DATA TEKNISI"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16140
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   16140
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6960
      TabIndex        =   23
      Top             =   1680
      Width           =   3975
   End
   Begin VB.ComboBox text7 
      Appearance      =   0  'Flat
      DataSource      =   "Adodc2"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6960
      TabIndex        =   21
      Top             =   2880
      Width           =   3975
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10920
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Textfoto 
      Height          =   375
      Left            =   10440
      TabIndex        =   20
      Top             =   7080
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   11400
      TabIndex        =   19
      Top             =   2880
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Foto"
      Enabled         =   0   'False
      Height          =   375
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton cmdtambah 
      BackColor       =   &H00808000&
      Caption         =   "TAMBAH"
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdubah 
      BackColor       =   &H00808000&
      Caption         =   "UBAH"
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   16335
      Begin VB.Label Label5 
         BackColor       =   &H00808000&
         Caption         =   "DATA TEKNISI"
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
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "CARI"
      Height          =   495
      Left            =   14040
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   600
      TabIndex        =   8
      Top             =   3840
      Width           =   14775
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   9840
         TabIndex        =   11
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton cmdhapus 
         BackColor       =   &H00808000&
         Caption         =   "HAPUS"
         Height          =   495
         Left            =   3360
         MaskColor       =   &H0000FFFF&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdsimpan 
         BackColor       =   &H00808000&
         Caption         =   "SIMPAN"
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   13560
      Top             =   7080
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
      RecordSource    =   "teknisi"
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
      Bindings        =   "teknisi.frx":0000
      Height          =   1815
      Left            =   600
      TabIndex        =   7
      Top             =   5040
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   3201
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
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   11400
      TabIndex        =   6
      Top             =   1680
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   2520
      TabIndex        =   2
      Top             =   2880
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   615
      Left            =   2520
      TabIndex        =   0
      Top             =   1680
      Width           =   3975
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   11400
      Top             =   7080
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
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "JENIS KELAMIN"
      Height          =   255
      Left            =   6960
      TabIndex        =   22
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Line Line3 
      X1              =   11160
      X2              =   11160
      Y1              =   1200
      Y2              =   3480
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ALAMAT"
      Height          =   255
      Left            =   11400
      TabIndex        =   18
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Line Line2 
      X1              =   2280
      X2              =   2280
      Y1              =   1200
      Y2              =   3480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1575
      Left            =   720
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   6720
      X2              =   6720
      Y1              =   1200
      Y2              =   3480
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CONTACT PERSON"
      Height          =   255
      Left            =   11400
      TabIndex        =   5
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "WORKZONE"
      Height          =   255
      Left            =   6960
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NAMA TEKNISI"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NIK"
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "teknisi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MENAMPILKAN DATA PADA DATABASE KE COMBO
Sub tambahcom()
Adodc2.ConnectionString = conn.ConnectionString
Adodc2.RecordSource = "select* from combo"

For Each gosong In Me.Controls
If TypeOf gosong Is ComboBox Then
gosong.Text = ""
With Adodc2.Recordset
    Do While Not .EOF
    On Error Resume Next
    Text7.AddItem !workzone
    .MoveNext
    Loop
End With
End If
Next
End Sub

'BUKA FOTO
'jika tombol pilih foto diklik
Private Sub Command1_Click()
CommonDialog1.ShowOpen
'munculkan dialog pilih foto
Textfoto = CommonDialog1.FileName
End Sub
Private Sub Textfoto_Change()
Image1.Picture = LoadPicture(Textfoto)
End Sub

'CARI
Private Sub Command3_Click()
Adodc1.Recordset.Filter = "nama like '%" + Me.Text5.Text + "%' or nik like '%" + Me.Text5.Text + "%'"
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
Combo1.AddItem "Laki-laki"
Combo1.AddItem "Perempuan"
Call tambahcom
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
Text7.Text = ""
Combo1.Text = ""
Textfoto.Text = ""
'Text1.SetFocus
End Sub

'HIDUPKAN FORM
Sub enabel()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text7.Enabled = True
Combo1.Enabled = True
Command1.Enabled = True
Text1.SetFocus
End Sub

'FORMAT TABEL
Sub table()
With DataGrid1
.Columns(0).Width = 2200
.Columns(1).Width = 3500
.Columns(2).Width = 2000
.Columns(3).Width = 2000

.Columns(0).Caption = "NIK "
.Columns(1).Caption = "NAMA TEKNISI"
.Columns(2).Caption = "JENIS KELAMIN"
.Columns(3).Caption = "WORKZONE"
.Columns(4).Caption = "ALAMAT"
.Columns(5).Caption = "KONTAK"
.Columns(6).Caption = "FOTO"

End With
End Sub

'PINDAH DATA DARI TABEL KE FORM
Private Sub DataGrid1_Click()
cmdtambah.Visible = True
Text1.Text = Adodc1.Recordset!nik
Text2.Text = Adodc1.Recordset!nama
Combo1.Text = Adodc1.Recordset!jenis_kelamin
Text7.Text = Adodc1.Recordset!workzone
Text3.Text = Adodc1.Recordset!alamat
Text4.Text = Adodc1.Recordset!cp
Textfoto.Text = Adodc1.Recordset!foto
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
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text7.Text = "" Or Combo1.Text = "" Then
MsgBox "LENGKAPI DAHULU DATA YANG AKAN DIINPUTKAN !", vbInformation, "PERHATIAN !"
Else
'jika semua form sudah terisi, lakukan operasi simpan
Adodc1.Recordset.AddNew 'untuk tambah record'
Adodc1.Recordset.Fields("nik") = Text1.Text
Adodc1.Recordset.Fields("nama") = Text2.Text
Adodc1.Recordset.Fields("jenis_kelamin") = Combo1.Text
Adodc1.Recordset.Fields("workzone") = Text7.Text
Adodc1.Recordset.Fields("alamat") = Text3.Text
Adodc1.Recordset.Fields("cp") = Text4.Text
Adodc1.Recordset.Fields("foto") = Textfoto.Text
Adodc1.Recordset.Update
MsgBox "DATA BERHASIL DITAMBAHKAN !", vbInformation, "INFORMASI !"
Call clear
End If
End Sub

'UBAH
Private Sub cmdubah_Click()
'jika ada inputan yang kosong, tampilkan pesan peringatan
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text7.Text = "" Or Combo1.Text = "" Then
MsgBox "PILIH DAHULU DATA YANG AKAN DIHAPUS !", vbInformation, "PERHATIAN !"
Else
'jika semua form sudah terisi, lakukan operasi simpan
'Adodc1.Recordset.Fields("nik") = Text1.Text
Adodc1.Recordset.Fields("nama") = Text2.Text
Adodc1.Recordset.Fields("jenis_kelamin") = Combo1.Text
Adodc1.Recordset.Fields("workzone") = Text7.Text
Adodc1.Recordset.Fields("alamat") = Text3.Text
Adodc1.Recordset.Fields("cp") = Text4.Text
Adodc1.Recordset.Fields("foto") = Textfoto.Text
Adodc1.Recordset.Update
MsgBox "DATA BERHASIL DIUBAH !", vbInformation, "INFORMASI !"
Call clear
End If
End Sub

'HAPUS
Private Sub cmdhapus_Click()
If Text1 = "" Or Text2 = "" Or Text3 = "" Or Text4 = "" Or Text7.Text = "" Or Combo1.Text = "" Then
MsgBox "PILIH DAHULU DATA YANG AKAN DIHAPUS !", vbInformation, "PERHATIAN !"
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






