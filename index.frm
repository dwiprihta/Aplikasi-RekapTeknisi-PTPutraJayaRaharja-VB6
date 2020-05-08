VERSION 5.00
Begin VB.Form index 
   BackColor       =   &H00FFFFFF&
   Caption         =   "APLIKASI REKAP DATA TEKNISI PT PUTRA JAYA RAHARJA"
   ClientHeight    =   8205
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   Picture         =   "index.frx":0000
   ScaleHeight     =   8205
   ScaleWidth      =   14550
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   20535
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "REKAP PASANG BARU"
         Height          =   255
         Left            =   6960
         TabIndex        =   10
         Top             =   1680
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PASANG BARU"
         Height          =   255
         Left            =   5640
         TabIndex        =   9
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8640
         TabIndex        =   8
         Top             =   120
         Width           =   11775
      End
      Begin VB.Image Image3 
         Height          =   975
         Left            =   5760
         Picture         =   "index.frx":573C1
         Stretch         =   -1  'True
         ToolTipText     =   "DATA TIKET CLOSE"
         Top             =   600
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   1095
         Left            =   7200
         Picture         =   "index.frx":58216
         Stretch         =   -1  'True
         ToolTipText     =   "DATA TIKET CLOSE"
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "PERBAIKAN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   5055
      End
      Begin VB.Line Line1 
         X1              =   5280
         X2              =   5280
         Y1              =   480
         Y2              =   1800
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "PASANG BARU"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5040
         TabIndex        =   6
         Top             =   120
         Width           =   3855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "REKAP PERBAIKAN"
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TIKET CLOSE"
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "TIKET OPEN"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label77 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "--:--:--"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   24
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   14880
         TabIndex        =   2
         Top             =   1080
         Width           =   5175
      End
      Begin VB.Label Label88 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FF80&
         BackStyle       =   0  'Transparent
         Caption         =   "--/--/----"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   24
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   15000
         TabIndex        =   1
         Top             =   480
         Width           =   5055
      End
      Begin VB.Image Image5 
         Height          =   975
         Left            =   600
         MousePointer    =   4  'Icon
         Picture         =   "index.frx":5B65B
         Stretch         =   -1  'True
         ToolTipText     =   "DATA TIKET OPEN"
         Top             =   600
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   975
         Left            =   2160
         Picture         =   "index.frx":5F00C
         Stretch         =   -1  'True
         ToolTipText     =   "DATA TIKET CLOSE"
         Top             =   600
         Width           =   975
      End
      Begin VB.Image Image6 
         Height          =   975
         Left            =   3720
         Picture         =   "index.frx":6263F
         Stretch         =   -1  'True
         ToolTipText     =   "REKAP LAPORAN TEKNISI"
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Menu MASTER 
      Caption         =   "FILE"
      Begin VB.Menu plg 
         Caption         =   "DATA PELANGGAN"
      End
      Begin VB.Menu DATA_TEKNISI 
         Caption         =   "DATA TEKNISI"
      End
      Begin VB.Menu DATA_ADMIN 
         Caption         =   "DATA ADMIN"
      End
   End
   Begin VB.Menu LOG_OUT 
      Caption         =   "LOG-OUT"
   End
End
Attribute VB_Name = "index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DATA_ADMIN_Click()
admin.Show
End Sub

Private Sub DATA_TEKNISI_Click()
teknisi.Show
End Sub

Private Sub Image1_Click()
sorting_laporan_psb.Show
End Sub

Private Sub Image2_Click()
tiket_close.Show
End Sub

Private Sub Image3_Click()
psb.Show
End Sub

Private Sub Image5_Click()
tiket.Show
End Sub

Private Sub Image6_Click()
sorting_laporan.Show
End Sub


Private Sub plg_Click()
pelanggan.Show
End Sub

'TAMPILKAN WAKTU
Private Sub Timer1_Timer()
Label77.Caption = Format(Now, "hh : mm : ss")
'Label88.Caption = Format(Now, "dd MMMM yyyy")
Label88.Caption = Format(Now, "dd MMMM yyyy")
   'Label2.Caption = Time
End Sub

Private Sub LOG_OUT_Click()
If MsgBox("APAKAH ANDA YAKIN AKAN KELUAR DARI APLIKASI INI ?", vbYesNo + vbDefaultButton2 + vbQuestion, "PERINGATAN!") = vbYes Then
End
End If
End Sub


Private Sub TIKET_MASUK_Click()

End Sub

Private Sub tiket_tutup_Click()

End Sub
