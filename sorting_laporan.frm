VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form sorting_laporan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SORTIR LAPORAN"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox ctahun 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Text            =   "TAHUN"
      Top             =   1800
      Width           =   6135
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808000&
      Caption         =   "TAMPILKAN LAPORAN KESELURUHAN"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   6135
   End
   Begin VB.ComboBox cbulan 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   "BULAN"
      Top             =   1080
      Width           =   6135
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   5880
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "SORTIR LAPORAN"
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2520
      Width           =   6135
   End
End
Attribute VB_Name = "sorting_laporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call Koneksi
RS.Open "select*from tiket_close where month(tgl_close)='" & Val(cbulan) & "' and year(tgl_close)='" & Val(ctahun) & "'", conn
If RS.EOF Then
MsgBox "DATA TIDAK DITEMUKAN !", vbInformation, "PERHATIAN !"

cbulan.SetFocus
Else
CR1.SelectionFormula = "Month({tiket_close.tgl_close}) = " & Val(cbulan) & " And Year({tiket_close.tgl_close}) = " & Val(ctahun) & ""
CR1.ReportFileName = App.Path & "\rekap.rpt"
CR1.WindowState = crptMaximized
CR1.RetrieveDataFiles
CR1.Action = 1
End If
End Sub
   

Private Sub Command2_Click()
xx = "\rekap.rpt"
cc = "*"
With CR1
    .ReportFileName = App.Path & xx
    .WindowState = crptMaximized
    .RetrieveDataFiles
    .Action = 1
End With
End Sub

Private Sub Form_Load()
ctahun.AddItem ("2015")
ctahun.AddItem ("2016")
ctahun.AddItem ("2017")
ctahun.AddItem ("2018")
ctahun.AddItem ("2019")
ctahun.AddItem ("2020")
ctahun.AddItem ("2021")
ctahun.AddItem ("2022")
ctahun.AddItem ("2023")
ctahun.AddItem ("2024")
ctahun.AddItem ("2025")
ctahun.AddItem ("2026")
ctahun.AddItem ("2027")
ctahun.AddItem ("2028")
ctahun.AddItem ("2029")
ctahun.AddItem ("2030")

cbulan.AddItem ("1")
cbulan.AddItem ("2")
cbulan.AddItem ("3")
cbulan.AddItem ("4")
cbulan.AddItem ("5")
cbulan.AddItem ("6")
cbulan.AddItem ("7")
cbulan.AddItem ("8")
cbulan.AddItem ("9")
cbulan.AddItem ("10")
cbulan.AddItem ("11")
cbulan.AddItem ("12")

End Sub
