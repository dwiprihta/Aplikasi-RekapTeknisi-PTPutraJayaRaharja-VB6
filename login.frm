VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN APLIKASI"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "BATAL"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "LOGIN"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4440
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   3480
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   3975
   End
   Begin VB.Image Image1 
      Height          =   1860
      Left            =   720
      Picture         =   "login.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3315
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PASSWORD"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "USERNAME"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'FORM LOGIN
'MENAMPILKAN FORM LOGIN
'by INDRI DWI S
'======================================================================

Private Sub Command1_Click()
'panggil modul koneksi
Call Koneksi
'cek jika form masih kosong
If Text1.Text = "" Then
MsgBox "FORM USERNAME ANDA MASIH KOSONG !", vbCritical, "Perhatian"
Text1.SetFocus
ElseIf Text2.Text = "" Then
MsgBox "FORM PASSWORD ANDA MASIH KOSONG !!!", vbCritical, "Perhatian"
Text2.SetFocus
Else

'cari data login di database admin
query = "select * from login where username='" & Text1.Text & "' and password='" & Text2.Text & "'"
RS.Open (query), conn
    If RS.EOF Then
    'tampilkan notif jika username atau password salah
    MsgBox "USERNAME ATAU PASSWORD ANDA SALAH !", vbExclamation, "Gagal !"
    'bersihkan inputan form
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus
    Else
    
    'jika berhasil login masuk ke menu admin
    MsgBox "ANDA BERHASIL LOGIN !", vbInformation, "LOGIN SUKSES !"
    index.Show
    'tutup form login
    Unload Me
    End If
End If
End Sub

Private Sub Command2_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub



