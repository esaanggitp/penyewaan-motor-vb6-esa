VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form fpeminjam 
   Caption         =   "Form Data Peminjam"
   ClientHeight    =   7575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10275
   LinkTopic       =   "Form5"
   ScaleHeight     =   7575
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3015
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5318
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame2 
      Caption         =   "Option"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   8295
      Begin VB.TextBox txtcari 
         Height          =   375
         Left            =   5400
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton btutup 
         Caption         =   "TUTUP"
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton binput 
         Caption         =   "INPUT"
         Height          =   375
         Left            =   600
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Peminjam"
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8295
      Begin VB.ComboBox cjk 
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox ttelp 
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Text            =   "Text5"
         Top             =   2880
         Width           =   2175
      End
      Begin VB.TextBox talamat 
         Height          =   405
         Left            =   1680
         TabIndex        =   7
         Text            =   "Text4"
         Top             =   2280
         Width           =   3975
      End
      Begin VB.TextBox tnama 
         Height          =   405
         Left            =   1680
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox tid 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label3 
         Caption         =   "Jenis Kelamin"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "No. Telepon"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Alamat"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Nama"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "No. Identitas"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "fpeminjam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub nonaktif()
Dim kontrol As Control
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.Enabled = False
Next
cjk.Enabled = False
End Sub

Sub isicsat()
cjk.AddItem "LAKI-LAKI"
cjk.AddItem "PEREMPUAN"
End Sub
Sub aktif()
Dim kontrol As Control
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.Enabled = True
Next
cjk.Enabled = True
End Sub
Sub bersih()
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.text = ""
Next
cjk = ""
End Sub

Sub tampilgrid()
Call koneksi
rspelanggan.Open "seLect* from pelanggan order by id_pelanggan", KON
Set grid.DataSource = rspelanggan
grid.ColWidth(0) = 0
grid.ColWidth(1) = 1600
grid.ColWidth(2) = 3000
grid.ColWidth(3) = 3000
grid.ColWidth(4) = 3000
grid.ColWidth(5) = 1900
End Sub

Private Sub binput_Click()
If binput.Caption = "INPUT" Then
binput.Caption = "SIMPAN"
btutup.Caption = "BATAL"
Call aktif
tid.SetFocus
ElseIf binput.Caption = "SIMPAN" Then
Call simpanuser
Call tampilgrid
binput.Caption = "INPUT"
btutup.Caption = "TUTUP"
Call bersih
Call nonaktif
ElseIf binput.Caption = "UPDATE" Then
Call updateuser
Call tampilgrid
Call bersih
Call nonaktif
binput.Caption = "INPUT"
btutup.Caption = "TUTUP"
End If
End Sub

Private Sub btutup_Click()
If btutup.Caption = "TUTUP" Then
Unload Me
menu.Show
ElseIf btutup.Caption = "BATAL" Then
Call bersih
Call nonaktif
btutup.Caption = "TUTUP"
binput.Caption = "INPUT"
End If
End Sub
Private Sub Form_Activate()
Call bersih
Call nonaktif
Call isicsat
Call tampilgrid
End Sub
Private Sub form_load()
Call koneksi
End Sub

Private Sub tid_KeyPress(KeyAscii As Integer)
Call koneksi
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
rspelanggan.Open "SELECT * FROM pelanggan WHERE id_pelanggan='" & tid & "'", KON
With rspelanggan
If .BOF And .EOF Then
psn = MsgBox("KD " + tid + " TDK ADA", vblnformation, "KONE")
tnama = ""
cjk = ""
talamat = ""
tnama.SetFocus
Else
tid.Enabled = False
tnama = .Fields("nama_karyawan")
talamat = .Fields("alamat")
cjk = .Fields("jenis_kelamin_karyawan")
binput.Caption = "UPDATE"
End If
End With
End If
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
a = grid.Row
kode = grid.TextMatrix(a, 1)
Call koneksi
rspelanggan.Open "select * from pelanggan ", KON
With rspelanggan
If KeyAscii = 6 Then
If Not (.BOF And .EOF) Then
h = MsgBox("Bener mau dihapus ?", vbQuestion + vbYesNo, "——TaNYa——")
If h = vbYes Then
hapus = "delete from user where id_pelanggan='" & kode & "'"
KON.Execute (hapus)
End If
End If
End If
End With
End Sub

Private Sub tnama_keypress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
cjk.SetFocus
End If
Call tampilgrid
grid.Refresh
End Sub

Sub simpanuser()
simpan = "insert into pelanggan values('" & tid.text & "','" & tnama.text & "','" & talamat.text & _
"','" & cjk.text & "','" & ttelp.text & "')"
KON.Execute simpan
End Sub

Sub updateuser()
Update = "update pelanggan set nama_pelanggan='" & tnama.text & "',jenis_kelamin_pelanggan='" & cjk.text & _
"',alamat_pelanggan='" & talamat.text & "' where id_pelanggan='" & tid.text & "'"
KON.Execute Update
End Sub


Sub sqluser()
SQL1 = "select * from pelanggan where nama_pelanggan like '%" & txtcari.text & "%' order by nama_pelanggan asc"
KON.Execute SQL1
End Sub
Sub tampiluser()
Call koneksi
rspelanggan.Open "seLect* from pelanggan where nama_pelanggan like '%" & txtcari.text & "%'", KON
Set grid.DataSource = rspelanggan
grid.ColWidth(0) = 0
grid.ColWidth(1) = 1600
grid.ColWidth(2) = 3000
grid.ColWidth(3) = 3000
grid.ColWidth(4) = 3000
grid.ColWidth(5) = 1900
End Sub

Private Sub txtcari_Change()
Call koneksi
Call tampiluser
Call sqluser
End Sub

