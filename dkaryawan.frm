VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form fkaryawan 
   BackColor       =   &H8000000C&
   Caption         =   "Data Karyawan"
   ClientHeight    =   8985
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14835
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   14835
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "CARI"
      Height          =   615
      Left            =   6000
      TabIndex        =   14
      Top             =   4080
      Width           =   4815
      Begin VB.TextBox txtcari 
         Height          =   405
         Left            =   600
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.TextBox tpass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   8760
      PasswordChar    =   "*"
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   720
      Width           =   3015
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   3015
      Left            =   600
      TabIndex        =   11
      Top             =   5280
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5318
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.ComboBox cjk 
      Height          =   315
      Left            =   3000
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox talamat 
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox tnama 
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Text            =   "Text2"
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox tid 
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   720
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1560
      TabIndex        =   4
      Top             =   3960
      Width           =   3615
      Begin VB.CommandButton btutup 
         Caption         =   "TUTUP"
         Height          =   375
         Left            =   2040
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton binput 
         Caption         =   "INPUT"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   6480
      TabIndex        =   12
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ALAMAT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   720
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "JENIS KELAMIN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NAMA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ID KARYAWAN"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "fkaryawan"
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
rskaryawan.Open "select* from karyawan order by id_karyawan", KON
Set grid.DataSource = rskaryawan
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
Sub simpanuser()
simpan = "insert into karyawan values('" & tid.text & "','" & tpass.text & "','" & tnama.text & _
"','" & talamat.text & "','" & cjk.text & "')"
KON.Execute simpan
End Sub

Sub updateuser()
Update = "update karyawan set nama_karyawan='" & tnama.text & "',jenis_kelamin_karyawan='" & cjk.text & _
"',password='" & tpss.text & "' where id_karyawan='" & tid.text & "'"
KON.Execute Update
End Sub
Private Sub grid_KeyPress(KeyAscii As Integer)
a = grid.Row
kode = grid.TextMatrix(a, 1)
Call koneksi
rskaryawan.Open "select * from karyawan ", KON
With rskaryawan
If KeyAscii = 6 Then
If Not (.BOF And .EOF) Then
h = MsgBox("Bener mau dihapus ?", vbQuestion + vbYesNo, "——TaNYa——")
If h = vbYes Then
hapus = "delete from user where id_karyawan='" & kode & "'"
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

Private Sub tpss_keypress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
cjk.SetFocus
End If
Call tampilgrid
grid.Refresh
End Sub
Private Sub tid_KeyPress(KeyAscii As Integer)
Call koneksi
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
rskaryawan.Open "SELECT * FROM karyawan WHERE id_karyawan='" & tid & "'", KON
With rskaryawan
If .BOF And .EOF Then
psn = MsgBox("KD " + tid + " TDK ADA", vblnformation, "KONE")
tnama = ""
cjk = ""
tpss = ""
tnama.SetFocus
Else
tid.Enabled = False
tnama = .Fields("nama_karyawan")
cjk = .Fields("jenis_kelamin_karyawan")
tpss = .Fields("password")
binput.Caption = "UPDATE"
End If
End With
End If
End Sub
Sub sqluser()
SQL1 = "select * from karyawan where nama_karyawan like '%" & txtcari.text & "%' order by nama_karyawan asc"
KON.Execute SQL1
End Sub
Sub tampiluser()
Call koneksi
rskaryawan.Open "seLect* from karyawan where nama_karyawan like '%" & txtcari.text & "%'", KON
Set grid.DataSource = rskaryawan
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
