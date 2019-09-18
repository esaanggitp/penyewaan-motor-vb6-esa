VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form fmotor 
   Caption         =   "DATA MOTOR"
   ClientHeight    =   8145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   12780
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cjenis 
      Height          =   315
      Left            =   2640
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Caption         =   "CARI"
      Height          =   735
      Left            =   5160
      TabIndex        =   17
      Top             =   4200
      Width           =   3855
      Begin VB.TextBox txtcari 
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   960
      TabIndex        =   14
      Top             =   4200
      Width           =   3495
      Begin VB.CommandButton btutup 
         Caption         =   "TUTUP"
         Height          =   375
         Left            =   1920
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton binput 
         Caption         =   "INPUT"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   1335
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   2655
      Left            =   840
      TabIndex        =   13
      Top             =   5040
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4683
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox tharga 
      Height          =   405
      Left            =   2640
      TabIndex        =   12
      Text            =   "Text9"
      Top             =   3480
      Width           =   2655
   End
   Begin VB.TextBox twarna 
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Text            =   "Text8"
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox tbuat 
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Text            =   "Text7"
      Top             =   2520
      Width           =   2655
   End
   Begin VB.TextBox merk 
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Text            =   "Text6"
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox noplat 
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Text            =   "Text4"
      Top             =   1080
      Width           =   2655
   End
   Begin VB.TextBox idmotor 
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Text            =   "Text3"
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label text17 
      Caption         =   "HARGA/HARI"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   720
      TabIndex        =   6
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label text16 
      Caption         =   "WARNA"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   720
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label text14 
      Caption         =   "Thn BUAT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   4
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label text13 
      Caption         =   "MERK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   3
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label text12 
      Caption         =   "JENIS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label text1 
      Caption         =   "NO PLAT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label text2 
      Caption         =   "ID MOTOR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1575
   End
End
Attribute VB_Name = "fmotor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub nonaktif()
Dim kontrol As Control
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.Enabled = False
Next
cjenis.Enabled = False
End Sub

Sub isijenis()
cjenis.AddItem "MTR_BEBEK"
cjenis.AddItem "MTR_METIC"
End Sub
Sub aktif()
Dim kontrol As Control
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.Enabled = True
Next
cjenis.Enabled = True
End Sub

Sub bersih()
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.text = ""
Next
cjenis = ""
End Sub

Sub tampilgrid()
Call koneksi
rsmotor.Open "seLect* from motor order by id_motor", KON
Set grid.DataSource = rsmotor
grid.ColWidth(0) = 0
grid.ColWidth(1) = 1000
grid.ColWidth(2) = 1000
grid.ColWidth(3) = 3000
grid.ColWidth(4) = 2000
grid.ColWidth(5) = 1900
grid.ColWidth(6) = 1900
End Sub

Private Sub binput_Click()
If binput.Caption = "INPUT" Then
binput.Caption = "SIMPAN"
btutup.Caption = "BATAL"
Call aktif
idmotor.SetFocus
ElseIf binput.Caption = "SIMPAN" Then
Call simpanmotor
Call tampilgrid
binput.Caption = "INPUT"
btutup.Caption = "TUTUP"
Call bersih
Call nonaktif
ElseIf binput.Caption = "UPDATE" Then
Call updatmotor
Call tampilgrid
Call bersih
Call nonaktif
binput.Caption = "INPUT"
btutup.Caption = "TUTUP"
End If
End Sub




Private Sub Form_Activate()
Call bersih
Call nonaktif
Call isijenis
Call tampilgrid
End Sub
Private Sub form_load()
Call koneksi
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
a = grid.Row
kode = grid.TextMatrix(a, 1)
Call koneksi
rsmotor.Open "select * from motor ", KON
With rsmotor
If KeyAscii = 6 Then
If Not (.BOF And .EOF) Then
h = MsgBox("Bener mau dihapus ?", vbQuestion + vbYesNo, "——TaNYa——")
If h = vbYes Then
hapus = "delete from motor where id_motor='" & idmotor & "'"
KON.Execute (hapus)
End If
End If
End If
End With
Call tampilgrid
grid.Refresh
End Sub

Private Sub tharga_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
binput.SetFocus
End If
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack) Then
Beep
KeyAscii = C
    End If
End Sub

Private Sub noplat_keypress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
cjenis.SetFocus
End If
End Sub


Private Sub idmotor_KeyPress(KeyAscii As Integer)
Call koneksi
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
rsbrg.Open "SELECT * FROM motor WHERE id_motor='" & idmotor & "'", KON
With rsmotor
If .BOF And .EOF Then
psn = MsgBox("KD " + idmotor + " TDK ADA", vblnformation, "KONE")
tnama = ""
noplat = ""
cjenis = ""
merk = ""
tbuat = ""
twarna = ""
tharga = ""
noplat.SetFocus
Else
idmotor.Enabled = False
noplat = .Fields("no_plat")
cjenis = .Fields("jenis")
merk = .Fields("merk")
tbuat = .Fields("thn_buat")
twarna = .Fields("warna")
tharga = .Fields("harga_motor")
binput.Caption = "UPDATE"
End If
End With
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



Sub sqlbrg()
SQL1 = "select * from motor where merk like '%" & txtcari.text & "%' order by merk asc"
KON.Execute SQL1
End Sub

Sub tampilbrg()
Call koneksi
rsmotor.Open "seLect* from motor where no_plat like '%" & txtcari.text & "%'", KON
Set grid.DataSource = rsmotor
grid.ColWidth(0) = 0
grid.ColWidth(1) = 1000
grid.ColWidth(2) = 1000
grid.ColWidth(3) = 3000
grid.ColWidth(4) = 2000
grid.ColWidth(5) = 1900
grid.ColWidth(6) = 1900
End Sub
Private Sub txtcari_Change()
Call koneksi
Call tampilbrg
Call sqlbrg
End Sub

Sub simpanmotor()
simpan = "insert into motor values('" & idmotor.text & "','" & noplat.text & "','" & cjenis.text & _
"','" & merk.text & "','" & tbuat.text & "','" & twarna.text & "','" & tharga.text & "')"
KON.Execute simpan
End Sub



Sub updatmotor()
Update = "update motor set no_plat='" & noplat.text & "',jenis='" & cjenis.text & _
"',merk='" & merk.text & "',thn_buat='" & tbuat.text & "',warna='" & twarna.text & "',harga_motor='" & tharga.text & "' where id_motor='" & idmotor.text & "'"
KON.Execute Update
End Sub
