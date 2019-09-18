VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ftransaksi 
   Caption         =   "Transaksi"
   ClientHeight    =   9105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14385
   LinkTopic       =   "Form2"
   ScaleHeight     =   9105
   ScaleWidth      =   14385
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tnamapelanggan 
      Height          =   375
      Left            =   10680
      TabIndex        =   40
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox tpelanggan 
      Height          =   375
      Left            =   10680
      TabIndex        =   38
      Text            =   "Text1"
      Top             =   720
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   8400
      TabIndex        =   37
      Top             =   1440
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
      Height          =   975
      Left            =   240
      TabIndex        =   19
      Top             =   7800
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1720
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame2 
      Caption         =   "Option "
      Height          =   975
      Left            =   240
      TabIndex        =   15
      Top             =   6720
      Width           =   8775
      Begin VB.CommandButton binput 
         Caption         =   "INPUT"
         Height          =   375
         Left            =   360
         TabIndex        =   35
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton btutup 
         Caption         =   "TUTUP"
         Height          =   375
         Left            =   4080
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton bbatal 
         Caption         =   "BATAL"
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton bsimpan 
         Caption         =   "SIMPAN"
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame tt 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin VB.TextBox tuser 
         Height          =   405
         Left            =   7920
         TabIndex        =   36
         Text            =   "Text1"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox tbayar 
         Height          =   375
         Left            =   7800
         TabIndex        =   31
         Text            =   "Text1"
         Top             =   5640
         Width           =   1815
      End
      Begin VB.TextBox tkembali 
         Height          =   375
         Left            =   7800
         TabIndex        =   30
         Text            =   "Text1"
         Top             =   6120
         Width           =   1815
      End
      Begin Crystal.CrystalReport cr 
         Left            =   9960
         Top             =   5280
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Timer Timer1 
         Left            =   9960
         Top             =   6120
      End
      Begin VB.TextBox tsubtotal 
         Height          =   375
         Left            =   1800
         TabIndex        =   28
         Text            =   "Text4"
         Top             =   6000
         Width           =   2055
      End
      Begin VB.TextBox thari 
         Height          =   375
         Left            =   1800
         TabIndex        =   26
         Text            =   "Text3"
         Top             =   5400
         Width           =   1455
      End
      Begin VB.TextBox tharga 
         Height          =   405
         Left            =   1800
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   4800
         Width           =   1455
      End
      Begin VB.TextBox twarna 
         Height          =   375
         Left            =   1800
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton blistbarang 
         Caption         =   "LIST"
         Height          =   495
         Left            =   7320
         TabIndex        =   21
         Top             =   2400
         Width           =   1215
      End
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   3480
         TabIndex        =   20
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox ttahun 
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Text            =   "Text7"
         Top             =   3840
         Width           =   1455
      End
      Begin VB.TextBox tmerk 
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Text            =   "Text6"
         Top             =   3240
         Width           =   1455
      End
      Begin VB.TextBox tjenis 
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Text            =   "Text5"
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox tnomor 
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Text            =   "Text4"
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox tmotor 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Text            =   "Text3"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox ttgl 
         Height          =   375
         Left            =   5760
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox tnotrans 
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lbayar 
         Caption         =   "Label12"
         Height          =   615
         Left            =   8040
         TabIndex        =   34
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Uang Bayar"
         Height          =   375
         Left            =   6240
         TabIndex        =   33
         Top             =   5640
         Width           =   1215
      End
      Begin VB.Label ttttt 
         Caption         =   "Kembali"
         Height          =   375
         Left            =   6240
         TabIndex        =   32
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Sub Total"
         Height          =   375
         Left            =   360
         TabIndex        =   29
         Top             =   6000
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Hari"
         Height          =   375
         Left            =   360
         TabIndex        =   27
         Top             =   5400
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Harga/Hari"
         Height          =   375
         Left            =   240
         TabIndex        =   25
         Top             =   4920
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Warna"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   4440
         Width           =   855
      End
      Begin VB.Label text 
         Caption         =   "Thn Buat"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Merk"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Jenis"
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Nomor Polisi"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Label text555 
         Caption         =   "Id Motor"
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Tanggal sewa"
         Height          =   495
         Left            =   4200
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   " No. Transaksi "
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label Label3 
      Caption         =   "ID Pelanggan"
      Height          =   255
      Left            =   10800
      TabIndex        =   39
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "ftransaksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub cetak()
Call koneksi
cr.SelectionFormula = "{sewa.id_sewa}='" & tnotrans & "'"
cr.ReportFileName = App.Path & "\cetak2.rpt"
cr.WindowState = crptNormal
cr.RetrieveDataFiles
cr.Action = 1
'cr.DiscardSavedData = False
End Sub
Sub semula()
Call bersih
Call nonaktif
bsimpan.Enabled = False
binput.Enabled = True
btutup.Enabled = True
bbatal.Enabled = False
blistbarang.Enabled = False
List1.Visible = False
List2.Visible = False
End Sub
Sub nonaktif()
Dim kontrol As Control
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.Enabled = False
Next
List1.Enabled = False
List2.Enabled = False
End Sub
Sub aktif()
Dim kontrol As Control
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.Enabled = True
Next
List1.Enabled = True
List2.Enabled = True
ttgl.Enabled = False
End Sub

Sub bersih()
Dim kontrol As Control
lbayar.Caption = " Bayar "
Call hapusTEMP
For Each kontrol In Me.Controls
If TypeOf kontrol Is TextBox Then kontrol.text = ""
Next
End Sub

Sub isilist2()
rspelanggan.Open "select * from pelanggan", KON
List2.Clear
Do While Not rspelanggan.EOF
List2.AddItem rspelanggan!id_pelanggan & Space(5) & rspelanggan!nama_pelanggan
rspelanggan.MoveNext
Loop
End Sub
Private Sub List2_Click()
plg = "select * from pelanggan where id_pelanggan='" & Left(List2, 7) & "'"
Set rspelanggan = KON.Execute(plg)
tpelanggan = rspelanggan!id_pelanggan
tnamapelanggan = rspelanggan!nama_pelanggan
List1.Visible = True
List2.Visible = False
End Sub
Private Sub timer1_timer()
ttgl = Format(Date, "DD/MM/YYYY")
End Sub
Private Sub list1_Click()
mtr = "select * from motor where id_motor='" & Left(List1, 7) & "'"
Set rsmotor = KON.Execute(mtr)
tmotor = rsmotor!id_motor
tnomor = rsmotor!no_plat
tjenis = rsmotor!jenis
tmerk = rsmotor!merk
ttahun = rsmotor!thn_buat
twarna = rsmotor!warna
tharga = rsmotor!harga_motor
thari.SetFocus
List1.Visible = False
End Sub
Private Sub blistbarang_Click()

List2.Visible = True
End Sub
Private Sub bbatal_Click()
Call semula
Call hapusTEMP
End Sub
Private Sub binput_Click()
Call aktif
tuser.Enabled = False
tnotrans.Enabled = False
binput.Enabled = False
btutup.Enabled = False
bsimpan.Enabled = True
bbatal.Enabled = True
blistbarang.Enabled = True
Call hapusTEMP
Call bikinTEMP
Call nomor
tmotor.SetFocus
End Sub
Sub isilist()
rsmotor.Open "select * from motor", KON
List1.Clear
Do While Not rsmotor.EOF
List1.AddItem rsmotor!id_motor & Space(15) & rsmotor!no_plat & Space(20) & rsmotor!jenis & Space(10) & rsmotor!merk & Space(10) & rsmotor!thn_buat & Space(10) & rsmotor!warna & Space(10) & rsmotor!harga_motor
rsmotor.MoveNext
Loop
End Sub
Sub tampilgrid()
Call koneksi
rstemp.Open "select * from TEMP", KON
Set grid.DataSource = rstemp
grid.ColWidth(0) = 100
grid.ColWidth(1) = 1000
grid.ColWidth(2) = 1000
grid.ColWidth(3) = 1500
grid.ColWidth(4) = 1000
grid.ColWidth(5) = 1500
grid.ColWidth(6) = 1500
grid.ColWidth(7) = 1500
grid.ColWidth(8) = 1500
grid.ColWidth(9) = 1500
grid.ColWidth(10) = 1500
grid.TextMatrix(0, 1) = "ID pelanggan"
grid.TextMatrix(0, 2) = "ID motor"
grid.TextMatrix(0, 3) = "NO Polisi"
grid.TextMatrix(0, 4) = "Jenis"
grid.TextMatrix(0, 5) = "Merk"
grid.TextMatrix(0, 6) = "Warna"
grid.TextMatrix(0, 7) = "harga/hari"
grid.TextMatrix(0, 8) = "Hari"
grid.TextMatrix(0, 9) = "Sub total"
End Sub


Private Sub Form_Activate()
tnotrans.Enabled = False
Call semula
tuser = menu.stbar.Panels(2).text
End Sub
Private Sub form_load()
Call koneksi
Call isilist
Call isilist2
End Sub

Private Sub bsimpan_Click()
Call simpantransaksijual
Call simpandetailjual
X = MsgBox("cetak?", vbYesNo, "cetak")
If X = vbYes Then
Call cetak
Y = "delete from temp"
KON.Execute (Y)
Call tampilgrid
Call semula
Else
Call tampilgrid
Call semula
End If
End Sub
Private Sub btutup_Click()
Call hapusTEMP
Unload Me
menu.Show
End Sub
Sub nomor()
ttgl = Format(Date, "DD/MM/YYYY")
Dim cari As String
Call koneksi
rstrans.Open "SELECT * FROM sewa ORDER BY id_sewa DESC ", KON
With rstrans
If .EOF Then
tnotrans = Format(Date, "yymm") + "001"
ElseIf Left(rstrans!id_sewa, 4) <> Format(Date, "yymm") Then
tnotrans = Format(Date, "yymm") + "001"
Else
NO = Val(.Fields("id_sewa")) + 1
tnotrans = Format(Date, "yymm") + Right("000" + NO, 3)
End If
End With
End Sub
Private Sub thari_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call koneksi
rsmotor.Open "select * from motor where id_motor='" & tmotor & "'", KON
If Val(thari) > rsmotor!harga_motor Then
thari.SetFocus
Call koneksi
Exit Sub
Else
tsubtotal = Val(thari) * Val(tharga)
Call simpanTEMP
Call tampilgrid
Call isilist
ttl = U
For a = 1 To (grid.Rows - 1)
X = Val(grid.TextMatrix(a, 9))
ttl = ttl + X
Next a
lbayar.Caption = ttl
t = MsgBox("yakin??", vbQuestion + vbtesno, "konfirmasi")
If t = vbYes Then
tmotor = ""
tmotor.SetFocus
tnomor = ""
tjenis = ""
tmerk = ""
ttahun = ""
twarna = ""
tharha = ""
thari = ""
tsubtotal = ""
Else
Me.Refresh
grid.Refresh
tbayar.SetFocus
End If
End If
End If
End Sub
Private Sub tbayar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Val(lbayar) > Val(tbayar) Then
MsgBox "uang bayar kurang"
tbayar.SetFocus
tkembali.Enabled = False
Else
tkembali.Enabled = True
tkembali = Val(tbayar) - Val(lbayar)
bsimpan.SetFocus
End If
End If
End Sub

Sub bikinTEMP()
bikin = "create table TEMP(id_pelanggan int(5),id_motor int(5),no_plat varchar(10),jenis varchar(10),merk varchar(10),sat varchar(10),hrg double,qty int, subttl double)"
KON.Execute (bikin)
tampilgrid
End Sub

Sub hapusTEMP()
hapus = "drop table if exists TEMP"
KON.Execute (hapus)
End Sub

Sub simpanTEMP()
simpan = "insert into TEMP() vaLues('" & tpelanggan & "','" & tmotor & "','" & tnomor & "','" & tjenis & _
"','" & tmerk & "','" & twarna & "','" & tharga & "','" & thari & "','" & tsubtotal & "')"
KON.Execute (simpan)
End Sub

Sub simpantransaksijual()
ttgl = Format(Date, "YYYY/NM/DD")
simpan = "INSERT INTO sewa() VALUES('" & tnotrans & "','" & tmotor & _
"','" & ttgl & "','" & Val(lbayar) & "','" & tuser & "','" & tpelanggan & "')"
KON.Execute (simpan)
End Sub
Sub simpandetailjual()
Dim simpan, fak, tuser As String
Dim jumlah As Integer
Dim subtotal As Double
For a = 1 To (grid.Rows - 1)
fak = tnotrans
tuser = grid.TextMatrix(a, 2)
tjumlah = grid.TextMatrix(a, 8)
subtotal = grid.TextMatrix(a, 9)
simpan = "insert into detailsewa()vaLues('" & fak & "','" & Val(thari) & _
"','" & Val(subtotal) & "','" & tuser & "')"
Set rsdetail = KON.Execute(simpan)
Next a
End Sub
