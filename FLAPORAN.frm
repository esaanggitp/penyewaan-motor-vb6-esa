VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form flaporan 
   Caption         =   "LAPORAN"
   ClientHeight    =   5070
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport crmotor 
      Left            =   2640
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crpenyewaan 
      Left            =   3240
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Caption         =   "Bulanan"
      Height          =   1335
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   3735
      Begin VB.ComboBox ctahun 
         Height          =   315
         Left            =   2160
         TabIndex        =   14
         Text            =   "Pilih"
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cbulan 
         Height          =   315
         Left            =   2160
         TabIndex        =   13
         Text            =   "Pilih"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Tahun"
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   11
         Text            =   "Bulan"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mingguan"
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   3735
      Begin VB.ComboBox cmingguakhir 
         Height          =   315
         Left            =   1920
         TabIndex        =   9
         Text            =   "Pilih"
         Top             =   720
         Width           =   1575
      End
      Begin VB.ComboBox cmingguawal 
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Text            =   "Pilih"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Tanggal Akhir"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "Tanggal Awal"
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Harian"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3735
      Begin VB.ComboBox charian 
         Height          =   315
         Left            =   2040
         TabIndex        =   4
         Text            =   "Pilih"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   "Tanggal"
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton bkeluar 
      Caption         =   "keluar"
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton bcetak 
      Caption         =   "CEK DATA"
      Height          =   735
      Left            =   4440
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "flaporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bcetak_Click()
crmotor.ReportFileName = App.Path & "\laporan_motor.rpt"
crmotor.WindowState = crptMaximized
crmotor.RetrieveDataFiles
crmotor.Action = 1
End Sub
Private Sub bkeluar_Click()
Unload Me
menu.Show
End Sub

Private Sub cbulan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
Private Sub charian_Click()
crpenyewaan.SelectionFormula = "Totext ({sewa.tgl_pinjam})='" & charian & "'"
crpenyewaan.ReportFileName = App.Path & "\laporan_harian.rpt"
crpenyewaan.WindowState = crptMaximized
crpenyewaan.RetrieveDataFiles
crpenyewaan.Action = 1
End Sub
Private Sub charian_KeyPress(KeyAscii As Integer)
If charian = " " Or KeyAscii = 27 Then Unload Me
End Sub
Private Sub cmingguakhir_Click()
If cmingguawal = " " Then
MsgBox "fanggal awal kosong", , "Informasi"
cmingguawal.SetFocus
Exit Sub
End If
crpenyewaan.SelectionFormula = "{sewa.tgl_pinjam} in date (" & cmingguawal.text & _
") to date (" & cmingguakhir.text & ")"
crpenyewaan.ReportFileName = App.Path & "\laporan_mingguan.rpt"
crpenyewaan.WindowState = crptMaximized
crpenyewaan.RetrieveDataFiles
crpenyewaan.Action = 1
End Sub

Private Sub cmingguawal_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub ctahun_Click()
Call koneksi
rstrans.Open "select * from sewa where month(tgl_pinjam)=' " & Val(cbulan) & _
" ' and year(tgl_pinjam)=' " & (ctahun) & " ' ", KON
If rstrans.EOF Then
MsgBox "Data tidak ditemukan"
Exit Sub
cbulan.SetFocus
End If
crpenyewaan.SelectionFormula = "Month({sewa.tgl_pinjam})=" & Val(cbulan.text) & _
" and Year({sewa.tgl_pinjam})=" & Val(ctahun.text)
crpenyewaan.ReportFileName = App.Path & "\laporan_bulanan.rpt"
crpenyewaan.WindowState = crptMaximized
crpenyewaan.RetrieveDataFiles
crpenyewaan.Action = 1
End Sub

Private Sub form_load()
Call koneksi
rstrans.Open "Select Distinct tgl_pinjam From sewa order By 1", KON
rstrans.Requery
Do Until rstrans.EOF
charian.AddItem rstrans!tgl_pinjam
cmingguawal.AddItem Format(rstrans!tgl_pinjam, "YYYY ,MM, DD")
cmingguakhir.AddItem Format(rstrans!tgl_pinjam, "YYYY ,MM, DD")
rstrans.MoveNext
Loop
For i = 1 To 12
cbulan.AddItem i
Next i
For i = 10 To 20
ctahun.AddItem 2000 + i
Next i
End Sub
