VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form menu 
   Caption         =   "Peminjaman Sepeda Motor"
   ClientHeight    =   5370
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8820
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   240
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4920
      Top             =   4800
   End
   Begin MSComctlLib.StatusBar stbar 
      Align           =   2  'Align Bottom
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   4635
      Width           =   8820
      _ExtentX        =   15558
      _ExtentY        =   1296
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "SELAMAT DATANG"
      BeginProperty Font 
         Name            =   "Goudy Stout"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   6375
   End
   Begin VB.Menu motor 
      Caption         =   "File"
      Begin VB.Menu datamotor 
         Caption         =   "Data Motor"
      End
      Begin VB.Menu datakaryawan 
         Caption         =   "Datar Karyawan"
      End
      Begin VB.Menu customer 
         Caption         =   "Data Customer"
      End
   End
   Begin VB.Menu transaksi 
      Caption         =   "Transaksi"
   End
   Begin VB.Menu laporan 
      Caption         =   "Laporan"
   End
   Begin VB.Menu logout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub customer_Click()
menu.Hide
fpeminjam.Show
End Sub

Private Sub datakaryawan_Click()
menu.Hide
fkaryawan.Show
End Sub

Private Sub datamotor_Click()
menu.Hide
fmotor.Show
End Sub

Private Sub laporan_Click()
menu.Hide
flaporan.Show
End Sub

Private Sub logout_Click()
Me.Visible = False
flogin.Show
flogin.tuser.text = ""
flogin.tpass.text = ""
End Sub
Private Sub timer1_timer()
menu.stbar.Panels(1) = Format(Date, "dd/nm/yyyy")
menu.stbar.Panels(3) = Time()
End Sub

Private Sub Timer2_Timer()
Label1.ForeColor = RGB(Rnd * 250, Rnd * 250, Rnd * 250)
End Sub

Private Sub transaksi_Click()
menu.Hide
ftransaksi.Show
End Sub
