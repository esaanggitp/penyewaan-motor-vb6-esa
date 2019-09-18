VERSION 5.00
Begin VB.Form flogin 
   BackColor       =   &H8000000C&
   Caption         =   "LOGIN KARYAWAN"
   ClientHeight    =   5025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ctutup 
      Caption         =   "TUTUP"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton clogin 
      Caption         =   "LOGIN"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox tpass 
      Height          =   495
      Left            =   3960
      TabIndex        =   1
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox tuser 
      Height          =   495
      Left            =   3960
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
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
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   5
      Top             =   2040
      Width           =   2055
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
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
End
Attribute VB_Name = "flogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b As Byte
Private Sub ctutup_Click()
End
End Sub
Private Sub Form_Activate()
tuser.Enabled = True
tpass.Enabled = False
clogin.Enabled = False
tuser.SetFocus
tuser.MaxLength = 10
tpass.PasswordChar = "*"
End Sub

Private Sub clogin_Click()
Call koneksi
rskaryawan.Open " select* from karyawan where id_karyawan=' " & tuser & _
 "' and password='" & tpass & "'", KON
If rskaryawan.EOF Then
b = b + 1
If 1 - b = 0 Then       'chr(3) = kalimat setelahnya ctomatis he bawah
MsgBox "Kesempatan ke" & b & " SaLah" & Chr(10) & _
"password '" & tpass.text & "' tidak dikenal"
tpass = "" 'KUCI harus RAPAT
tpass.SetFocus
ElseIf 2 - b = 0 Then
MsgBox "Kesempatan ke " & b & " Salah" & Chr(l0) & _
"password '" & tpass & "' tidak dikenal"
tpass = ""
tpass.SetFocus
ElseIf 3 - b = 0 Then
MsgBox "Kesempatan ke " & b & " Salah" & Chr(l0) & _
"password '" & tpass & "' tidak dikenal" & Chr(10) & _
"Kesempatan habis, ULangi dari awal"
Unload Me
End If
Else
formsplash.Show
Me.Visible = False
menu.stbar.Panels(2) = tuser
End If
End Sub


Private Sub tpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
clogin.Enabled = True
clogin.SetFocus
End If
End Sub



Private Sub tuser_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
Call koneksi
rskaryawan.Open "select id_karyawan from karyawan where id_karyawan='" & tuser & "'", KON
If rskaryawan.EOF Then
    a = a + 1
    If 1 - a = 0 Then

    MsgBox "Kesempatan Ke" & a & "salah" & Int(10) & "Nama'" & tuser & _
    "' tidak dikenal"
    tuser = ""
    tuser.SetFocus
ElseIf 2 - a = 0 Then
    MsgBox "Kesempatan ke " & a & " salah" & Int(10) & "Nama'" & tuser & _
    "' tidak dikenal"
    tuser = ""
    tuser.SetFocus
ElseIf 2 - a = 0 Then
    MsgBox "Kesempatan ke " & a & " salah" & Int(10) & "Nama'" & tuser & _
    "'tidak dikenal" & Int(10) & "kesempatan habis,ulangi dari awal"
    Unload Me
End If
Else
tuser.Enabled = False
tpass.Enabled = True
tpass.SetFocus
End If
End If
End Sub
