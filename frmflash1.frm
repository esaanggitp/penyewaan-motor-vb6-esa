VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form formsplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   2640
         Top             =   2280
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   495
         Left            =   2640
         TabIndex        =   7
         Top             =   1560
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label5 
         Caption         =   "Rental Motor Purwokerto"
         Height          =   495
         Left            =   2880
         TabIndex        =   12
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "Di"
         Height          =   375
         Left            =   3480
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Selamat Datang "
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   495
         Left            =   3240
         TabIndex        =   9
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   3240
         TabIndex        =   8
         Top             =   2400
         Width           =   735
      End
      Begin VB.Image imgLogo 
         Height          =   2625
         Left            =   120
         Stretch         =   -1  'True
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   4
         Top             =   3060
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   3270
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         Caption         =   "Warning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5970
         TabIndex        =   5
         Top             =   2700
         Width           =   885
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Platform"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5580
         TabIndex        =   6
         Top             =   2340
         Width           =   1275
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "LicenseTo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6855
      End
   End
End
Attribute VB_Name = "formsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub timer1_timer()
ProgressBar1.Value = ProgressBar1.Value + 2
Label1.Caption = ProgressBar1.Value & "%"
 If ProgressBar1.Value = ProgressBar1.Max Then
  Unload Me
  menu.Show
 End If
End Sub
