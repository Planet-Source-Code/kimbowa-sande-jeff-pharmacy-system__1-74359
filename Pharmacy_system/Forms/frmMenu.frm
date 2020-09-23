VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Main Menu"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   720
   ClientWidth     =   15135
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   15135
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command6 
      Appearance      =   0  'Flat
      Caption         =   "Exit System"
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   9480
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   4200
      TabIndex        =   5
      Top             =   8160
      Width           =   6255
      Begin VB.Label Label3 
         BackColor       =   &H00FF80FF&
         Caption         =   "username"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         TabIndex        =   7
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "respect"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   15195
      TabIndex        =   0
      Top             =   -120
      Width           =   15255
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   1080
         Top             =   360
      End
      Begin VB.Label l2 
         BackStyle       =   0  'Transparent
         Caption         =   "time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   615
         Left            =   9720
         TabIndex        =   4
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9000
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
      Begin VB.Label l3 
         BackStyle       =   0  'Transparent
         Caption         =   "date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   5040
         TabIndex        =   2
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.Image Image5 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   3120
         MousePointer    =   3  'I-Beam
         Picture         =   "frmMenu.frx":08CA
         Top             =   240
         Width           =   540
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   570
         Left            =   1440
         MousePointer    =   3  'I-Beam
         Picture         =   "frmMenu.frx":1594
         Stretch         =   -1  'True
         Top             =   240
         Width           =   705
      End
      Begin VB.Image Image7 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   2400
         MousePointer    =   3  'I-Beam
         Picture         =   "frmMenu.frx":9466
         Top             =   240
         Width           =   540
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   675
         Left            =   120
         MousePointer    =   3  'I-Beam
         Picture         =   "frmMenu.frx":9D30
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1035
      End
      Begin VB.Image Image6 
         Height          =   3075
         Left            =   120
         Picture         =   "frmMenu.frx":A6A9
         Stretch         =   -1  'True
         Top             =   -1680
         Width           =   15480
      End
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "NOTE: Please be sure when using MEDIZONE PHARMACY SYSTEM because all transactions you carry out are being tracked. Enjoy your stay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   4200
      TabIndex        =   9
      Top             =   7560
      Width           =   6135
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "USER LOGIN TRACK"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   7080
      Width           =   3015
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   4995
      Left            =   4320
      Picture         =   "frmMenu.frx":10FC5
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   6090
   End
   Begin VB.Image Image1 
      Height          =   11040
      Left            =   0
      Picture         =   "frmMenu.frx":1ED11
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15480
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnunewstaff 
         Caption         =   "New Staff"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnumsales 
      Caption         =   "S&ales"
      Begin VB.Menu mnumsale 
         Caption         =   "Make Sales"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuview 
         Caption         =   "View Sales"
         Begin VB.Menu mnuoption 
            Caption         =   "&Optional"
         End
         Begin VB.Menu mnuall 
            Caption         =   "&All"
         End
      End
      Begin VB.Menu mnubill 
         Caption         =   "Bill"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnustock 
      Caption         =   "&Stock"
      Begin VB.Menu mnunewstock 
         Caption         =   "Add New Stock"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnubalance 
         Caption         =   "Stock Balance"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuedit 
         Caption         =   "Edit Stock Master"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnucalc 
      Caption         =   "Calculator"
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuallsale_Click()
RptAllsales.Show vbModal
End Sub



Private Sub Command2_Click()
'Shell "calc.exe", vbNormalFocus
End Sub



Private Sub Command6_Click()
End
End Sub

Private Sub Form_Load()
l3 = Date
Label2.Caption = "Your Most Welcome"
Label3.Caption = frmLogin.txtUserName.Text

End Sub

Private Sub Image2_Click()
FrmAdminPass.Show vbModal
End Sub

Private Sub Image4_Click()
frmSale.Show vbModal
End Sub

Private Sub Image5_Click()
frmMaster.Show vbModal
End Sub

Private Sub Image7_Click()
frmFind.Show vbModal
End Sub

Private Sub mnuall_Click()
FrmAdminvall.Show vbModal
End Sub

Private Sub mnubalance_Click()

rptstockbal.Show vbModal
End Sub

Private Sub mnucalc_Click()
Shell "calc.exe", vbNormalFocus
End Sub

Private Sub mnuedit_Click()
FrmAdminedit.Show vbModal
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnumsale_Click()
frmSale.Show vbModal
End Sub

Private Sub mnunewstaff_Click()
FrmAdminPass.Show vbModal
End Sub

Private Sub mnunewstock_Click()
frmMaster.Show vbModal
End Sub

Private Sub mnuoption_Click()
FrmAdminvsales.Show vbModal
End Sub

Private Sub Timer1_Timer()
l2 = Time
End Sub
