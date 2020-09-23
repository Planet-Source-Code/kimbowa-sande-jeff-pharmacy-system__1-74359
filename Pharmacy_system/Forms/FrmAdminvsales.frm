VERSION 5.00
Begin VB.Form FrmAdminvsales 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrator Login"
   ClientHeight    =   2655
   ClientLeft      =   2370
   ClientTop       =   3975
   ClientWidth     =   6405
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6405
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   2280
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   2295
      Left            =   1800
      TabIndex        =   4
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1320
         Width           =   2925
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "$"
         TabIndex        =   0
         Top             =   720
         Width           =   2925
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   " Administrators username and password required to view sales details."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   0
         Width           =   3735
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "&User Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1080
      End
   End
   Begin VB.Image Image2 
      Height          =   2655
      Left            =   0
      Picture         =   "FrmAdminvsales.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   0
      Picture         =   "FrmAdminvsales.frx":691C
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   6495
   End
End
Attribute VB_Name = "FrmAdminvsales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)
Dim flag As Boolean

rs.Open "Select Username, password from Admin", con, adOpenKeyset, adLockOptimistic

While rs.EOF = False
If Me.txtUserName = rs!UserName And Me.txtPassword = rs!Password Then
flag = True
End If
rs.MoveNext
Wend
If flag = True Then
Unload Me
frmdate.Show vbModal
Else
MsgBox "Invalid username or password", vbInformation, "Error"
End If
End Sub

