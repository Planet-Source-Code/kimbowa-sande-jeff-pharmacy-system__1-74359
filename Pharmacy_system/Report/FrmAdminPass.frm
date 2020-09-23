VERSION 5.00
Begin VB.Form FrmAdminPass 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6375
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   2040
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   1800
      TabIndex        =   2
      Top             =   240
      Width           =   4455
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
         Left            =   1320
         TabIndex        =   4
         Top             =   480
         Width           =   2925
      End
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
         TabIndex        =   3
         Top             =   1080
         Width           =   2925
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
         TabIndex        =   6
         Top             =   600
         Width           =   1080
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
         TabIndex        =   5
         Top             =   1200
         Width           =   960
      End
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   0
      Picture         =   "FrmAdminPass.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "FrmAdminPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
frmNewlog.Show vbModal
Else
MsgBox "Invalid username or password", vbInformation, "Error"
End If

End Sub

