VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login to Acess the System"
   ClientHeight    =   2655
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1568.662
   ScaleMode       =   0  'User
   ScaleWidth      =   5943.526
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4560
      TabIndex        =   4
      Top             =   2040
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "LOGIN"
      Default         =   -1  'True
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   2040
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   1680
      TabIndex        =   0
      Top             =   240
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
         TabIndex        =   2
         Top             =   1080
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
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Width           =   2925
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   1935
         Left            =   0
         Picture         =   "frmLogin.frx":08CA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   4215
      End
   End
   Begin VB.Image Image3 
      Height          =   2655
      Left            =   0
      Picture         =   "frmLogin.frx":0C7A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   2895
      Left            =   -120
      Picture         =   "frmLogin.frx":13C7
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   6735
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdOK_Click()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)
Dim flag As Boolean

rs.Open "Select Username, password from login", con, adOpenKeyset, adLockOptimistic

While rs.EOF <> True
If Me.txtUserName = rs!UserName And Me.txtPassword = rs!Password Then
flag = True
End If
rs.MoveNext
Wend
If flag = True Then
Me.Hide
frmMenu.Show

Else
MsgBox "Invalid username or password", vbInformation, "Error"
End If

End Sub

Private Sub Form_Load()
Main
End Sub


