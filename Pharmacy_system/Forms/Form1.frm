VERSION 5.00
Begin VB.Form frmNewlog 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Staff Registration"
   ClientHeight    =   3735
   ClientLeft      =   1695
   ClientTop       =   2625
   ClientWidth     =   8925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8925
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      TabIndex        =   12
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   11
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "New Staff Registration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3255
      Left            =   3120
      TabIndex        =   5
      Top             =   0
      Width           =   5775
      Begin VB.TextBox txtusername 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Width           =   4215
      End
      Begin VB.TextBox txtans 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   2760
         Width           =   4215
      End
      Begin VB.TextBox txtquest 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   2160
         Width           =   4215
      End
      Begin VB.TextBox txtpass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   1560
         Width           =   4215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Answer"
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
         Left            =   480
         TabIndex        =   10
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Question"
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
         Left            =   360
         TabIndex        =   9
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         TabIndex        =   8
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name"
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
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Image Image2 
      Height          =   3735
      Left            =   -120
      Picture         =   "Form1.frx":08CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   4455
      Left            =   2520
      Picture         =   "Form1.frx":27E5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmNewlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
con.Open (Constring)

rs.Open "Select * from Login", con, adOpenKeyset, adLockOptimistic

If Me.txtname = "" Then
MsgBox "Enter name in the name field", vbOKOnly, "Empty field"
Exit Sub
End If

If txtUserName = "" Then
MsgBox "Enter username in the username field", vbOKOnly, "Empty field"
Exit Sub
End If

If txtpass = "" Then
MsgBox "Enter password in the password field", vbOKOnly, "Empty field"
Exit Sub
End If
If txtquest = "" Then
MsgBox "Enter Secret question in the question field", vbOKOnly, "Empty field"
Exit Sub
End If
If txtans = "" Then
MsgBox "Enter Secret answer in the answer field", vbOKOnly, "Empty field"
Exit Sub
End If


With rs
    .AddNew
    .Fields("Fullname") = Me.txtname
    .Fields("Username") = Me.txtUserName
    .Fields("Password") = Me.txtpass
    .Fields("Secquest") = Me.txtquest
    .Fields("Secanswer") = Me.txtans
    .Update
    .close
End With
Me.txtname = ""
Me.txtUserName = ""
Me.txtpass = ""
Me.txtquest = ""
Me.txtans = ""
Set rs = Nothing
Set con = Nothing
End Sub

Private Sub Form_Load()
Me.Height = 4995
Me.Width = 9165
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2.4

End Sub

