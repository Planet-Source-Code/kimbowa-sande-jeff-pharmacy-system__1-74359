VERSION 5.00
Begin VB.Form FrmAdminedit 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrator Login"
   ClientHeight    =   2670
   ClientLeft      =   2370
   ClientTop       =   4065
   ClientWidth     =   6375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6375
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4560
      TabIndex        =   3
      Top             =   2040
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   2175
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
         Top             =   1440
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
         Top             =   840
         Width           =   2925
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " Administrator rights required  to edit stock."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   0
         Width           =   4215
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
         Top             =   1560
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
         Top             =   960
         Width           =   1080
      End
   End
   Begin VB.Image Image2 
      Height          =   2775
      Left            =   -120
      Picture         =   "FrmAdminedit.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   2895
      Left            =   0
      Picture         =   "FrmAdminedit.frx":0979
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   6615
   End
End
Attribute VB_Name = "FrmAdminedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
'to close this admin login form on clicking
Unload Me
End Sub

Private Sub cmdOK_Click()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)
Dim flag As Boolean

'to compare Administrators username and password for login from table[Admin]
rs.Open "Select Username, password from Admin", con, adOpenKeyset, adLockOptimistic

'EOF =End of file to check admin table upto the last record
While rs.EOF = False
If Me.txtUserName = rs!UserName And Me.txtPassword = rs!Password Then
flag = True
End If
rs.MoveNext
Wend
If flag = True Then
Unload Me
frmEdit.Show vbModal
Else
'message to give when wrong details are entered

MsgBox "Invalid! Check your username & password PLEASE", vbInformation, "Authentication Failure"
End If
End Sub

