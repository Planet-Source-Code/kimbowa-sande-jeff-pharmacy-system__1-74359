VERSION 5.00
Begin VB.Form frmFind 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   3990
   ClientLeft      =   3435
   ClientTop       =   2625
   ClientWidth     =   5010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   5010
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Search Results"
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton close 
         Caption         =   "Close"
         Height          =   615
         Left            =   3840
         TabIndex        =   11
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtexdate 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtbal 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox txtshelf 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox txtpdate 
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Production Date"
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
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Shelf"
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
         Left            =   1200
         TabIndex        =   5
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry Date"
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
         Left            =   720
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Balance"
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
         Left            =   360
         TabIndex        =   3
         Top             =   2280
         Width           =   1335
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmfind.frx":0000
      Left            =   1200
      List            =   "frmfind.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   3975
      Left            =   0
      Picture         =   "frmfind.frx":0004
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5130
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub close_Click()
Unload Me
End Sub

Private Sub Combo1_Click()
'make frame1 visible
Frame1.Visible = 1
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

'searches drug from table master depending selection in the combo
rs.Open "Select * from Master where DrugName = '" & Combo1 & "'", con, adOpenKeyset, adLockOptimistic
If rs.EOF = False And rs.BOF <> True Then
Me.txtpdate = rs!MfdDate
Me.txtbal = rs!Qty
Me.txtexdate = rs!ExpDate
Me.txtshelf = rs!Shelf
End If
If Val(txtbal.Text) = 0 Then
MsgBox Me.Combo1 & " is not available in stock", vbInformation, "Stock Query"
Else: MsgBox "There are " & Val(txtbal.Text) & " " & Me.Combo1 & "(s) in stock", vbInformation, "Stock Query"
End If
Set con = Nothing

Set rs = Nothing
End Sub

Private Sub Combo1_GotFocus()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

rs.Open "Select DrugName from Master order by DrugName", con, adOpenKeyset, adLockOptimistic

While rs.EOF <> True And rs.BOF <> True
Combo1.AddItem rs!DrugName
rs.MoveNext
Wend
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Unload Me
End Sub

Private Sub Form_Load()
Me.Combo1.Clear
End Sub
