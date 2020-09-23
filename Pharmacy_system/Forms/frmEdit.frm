VERSION 5.00
Begin VB.Form frmEdit 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editing Drug Details"
   ClientHeight    =   3480
   ClientLeft      =   1695
   ClientTop       =   3390
   ClientWidth     =   7470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   7470
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Drug Details"
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
      Height          =   3975
      Left            =   2640
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton cmddelete 
         Caption         =   "&Delete"
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
         Left            =   1800
         TabIndex        =   11
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox txtexpiry 
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1540
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
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   920
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
         Left            =   1680
         TabIndex        =   4
         Top             =   2160
         Width           =   1335
      End
      Begin VB.ComboBox cmbpid 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   2895
      End
      Begin VB.CommandButton cmdsave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
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
         Left            =   240
         MaskColor       =   &H00808080&
         TabIndex        =   2
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "E&xit"
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
         Left            =   3360
         TabIndex        =   1
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         X1              =   240
         X2              =   4560
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line1 
         BorderStyle     =   2  'Dash
         BorderWidth     =   3
         X1              =   240
         X2              =   4560
         Y1              =   2640
         Y2              =   2640
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
         Left            =   1080
         TabIndex        =   10
         Top             =   2280
         Width           =   615
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
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
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
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Drug Name"
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
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   3720
      Left            =   0
      Picture         =   "frmEdit.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4155
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbpid_Click()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

rs.Open "Select * from Master where DrugName = '" & Me.cmbpid & "'", con, adOpenKeyset, adLockOptimistic

If rs.EOF <> True And rs.BOF <> True Then
Me.txtexpiry = rs.Fields("ExpDate")
Me.txtpdate = rs.Fields("MfdDate")
Me.txtshelf = rs.Fields("Shelf")
End If
End Sub

Private Sub cmbpid_GotFocus()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

rs.Open "Select DrugName from Master", con, adOpenKeyset, adLockOptimistic
While rs.EOF <> True
cmbpid.AddItem rs!DrugName
rs.MoveNext
Wend

End Sub

Private Sub cmddelete_Click()
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

rs.Open "Delete * from Master where DrugName = '" & Me.cmbpid & "'", con, adOpenKeyset, adLockOptimistic

    Me.cmbpid.Clear
    Me.txtshelf = ""
    Me.txtpdate = ""
    Me.txtexpiry = ""
    MsgBox "Item Deleted", vbInformation, "Deletion"
Set rs = Nothing
Set con = Nothing

End Sub

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

rs.Open "Select * from Master where DrugName = '" & Me.cmbpid & "'", con, adOpenKeyset, adLockOptimistic
If rs.EOF <> True And rs.BOF <> True Then

With rs
    .Fields("DrugName") = Me.cmbpid
    .Fields("Shelf") = Me.txtshelf
    .Fields("MfdDate") = Me.txtpdate
    .Fields("ExpDate") = Me.txtexpiry
    .Update
    .close
End With
    Me.cmbpid.Clear
     Me.txtshelf = ""
    Me.txtpdate = ""
    Me.txtexpiry = ""
       
    MsgBox "item qty update"
Set rs = Nothing
Set con = Nothing

End If
End Sub

Private Sub Form_Load()

Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2.4
End Sub

