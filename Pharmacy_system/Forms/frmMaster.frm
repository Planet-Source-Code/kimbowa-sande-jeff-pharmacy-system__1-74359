VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMaster 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Products"
   ClientHeight    =   4425
   ClientLeft      =   1305
   ClientTop       =   1965
   ClientWidth     =   9000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   9000
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   7800
      TabIndex        =   10
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   500
      Left            =   5400
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      Caption         =   "Drug Description"
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
      Height          =   3855
      Left            =   3000
      TabIndex        =   3
      Top             =   0
      Width           =   6015
      Begin MSComCtl2.DTPicker DTPed 
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   1800
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20840448
         CurrentDate     =   39485
      End
      Begin MSComCtl2.DTPicker DTPpd 
         Height          =   375
         Left            =   1680
         TabIndex        =   11
         Top             =   1200
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   20840448
         CurrentDate     =   39485
      End
      Begin VB.TextBox txtdname 
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
         TabIndex        =   0
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox txtqty 
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
         TabIndex        =   1
         Top             =   2400
         Width           =   4215
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
         TabIndex        =   2
         Top             =   3120
         Width           =   4215
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
         Left            =   960
         TabIndex        =   8
         Top             =   3240
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         TabIndex        =   7
         Top             =   2520
         Width           =   855
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
         Left            =   480
         TabIndex        =   6
         Top             =   1920
         Width           =   1095
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
         TabIndex        =   5
         Top             =   1320
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
         Left            =   480
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Image Image2 
      Height          =   12720
      Left            =   0
      Picture         =   "frmMaster.frx":0000
      Stretch         =   -1  'True
      Top             =   -1080
      Width           =   8490
   End
   Begin VB.Image Image1 
      Height          =   4455
      Left            =   0
      Picture         =   "frmMaster.frx":8B9A2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "frmMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdsave_Click()
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

rs.Open "Select * from Master where DrugName= '" & txtdname & "'", con, adOpenKeyset, adLockOptimistic

If rs.EOF <> True And rs.BOF <> True Then
With rs
    .Fields("Qty") = rs.Fields("qty") + Val(Me.txtqty.Text)
    .Update
    .close
End With
Me.txtdname = ""
Me.txtqty = ""
Me.txtshelf = ""
Set rs = Nothing
Set con = Nothing

Else

rs1.Open "Select * from Master", con, adOpenKeyset, adLockOptimistic
With rs1
    .AddNew
    .Fields("DrugName") = Me.txtdname
    .Fields("MfdDate") = Me.DTPpd
    .Fields("ExpDate") = Me.DTPed
    .Fields("Shelf") = Me.txtshelf
    .Fields("Qty") = Me.txtqty
    .Update
    .close

Me.txtdname = ""
Me.txtqty = ""
Me.txtshelf = ""
End With
Set rs1 = Nothing
Set con = Nothing
End If
End Sub

Private Sub Form_Load()
Me.Height = 4999
Me.Width = 9165
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2.4
End Sub

