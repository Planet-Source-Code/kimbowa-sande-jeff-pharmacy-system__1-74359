VERSION 5.00
Begin VB.Form frmBill 
   BackColor       =   &H00800000&
   Caption         =   "Bill Entry"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8850
   Begin VB.CommandButton cmdexit 
      Appearance      =   0  'Flat
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
      Left            =   7320
      TabIndex        =   21
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print Bill"
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
      Left            =   5880
      TabIndex        =   16
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
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
      TabIndex        =   15
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bill"
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
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txtuprice 
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
         Left            =   3480
         TabIndex        =   22
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtAdd 
         Height          =   735
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   19
         Top             =   2520
         Width           =   4215
      End
      Begin VB.TextBox txtcustname 
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
         TabIndex        =   17
         Top             =   2040
         Width           =   4215
      End
      Begin VB.TextBox txtpname 
         Enabled         =   0   'False
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
         TabIndex        =   7
         Top             =   1560
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
         TabIndex        =   6
         Top             =   3360
         Width           =   735
      End
      Begin VB.TextBox txtexpiry 
         Enabled         =   0   'False
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
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtpdate 
         Enabled         =   0   'False
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
         TabIndex        =   4
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtbillno 
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
         TabIndex        =   3
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox cmbpid 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtprice 
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
         Left            =   4920
         TabIndex        =   1
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Unit Price"
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
         Left            =   2520
         TabIndex        =   23
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Address"
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
         TabIndex        =   20
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Custormer Name"
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
         TabIndex        =   18
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Bill No"
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
         Left            =   3960
         TabIndex        =   14
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label5 
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
         TabIndex        =   13
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label4 
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
         Left            =   2880
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
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
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Drug ID"
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
         Left            =   840
         TabIndex        =   10
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
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
         TabIndex        =   9
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Price"
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
         Left            =   4440
         TabIndex        =   8
         Top             =   3480
         Width           =   495
      End
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   0
      Picture         =   "frmBill.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2715
   End
End
Attribute VB_Name = "frmBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbpid_Click()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

rs.Open "Select * from Master where DrugID = '" & Me.cmbpid & "'", con, adOpenKeyset, adLockOptimistic

If rs.EOF <> True And rs.BOF <> True Then
Me.txtexpiry = rs.Fields("ExpDate")
Me.txtpdate = rs.Fields("MfdDate")
Me.txtpname = rs.Fields("DrugName")
'Me.txtprice = rs.Fields("Price")
'Me.txtqty = rs.Fields("Qty")
'Me.txtshelf = rs.Fields("Shelf")
End If

End Sub

Private Sub cmbpid_GotFocus()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

rs.Open "Select DrugID from Master", con, adOpenKeyset, adLockOptimistic
While rs.EOF <> True
cmbpid.AddItem rs!DrugID
rs.MoveNext
Wend
End Sub

Private Sub cmdexit_Click()
'DataEnvironment1.Bill.close
Unload Me
End Sub

Private Sub cmdprint_Click()
''Set rs = New ADODB.Recordset
''Set con = New ADODB.Connection
''con.Open (Constring)
''
''If DataEnvironment1.Bill.State = 1 Then
''DataEnvironment1.Bill.close
''End If
''DataEnvironment1.Bill.Open
''BillNo = Me.txtbillno.Text
''DataEnvironment1.Bill BillNo
''If DataEnvironment1.rsSales.RecordCount = 0 Then
''MsgBox "No such Record Exists", vbExclamation
''Exit Sub
''End If
Me.Hide
billrpt.Show vbModal
End Sub

Private Sub cmdrefresh_Click()
Me.txtexpiry = ""
Me.txtexpiry = ""
Me.txtpdate = ""
Me.txtshelf = ""
Me.txtcustname = ""
Me.txtAdd = ""
Me.txtpname = ""
Me.txtqty = ""
Me.txtprice = ""
Me.txtuprice = ""
Me.cmbpid.Clear
End Sub

Private Sub cmdsave_Click()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

rs.Open "Select * from Bill", con, adOpenKeyset, adLockOptimistic

With rs
    .AddNew
    .Fields("CustName") = Me.txtcustname
    .Fields("Address") = Me.txtAdd
    .Fields("Description") = Me.txtpname
    .Fields("Qty") = Me.txtqty
    .Fields("UnitPrice") = Me.txtprice
    .Fields("TotalPrice") = Me.txtuprice
    .Fields("BillNo") = Me.txtbillno
    .Update
    .close
End With
Me.txtcustname = ""
Me.txtAdd = ""
Me.txtpname = ""
Me.txtqty = ""
Me.txtprice = ""
Me.txtuprice = ""
Me.txtexpiry = ""
Me.txtexpiry = ""
Me.txtpdate = ""
'Me.txtbillno = ""
Me.cmbpid.Clear

'Set con = Nothing
'Set rs = Nothing

End Sub

Private Sub Form_Load()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

con.Execute "delete * from bill"

Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2.4

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'con.close
End Sub

Private Sub txtuprice_LostFocus()
Me.txtprice = Val(Me.txtqty) * Val(Me.txtuprice)
End Sub
