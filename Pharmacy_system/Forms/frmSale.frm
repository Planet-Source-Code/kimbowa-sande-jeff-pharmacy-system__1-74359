VERSION 5.00
Begin VB.Form frmSale 
   BackColor       =   &H80000016&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Medizone Pharmacy System (Drug Sale)"
   ClientHeight    =   3795
   ClientLeft      =   1695
   ClientTop       =   3015
   ClientWidth     =   9195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   9195
   Begin VB.TextBox txtbalance 
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
      Left            =   4320
      TabIndex        =   15
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Product Description"
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
      Height          =   3615
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   6375
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
         Left            =   2760
         TabIndex        =   21
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtTDate 
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
         Left            =   4800
         TabIndex        =   19
         Top             =   2400
         Width           =   1455
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
         Left            =   4440
         TabIndex        =   18
         Top             =   3000
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
         Left            =   1080
         TabIndex        =   17
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txttotalprice 
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
         Left            =   4800
         TabIndex        =   13
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtunitprice 
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
         Left            =   4800
         TabIndex        =   11
         Top             =   1440
         Width           =   1455
      End
      Begin VB.ComboBox cmbpid 
         BackColor       =   &H8000000E&
         Height          =   315
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   3855
      End
      Begin VB.TextBox txtshelf 
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
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   1455
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   2055
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1440
         Width           =   2055
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
         Left            =   1560
         TabIndex        =   2
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label10 
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
         Height          =   255
         Left            =   3840
         TabIndex        =   20
         Top             =   2520
         Width           =   975
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
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Price"
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
         Left            =   3840
         TabIndex        =   14
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
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
         Left            =   3840
         TabIndex        =   12
         Top             =   1560
         Width           =   1095
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
         TabIndex        =   10
         Top             =   480
         Width           =   1335
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
         TabIndex        =   9
         Top             =   960
         Width           =   1455
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
         TabIndex        =   8
         Top             =   1560
         Width           =   1455
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
         Left            =   120
         TabIndex        =   7
         Top             =   2160
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
         Left            =   3840
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Image Image2 
      Height          =   4095
      Left            =   -240
      Picture         =   "frmSale.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   6000
      Left            =   -480
      Picture         =   "frmSale.frx":B4F6
      Stretch         =   -1  'True
      Top             =   -480
      Width           =   10155
   End
End
Attribute VB_Name = "frmSale"
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
'Me.txtpname = rs.Fields("DrugName")
'Me.txtprice = rs.Fields("Price")
'Me.txtqty = rs.Fields("Qty")
Me.txtshelf = rs.Fields("Shelf")
Me.txtbalance = rs.Fields("Qty")
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

Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdprint_Click()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)
Unload Me
If DataEnvironment1.Bill.State = 1 Then
DataEnvironment1.Bill.Close
End If
DataEnvironment1.Bill.Open
DataEnvironment1.custormer
billrpt.Show vbModal
End Sub

Private Sub cmdsave_Click()
'used to make connection to the database
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
Set con = New ADODB.Connection
Set rsbill = New ADODB.Recordset
con.Open (Constring)

rs.Open "Select * from Master where DrugName = '" & Me.cmbpid & "'", con, adOpenKeyset, adLockOptimistic
If rs.EOF <> True And rs.BOF <> True Then

With rs
    .Fields("Qty") = rs.Fields("Qty") - Val(Me.txtqty)
If rs.Fields("qty") <= -1 Then
MsgBox "THAT ITEM IS NOT AVIALABLE ", vbInformation
Exit Sub
End If
    .Update
End With
MsgBox "item qty update"
Set rs = Nothing

rs1.Open "Select * from Sales", con, adOpenKeyset, adLockOptimistic

With rs1
    .AddNew
    .Fields("DrugName") = Me.cmbpid
    .Fields("Price") = Me.txtunitprice
    .Fields("Tprice") = Me.txttotalprice
    .Fields("Qty") = Me.txtqty
    .Fields("Shelf") = Me.txtshelf
    .Fields("ProdDate") = Me.txtpdate
    .Fields("ExpDate") = Me.txtexpiry
    .Fields("seller") = frmLogin.txtUserName.Text
    .Fields("Selldate") = Me.txtTDate
    .Update
    .Close
End With

rsbill.Open "Select * from Bill", con, adOpenKeyset, adLockOptimistic

With rsbill
    .AddNew
    .Fields("Description") = Me.cmbpid
    .Fields("Qty") = Me.txtqty
    .Fields("UnitPrice") = Me.txtunitprice
    .Fields("TotalPrice") = Me.txttotalprice
    .Update
    .Close
End With
    Me.cmbpid.Clear
    Me.txtunitprice = ""
    Me.txttotalprice = ""
    Me.txtqty = ""
    Me.txtshelf = ""
    Me.txtpdate = ""
    Me.txtexpiry = ""
    txtbalance = ""
Set rs = Nothing
Set con = Nothing

End If
End Sub

Private Sub Form_Load()
Set con = New ADODB.Connection
con.Open (Constring)

con.Execute "Delete * from Bill"
cmbpid.Clear
Me.txtTDate = Date
Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2.4

End Sub





Private Sub txtunitprice_LostFocus()
Me.txttotalprice.Text = Val(Me.txtqty) * Val(Me.txtunitprice)
End Sub
