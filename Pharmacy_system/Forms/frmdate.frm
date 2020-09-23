VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmdate 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Date Selection"
   ClientHeight    =   2550
   ClientLeft      =   2760
   ClientTop       =   3780
   ClientWidth     =   6135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6135
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Select your Date choice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton cmdexit 
         Caption         =   "Exit"
         Height          =   495
         Left            =   3720
         TabIndex        =   6
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton cmdload 
         Caption         =   "Generate"
         Height          =   495
         Left            =   1320
         TabIndex        =   1
         Top             =   1680
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20840449
         CurrentDate     =   39459
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20840449
         CurrentDate     =   39459
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Up-To"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim Fdate As Date
Dim Ldate As Date


Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdload_Click()
Set rs = New ADODB.Recordset
Set con = New ADODB.Connection
con.Open (Constring)

If DataEnvironment1.Bill.State = 1 Then
DataEnvironment1.Bill.close
End If
DataEnvironment1.Bill.Open
Fdate = DTPicker1.Value
Ldate = DTPicker2.Value
DataEnvironment1.Sales Fdate, Ldate
If DataEnvironment1.rsSales.RecordCount = 0 Then
MsgBox "No such Record Exists", vbExclamation
Exit Sub
End If
Me.Hide
Rptsales.Show vbModal
End Sub

Private Sub Form_Load()

Left = (Screen.Width - Width) / 2
Top = (Screen.Height - Height) / 2.7

End Sub

