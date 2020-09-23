VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   4065
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   6570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   4050
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6585
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   600
         Top             =   1560
      End
      Begin VB.Label Lbl1 
         Caption         =   "Label1"
         Height          =   615
         Left            =   1920
         TabIndex        =   2
         Top             =   3120
         Width           =   4215
      End
      Begin VB.Image Image1 
         Height          =   4035
         Left            =   0
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6570
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Developed by:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   2520
         TabIndex        =   1
         Top             =   1200
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    frmLogin.Show
End Sub

Private Sub Timer1_Timer()
i = i + 1
If i = 1 Then
Me.Lbl1.Caption = "Please wait..."
ElseIf i = 3 Then
Me.Lbl1.Caption = "Loading Database..."
ElseIf i = 5 Then
Me.Lbl1.Caption = "Creating Application environment..."
ElseIf i = 7 Then
Me.Lbl1.Caption = "Almost done..."
ElseIf i = 9 Then
Me.Lbl1.Caption = "Welcome"
ElseIf i = 11 Then
Unload Me
frmLogin.Show
End If
End Sub

Private Sub Timer2_Timer()
pb.Value = pb.Value + 9
If pb.Value = 255 Then
Timer2.Enabled = False
End If
End Sub
