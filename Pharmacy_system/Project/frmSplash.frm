VERSION 5.00
Begin VB.Form frmSplash 
   Caption         =   "Welcome To Medizone Pharmacy System "
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4320
      Top             =   360
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   0
      Top             =   3600
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   4935
      Left            =   -240
      Picture         =   "frmSplash.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   7890
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
Me.Lbl1.Caption = "Please be Patient..."
ElseIf i = 3 Then
Me.Lbl1.Caption = "Loading Database..."
ElseIf i = 5 Then
Me.Lbl1.Caption = "Initialising Application environment..."
ElseIf i = 7 Then
Me.Lbl1.Caption = "About to get Started..."
ElseIf i = 9 Then
Me.Lbl1.Caption = "Thank you!WELCOME"
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

