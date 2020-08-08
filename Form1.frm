VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mengatur Urutan Form yang Tampil"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5715
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5715
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  If Command1.Caption = "Form2" Then
     Command1.Caption = "Form3"
     Form2.ZOrder 0 'Form2 di atas
     Form3.ZOrder 1 'Form3 di bawah
  Else
     Command1.Caption = "Form2"
     Form3.ZOrder 0 'Form3 di atas
     Form2.ZOrder 1 'Form2 di bawah
  End If
End Sub

Private Sub Form_Load()
  Form2.Show
  Form3.Show
  Me.Move 6000, 5000
  Command1.Caption = "Form2"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
UnloadMode As Integer)
  End
End Sub


