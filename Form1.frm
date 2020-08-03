VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Menampilkan Kotak Dialog Find File"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5685
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cari"
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Call ShowFindDialog("C:\")
End Sub

