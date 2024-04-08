VERSION 5.00
Begin VB.Form frmParameter 
   Caption         =   "Set Parameter"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   1425
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label lblDescription 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmParameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOk_Click()
    Me.Hide
End Sub
