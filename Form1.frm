VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   1560
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2280
      TabIndex        =   2
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "Encrypted Text:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Key 2:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Key 1:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Input Text:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim crypto As New clsEncrypt
Set crypto = New clsEncrypt

    crypto.inputString = Text1.Text
    crypto.Key3DES1 = Text2.Text
    crypto.key3DES2 = Text3.Text
    crypto.mode = mode_ENCRYPT
    Text4.Text = crypto.encrypt3DES
End Sub
