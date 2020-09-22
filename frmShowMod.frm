VERSION 5.00
Begin VB.Form frmShowMod 
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEnd 
      Height          =   2775
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   3480
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   6360
      Width           =   975
   End
   Begin VB.TextBox txtStart 
      Height          =   2775
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
   Begin VB.Label Label2 
      Caption         =   "Modified Value:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Previous Value:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmShowMod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
    
End Sub


