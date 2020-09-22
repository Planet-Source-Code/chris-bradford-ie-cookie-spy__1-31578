VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cookie Spy"
   ClientHeight    =   1020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1020
   ScaleWidth      =   2805
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton btnStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stopped"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnStart_Click()
    'swap buttons and notification message
    
    btnStart.Enabled = False
    Command2.Enabled = True
    Label1.Caption = "Monitoring cookies"
    
    'close the results window, so it can load the results
    'properly when the stop button is clicked
    Unload frmCompare
    
    'inventory the cookies into our 2st array
    CookieInventory StartInv()

End Sub




Private Sub Command2_Click()
    'swap buttons and notification message
    
    Command2.Enabled = False
    btnStart.Enabled = True
    Label1.Caption = "Stopped"
    
    'inventory the cookies into our 2nd array
    CookieInventory EndInv()
    
    'open the results window
    frmCompare.Show
    
    
End Sub



Private Sub Form_Unload(Cancel As Integer)
    'call the exit function in modMain
    goodbye
    
End Sub
