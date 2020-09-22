VERSION 5.00
Begin VB.Form frmCompare 
   Caption         =   "Spy Results"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5580
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstMod 
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.ListBox lstModified 
      Height          =   2010
      Left            =   0
      TabIndex        =   1
      Top             =   2880
      Width           =   5415
   End
   Begin VB.ListBox lstNew 
      Height          =   2010
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label2 
      Caption         =   "Modified Cookies"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "New Cookies"
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
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
        
    'the two cookie arrays are compared when this form is loaded
        
    'clear the listboxes
    lstNew.Clear
    lstModified.Clear
    
    
    'lets do the comparisons between arrays
    
    For x = 1 To UBound(EndInv, 2)          'loop through 2nd array
        
        tmpFoundIt = False                  'test boolean for new cookie
        For y = 1 To UBound(StartInv, 2)    'loop through 1st array
        
            If EndInv(1, x) = StartInv(1, y) Then   'find matching item
                
                'found a matching cookie
                tmpFoundIt = True
                
                'test for modification
                If EndInv(2, x) <> StartInv(2, y) Then
                    'store modified item
                    lstModified.AddItem EndInv(1, x)
                    lstModified.ItemData(lstModified.NewIndex) = x
                    
                    'store modified item from starting array
                    lstMod.AddItem StartInv(1, y)
                    lstMod.ItemData(lstModified.NewIndex) = y
                    
                End If
                
                Exit For        'jump out of this loop
            End If
        Next y
         
        'check for new item
        If tmpFoundIt = False Then
            'add new cookie to "new cookies" listbox
            lstNew.AddItem EndInv(1, x)
            lstNew.ItemData(lstNew.NewIndex) = x
        End If
        
    Next x
    
End Sub


Private Sub lstModified_DblClick()
    
    'sanitize the strings for multiline display
    tmpStartText = Replace(StartInv(2, lstMod.ItemData(lstModified.ListIndex)), Chr$(10), vbCrLf)
    tmpEndText = Replace(EndInv(2, lstModified.ItemData(lstModified.ListIndex)), Chr$(10), vbCrLf)
    
    'instanciate a new form
    tmpBound = UBound(frmMod) + 1
    ReDim Preserve frmMod(tmpBound)
    Set frmMod(tmpBound) = New frmShowMod
    
    'set form and information
    frmMod(tmpBound).txtStart.Text = tmpStartText
    frmMod(tmpBound).txtEnd.Text = tmpEndText
    frmMod(tmpBound).Caption = EndInv(1, lstModified.ItemData(lstModified.ListIndex))
    
    frmMod(tmpBound).Show
    
End Sub

Private Sub lstNew_DblClick()
    
    'sanitize the string for multiline display
    tmpText = Replace(EndInv(2, lstNew.ItemData(lstNew.ListIndex)), Chr$(10), vbCrLf)
    
    'instanciate a new form
    tmpBound = UBound(frmNew) + 1
    ReDim Preserve frmNew(tmpBound)
    Set frmNew(tmpBound) = New frmShowNew
    
    'set form and information
    frmNew(tmpBound).txtCookie.Text = tmpText
    frmNew(tmpBound).Caption = EndInv(1, lstNew.ItemData(lstNew.ListIndex))
    
    frmNew(tmpBound).Show
    
End Sub
