Attribute VB_Name = "modMain"
'** Cookie Spy Utility
'** (c)2002, Christopher Bradford

'** http://www.snowjournal.com

'** Distributed as "open source" under the GPL
'** Feel free to distribute, modify, and use in accordance with the GPL

'HOW TO USE THIS PROGRAM

'Its pretty simple... click the "start" button when you're ready to begin
'monitoring your cookies.  Then, browse the web in the area of interest.
'When finished, click the "stop" button and the results will be displayed for you.
'Double clicking on a cookie from the results window will display the details.
'You can open as many detail windows as needed, and they can remain open for
'comparison in subsequent cookie monitoring sessions


'** I M P O R T A N T **

'set the following constant to location of your cookies
'include the trailing backslash

Public Const CookieDir = "C:\Documents and Settings\chris\Cookies\"

'**

Public StartInv() As String
Public EndInv() As String
Public frmMod() As Form
Public frmNew() As Form

Sub main()
    
    'initialize the cookie arrays
    ReDim StartInv(2, 0)
    ReDim EndInv(2, 0)
    
    'initialize the forms arrays
    ReDim frmMod(0)
    ReDim frmNew(0)
        
    frmMain.Show
    
End Sub

Public Sub CookieInventory(ByRef inv() As String)

    ReDim inv(2, 0)
    
    strFile = Dir(CookieDir + "*.txt")
    Do
        
        tmpBound = UBound(inv, 2) + 1
        ReDim Preserve inv(2, tmpBound)
        
        inv(1, tmpBound) = strFile
        
        Open CookieDir + strFile For Input As #1
        
        tmpFile = ""
        While Not EOF(1)
            Input #1, a$
            tmpFile = tmpFile + a$
        Wend
        
        Close #1
        
        inv(2, tmpBound) = tmpFile
        
        strFile = Dir               'go get the next one
        
    Loop While strFile <> ""
    
End Sub

Public Sub goodbye()

    For Each Form In Forms
        Unload Form
    Next

End Sub
