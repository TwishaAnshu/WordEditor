VERSION 5.00
Begin VB.Form frmEdit 
   Caption         =   "Form1"
   ClientHeight    =   8184
   ClientLeft      =   192
   ClientTop       =   840
   ClientWidth     =   17352
   LinkTopic       =   "Form1"
   ScaleHeight     =   8184
   ScaleWidth      =   17352
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEdit 
      Height          =   7812
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "fmMEdit.frx":0000
      Top             =   1200
      Width           =   16332
   End
   Begin VB.Menu mnuFile 
      Caption         =   "F&ile"
      Begin VB.Menu New 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu save 
         Caption         =   "save"
      End
      Begin VB.Menu mnusave 
         Caption         =   "SAVE AS"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&XIT"
      End
   End
   Begin VB.Menu Edit 
      Caption         =   "E&dit"
      Begin VB.Menu cut 
         Caption         =   "cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu selectall 
         Caption         =   "selectall"
         Shortcut        =   ^A
      End
      Begin VB.Menu copy 
         Caption         =   "copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu Font 
      Caption         =   "Font"
      Begin VB.Menu style 
         Caption         =   "style"
         Begin VB.Menu arial 
            Caption         =   "arial"
         End
         Begin VB.Menu comicSans 
            Caption         =   "comic sans"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu size 
         Caption         =   "size"
         Begin VB.Menu thirty 
            Caption         =   "30"
         End
         Begin VB.Menu fifty 
            Caption         =   "50"
         End
      End
      Begin VB.Menu extra 
         Caption         =   "extra features"
         Begin VB.Menu bold 
            Caption         =   "bold"
         End
         Begin VB.Menu italics 
            Caption         =   "italics"
         End
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim OldBody, NewBody, LMNO As String
Dim total As Double

Dim fname, path, ans As String
' Public Check()

'OldBody = Val(Len(txtEdit.SelLength))
'NewBody = Val(Len(txtEdit.SelLength) + OldBody)

'If NewBody = Val(OldBody / 2) Then
'MsgBox ("no change ")


'Else
'MsgBox (" change ")

'End If



'End

'Public CHECK()
'OldBody = Val(Len(LMNO))
'NewBody = Val(Len(Total))

'If OldBody = NewBody Then
'End
'End If

'If OldBody <> NewBody Then
'fname = Trim(UCase(InputBox("Enter the name of the file to save:", " First SAVEAS", "bob")))
'path = "C:\Users\twish\Desktop\CP1 NEW NEW\MEdit\" + fname + ".txt"
'Open path For Output As #1
'Print #1, txtEdit
'Close #1

'End If


Private Sub arial_Click()
txtEdit.Font = "arial"
End Sub

Private Sub blue_Click()
txtEdit.ForeColor = "blue"
End Sub

Private Sub bold_Click()
txtEdit.FontBold = True
End Sub

Private Sub comicSans_Click()
txtEdit.Font = "comic sans"
End Sub

Private Sub copy_Click()
Clipboard.SetText txtEdit.SelText
paste.Enabled = True

End Sub

Private Sub cut_Click()
Clipboard.SetText txtEdit.SelText
txtEdit.SelText = ""
paste.Enabled = True

End Sub

Private Sub fifty_Click()
txtEdit.FontSize = "50"

End Sub

Private Sub italics_Click()
txtEdit.FontItalic = True
End Sub

Private Sub LMNO_Click()

End Sub

Private Sub New_Click()



'Check



ans = MsgBox("Are you sure you want to open a new file? ", vbYesNo, "NEW")
If ans = vbYes Then
txtEdit = ""

selectall.Enabled = True


LMNO = Val(Len(txtEdit))


OldBody = Len(LMNO)
NewBody = Len(total)


If OldBody = NewBody Then
End
End If

If OldBody <> NewBody Then
fname = Trim(UCase(InputBox("Enter the name of the file to save:", " First SAVEAS", "bob")))
path = "C:\Users\twish\Desktop\CP1 NEW NEW\MEdit\" + fname + ".txt"
Open path For Output As #1
Print #1, txtEdit
Close #1
End If
End If

If ans = vbNo Then
txtEdit.SetFocus

End If




End Sub

Private Sub paste_Click()
txtEdit.SelText = Clipboard.GetText()

End Sub


Private Sub Form_Resize()
txtEdit.Width = frmEdit.ScaleWidth
txtEdit.Height = frmEdit.ScaleHeight

End Sub

Private Sub mnuExit_Click()
If mnusave.Enabled = True Then
fname = Trim(UCase(InputBox("Enter the name of the file to save:", " First SAVEAS", "bob")))
path = "C:\Users\twish\Desktop\CP1 NEW NEW\MEdit\" + fname + ".txt"
Open path For Output As #1
Print #1, txtEdit
Close #1

End If


End
End Sub

Private Sub mnuOpen_Click()


ans = vbNo
Do While ans = vbNo
fname = UCase$(Trim$(InputBox("FileName", "OpenFile")))
path = "C:\Users\twish\Desktop\CP1 NEW NEW\MEdit\" + fname + ".txt"
ans = MsgBox(path, vbYesNo, " Is this the path?")

Loop
If ans = vbYes Then
Open path For Input As #1
Dim filesize As Integer
filesize = LOF(1)
txtEdit = Input$(filesize, #1)
Close #1
End If
End Sub

Private Sub mnuSave_Click()
fname = Trim(UCase(InputBox("Enter the name of the file to save:", " First SAVEAS", "bob")))
path = "C:\Users\twish\Desktop\CP1 NEW NEW\MEdit\" + fname + ".txt"
Open path For Output As #1
Print #1, txtEdit
Close #1





End Sub

Private Sub red_Click()
txtEdit.Text = "red"
End Sub

Private Sub save_Click()
path = "C:\Users\twish\Desktop\CP1 NEW NEW\MEdit\" + fname + ".txt"
Open path For Output As #1
Print #1, txtEdit
Close #1

End Sub

Private Sub selectall_Click()
 With txtEdit
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub thirty_Click()
txtEdit.FontSize = "30"
End Sub

Private Sub txtEdit_Change()
total = Len(txtEdit)


End Sub

Private Sub txtEdit_GotFocus()
LMNO = Val(Len(txtEdit))
End Sub
