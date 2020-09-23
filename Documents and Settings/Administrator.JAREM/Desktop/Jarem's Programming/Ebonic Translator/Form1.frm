VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Ebonic Translator"
   ClientHeight    =   5070
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Enter Text Here:"
      Height          =   2610
      Left            =   15
      TabIndex        =   4
      Top             =   -30
      Width           =   5940
      Begin VB.TextBox Text1 
         Height          =   2325
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "Form1.frx":0000
         Top             =   210
         Width           =   5730
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Translation:"
      Height          =   2115
      Left            =   15
      TabIndex        =   2
      Top             =   2580
      Width           =   5940
      Begin VB.TextBox Text2 
         Height          =   1875
         Left            =   105
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   195
         Width           =   5730
      End
   End
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   300
      Left            =   1665
      TabIndex        =   1
      Top             =   4755
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
      Max             =   656
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Translate!"
      Enabled         =   0   'False
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Top             =   4770
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Word() As String

Private Sub Command1_Click()
Translate
End Sub

Private Sub Form_Resize()
On Error Resume Next
Frame2.Height = Me.Height / 2
Frame1.Top = Frame2.Height - 30
Frame1.Height = Frame2.Height - 750

Frame2.Width = Me.Width - 180
Frame1.Width = Me.Width - 180

Text1.Height = Frame2.Height - 285
Text1.Width = Frame2.Width - 210

Text2.Height = Frame1.Height - 240
Text2.Width = Frame1.Width - 210

Progress.Top = Me.Height - 720
Command1.Top = Me.Height - 720

Progress.Width = Me.Width - 1815
Progress.Left = Command1.Width + 100


End Sub

Private Sub ProgressBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub
Public Sub Translate()
'On Error GoTo err
Dim fnum As Integer
Dim Temp
Dim Txt
Dim lngindex As Integer
Dim Words
Dim Wordnum As Integer

fnum = FreeFile

Open App.Path & "\Ebonics.dic" For Input As fnum
Text1 = Text1 & " " 'add a space just in case
                    'its just one word
Words = Split(Text1, " ")
Text2 = ""

Progress.Value = 0
For lngindex = 1 To 656
Progress.Value = Progress + 1
Input #fnum, Txt
Temp = Split(Txt, vbTab)
Trim (Temp(0))
Trim (Temp(1))

For Wordnum = 0 To UBound(Words)
If LCase(Words(Wordnum)) = LCase(Temp(0)) Then
Words(Wordnum) = Temp(1)
End If
Next Wordnum


Next lngindex

For Wordnum = 0 To UBound(Words)
Text2 = Text2 & " " & Words(Wordnum)
Text2 = Trim(Text2)
Next Wordnum






Close fnum


Exit Sub
err:
Close fnum
MsgBox err.Description, vbCritical
End Sub
Private Sub Text1_GotFocus()
If Left(Text1, 3) = "Tip" Then
Command1.Enabled = True
Text1.Text = ""
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1.Value = True
End If
End Sub
