VERSION 5.00
Begin VB.Form Main 
   Caption         =   "Converter"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command12 
      Caption         =   "Convert decimal to hexadecimal"
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   1440
      Width           =   3375
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Convert hexadecimal to decimal"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   3375
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Convert binary ASCII code to character"
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   2880
      Width           =   3375
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Convert character to binary ASCII code"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   3375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Convert hexadecimal ASCII code to character"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   2400
      Width           =   3375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Convert character to hexadecimal ASCII code"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   3375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Convert ASCII code to character"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Convert character to ASCII code"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Convert hexadecimal to binary"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Convert binary to hexadecimal"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Convert decimal to binary"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convert binary to decimal"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3600
      TabIndex        =   1
      Text            =   "Output"
      Top             =   120
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Input"
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label1 
      Height          =   135
      Left            =   3240
      TabIndex        =   14
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TmpTXT1

Private Sub Command1_Click()
On Error GoTo Err
strDec = Text1.Text
Dim i%, lngtmp1&, lngtmp2&, byttmp As Byte
For i = 1 To Len(strDec)
byttmp = Asc(Mid$(strDec, i, 1)) - Asc("0")
If (byttmp = 1) Then
lngtmp2 = 2 ^ (Len(strDec) - i)
lngtmp1 = lngtmp1 + lngtmp2
End If
Next i
Text2.Text = lngtmp1
Exit Sub
Err:
MsgBox "There was an error", vbCritical, "Error"
List.List1.AddItem Time & ": Error " & Err.Number & ", " & Err.Description
End Sub

Private Sub Command10_Click()
On Error GoTo Err
Command1_Click
TmpTXT1 = Text1.Text
Text1.Text = Text2.Text
Command6_Click
Text1.Text = TmpTXT1
Exit Sub
Err:
MsgBox "There was an error", vbCritical, "Error"
List.List1.AddItem Time & ": Error " & Err.Number & ", " & Err.Description
End Sub

Private Sub Command11_Click()
On Error GoTo Err
Dim H As String
H = Text1.Text
Dim Tmp$
Dim lo1 As Integer, lo2 As Integer
Dim hi1 As Long, hi2 As Long
Const Hx = "&H"
Const BigShift = 65536
Const LilShift = 256, Two = 2
Tmp = H
If UCase(Left$(H, 2)) = "&H" Then Tmp = Mid$(H, 3)
Tmp = Right$("0000000" & Tmp, 8)
If IsNumeric(Hx & Tmp) Then
lo1 = CInt(Hx & Right$(Tmp, Two))
hi1 = CLng(Hx & Mid$(Tmp, 5, Two))
lo2 = CInt(Hx & Mid$(Tmp, 3, Two))
hi2 = CLng(Hx & Left$(Tmp, Two))
Text2.Text = CCur(hi2 * LilShift + lo2) * BigShift + (hi1 * LilShift) + lo1
End If
Exit Sub
Err:
MsgBox "There was an error", vbCritical, "Error"
List.List1.AddItem Time & ": Error " & Err.Number & ", " & Err.Description
End Sub

Private Sub Command12_Click()
On Error GoTo Err
Decnum = Text1.Text
Dim NextHexDigit As Double
Dim HexNum As String
HexNum = ""
While Decnum <> 0
NextHexDigit = Decnum - (Int(Decnum / 16) * 16)
If NextHexDigit < 10 Then
HexNum = Chr(Asc(NextHexDigit)) & HexNum
Else
HexNum = Chr(Asc("A") + NextHexDigit - 10) & HexNum
End If
Decnum = Int(Decnum / 16)
Wend
If HexNum = "" Then HexNum = "0"
Text2.Text = HexNum
Exit Sub
Err:
MsgBox "There was an error", vbCritical, "Error"
List.List1.AddItem Time & ": Error " & Err.Number & ", " & Err.Description
End Sub

Private Sub Command2_Click()
On Error GoTo Err
Dec = Text1.Text
Dim Temp As String, Retrn As String ' as string so that we don't Get number limitations
Do
Temp = Str(Dec Mod 2)
Retrn = Temp & Retrn
Dec = IIf(Right(Str(Dec), 2) = ".5", Dec - 0.5, IIf(Dec Mod 2 > 0, Dec - 1, Dec)) / 2
Loop Until Dec = 0
Text2.Text = Val(Retrn)
Exit Sub
Err:
MsgBox "There was an error", vbCritical, "Error"
List.List1.AddItem Time & ": Error " & Err.Number & ", " & Err.Description
End Sub

Private Sub Command3_Click()
On Error GoTo Err
Command1_Click
TmpTXT1 = Text1.Text
Text1.Text = Text2.Text
Command12_Click
Text1.Text = TmpTXT1
Exit Sub
Err:
MsgBox "There was an error", vbCritical, "Error"
List.List1.AddItem Time & ": Error " & Err.Number & ", " & Err.Description
End Sub

Private Sub Command4_Click()
On Error GoTo Err
Command11_Click
TmpTXT1 = Text1.Text
Text1.Text = Text2.Text
Command2_Click
Text1.Text = TmpTXT1
Exit Sub
Err:
MsgBox "There was an error", vbCritical, "Error"
List.List1.AddItem Time & ": Error " & Err.Number & ", " & Err.Description
End Sub

Private Sub Command5_Click()
On Error GoTo Err
If Len(Text1.Text) > 1 Then
MsgBox "Sorry, can only convert 1 character to ASCII at a time.", vbCritical, "Too long"
List.List1.AddItem Time & ": Error " & "00" & ", " & "Too much characters"
Exit Sub
End If
Text2.Text = Asc(Text1.Text)
Exit Sub
Err:
MsgBox "There was an error", vbCritical, "Error"
List.List1.AddItem Time & ": Error " & Err.Number & ", " & Err.Description
End Sub

Private Sub Command6_Click()
On Error GoTo Err
Text2.Text = Chr(Text1.Text)
Exit Sub
Err:
MsgBox "There was an error", vbCritical, "Error"
List.List1.AddItem Time & ": Error " & Err.Number & ", " & Err.Description
End Sub

Private Sub Command7_Click()
On Error GoTo Err
Command5_Click
TmpTXT1 = Text1.Text
Text1.Text = Text2.Text
Command12_Click
Text1.Text = TmpTXT1
Exit Sub
Err:
MsgBox "There was an error", vbCritical, "Error"
List.List1.AddItem Time & ": Error " & Err.Number & ", " & Err.Description
End Sub

Private Sub Command8_Click()
On Error GoTo Err
Command11_Click
TmpTXT1 = Text1.Text
Text1.Text = Text2.Text
Command6_Click
Text1.Text = TmpTXT1
Exit Sub
Err:
MsgBox "There was an error", vbCritical, "Error"
List.List1.AddItem Time & ": Error " & Err.Number & ", " & Err.Description
End Sub

Private Sub Command9_Click()
On Error GoTo Err
Command5_Click
TmpTXT1 = Text1.Text
Text1.Text = Text2.Text
Command2_Click
Text1.Text = TmpTXT1
Exit Sub
Err:
MsgBox "There was an error", vbCritical, "Error"
List.List1.AddItem Time & ": Error " & Err.Number & ", " & Err.Description
End Sub

Private Sub Label1_DblClick()
MsgBox "Good job, you found the secret list!!!", vbExclamation, "LOL :D"
List.Show
End Sub
