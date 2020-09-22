VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encode / Decode"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Encode"
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Decode"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   600
      Width           =   4215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is a simple encryption/decryption routine I wrote for storing a password for a
'program I was working on - this then turned into a useful function for a client/server
'program I was writing that required all data between the client and the server to
'include a key so that no-one could send spurious data to either without knowing the
'key.  The Key was the current time in minutes that was encrypted and sent with the data
'which was then decrypted at the other end and compared to the current time in minutes
'and then checked to see if it was -1,0 or +1 minutes correct.

'The encyrption/decryption is by no means hard to crack - but it is at least better
'than plain text passwords and will throw off all but the most curious / determined
'user.

'The program works by simply breaking down characters to be encrypted to their ASCII value
'and then shifting the value by a random number between 1 and 9
'and storing this number at the front of the encrypted text.
'The decryption works in the same way but reverses the process shifting the ascii value
'back by the number stored at the front of the text

'Because the number itself is a character it can be encrypted too - which means the more
'times you run the encryption routine - the longer it will take for someone who doesn't
'know the algorithm to decode.

'It is set at 100 times - which takes roughly a second - with a standard sentance

Private Sub Command1_Click()
For a = 1 To 100 'Run decode function 100 times
'get the number at the beginning of the text to pass to decoder
Text2.Text = DecodeText(Text1.Text, CInt(Left$(Text1.Text, 1)))
'pass text1.text to the decode function and the number stored at the front of the text
Text1.Text = Text2.Text 'make sure decoding just decoded text
Next a
Text1.Text = ""
'clear text1.text
End Sub

Private Sub Command2_Click()
Dim EncoderShift As Integer
For a = 1 To 100 'run 100 times
GetEncodeShift:
Randomize Timer 'make sure random number is actually random
EncoderShift = Int(Rnd * 9) 'get number between 0 and 9

If EncoderShift = 0 Then GoTo GetEncodeShift 'if number is 0 - text wont be replaced
Text2.Text = EncoderShift & EncodeText(Text1.Text, EncoderShift)
'add ecodershift value to beginning of string and then encode text
'pass text1.text to the encoder and the random number EncoderShift
Text1.Text = Text2.Text ' make sure the encoder is working with the just encoded text
Next
Text1.Text = ""
'clear text1.text
End Sub

Public Function EncodeText(ecStr As String, ecShift As Integer) As String
Dim ecNewStr As String
Dim ecStrID As Integer
ecNewStr = "" 'clear variable
For ecStrID = 1 To Len(ecStr) 'loop through each letter to be encoded
ecNewStr = ecNewStr + CoreEncodeShift(Mid$(ecStr, ecStrID, 1), ecShift, 1)
'add encoded letter to string to return
Next
EncodeText = ecNewStr 'return encoded string
End Function

Public Function DecodeText(dcStr As String, ecShift As Integer) As String
Dim dcNewStr As String
Dim dcOldStr As String
Dim dcStrID As Integer
dcNewStr = "" 'clear variable
dcOldStr = Mid$(dcStr, 2) 'remove the number key at the beginning of the text
For dcStrID = 1 To Len(dcOldStr)
dcNewStr = dcNewStr + CoreEncodeShift(Mid$(dcOldStr, dcStrID, 1), ecShift, 0)
'add decoded character to string
Next
DecodeText = dcNewStr 'return decoded text
End Function

Public Function CoreEncodeShift(cesChar As String, cesShift, cesEncode As Integer) As String
Dim cesAscEnd As Integer
Dim cesAsc As Integer
Select Case cesEncode

Case 0
'Decode String

cesAsc = Asc(cesChar) - cesShift 'subtract character ASCII value by cesShift

If cesAsc < 32 Then 'If character is under 32 it wont display
'these characters are things like vbCrLf and other unprintable character that
'would cause program errors

    cesAscEnd = 32 - cesAsc 'we work out how many under 32 the subtraction equals
    CoreEncodeShift = Chr$(126 - cesAscEnd) 'and subtract it from 126
    '126 is the top end of the printable characters we are working with
    'return the found character
Else
    CoreEncodeShift = Chr$(cesAsc) 'otherwise return the Character found by
    'doing the subtraction

End If

Case 1
'Encode String

cesAsc = Asc(cesChar) + cesShift 'add cesShift on to ASCII character value

If cesAsc > 126 Then 'if over 126 then past the printable characters we are working with
'so find out how many over 126 the value is
 '   Dim cesAscEnd As Integer
    cesAscEnd = cesAsc - 126 'and subtract it from 126 which is the top end of the
    'printable characters we want to work with
    CoreEncodeShift = Chr$(32 + cesAscEnd) 'return encoded character
Else
    CoreEncodeShift = Chr$(cesAsc) 'return encoded character

End If
Case Else
End Select
End Function
