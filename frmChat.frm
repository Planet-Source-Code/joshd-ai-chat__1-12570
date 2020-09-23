VERSION 5.00
Begin VB.Form frmChat 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "AI Chat"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   289
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   383
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtChat 
      Height          =   3735
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   5535
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go >"
      Height          =   300
      Left            =   5160
      TabIndex        =   1
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox txtSay 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   4935
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UserName As String
Dim ResponseCount As Integer
Dim I As Integer
Private Sub cmdGo_Click()
Dim AIReply As String
If UserName = "" Then       'The first time the button is pressed
    UserName = txtSay.Text   'Whatever the user types becomes their name
    txtChat.Text = txtChat.Text & vbNewLine & vbNewLine & UserName & ">> " & UserName
    txtChat.Text = txtChat.Text & vbNewLine & vbNewLine & "Mr Computer>> Hiya " & UserName & ". Talk."
Else
    'Show what the user typed in the main text box
    txtChat.Text = txtChat.Text & vbNewLine & vbNewLine & UserName & ">> " & txtSay.Text
    For I = 1 To ResponseCount      'For each reponse from chat.dat, if all three keywords are in the string then reply is the reply for that line
        If InStr(1, LCase(txtSay.Text), AIAnswer(I).KeyWord(1)) <> 0 And InStr(1, LCase(txtSay.Text), AIAnswer(I).KeyWord(1)) <> 0 And InStr(1, LCase(txtSay.Text), AIAnswer(I).KeyWord(2)) <> 0 Then
            AIReply = AIAnswer(I).AnswerText
            Exit For                'Stop looking
        End If
    Next I
    Dim TextToFile As String
    If AIReply = "" Then    'If computer hasn't found a reply yet
        TextToFile = txtSay.Text    'Must use a variable when writing to a file
        AIReply = "Huh?"
        Open App.Path & "/Huh.txt" For Append As #1
        Print #1, TextToFile    'Log the question (so you can add it to chat.dat)
        Close #1
    End If
    txtChat.Text = txtChat.Text & vbNewLine & vbNewLine & "Mr Computer>> " & AIReply    'Put the computer's response in the main window
End If
txtChat.SelStart = Len(txtChat) 'Look at the bottom (latest) of the main chat box.
End Sub

Private Sub Form_Load()
ResponseCount = 0
Open App.Path & "/Chat.dat" For Input As #1
While Not EOF(1) 'Keep loading until we get to the bottom of the file
    ResponseCount = ResponseCount + 1   'Keep track of the number of responses
    Input #1, AIAnswer(ResponseCount).KeyWord(1), AIAnswer(ResponseCount).KeyWord(2), AIAnswer(ResponseCount).KeyWord(3), AIAnswer(ResponseCount).AnswerText
Wend
Close #1
txtChat.Text = "Mr Computer>> Hello! What is your name?"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Open App.Path & "/Log.txt" For Append As #1 'Log the conversation to log.txt
    Print #1, "::::::::::::::" & Date & "::::::::::::::"
    Print #1, txtChat.Text
Close #1
End Sub

Private Sub txtSay_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then    'If the user pressed enter
    KeyCode = 0 'Stop the beep
    cmdGo_Click
End If
End Sub
