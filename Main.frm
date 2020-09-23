VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Secure Me (+) Update! 2.7"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2140
      TabIndex        =   3
      Text            =   "21,23,139"
      Top             =   600
      Width           =   870
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4440
      TabIndex        =   5
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   285
      Left            =   3120
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Exit"
      Height          =   255
      Left            =   5160
      TabIndex        =   12
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "About"
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save log to disk"
      Height          =   255
      Left            =   2520
      TabIndex        =   10
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "View Soft Log"
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "View Hard Log"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   4800
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2400
      Top             =   1200
   End
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   3120
      Top             =   1200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox newlog 
      Height          =   1800
      Left            =   0
      TabIndex        =   17
      Top             =   860
      Width           =   5920
      _ExtentX        =   10451
      _ExtentY        =   3175
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Main.frx":0442
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2100
      TabIndex        =   1
      Text            =   "You do not have permission to access this service and are being reported to your ISP for malicious activatie."
      Top             =   0
      Width           =   3675
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Send Intruder Message:"
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Send me a visual alert when a intruder tries to connect to my computer."
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   5655
   End
   Begin VB.ListBox List1 
      Height          =   3180
      Left            =   4320
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Width           =   5895
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   495
      Left            =   2400
      TabIndex        =   16
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "1"
      Height          =   495
      Left            =   2400
      TabIndex        =   15
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   2400
      TabIndex        =   14
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Port to listen for connection:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   630
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'oh I am meth0s
'eat me!

Private Sub Command1_Click()
On Error GoTo er
'this is our start button when its clicked
Dim helloa As String
Dim hellob As String
Dim helloc As Integer
Dim hellod As String
Dim hello1 As String
'we are going to dim everything

hellod = ""
helloa = ""
hellob = ""
helloc = 0
hello1 = ""
Timer2.Enabled = False
List1.Clear
'we are going to clear these

helloa = Len(Text3.Text)
'saying helloa is however many characters are in text3

For i = 1 To helloa
'For every character in helloa.

helloc = helloc + 1
'You will understand this later on in the code.

hellob = Mid(Text3.Text, i, 1)
'hellob is going to be the character of i which is the current character.

If hellob = "," Then
'if hellob is , we are going to parse it.

hellob = ""
'we are clearing away the , so that we only have the number.

hellod = hellod + hellob
'hellod is going to be the same as it use to be just add hellob to it.

If hellod = "" Then
'if hellod is nothing. we dont want to add it.

Else
'if its not.

List1.AddItem hellod
'lets go ahead and add it to our listbox

End If

hellod = ""
'clearing hellod so we start over on the next port.

Else
'if its not a , then add it to hello d
hellod = hellod + hellob

If helloc = helloa Then List1.AddItem hellod
'remember up there. helloc = helloc + 1 well this is where we use it.
'if we are on the last character. we will see if its the last port.
'so we can add it without having to put the , for the parsing.

End If
Next i
'go to the next character.

Call xListKillDupes(List1) 'calls sub from module


'now we will count all the winsock controls we need to load.
hello1 = List1.ListCount

For c = 0 To List1.ListCount - 1
Load Winsock(c + 1) 'loading a new winsock control in
List1.ListIndex = c 'going to the current port in the listbox
Winsock(c).LocalPort = List1.Text 'assigning this new controls port.
Winsock(c).Listen 'making the new control listen on its new port.
Winsock(c).Tag = List1.Text 'so when they connect we know what port they connect on.
Next c
Check1.Enabled = False 'disabling check box so nothing messes up!
Check2.Enabled = False 'disabling other check box so nothing messes up!
Text2.Enabled = False 'disabling text2 so nothing messes up!
Text3.Enabled = False 'disabling text3 so nothing messes up!
Command1.Enabled = False 'disabling start button so nothing can mess up!
Command2.Enabled = True 'Enabling the stop button so you can stop it!
newlog.Text = ""
newlog.SelColor = &H80000002
newlog.SelText = "!!!"
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & " Started on "
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Format(Now, "General Date")
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & " !!!" & vbCrLf
newlog.SelColor = &H80&
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & "!!! "
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & "Listening for connections on TCP port(s) "
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Text3.Text
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & " !!!" & vbCrLf
newlog.SelColor = &H80&
er:
If Err.Number = 10055 Then
MsgBox "You have ran out of buffer space." & vbCrLf & "Some of the ports you have tried to open" & vbCrLf & "Where not opened up.", vbCritical, "Warning - Win9x"
Check1.Enabled = False 'disabling check box so nothing messes up!
Check2.Enabled = False 'disabling other check box so nothing messes up!
Text2.Enabled = False 'disabling text2 so nothing messes up!
Text3.Enabled = False 'disabling text3 so nothing messes up!
Command1.Enabled = False 'disabling start button so nothing can mess up!
Command2.Enabled = True 'Enabling the stop button so you can stop it!
newlog.SelColor = &H80000002
newlog.SelText = "!!!"
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & " Started on "
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Format(Now, "General Date")
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & " !!!" & vbCrLf
newlog.SelColor = &H80&
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & "!!! "
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & "Listening for connections on tcp port(s) "
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Text3.Text
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & " !!!" & vbCrLf
newlog.SelColor = &H80&
ElseIf Err.Number = 0 Then
Check1.Enabled = False 'disabling check box so nothing messes up!
Check2.Enabled = False 'disabling other check box so nothing messes up!
Text2.Enabled = False 'disabling text2 so nothing messes up!
Text3.Enabled = False 'disabling text3 so nothing messes up!
Command1.Enabled = False 'disabling start button so nothing can mess up!
Command2.Enabled = True 'Enabling the stop button so you can stop it!
Else
MsgBox Err.Description, vbCritical, "Error"
End
End If
End Sub

Private Sub Command2_Click()
List1.Clear
Timer2.Enabled = False
' this is our stop button when its clicked
Check1.Enabled = True 'Enabling Check1 so that it can be used again
Check2.Enabled = True 'Enabling Check2 so it can be used again
Text2.Enabled = True 'Enabling text2 so it can be used again
Text3.Enabled = True 'Enabling text3 so it can be used again
Command1.Enabled = True 'Enabling the start button so it can be used again!
Command2.Enabled = False 'Disabling the stop button so nothing can mess up!
For i = Winsock.LBound + 1 To Winsock.UBound
Winsock(i).Close
Unload Winsock(i)
Next i
Winsock(0).Close
'A better winsock unloader. It was having trouble previously.
newlog.SelColor = &H80000002
newlog.SelText = "!!!"
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & newlog.SelText & " Stopped on "
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Format(Now, "General Date")
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & " !!!" & vbCrLf
newlog.SelColor = &H80&
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & "!!! "
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & "Closing TCP port(s) "
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Text3.Text
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & " !!!" & vbCrLf
newlog.SelColor = &H80&
End Sub

Private Sub Command3_Click()
Label4.Caption = 1
newlog.SelColor = &H80000002
newlog.SelText = "!!! "
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Format(Now, "General Date")
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & " - Hard log is on"
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & " !!!" & vbCrLf
newlog.SelColor = &H80&
End Sub

Private Sub Command4_Click()
Label4.Caption = 2
newlog.SelColor = &H80000002
newlog.SelText = "!!! "
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Format(Now, "General Date")
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & " - Soft log is on"
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & " !!!" & vbCrLf
newlog.SelColor = &H80&
End Sub

Private Sub Command5_Click()
Open "Log.txt" For Output As #1
    Print #1, newlog.Text
Close #1
End Sub

Private Sub Command6_Click()
MsgBox "Programmed by meth0s" & vbCrLf & vbCrLf & "ICQ#4159440" & vbCrLf & "Mail:hackers8mymind@china.com" & vbCrLf & vbCrLf & "Thanks for listbox help fire", vbSystemModal, "About"
End Sub

Private Sub Command7_Click()
End
End Sub

Private Sub Form_Load()
Text3.ToolTipText = "Ports are opened in order of left to right. Seperate each port with the , key"
End Sub

Private Sub Form_Terminate()
'Making shure our program closes and doesnt hang in memory
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Making shure our program closes and doesnt hang in memory
End
End Sub

Private Sub Logtextfile_Change()
'When ever the log text box changes make shure it auto scrolls to the bottom line!
    On Error Resume Next
    Logtextfile.SelLength = 0
    If Len(Logtextfile.Text) > 0 Then
        If Right$(Logtextfile.Text, 1) = vbCrLf Then
            Logtextfile.SelStart = Len(Logtextfile.Text) - 1
            Exit Sub
        End If
        Logtextfile.SelStart = Len(Logtextfile.Text)
    End If
End Sub

Private Sub text3_KeyPress(KeyAscii As Integer)
    Dim Numbers As Integer
    Dim Msg As String
    Numbers = KeyAscii


    If (((Numbers < 48 Or Numbers > 57) And Numbers <> 8) And Numbers <> 44) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Timer1_Timer()
'Timer1 ones job is to update our label1!
Label1.Caption = Format(Now, "General Date") + " - " + Winsock(Index).LocalIP + " on " + Winsock(Index).LocalHostName ' so every 1/1000th's of a second it updates the label to the new time + the IP of your computer and the host name of your computer!
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
'timer2's job is to disconnect the intruder from yur computer 2 seconds after he has connected!
Winsock(Label3.Caption).Close 'Disconnecting the intruder from your computer
Winsock(Label3.Caption).Listen 'Listening for another intruder to connect to your computer!
If Check2.Value = vbChecked Then MsgBox Label5.Caption + " has been disconnected from your computer!", vbSystemModal, "ALERT!" 'If check2 is checked then send me a msgbox saying that he is being disconnected!
Timer2.Enabled = False 'Disabling this because there is no more intruder!
End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
On Error Resume Next
'This is when someone trie's to connect weather we allow them to or not and what happens when they try!
Winsock(Index).Close 'Under winsock contrl if someone wants to connect you must close the port to alow them!
Winsock(Index).Accept requestID 'Accepting there request to connect!
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & "!!! "
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & "Warning "
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Format(Now, "General Date")
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & " - "
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Winsock(Index).RemoteHostIP
'newlog.SelText = newlog.SelText & NameByAddr(Winsock(Index).RemoteHostIP)
'We arent using this yet.
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & ":"
newlog.SelColor = &H80&
newlog.SelText = newlog.SelText & Winsock(Index).Tag
newlog.SelColor = &H4040&
newlog.SelText = newlog.SelText & " Connection Attempted"
newlog.SelColor = &H80000002
newlog.SelText = newlog.SelText & " !!!" & vbCrLf

If Check1.Value = vbChecked Then Winsock(Index).SendData Text2.Text 'If check1 is checked then send a message to the intruder... the message is text2.text
If Check2.Value = vbChecked Then MsgBox "Computer " + Winsock(Index).RemoteHostIP + ":" + Winsock(Index).Tag + " has connected!" + vbCrLf + "Disconnecting them from your computer!", vbSystemModal, "ALERT!" 'If check2 is checked then send us a msgbox saying that the intruder *IP* has tried to connnect to us!
Timer2.Enabled = True 'Enabling timer2 so we can disconnect him from our computer!
Label3.Caption = Winsock.Item(Index).Index
Label5.Caption = Winsock(Index).RemoteHostIP
End Sub


Private Sub Winsock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'If the intruder trie's to send us any data
Dim incomingdata As String 'Dim the string as a string =)
If Label4.Caption = 1 Then
Winsock(Index).GetData incomingdata
Else
End If
End Sub
