VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "THE CHEAPEST FTP CLIENT IN THE WORLD  :)"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton txt_GetList 
      Caption         =   "Get Remote Directory Listing"
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   960
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog dialog 
      Left            =   2280
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txt_dir 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Be sure to respect case, as some servers are case sensitive."
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox txt_Pass 
      Height          =   285
      Left            =   4200
      TabIndex        =   7
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txt_User 
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Top             =   360
      Width           =   1815
   End
   Begin VB.TextBox txt_Host 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton cmd_PutFileList 
      Caption         =   ">>"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2160
      Width           =   495
   End
   Begin VB.ListBox List1 
      Height          =   2985
      ItemData        =   "Form1.frx":0442
      Left            =   120
      List            =   "Form1.frx":0444
      TabIndex        =   3
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton cmd_Browse 
      Caption         =   "Browse"
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox txt_Commands 
      Height          =   3015
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0446
      Top             =   1920
      Width           =   3135
   End
   Begin VB.CommandButton cmd_RunCommand 
      Caption         =   "Run Command"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remote Directory :"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Login :"
      Height          =   255
      Left            =   2160
      TabIndex        =   14
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Password :"
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remote Host :"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "File Upload :"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Command Window :"
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   1680
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This is just a cheap and simple prog i put together in my line
' of experiments with the shell command and command line progs that
' ship with windows. This is not really the way to do it. But in
' some situations, the simplicity of this may come handy, as you
' dont need and special ocx'es or whatever. And ftp.exe exists on
' winNT just as well.
' You could even include the ftp.exe as its very small, but im not sure
' if it will work on win95. The great thing about ftp.exe, is that it
' can read all the commands out of a file, which here we create
' with VB.
' The big backdraw is that you dont get any error codes back, so you
' dont know if all went well or not, so you must really be sure that it
' will work.
' You can transfer entire directories and get entire direcories with
' ftp.exe
' type ftp -h at the command line to see some of the input it supports
' and ? or ? command once in the session to see the internal commands and a
' short definition. Then you'll just have to play around to see what they do.

' if you want to list the root directory then put ./ as directory
Dim countIT As Integer

Private Sub cmd_RunCommand_Click()

Dim CommandFile As Integer
CommandFile = FreeFile
Open App.Path & "\Commands.txt" For Binary As CommandFile
Put #CommandFile, , txt_Commands.Text
Close CommandFile
Shell ("command.com /c" & App.Path & "\RunCommands.bat"), vbHide
MsgBox "Commands written to file and ftp.exe now Executing the commands."
Do While FileLen(App.Path & "\OK.txt") < 10
Form1.MousePointer = vbHourglass
Loop
MsgBox "All command executed."
Shell ("command.com /c echo 0 > " & App.Path & "\Ok.txt"), vbHide
Shell ("command.com /c echo 0 > " & App.Path & "\Commands.txt"), vbHide
Shell ("command.com /c echo 0 > " & App.Path & "\List.txt"), vbHide
Form1.MousePointer = 0

End Sub

Private Sub cmd_Browse_Click()

dialog.InitDir = "."
dialog.ShowOpen
List1.AddItem (dialog.FileName)

End Sub

Private Sub cmd_PutFileList_Click()

Dim MyString As String
For x = 0 To List1.ListCount - 1
MyString = MyString & "put " & List1.List(x) & vbNewLine
Next x
txt_Commands.Text = "open " & txt_Host.Text & vbNewLine & txt_User.Text & vbNewLine & txt_Pass.Text & vbNewLine & "cd " & txt_dir.Text & vbNewLine & MyString & "bye"
List1.Clear
MsgBox "Commands for transfering files created, click the |Run Command| button to transfer the files."

End Sub

Private Sub txt_GetList_Click()

List1.Clear
Shell ("command.com /c echo 0 > " & App.Path & "\Ok.txt"), vbHide
Dim ListFile As Integer
' We set up the command file like above
txt_Commands.Text = "open " & txt_Host.Text & vbNewLine & txt_User.Text & vbNewLine & txt_Pass.Text & vbNewLine & "cd " & txt_dir.Text & vbNewLine & "mls *.*" & vbNewLine & "List.txt" & vbNewLine & "y" & vbNewLine & "bye"
MsgBox "Command has been created, click ok to get listing"
' We write the command file and tell ftp to execute it
ListFile = FreeFile
Open App.Path & "\Commands.txt" For Binary As ListFile
Put #ListFile, , txt_Commands.Text
Close ListFile
Shell ("command.com /c" & App.Path & "\RunCommands.bat"), vbHide
GetListing

End Sub

Private Sub GetListing()

Dim LListFile As Integer
Dim TmpText As String
LListFile = FreeFile
' Before i used a timer to wait a bit to make sure that the
' ftp.exe had finished its work, but sometimes it would not
' have enough time to finish its job and the list.txt would
' either be empty of OWNED by ftp.exe
' By using the .bat file i can know exactly when the prog
' has finished, because it wont write anything to the file
' BEFORE ftp has finished, so i simply loop untill the file
' is not 0 lenght anymore. I think there is a way to know
' if a app launched with the shell command has terminated
' but firstly i dont know how, and secondly, this quite simple
' to pull off by simlpy runnin a command and then testing the
' lenght of the file untill its not NULL anymore.
txt_Commands.Text = ""
Do While FileLen(App.Path & "\OK.txt") < 10
Form1.MousePointer = vbHourglass
Loop
If FileLen(App.Path & "\OK.txt") > 0 Then
Open App.Path & "\List.txt" For Input As LListFile
Do Until EOF(LListFile)
Input #LListFile, TmpText
TmpText = Replace(TmpText, "./", "")
List1.AddItem TmpText
Loop
Close LListFile
For x = 0 To List1.ListCount - 1 Step 2
TmpText = TmpText & "get " & List1.List(x) & vbNewLine
Next x
List1.Clear
MsgBox "Got the Directory Listing, now just click the |Run Command| button to get the files."
txt_Commands.Text = "open " & txt_Host.Text & vbNewLine & txt_User.Text & vbNewLine & txt_Pass.Text & vbNewLine & "cd " & txt_dir.Text & vbNewLine & TmpText & "bye"
End If
Form1.MousePointer = 0
Shell ("command.com /c echo 0 > " & App.Path & "\Ok.txt"), vbHide
Shell ("command.com /c echo 0 > " & App.Path & "\Commands.txt"), vbHide
Shell ("command.com /c echo 0 > " & App.Path & "\List.txt"), vbHide
End Sub
' Thats all :) Just some fun for messin around meaninglessly.
' But this might give you some ideas of what you could do with other
' progs that support command line parameters.
