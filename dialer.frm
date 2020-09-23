VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dialer -=HaCtOrInDp=-"
   ClientHeight    =   3720
   ClientLeft      =   3225
   ClientTop       =   3525
   ClientWidth     =   9645
   Icon            =   "dialer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command18 
      Caption         =   "Stop"
      Height          =   255
      Left            =   7560
      TabIndex        =   39
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Start"
      Height          =   255
      Left            =   5400
      TabIndex        =   38
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command16 
      Caption         =   "Stop"
      Height          =   255
      Left            =   2640
      TabIndex        =   37
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command15 
      Caption         =   "Start"
      Height          =   255
      Left            =   2640
      TabIndex        =   36
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   2280
      Top             =   3120
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   3720
      TabIndex        =   35
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   3720
      TabIndex        =   34
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   2040
      TabIndex        =   31
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2640
      TabIndex        =   29
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6600
      Top             =   240
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   5400
      TabIndex        =   26
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   7200
      TabIndex        =   24
      Text            =   "Yahoo! Mail Alert"
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   3000
      Top             =   120
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   3720
      TabIndex        =   20
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2040
      TabIndex        =   18
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Top             =   4800
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   720
      Top             =   4800
   End
   Begin VB.CommandButton Command14 
      Caption         =   "#"
      Height          =   495
      Left            =   1320
      TabIndex        =   14
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command13 
      Caption         =   "0"
      Height          =   495
      Left            =   720
      TabIndex        =   13
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command12 
      Caption         =   "*"
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   2400
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   720
      Top             =   4440
   End
   Begin VB.CommandButton Command11 
      Caption         =   "cancel"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Dial"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H0000FF00&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   720
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Yahoo! Mail (Yahoo! Messenger must be running)"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5160
      TabIndex        =   40
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "For this to work you must be connected to the internet, but not be using your modem"
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   5160
      TabIndex        =   41
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Interval (minutes)"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3720
      TabIndex        =   33
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Number to Call"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2040
      TabIndex        =   32
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Number to Call"
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   30
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   " Continuous Dialer"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2040
      TabIndex        =   28
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   5040
      X2              =   1920
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line4 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   1920
      X2              =   1920
      Y1              =   120
      Y2              =   3600
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   2
      X1              =   5040
      X2              =   5040
      Y1              =   120
      Y2              =   3600
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "No Mail"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   7320
      TabIndex        =   27
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Number to Call"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5400
      TabIndex        =   25
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail Caller"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5400
      TabIndex        =   23
      Top             =   120
      Width           =   1695
   End
   Begin VB.Line Line2 
      X1              =   5160
      X2              =   5160
      Y1              =   3360
      Y2              =   360
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   2040
      Y1              =   360
      Y2              =   3360
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   735
      Left            =   3960
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   2520
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "AM/PM all caps"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3720
      TabIndex        =   19
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Time, 12 Hour Format"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2040
      TabIndex        =   17
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Timed Dial"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2040
      TabIndex        =   16
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

DefInt A-Z

' This flag is set when the user chooses Cancel.
Dim CancelFlag
Private Sub Command1_Click()
Text1.Text = Text1.Text & "1"
End Sub

Private Sub Command15_Click()
Text10.Text = Text9.Text
Timer5.Enabled = True

End Sub

Private Sub Command16_Click()
Timer5.Enabled = False
End Sub

Private Sub Command17_Click()

If Option1 = True Then Timer4.Enabled = True

End Sub

Private Sub Command18_Click()

If Option1 = True Then Timer4.Enabled = False


End Sub




Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        ' Put your code here to emulate systray icons events. Note: If there is any
        ' code for MouseMove, MouseDown, or MouseUp then a Double Click won't be
        ' caught.
        ' Only uncomment the events that your app will use, so as to avoid any
        ' strange errors.
      If RunningInTray Then
        Select Case X
            'Case 7680   ' MouseMove
            'Case 7695   ' Left MouseDown
            'Case 7710   ' Left MouseUp
            Case 7725   ' Left DoubleClick
                Me.WindowState = vbNormal   ' Or vbMaximized if you feel like it.
                Me.Show
                RemoveIcon Me
            'Case 7740   ' Right MouseDown
            'Case 7755   ' Right MouseUp
            'Case 7770   ' Right DoubleClick
        End Select
      End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ' Remove the icon when this form unload. Don't forget to unload this form!
    RemoveIcon Me 'Add your form's name here for the sub to work.
End Sub
Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        ' This code hides the Form and puts the icon in the tray. Feel free to move
        ' it around if you like.
        Me.Hide
        ShowIcon Me
    End If
End Sub

Private Sub Command10_Click()

Dial Text1.Text


End Sub

Private Sub Command11_Click()
CancelFlag = True
End Sub

Private Sub Command12_Click()
Text1.Text = Text1.Text & "*"
End Sub

Private Sub Command13_Click()
Text1.Text = Text1.Text & "0"
End Sub

Private Sub Command14_Click()
Text1.Text = Text1.Text & "#"
End Sub

Private Sub Command2_Click()
Text1.Text = Text1.Text & "2"
End Sub

Private Sub Command3_Click()
Text1.Text = Text1.Text & "3"
End Sub

Private Sub Command4_Click()
Text1.Text = Text1.Text & "4"
End Sub

Private Sub Command5_Click()
Text1.Text = Text1.Text & "5"
End Sub

Private Sub Command6_Click()
Text1.Text = Text1.Text & "6"
End Sub

Private Sub Command7_Click()
Text1.Text = Text1.Text & "7"
End Sub

Private Sub Command8_Click()
Text1.Text = Text1.Text & "8"
End Sub

Private Sub Command9_Click()
Text1.Text = Text1.Text & "9"
End Sub

Private Sub Form_Load()
MSComm1.InputLen = 0
OnTop Me
End Sub
 Sub Dial(Number$)
 Dim DialString$, FromModem$, dummy

    
    DialString$ = "ATDT" + Number$ + ";" + vbCr

   
    MSComm1.CommPort = Text2.Text
    MSComm1.Settings = "9600,N,8,1"
    
   
    On Error Resume Next
    MSComm1.PortOpen = True
    If Err Then
       MsgBox "COM2: not available. Change the CommPort property to another port."
       Exit Sub
    End If
    
    
    MSComm1.InBufferCount = 0
    
   
    MSComm1.Output = DialString$
    
    ' Wait for "OK" to come back from the modem.
    Do
       dummy = DoEvents()
       ' If there is data in the buffer, then read it.
       If MSComm1.InBufferCount Then
          FromModem$ = FromModem$ + MSComm1.Input
          ' Check for "OK".
          If InStr(FromModem$, "OK") Then
             ' Notify the user to pick up the phone.
             Beep
          
             Exit Do
          End If
       End If
        
       ' Did the user choose Cancel?
       If CancelFlag Then
          CancelFlag = False
          Exit Do
       End If
    Loop
    
    ' Disconnect the modem.
    MSComm1.Output = "ATH" + vbCr
    
    ' Close the port.
    MSComm1.PortOpen = False
End Sub



Private Sub Label4_Change()
If Label4.Caption = Text3.Text And Label8.Caption = Text4.Text Then Dial Text7.Text
End Sub

Private Sub Timer1_Timer()
If Text1.Text = "" Then
Command10.Enabled = False
Else: Command10.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Dim port, X, instring
port = 1
PortinG:
MSComm1.CommPort = port
MSComm1.PortOpen = True


Form1.MSComm1.Settings = "9600,N,8,1"
    MSComm1.Output = "AT" + Chr$(13)
    X = 1


    Do: DoEvents
        X = X + 1
        If X = 1000 Then MSComm1.Output = "AT" + Chr$(13)
        If X = 2000 Then MSComm1.Output = "AT" + Chr$(13)
        If X = 3000 Then MSComm1.Output = "AT" + Chr$(13)
        If X = 4000 Then MSComm1.Output = "AT" + Chr$(13)
        If X = 5000 Then MSComm1.Output = "AT" + Chr$(13)
        If X = 6000 Then MSComm1.Output = "AT" + Chr$(13)


        If X = 7000 Then
            MSComm1.PortOpen = False
            port = port + 1
            GoTo PortinG:


            If MSComm1.CommPort >= 5 Then
errr:
                MsgBox "Can't Find Modem!"
                GoTo done:
            End If
        End If
    Loop Until MSComm1.InBufferCount >= 2
    instring = MSComm1.Input
    MSComm1.PortOpen = False

  Text2.Text = port



done:
Timer2.Enabled = False
End Sub

Private Sub Timer3_Timer()
Label4.Caption = Left(time, Len(time) - 6)
'Label7.Caption = Right(Time, 2)
Label8.Caption = Right(time, 2)
'Label8.Caption = Left(Time, 2)
End Sub


Private Sub Timer4_Timer()
'Dim x as long
Dim X&

'Find window api: pass a null string, text1.text has the win
'name we are looking for...
X = FindWindow(vbNullString, Text5.Text)

'if it was false then we didnt find a window
If X = 0 Then
    Label7.Caption = "No Mail"
Else
'we found the window
    Dial Text6.Text
    Label6.Caption = "Mail"
    End If
End Sub

Private Sub Timer5_Timer()
Text10.Text = (Text10.Text) - 1
If Text10.Text = "0" Then
  Dial Text8.Text
  Text10.Text = Text9.Text
End If

End Sub
