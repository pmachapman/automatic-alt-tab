VERSION 4.00
Begin VB.Form FormMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Automatic Alt Tab"
   ClientHeight    =   1005
   ClientLeft      =   2475
   ClientTop       =   5205
   ClientWidth     =   2310
   Height          =   1515
   Icon            =   "FormMain.frx":0000
   Left            =   2415
   LinkTopic       =   "FormMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1005
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   Top             =   4755
   Width           =   2430
   Begin VB.CommandButton CommandExit 
      Cancel          =   -1  'True
      Caption         =   "&Exit"
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton CommandStart 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox TextInterval 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Text            =   "30"
      Top             =   120
      Width           =   615
   End
   Begin VB.Timer TimerMain 
      Enabled         =   0   'False
      Left            =   0
      Top             =   0
   End
   Begin VB.Label LabelInterval 
      Caption         =   "&Interval (seconds):"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   145
      Width           =   1455
   End
End
Attribute VB_Name = "FormMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' Require variable declaration
Option Explicit

' Exit button click event handler
Private Sub CommandExit_Click()
    Unload Me
End Sub

' Start/Stop button click event handler
Private Sub CommandStart_Click()
    ' Take action depending on whether the
    ' timer is in progress
    If CommandStart.Caption = "&Start" Then
        ' Make sure the vale is valid
        If Val(TextInterval.Text) > 0 Then
            ' Set the inveral, and start the timer
            TimerMain.Interval = Val(TextInterval.Text) * 1000
            TimerMain.Enabled = True
            TextInterval.Enabled = False
            CommandStart.Caption = "&Stop"
        Else
            ' Show an error
            MsgBox "You must enter a valid number of seconds", vbOKOnly + vbExclamation, "Error Setting Interval"
            TextInterval.SetFocus
        End If
    Else
        ' Stop the timer
        TimerMain.Enabled = False
        TextInterval.Enabled = True
        CommandStart.Caption = "&Start"
    End If
End Sub

' Form load event handler
Private Sub Form_Load()
    ' Select the interval for editing
    TextInterval.SelStart = 0
    TextInterval.SelLength = Len(TextInterval.Text)
End Sub

' Form unload event handler
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

' Alt-Tab timer event handler
Private Sub TimerMain_Timer()
    ' Press Alt
    keybd_event VK_ALT, 0, 0, 0
    DoEvents

    ' Press Esc
    keybd_event VK_ESCAPE, 1, 0, 0
    DoEvents

    ' Release Alt
    keybd_event VK_ALT, 0, KEYEVENTF_KEYUP, 0
    DoEvents
    
End Sub
