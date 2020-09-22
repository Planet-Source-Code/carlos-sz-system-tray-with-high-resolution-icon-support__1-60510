VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System tray XP"
   ClientHeight    =   2010
   ClientLeft      =   225
   ClientTop       =   435
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   3105
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.CommandButton Command3 
      Caption         =   "Tooltip"
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Balloon (no sound)"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Balloon"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Menu mSysPopup 
      Caption         =   "SysPopup"
      Visible         =   0   'False
      Begin VB.Menu mShow 
         Caption         =   "Show"
      End
      Begin VB.Menu mSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this API is necessary to make sure that menu will disappear if user clicks outside of it
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Sub Command1_Click()
    
    'How to show a balloon
    TrayBalloon Form1, "This is a custom balloon with sound!", "System tray XP", NIIF_INFO
    
End Sub

Private Sub Command2_Click()

    'How to show a silent balloon
    TrayBalloon Form1, "This is a custom balloon WITHOUT sound!", "System tray XP", NIIF_INFO Or NIIF_NOSOUND

End Sub

Private Sub Command3_Click()

    'How to change just the tooltip text
    TrayTip Form1, "Tray with high resolution icon!"

End Sub

Private Sub Form_Load()

    'Load the system tray feature
    TrayAddIcon Form1, App.Path & "\xptray.ico", "XP Tray"

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'All mouse events including balloon click
    Dim Result As Long
    Dim cEvent As Single
    cEvent = X / Screen.TwipsPerPixelX

    Select Case cEvent

    Case MouseMove
        'Debug.Print "MouseMove"
    Case LeftUp
        Debug.Print "Left Up"
    Case LeftDown
        Debug.Print "LeftDown"
        Form1.WindowState = 0
        Form1.Show
    Case LeftDbClick
        Debug.Print "LeftDbClick"
    Case MiddleUp
        Debug.Print "MiddleUp"
    Case MiddleDown
        Debug.Print "MiddleDown"
    Case MiddleDbClick
        Debug.Print "MiddleDbClick"
    Case RightUp
        Debug.Print "RightUp"
    Case RightDown
        Debug.Print "RightDown"
        'make sure that menu will disappear if user clicks outside of it
        Result = SetForegroundWindow(Me.hwnd)
        'now show it
        Me.PopupMenu Me.mSysPopup
    Case RightDbClick
        Debug.Print "RightDbClick"
    Case BalloonClick
        Debug.Print "Balloon Click"

    End Select

End Sub

Private Sub Form_Resize()

    'this is necessary to assure that the minimized window is hidden
    If Me.WindowState = vbMinimized Then Me.Hide

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Remove system tray
    TrayRemoveIcon

End Sub

Private Sub mShow_Click()
    
    Form1.WindowState = 0
    Form1.Show

End Sub

Private Sub mExit_Click()

    Unload Me

End Sub
