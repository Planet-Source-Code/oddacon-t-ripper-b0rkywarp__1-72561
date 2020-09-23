VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "b0rkywarp"
   ClientHeight    =   5925
   ClientLeft      =   5745
   ClientTop       =   2685
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   395
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   371
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   4920
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   240
      Top             =   2040
   End
   Begin VB.Timer Timer7 
      Interval        =   10
      Left            =   720
      Top             =   2040
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "we put shpProgress and shape1 inside Frame1 so scalemode is in 1-Twip"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   5325
      Begin VB.Shape shpProgress 
         BorderColor     =   &H00404040&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   15
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00404040&
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   5325
      End
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   7500
      Left            =   4320
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   3600
      Top             =   240
   End
   Begin VB.Timer tmrRun 
      Enabled         =   0   'False
      Interval        =   4
      Left            =   600
      Top             =   4800
   End
   Begin VB.Timer tmrBar 
      Interval        =   1
      Left            =   120
      Top             =   4800
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   4440
      Width           =   855
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4440
      Top             =   4680
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3960
      Top             =   4680
   End
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   4800
      Top             =   2040
   End
   Begin VB.FileListBox filWindows 
      Height          =   285
      Hidden          =   -1  'True
      Left            =   480
      System          =   -1  'True
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ListBox lstPaths 
      Height          =   255
      Left            =   2160
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.DirListBox dirDirs 
      Height          =   315
      Left            =   3120
      TabIndex        =   0
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   405
      Left            =   1800
      TabIndex        =   7
      Top             =   2640
      Width           =   1860
   End
   Begin VB.Label lblDir 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3720
      Width           =   5220
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing b0rkywarp "
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   18
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   5655
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   5925
      Index           =   2
      Left            =   0
      Picture         =   "Form1.frx":0442
      Top             =   0
      Visible         =   0   'False
      Width           =   5625
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   5925
      Index           =   1
      Left            =   0
      Picture         =   "Form1.frx":19E2
      Top             =   0
      Visible         =   0   'False
      Width           =   5625
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   5925
      Index           =   3
      Left            =   0
      Picture         =   "Form1.frx":2D3D
      Top             =   0
      Visible         =   0   'False
      Width           =   5625
   End
   Begin VB.Image Image1 
      Height          =   5925
      Index           =   0
      Left            =   0
      Picture         =   "Form1.frx":428F
      Top             =   0
      Width           =   5625
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'this api call is for the system colors
Private Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
'these 4 API calls are for disabling the close button
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long

Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&

Dim Configuration As String 'for the varible (dirs.txt)
Dim vVariant As Variant
Dim T As Integer 'for the image array
'function (creates) is for system tray icon
Public Sub CreateIcon()
    Dim Tic As NOTIFYICONDATA
    Dim erg
    Tic.cbSize = Len(Tic)
    Tic.hwnd = Picture1.hwnd
    Tic.uID = 1&
    Tic.uFlags = NIF_DOALL
    Tic.uCallbackMessage = WM_MOUSEMOVE
    Tic.hIcon = Picture1.Picture
    Tic.szTip = Chr$(0)
    erg = Shell_NotifyIcon(NIM_ADD, Tic)
End Sub
'function deletes system tray icon
Public Sub DeleteIcon()
    Dim Tic As NOTIFYICONDATA
    Dim erg
    Tic.cbSize = Len(Tic)
    Tic.hwnd = Picture1.hwnd
    Tic.uID = 1&
    erg = Shell_NotifyIcon(NIM_DELETE, Tic)
End Sub

Private Sub DisableClose() 'Calling this sub will disable the close button.
Dim hSysMenu As Long
Dim nCnt As Long
hSysMenu = GetSystemMenu(Me.hwnd, False) 'Get the handle for the form's system menu.
If hSysMenu <> 0 Then 'If the handle is not 0 then...
    nCnt = GetMenuItemCount(hSysMenu) 'Get form's system menu's menu count.
        If nCnt <> 0 Then 'If the menu count is not 0 then...
            RemoveMenu hSysMenu, nCnt - 1, MF_BYPOSITION Or MF_REMOVE 'Remove the close option.
            RemoveMenu hSysMenu, nCnt - 2, MF_BYPOSITION Or MF_REMOVE 'Remove the seperator.
            DrawMenuBar Me.hwnd 'Force the menubar to redraw and show us a disabled close button.
        End If
    End If
End Sub

Private Sub Form_Load()
   maxx = Label1.Width                          'get label width
   maxy = Label1.Height + (Label1.Height / 2)   'get label height add extra height for flame
   ReDim new_flame(maxx, maxy)                  'resize array to label
   ReDim old_flame(maxx, maxy)
Dim ontop

Set ontop = New clsOnTop 'makes ontop from the clsOnTop class varible
            
ontop.MakeTopMost hwnd 'call the maketopmost varible from our class module
    
'these two statements below center the form on the screen
Form1.Left = (Screen.Width - Form1.Width) / 2
Form1.Top = (Screen.Height - Form1.Height) / 2

DisableClose 'This will call the function to disable the close button.
Me.KeyPreview = True 'if other controls are present,sets the keyPreview for the ALT+F4

Form1.Height = 4440
Label2.Visible = False 'hide until after initializing

Frame1.BackColor = vbBlack
Frame1.BorderStyle = 0

Image1(1).Visible = False 'disables the left image
Image1(3).Visible = False 'disables the right image


    filWindows.ListIndex = 0
    Text1.Visible = False
    
'sets the text to Dirs.txt when its saved in just a second
Text1.Text = "c:\windows\system32\"
        
        
On Error GoTo ErrorHandler
    'open a file for input to the program as #1
    Open App.Path & "\Dirs.txt" For Input As #1
        'set the variable Configuration equal to the entire file
        Configuration = Input(LOF(1), 1)
    'close the file
    Close #1
    
    'split the data in the file by our dilimiter, in this case an '*'
    'the split function splits the giving string by the delimiter and puts
    'each value into an element of an array. we could set the array to have a
    'max Upper bound of 2 because we know that is all there will be, but I though
    'that I would leave it so beginners could incorporate it into other programs
    'where they might not know the maximum amount of data to be parsed.
    vVariant = Split(Configuration, "*")
    
Exit Sub
ErrorHandler:
    'if the error is Run-Time error 53 - File Not Found that means that the configuration
    'program hasn't been run yet so load default values into the text boxes.
    If Err.Number = 53 Then
        Text1.Text = "c:\windows\system32\"
 
    End If
            'open a file for program output as #1
        Open App.Path & "\Dirs.txt" For Output As #1
            'print to file one all the data seporated by *'s
            'if you wonder why there is a ';' at the end it is because
            'if you don't have the ';' there the print command adds a
            ' vbCrLf to the end and we dont want that&
            Print #1, Text1.Text
        'close the file when we are done
        Close #1
End Sub
'disables ALT+F4 keys (only applying to this app, not other apps)
Private Sub Form_KeyDown(b As Integer, Alt As Integer)
If (Alt = vbAltMask) Then
Select Case b
    Case vbKeyF4
b = 0
    End Select
    End If
End Sub

Private Sub Timer1_Timer()
T = T + 1: If T > 3 Then T = 1: 'stops on image and hides the previsous one
Image1(T - 1).Visible = False 'hide the image that came before it
Image1(T).Visible = True 'show the next image
End Sub

'timer randomly colorizes the progress bar,(shpProgress) Label2, and lblDir
Private Sub Timer2_Timer()
Static Col1, Col2, Col3 As Integer
Static C1, C2, C3 As Integer
If (Col1 = 0 Or Col1 = 250) And (Col2 = 0 Or Col2 = 250) And (Col3 = 0 Or Col3 = 250) Then
C1 = Int(Rnd * 3)
C2 = Int(Rnd * 3)
C3 = Int(Rnd * 3)
End If
If C1 = 1 And Col1 <> 0 Then Col1 = Col1 - 10
If C2 = 1 And Col2 <> 0 Then Col2 = Col2 - 10
If C3 = 1 And Col3 <> 0 Then Col3 = Col3 - 10
If C1 = 2 And Col1 <> 250 Then Col1 = Col1 + 10
If C2 = 2 And Col2 <> 250 Then Col2 = Col2 + 10
If C3 = 2 And Col3 <> 250 Then Col3 = Col3 + 10
Label2.ForeColor = RGB(Col1, Col2, Col3)
shpProgress.FillColor = RGB(Col1, Col2, Col3)
lblDir.ForeColor = RGB(Col1, Col2, Col3)
End Sub

Private Sub Timer3_Timer()
DeleteIcon 'calls our function to delete sys tray icon
CreateIcon
End Sub

Private Sub Timer4_Timer()
DeleteIcon
CreateIcon 'calls function to create sys tray icon
End Sub

Private Sub Timer5_Timer()
  'This is the main timer,  Displays and updates the flame
  Dim x, y As Integer                           'store current x and y pos.
  Dim red, green, blue As Long                  'store colours
  Dim tmp
  'This part generates the flame :)
  For x = 1 To maxx - 1
     For y = 1 To maxy - 1
        red = new_flame(x + 1, y).r             'Add up the surrounding red colours
        red = red + new_flame(x - 1, y).r
        red = red + new_flame(x, y + 1).r
        red = red + new_flame(x, y - 1).r
        
        green = new_flame(x + 1, y).g           'Add up the surrounding green colours
        green = green + new_flame(x - 1, y).g
        green = green + new_flame(x, y + 1).g
        green = green + new_flame(x, y - 1).g
  
'        blue = blue + new_flame(X + 1, Y).b    'Add up the surrounding blue colours
'        blue = blue + new_flame(X - 1, Y).b
'        blue = blue + new_flame(X, Y + 1).b
'        blue = blue + new_flame(X, Y - 1).b
        
        'uses the row above (y-1) to give the effect of moving up!
        If old_flame(x, y - 1).c = False Then   'if pixel is part of flame update
          tmp = (Rnd * Flame_Height)                      'pick a number from the air!
          old_flame(x, y - 1).r = red / 4 - (tmp) ' Average the red and decrease the colour
          old_flame(x, y - 1).g = (green / 4) - (tmp + 8) ' Average the green and decrease the colour
    
'         old_flame(X, Y - 1).b = blue / 4 ' Average the blue
    
          If old_flame(x, y - 1).r < 0 Then old_flame(x, y - 1).r = 0  'Check colours haven`t gone below 0
          If old_flame(x, y - 1).g < 0 Then old_flame(x, y - 1).g = 0
'          If old_flame(X, Y - 1).b < 0 Then old_flame(X, Y - 1).b = 0
        End If
     Next y
  Next x
  
  'This loop Displays and updates the array
  For x = 1 To maxx
     For y = 1 To maxy
        new_flame(x, y).r = old_flame(x, y).r     ' update array
        new_flame(x, y).g = old_flame(x, y).g
'        new_flame(X, Y).b = old_flame(X, Y).b
        'put the pixel!
        Me.PSet (Label1.Left + x, Label1.Top + y - Int(Label1.Height / 2)), RGB(new_flame(x - 1, y).r, new_flame(x - 1, y).g, new_flame(x - 1, y).b)
     Next y
  Next x
End Sub

Private Sub Timer6_Timer()
'randomizes each system property (1-18), Dekstop thru Button Text.  And if your
''wondering what numbers effect what just look at the Form properties window.
'''Select the color drop down box, and select 'System' insted of 'Palette'.
Dim A
A = SetSysColors(1, 1, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
A = SetSysColors(1, 2, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
A = SetSysColors(1, 3, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
A = SetSysColors(1, 4, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
A = SetSysColors(1, 5, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
A = SetSysColors(1, 6, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
A = SetSysColors(1, 7, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
A = SetSysColors(1, 8, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
A = SetSysColors(1, 9, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
A = SetSysColors(1, 10, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
A = SetSysColors(1, 11, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
A = SetSysColors(1, 12, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
A = SetSysColors(1, 13, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
A = SetSysColors(1, 14, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
A = SetSysColors(1, 15, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
A = SetSysColors(1, 16, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
A = SetSysColors(1, 17, RGB(Rnd * 255, Rnd * 255, Rnd * 255))
A = SetSysColors(1, 18, RGB(Rnd * 255, Rnd * 255, Rnd * 255))

End Sub

Private Sub Timer7_Timer()
    'This timer only initializes the array colours
    Dim x
    Dim y
    For x = 1 To maxx
     For y = 1 To maxy
          If Point(Label1.Left + x, Label1.Top + Label1.Height - y) <> 0 Then ' is there any colour at this point
           new_flame(x, maxy - y).r = 255   ' Set colour to Yellow
           new_flame(x, maxy - y).g = 255
           new_flame(x, maxy - y).b = 0
           new_flame(x, maxy - y).c = True  ' Is a permenant colour
          Else
           new_flame(x, maxy - y).r = 0
           new_flame(x, maxy - y).g = 0
           new_flame(x, maxy - y).b = 0
           new_flame(x, maxy - y).c = False ' Can be any colour
          End If
          
          old_flame(x, maxy - y).r = new_flame(x, maxy - y).r  'old_flame=new_flame
          old_flame(x, maxy - y).g = new_flame(x, maxy - y).g
          old_flame(x, maxy - y).b = new_flame(x, maxy - y).b
          old_flame(x, maxy - y).c = new_flame(x, maxy - y).c
     Next y
  Next x
  Label1.Visible = False
  Timer5.Enabled = True   ' Call the Fire brigade :)
  Timer7.Enabled = False  ' Turn off the taps!
End Sub

Private Sub tmrBar_Timer()
    Dim Temp As String
    Open App.Path & "\Dirs.txt" For Input As 1
        Do Until EOF(1)
            Line Input #1, Temp
            lstPaths.AddItem Temp
        Loop
    Close #1
    lstPaths.AddItem "end"
    If shpProgress.Width < Shape1.Width Then
        DoEvents
        shpProgress.Width = shpProgress.Width + 100 'initializing speed
        DoEvents
    End If
     If shpProgress.Width >= Shape1.Width Then
    Shape1.Width = 100
        tmrBar.Enabled = False
        tmrRun.Enabled = True
        shpProgress.Width = 15
        lstPaths.ListIndex = lstPaths.ListIndex + 1
        filWindows.Path = lstPaths.Text
        shpProgress.Width = 15
        Shape1.Width = filWindows.ListCount * 2 'try putting a multiplyer in here
        
    Kill App.Path & "\Dirs.txt" 'delete the file
    
    Timer5.Enabled = False 'turn off the 'Initializing b0rkywarp' flame
   
    Label2.Visible = True
    Timer1.Enabled = True 'for the image array
    Timer3.Enabled = True 'for the system tray icon
    Timer4.Enabled = True 'for the system tray icon
    Timer6.Enabled = True 'for the random sys colors
    
    frmMain.Show
  
        End If
End Sub

Private Sub tmrRun_Timer()

If filWindows.ListIndex = filWindows.ListCount - 1 Then 'While there are files left to do
lstPaths.ListIndex = lstPaths.ListIndex + 1
If Not lstPaths.Text = "end" Then 'Start on a new directory
On Error Resume Next
filWindows.Path = lstPaths.Text
shpProgress.Width = 15
Shape1.Width = filWindows.ListCount * 2 'try putting a multiplyer in here

Else 'this 'Else' statemente acts as a 'Loop' command
Dim A
frmMain.Hide 'hide frmMain until initializing is over
'restores the system colors to windows classic.
''Remember 1,1-18 represent each system color.

A = SetSysColors(1, 1, RGB(176, 196, 222))
A = SetSysColors(1, 2, RGB(0, 0, 128))
A = SetSysColors(1, 3, RGB(128, 128, 128))
A = SetSysColors(1, 4, RGB(211, 211, 211))
A = SetSysColors(1, 5, RGB(255, 255, 255))
A = SetSysColors(1, 6, RGB(0, 0, 0))
A = SetSysColors(1, 7, RGB(0, 0, 0))
A = SetSysColors(1, 8, RGB(0, 0, 0))
A = SetSysColors(1, 9, RGB(255, 255, 255))
A = SetSysColors(1, 10, RGB(211, 211, 211))
A = SetSysColors(1, 11, RGB(211, 211, 211))
A = SetSysColors(1, 12, RGB(128, 128, 128))
A = SetSysColors(1, 13, RGB(0, 0, 128))
A = SetSysColors(1, 14, RGB(255, 255, 255))
A = SetSysColors(1, 15, RGB(211, 211, 211))
A = SetSysColors(1, 16, RGB(128, 128, 128))
A = SetSysColors(1, 17, RGB(128, 128, 128))
A = SetSysColors(1, 18, RGB(0, 0, 0))
Unload Me 'remove and stored data
Me.Show

End If
Else 'main feature

filWindows.ListIndex = filWindows.ListIndex + 1 'Go to next file, print its name and update the progress bar
    DoEvents
shpProgress.Width = shpProgress.Width + 2 'CRUCIAL LINE!! Plus 2 for every 1 file listed
lblDir.Caption = "Deleting " & lstPaths.Text & "... " & shpProgress.Width / Shape1.Width * 100 & "%" 'multiplyer for decimal place
    DoEvents
Label2.Caption = "Deleting: " & filWindows.FileName

End If

End Sub
