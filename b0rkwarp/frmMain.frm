VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   Caption         =   "b0rkywarp"
   ClientHeight    =   6210
   ClientLeft      =   4650
   ClientTop       =   2475
   ClientWidth     =   7935
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrmunchi 
      Interval        =   800
      Left            =   1800
      Top             =   2040
   End
   Begin VB.Image munchi 
      Height          =   495
      Index           =   11
      Left            =   7320
      Top             =   5520
      Width           =   495
   End
   Begin VB.Image munchi 
      Height          =   495
      Index           =   10
      Left            =   120
      Top             =   5520
      Width           =   495
   End
   Begin VB.Image munchi 
      Height          =   495
      Index           =   9
      Left            =   7320
      Top             =   4440
      Width           =   495
   End
   Begin VB.Image munchi 
      Height          =   495
      Index           =   8
      Left            =   120
      Top             =   4440
      Width           =   495
   End
   Begin VB.Image munchi 
      Height          =   495
      Index           =   7
      Left            =   7320
      Top             =   2280
      Width           =   495
   End
   Begin VB.Image munchi 
      Height          =   495
      Index           =   6
      Left            =   120
      Top             =   120
      Width           =   495
   End
   Begin VB.Image munchi 
      Height          =   495
      Index           =   0
      Left            =   7320
      Top             =   120
      Width           =   495
   End
   Begin VB.Image munchipic 
      Height          =   480
      Index           =   0
      Left            =   2400
      Picture         =   "frmMain.frx":0000
      Top             =   2760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image munchipic 
      Height          =   480
      Index           =   1
      Left            =   3120
      Picture         =   "frmMain.frx":030A
      Top             =   2760
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image munchi 
      Height          =   495
      Index           =   1
      Left            =   120
      Top             =   2280
      Width           =   495
   End
   Begin VB.Image munchi 
      Height          =   495
      Index           =   2
      Left            =   7320
      Top             =   3360
      Width           =   495
   End
   Begin VB.Image munchipic 
      Height          =   480
      Index           =   2
      Left            =   3120
      Picture         =   "frmMain.frx":0614
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image munchipic 
      Height          =   480
      Index           =   3
      Left            =   2400
      Picture         =   "frmMain.frx":091E
      Top             =   3360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image munchi 
      Height          =   495
      Index           =   3
      Left            =   120
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image munchi 
      Height          =   495
      Index           =   4
      Left            =   7320
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image munchi 
      Height          =   495
      Index           =   5
      Left            =   120
      Top             =   3360
      Width           =   495
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Private Type PALETTEENTRY
   peRed As Byte
   peGreen As Byte
   peBlue As Byte
   peFlags As Byte
End Type
Private Type LOGPALETTE
   palVersion As Integer
   palNumEntries As Integer
   palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
End Type
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104
Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
'API CALLS
Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "GDI32" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "GDI32" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "GDI32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SelectPalette Lib "GDI32" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "GDI32" (ByVal hDC As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long

Private Type PicBmp 'for munchi image
   Size As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type
Private lngDC As Long ' hDC of the screen, available to every sub/function, wich allows us to call ReleaseDC(0, lngDC) in cExit
Private blnLoop As Boolean
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
  Dim r As Long
   Dim Pic As PicBmp
   Dim IPic As IPicture
   Dim IID_IDispatch As GUID
   With IID_IDispatch
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With
   With Pic
      .Size = Len(Pic)          ' Length of structure.
      .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
      .hBmp = hBmp              ' Handle to bitmap.
      .hPal = hPal              ' Handle to palette (may be null).
   End With
   r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
   Set CreateBitmapPicture = IPic
End Function

Private Sub Form_Load()
Form1.Enabled = True

    Dim intX As Integer, intY As Integer
    Dim intI As Integer, intJ As Integer
    Dim intWidth As Integer, intHeight As Integer
    
    intWidth = Screen.Width / Screen.TwipsPerPixelX 'Screenwidth
    intHeight = Screen.Height / Screen.TwipsPerPixelY 'Screenheight
    
    frmMain.Width = Screen.Width  ' Set formwidth to screenwidth
    frmMain.Height = Screen.Height  ' Set formheight to screenheight
    
    Me.KeyPreview = True 'if other controls are present,sets the keyPreview for the ALT+F4.
    
    lngDC = GetDC(0) ' GetDC(0) to get the hDC of the screen
    
    Call BitBlt(hDC, 0, 0, intWidth, intHeight, lngDC, 0, 0, vbSrcCopy) ' BitBlt screen onto form
    frmMain.Visible = vbTrue ' Make form visible
    frmMain.AutoRedraw = vbFalse ' Set autoredraw to 0 (or your graphics-card might cause a reboot)
    
    Randomize
    
    blnLoop = vbTrue
    Do While blnLoop = vbTrue
        intX = (intWidth - 128) * Rnd
        intY = (intHeight - 128) * Rnd
        
        intI = 2 * Rnd - 1 ' Horizontal displacement
        intJ = 2 * Rnd - 1 ' Vertical displacement
        
        ' Move a part of the screen 1 pixel in a semi-random direction, to get the "melting" effect
        Call BitBlt(frmMain.hDC, intX + intI, intY + intJ, 128, 128, frmMain.hDC, intX, intY, vbSrcCopy)
        
        DoEvents
    Loop

    Set frmMain = Nothing ' Remove form from memory
    Call ReleaseDC(0, lngDC) ' Release the screen-hDC
    End
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
Private Sub form_click()
'set here for coding purppose's, remove statement for app to become l33t =)
Dim A
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
A = SetSysColors(1, 11, RGB(128, 128, 128))
A = SetSysColors(1, 12, RGB(211, 211, 211))
A = SetSysColors(1, 13, RGB(0, 0, 128))
A = SetSysColors(1, 14, RGB(255, 255, 255))
A = SetSysColors(1, 15, RGB(211, 211, 211))
A = SetSysColors(1, 16, RGB(128, 128, 128))
A = SetSysColors(1, 17, RGB(128, 128, 128))
A = SetSysColors(1, 18, RGB(0, 0, 0))
Unload Me 'removes any stored data
End
End Sub
Private Sub form_queryunload(Cancel As Integer, UnloadMode As Integer)
'if users cloes app by any means this restores the system color to 'windows classic'
'Remember each line represents a different system color property
Dim A
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
A = SetSysColors(1, 11, RGB(128, 128, 128))
A = SetSysColors(1, 12, RGB(211, 211, 211))
A = SetSysColors(1, 13, RGB(0, 0, 128))
A = SetSysColors(1, 14, RGB(255, 255, 255))
A = SetSysColors(1, 15, RGB(211, 211, 211))
A = SetSysColors(1, 16, RGB(128, 128, 128))
A = SetSysColors(1, 17, RGB(128, 128, 128))
A = SetSysColors(1, 18, RGB(0, 0, 0))
Unload Me
End Sub

'timer for random icons
Private Sub tmrmunchi_Timer()
Randomize
Dim T
For T = 0 To 2 ' three images
'makes the munchipic(0) (THE SKULL) icon appear first, then switch to munchipic(0) (THE TOUNGE) then loop
If munchi(T).Picture = munchipic(2).Picture Then munchi(T).Picture = munchipic(0).Picture Else munchi(T).Picture = munchipic(2).Picture
munchi(T).Left = munchi(T).Left + Val(Int(Rnd * 500) + 200) 'edits the speed, left to right of icons
If munchi(T).Left >= Screen.Width + munchi(T).Width Then
    munchi(T).Top = Int(Rnd * Screen.Height)
    munchi(T).Left = 0
End If
munchi(T).ZOrder
Next T
For T = 3 To 5 ' three images
'makes munchipic(3) (THE SMILEY FACE) appear first, then switchs to munchipic(1) (THE EYE) then loop
If munchi(T).Picture = munchipic(3).Picture Then munchi(T).Picture = munchipic(1).Picture Else munchi(T).Picture = munchipic(3).Picture
munchi(T).Left = munchi(T).Left - Val(Int(Rnd * 500) + 200) 'edits the speed, left to right of icons
If munchi(T).Left < Val(0 - 100 - munchi(T).Width) Then
    munchi(T).Top = Int(Rnd * Screen.Height)
    munchi(T).Left = Screen.Width
End If
munchi(T).ZOrder
Next T
For T = 6 To 8 'three images
'makes the munchipic(0) (THE TOUNGE) appear first, then switchs to munchipic(3) (THE SMILIY) then loop
If munchi(T).Picture = munchipic(0).Picture Then munchi(T).Picture = munchipic(3).Picture Else munchi(T).Picture = munchipic(0).Picture
munchi(T).Left = munchi(T).Left + Val(Int(Rnd * 500) + 200) 'edits the speed, left to right of icons.
If munchi(T).Left >= Screen.Width + munchi(T).Width Then
    munchi(T).Top = Int(Rnd * Screen.Height)
    munchi(T).Left = 0
End If
munchi(T).ZOrder
Next T
For T = 9 To 11 ' three images, total of 12, 0-11
'makes munchipic(1) (THE EYE) appear first, thne switchs to munchipic(2) (THE SKULL) then loop
If munchi(T).Picture = munchipic(1).Picture Then munchi(T).Picture = munchipic(2).Picture Else munchi(T).Picture = munchipic(1).Picture
munchi(T).Left = munchi(T).Left + Val(Int(Rnd * 500) + 200) 'edits the speed, left to right of icons
If munchi(T).Left >= Screen.Width + munchi(T).Width Then
    munchi(T).Top = Int(Rnd * Screen.Height)
    munchi(T).Left = 0
End If
munchi(T).ZOrder
Next T
End Sub
