VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Real Snow   Press ESC to QUIT"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9015
   ControlBox      =   0   'False
   ForeColor       =   &H000100FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   604
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   601
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3960
      Top             =   4320
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000100FF&
      Height          =   9000
      Left            =   0
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   600
      TabIndex        =   0
      Top             =   0
      Width           =   9000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Use DIB for fast GFX
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Type RGBQUAD
 rgbBlue As Byte
 rgbGreen As Byte
 rgbRed As Byte
 rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
 biSize           As Long
 biWidth          As Long
 biHeight         As Long
 biPlanes         As Integer
 biBitCount       As Integer
 biCompression    As Long
 biSizeImage      As Long
 biXPelsPerMeter  As Long
 biYPelsPerMeter  As Long
 biClrUsed        As Long
 biClrImportant   As Long
End Type

Private Type BITMAPINFO
 bmiHeader As BITMAPINFOHEADER
End Type

Private Const DIB_RGB_COLORS As Long = 0

'A type i use for the snow
Private Type Winter
 x As Long 'Coordiantes x
 y As Long 'Coordiantes x
 Freefall As Byte 'Is our snow falling
 NoMove As Byte 'No move sience ?
End Type

Dim snow() As Winter 'Hold the snow
Dim Buf() As RGBQUAD 'Hold our Picture
Dim Binfo       As BITMAPINFO   'The GetDIBits API needs some Infos

'How many molecules should this demo have
Private Const Molecules = 3000



Private Sub Form_Load()
 Dim i As Integer
 Dim x As Long
 Dim y As Long
 Dim ctr As Integer
 ReDim snow(Molecules)
 Dim Full As Integer
 
 On Error Resume Next
 Randomize
 Me.Show
Restart:
 'Clear Picture and draw our Testpic
 'Every pixel that shouldnt get deletet needs to have at least one blue
 Pic.Cls
 Pic.AutoRedraw = True
 Pic.Line (0, 580)-(600, 580), &H100FF, B '1 blue 0 green 256 red
 Pic.Line (80, 80)-(80, 100), &H100FF, B
 Pic.Line (80, 100)-(120, 100), &H100FF, B
 Pic.CurrentX = 300
 Pic.CurrentY = 200
 Pic.Print "Real Snow"
 Pic.Line (300, 290)-(450, 350), &H100FF
 Pic.Line (550, 280)-(430, 350), &H100FF
 Pic.Line (110, 150)-(400, 550), &H100FF

 'Create a buffer that holds our picture
 ReDim Buf(0 To Pic.ScaleWidth - 1, 0 To Pic.ScaleHeight - 1)
 'Set the infos for our apicall
 With Binfo.bmiHeader
 .biSize = 40
 .biWidth = Pic.ScaleWidth
 .biHeight = Pic.ScaleHeight
 .biPlanes = 1
 .biBitCount = 32
 .biCompression = 0
 .biClrUsed = 0
 .biClrImportant = 0
 .biSizeImage = Pic.ScaleWidth * Pic.ScaleHeight
 End With
 'Now get the Picture
 GetDIBits Pic.hdc, Pic.Image.Handle, 0, Binfo.bmiHeader.biHeight, Buf(0, 0), Binfo, DIB_RGB_COLORS

 'Calculate the random snow
 For i = 0 To Molecules
  Do
   x = Int(Rnd * 590) + 2
   
   'Some spezial for DIB
   'DIBs work from bottom to top
   'In our case the picture has a height of 600
   'so in DIB 600 is the top and 0 i the bottom
   y = Rnd * 579 + 20
  Loop Until Buf(x, y).rgbBlue = 0 'only take these coordinates if there is nothing under this pixel
  snow(i).x = x
  snow(i).y = y
 Next i


Do
For ctr = 0 To Molecules
 'get the coordinates / fastern than allways get it from snow(ctr)
 x = snow(ctr).x
 y = snow(ctr).y
 
 'Snow has not moved over 100 times so set a new snowflake
 If snow(ctr).NoMove > 100 Then
  Do
   x = Int(Rnd * 590) + 2
   y = Rnd * 20 + 579
   Full = Full + 1
   If Full = 1000 Then GoTo Restart 'we tried over 100 times but allways snow so restart
  Loop Until Buf(x, y).rgbBlue = 0
  Full = 0
 End If

 'clear actual pixel
 Buf(x, y).rgbBlue = 0
 Buf(x, y).rgbRed = 0
 Buf(x, y).rgbGreen = 0

 'Chek if there is nothing under our snow
 If Buf(x, y - 1).rgbBlue = 0 Then

  'Check if we reached the bottom
  If y > 1 Then y = y - 1
  'Now we are in freefall
  snow(ctr).Freefall = snow(ctr).Freefall + 1
  snow(ctr).NoMove = 0

  'If our whater has moved mor than one pixel lets move to a side
  If snow(ctr).Freefall = 4 Then
   snow(ctr).Freefall = 0 'Set back to 0
   i = CLng(Rnd) 'Randomize to get a direction

   If i = 1 Then '1= right
    'Move to the right if possible
    If x < 598 Then
     If Buf(x + 1, y).rgbBlue = 0 Then x = x + 1
    End If
   Else
    'Move to the left if possible
    If x > 1 Then
     If Buf(x - 1, y).rgbBlue = 0 Then x = x - 1
    End If
   End If

  End If


 Else 'There is something under us so move to the side
  i = CLng(Rnd)

  If i = 1 Then
   'right  but only lighter blue not white
   If x < 599 And y > 0 Then
    If Buf(x + 1, y - 1).rgbBlue = 0 And Buf(x + 1, y).rgbBlue = 0 Then
     x = x + 1
     y = y - 1
     snow(ctr).NoMove = 0
    Else
     snow(ctr).NoMove = snow(ctr).NoMove + 1
    End If
   End If
  Else
   'left
   If x > 0 And y > 0 Then
    If Buf(x - 1, y - 1).rgbBlue = 0 And Buf(x - 1, y).rgbBlue = 0 Then
     x = x - 1
     y = y - 1
     snow(ctr).NoMove = 0
    Else
     snow(ctr).NoMove = snow(ctr).NoMove + 1
    End If
   End If
  End If
 End If

 'Store coordinates
 snow(ctr).x = x
 snow(ctr).y = y
 'Set the new pixel
 Buf(x, y).rgbBlue = &HFF
 Buf(x, y).rgbGreen = &HFF
 Buf(x, y).rgbRed = &HFF

Next ctr


'Set new picture
SetDIBits Pic.hdc, Pic.Image.Handle, 0, Binfo.bmiHeader.biHeight, Buf(0, 0), Binfo, DIB_RGB_COLORS
Pic.Refresh
DoEvents
Loop

End Sub


Private Sub Pic_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
  Unload Me
  End
 End If
End Sub

