VERSION 5.00
Begin VB.Form FrmFlip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   Caption         =   "Flip Album"
   ClientHeight    =   11745
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   18630
   LinkTopic       =   "Form1"
   ScaleHeight     =   783
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1242
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "FrmFlip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ****************************************************
' FlipPages   Animate a page or picture being flipped.
' Flip Perspective and sine wave -> Page Curve
' ****************************************************
' mailto: tmax_visiber@yahoo.com


Const Pi As Single = 3.141593
' sndPlaySound constant
Const SND_NOWAIT = &H2000        'don't wait if the driver is busy
Const SND_ASYNC = &H1            'Play asynchronously
Const SND_NODEFAULT = &H2        'silence not default, if sound not found
Const SND_MEMORY = &H4           'lpszSoundName points to a memory file
Const SND_LOOP = &H8             'loop the sound until next sndPlaySound
Const SND_NOSTOP = &H10          'don't stop any currently playing sound
Const SND_SYNC = &H0             'play synchronously (default)
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Const COLORONCOLOR = 3           '**IMPORTANT **  settting for StretchBlt
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" ( _
      ByVal hDC As Long, _
      ByVal x As Long, _
      ByVal Y As Long, _
      ByVal nWidth As Long, _
      ByVal nHeight As Long, _
      ByVal hSrcDC As Long, _
      ByVal XSrc As Long, _
      ByVal YSrc As Long, _
      ByVal nSrcWidth As Long, _
      ByVal nSrcHeight As Long, _
      ByVal dwRop As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Type PageL      'Page parameter
        Left As Long
        Top As Long
        Width As Long
        Height As Long
End Type

' variant use by FlipPage
Dim R As Long                       ' Radius of Page (measure from center page bottom to the top left or top right)
Dim radY As Long                    ' Y pos of radius
Dim dw As Long                      ' Increase in width
Dim StartX, StartY As Long          ' Start of X position
Dim StartHeight, EndHeight As Long  ' Start of Y position
Dim OutWidth As Long                ' Output width
Dim OutYOffset As Long              ' Output Y position offset from origin

'
Dim Page As PageL       ' Page Display
Dim AutoFlip As Boolean ' Auto flip pages
Dim Dx As Long          ' Differ in X direction -> for auto direction flippage
Dim AppPath  As String  ' Application photo's path
Dim CurrentPhoto%       ' Current photo display

' Runtime control object
Dim File1 As FileListBox
Dim Pic1 As PictureBox
Dim Pic2 As PictureBox

' ******************************
' DblClick to alternate AutoFlip
' ******************************
Private Sub Form_DblClick()
AutoFlip = Not AutoFlip
FlipPage
End Sub

' *********************************************
' Loading all the runtime controls & parameters
' *********************************************
Private Sub Form_Load()
'AppPath = App.Path & "\images\"         'Images folder
AppPath = App.Path & IIf(Right(App.Path, 1) <> "\", "\", "") & "images\"    'Images folder
Set File1 = Me.Controls.Add("VB.filelistbox", "File1")
File1.Path = AppPath
File1.Pattern = "*.jpg"
Set Pic1 = Me.Controls.Add("VB.PictureBox", "pic1")
Pic1.AutoRedraw = True
Pic1.AutoSize = True
Pic1.ScaleMode = 3
Set Pic2 = Me.Controls.Add("VB.PictureBox", "pic2")
Pic2.AutoRedraw = True
Pic2.AutoSize = True
Pic2.ScaleMode = 3
Pic1.Picture = LoadPicture(AppPath & File1.List(CurrentPhoto))
Pic2.Picture = LoadPicture(AppPath & File1.List(CurrentPhoto + 1))
AutoFlip = False
CurrentPhoto = 0
Me.Caption = "FlipPage -[" & File1.List(CurrentPhoto) & "]"
End Sub

' ***********************************
' drag mouse to flip Left < - > Right
' ***********************************
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And (x > Page.Left And x < Page.Left + Page.Width And Y > Page.Top And Y < Page.Top + Page.Height) Then Dx = x
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 1 And (x > Page.Left And x < Page.Left + Page.Width And Y > Page.Top And Y < Page.Top + Page.Height) Then
  If Dx - x > 0 Then
    R2L
  Else
    L2R
  End If
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
sndPlaySound vbNullString, SND_ASYNC
End
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.Cls
Me.Width = 3 / 2 * Me.Height
Page.Width = 900                  ' Preset to 900
Page.Height = 2 / 3 * Page.Width  ' for 4R photo (4" x 6" ) 2/3 ratio
Page.Left = (Me.ScaleWidth - Page.Width) / 2
Page.Top = (Me.ScaleHeight - Page.Height) / 2
SetStretchBltMode Me.hDC, COLORONCOLOR
StretchBlt Me.hDC, Page.Left, Page.Top, Page.Width, Page.Height, Pic1.hDC, 0, 0, Pic1.ScaleWidth, Pic1.ScaleHeight, vbSrcCopy
Me.Refresh
R = 5 / 6 * Page.Width            'R = Sqr(Page.Height ^2 + (Page.Width / 2) ^2)
End Sub

' ***********************
' Flip from Right to Left
' ***********************
Private Sub R2L()
Pic1.Picture = LoadPicture(AppPath & File1.List(CurrentPhoto))
CurrentPhoto = CurrentPhoto - 1
If CurrentPhoto < 0 Then CurrentPhoto = File1.ListCount - 1
Pic2.Picture = LoadPicture(AppPath & File1.List(CurrentPhoto))

Me.Caption = "FlipPage -[" & File1.List(CurrentPhoto) & "]"
sndPlaySound App.Path & "\flip.wav", SND_SYNC Or SND_NODEFAULT Or SND_NOWAIT

For dw = Page.Width To 0 + Page.Width / 50 Step -20
    If dw >= Page.Width / 2 Then
        Blting True
    Else
        Blting False
    End If
    Delay (30)
Next dw

Me.Cls
SetStretchBltMode Me.hDC, COLORONCOLOR
StretchBlt Me.hDC, Page.Left, Page.Top, Page.Width, Page.Height, Pic2.hDC, 0, 0, Pic2.ScaleWidth, Pic2.ScaleHeight, vbSrcCopy
Me.Refresh
End Sub

' ***********************
' Flip from Left to Right
' ***********************
Private Sub L2R()

Pic2.Picture = LoadPicture(AppPath & File1.List(CurrentPhoto))
CurrentPhoto = CurrentPhoto + 1
If CurrentPhoto > File1.ListCount - 1 Then CurrentPhoto = 0
Pic1.Picture = LoadPicture(AppPath & File1.List(CurrentPhoto))

Me.Caption = "FlipPage -[" & File1.List(CurrentPhoto) & "]"
sndPlaySound App.Path & "\flip.wav", SND_SYNC Or SND_NODEFAULT Or SND_NOWAIT

For dw = 0 To Page.Width - Page.Width / 50 Step 20
    If dw <= Page.Width / 2 Then
        Blting False
    Else
        Blting True
    End If
    Delay (30)
Next dw

Me.Cls
SetStretchBltMode Me.hDC, COLORONCOLOR
StretchBlt Me.hDC, Page.Left, Page.Top, Page.Width, Page.Height, Pic1.hDC, 0, 0, Pic1.ScaleWidth, Pic1.ScaleHeight, vbSrcCopy
Me.Refresh
End Sub

' **********************************
' Compute the parameter for FlipSBlt
' **********************************
Sub Blting(Reverse As Boolean)
    radY = (R - Page.Height) * Sin((dw / Page.Width) * Pi)    ' radY = Sqr(R ^2 - (Page.Width / 2 - i)^2) - Page.Height
    StartX = Page.Left + dw
    StartY = Page.Top - radY
    StartHeight = Page.Height
    EndHeight = Page.Height
    OutWidth = (Page.Width / 2) - dw
    OutYOffset = radY
    Me.Cls
    SetStretchBltMode Me.hDC, COLORONCOLOR
    StretchBlt Me.hDC, Page.Left, Page.Top, Page.Width / 2, Page.Height, Pic1.hDC, 0, 0, Pic1.ScaleWidth / 2, Pic1.ScaleHeight, vbSrcCopy
    StretchBlt Me.hDC, Page.Left + Page.Width / 2, Page.Top, Page.Width / 2, Page.Height, Pic2.hDC, Pic2.ScaleWidth / 2, 0, Pic2.ScaleWidth / 2, Pic2.ScaleHeight, vbSrcCopy
    If Not Reverse Then
      Call FlipSBlt(Me.hDC, StartX, StartY, OutWidth, StartHeight, EndHeight, OutYOffset, Pic2.hDC, Pic2.ScaleWidth / 2, Pic2.ScaleHeight, False)
    Else
      Call FlipSBlt(Me.hDC, StartX, StartY, OutWidth, StartHeight, EndHeight, OutYOffset, Pic1.hDC, Pic1.ScaleWidth / 2, Pic1.ScaleHeight, True)
   End If
   Me.Refresh
End Sub

' *******************************
' Perspective Blt with sine curve
' *******************************
Sub FlipSBlt(ByVal outDC As Long, ByVal outX As Long, ByVal outY As Long, _
    ByVal OutWidth As Long, ByVal outStartHeight As Long, ByVal outEndHeight As Long, _
    ByVal outYOff As Long, ByVal inDC As Long, ByVal inWidth As Long, ByVal inHeight As Long, Optional Reverse As Boolean = False)
Dim loopx As Long
Dim InterpPos As Single
Dim InterpH As Long
Dim StartLoop As Long
Dim EndLoop As Long
Dim rady1 As Long
If OutWidth = 0 Then Exit Sub
StartLoop = 0
EndLoop = OutWidth
If OutWidth < 0 Then
    StartLoop = OutWidth
    EndLoop = 0
End If
SetStretchBltMode outDC, COLORONCOLOR
For loopx = StartLoop To EndLoop
    InterpPos = loopx / OutWidth
    InterpH = InterpPos * (outEndHeight - outStartHeight)
    rady1 = outEndHeight / 20 * Sin((InterpPos) * 3.14159)
    If Not Reverse Then
      StretchBlt outDC, loopx + outX, outY + (InterpPos * outYOff) - rady1, 1, outStartHeight + InterpH, inDC, InterpPos * inWidth, 0, 1, inHeight, vbSrcCopy
    Else
      StretchBlt outDC, loopx + outX, outY + (InterpPos * outYOff) - rady1, 1, outStartHeight + InterpH, inDC, (2 - InterpPos) * inWidth, 0, 1, inHeight, vbSrcCopy
    End If
Next loopx

End Sub

' ********
' AutoFlip
' ********
Sub FlipPage()
Do While AutoFlip
    L2R
    Delay 1200
Loop
End Sub

' **********
' Time delay
' **********
Sub Delay(tSet As Long)
Dim tStart, tEnd As Long
tStart = GetTickCount
Do While tEnd < tSet
    tEnd = GetTickCount - tStart
    DoEvents
Loop
End Sub

