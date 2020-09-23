VERSION 5.00
Begin VB.Form frmWhichWord 
   Caption         =   "Which Word?"
   ClientHeight    =   6675
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9405
   Icon            =   "frmWhatWord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   445
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   627
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdResetHOF 
      Caption         =   "RESET Hall of Fame"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   225
      TabIndex        =   11
      Top             =   5220
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CommandButton cmdHOF 
      Cancel          =   -1  'True
      Caption         =   "CLOSE Hall of Fame Screen"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2565
      TabIndex        =   10
      Top             =   5220
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6135
      Top             =   4575
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   6015
      Left            =   6105
      TabIndex        =   3
      Top             =   135
      Width           =   3315
      Begin VB.ListBox lstWords 
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4920
         Left            =   0
         MultiSelect     =   1  'Simple
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   3210
      End
      Begin VB.ListBox lstScore 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   15
         MultiSelect     =   1  'Simple
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   4965
         Width           =   3210
      End
   End
   Begin VB.PictureBox picRounds 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1230
      Left            =   30
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6855
      Visible         =   0   'False
      Width           =   6030
   End
   Begin VB.CommandButton cmdBlast 
      Caption         =   "Blast !"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6135
      TabIndex        =   1
      Top             =   6180
      Width           =   1215
   End
   Begin VB.PictureBox picBoard 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   30
      ScaleHeight     =   398
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   398
      TabIndex        =   0
      Top             =   120
      Width           =   6000
      Begin VB.ListBox lstHOF 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4380
         ItemData        =   "frmWhatWord.frx":030A
         Left            =   135
         List            =   "frmWhatWord.frx":030C
         MultiSelect     =   1  'Simple
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   255
         Visible         =   0   'False
         Width           =   5640
      End
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      Caption         =   "LeveL:  1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   7425
      TabIndex        =   8
      Top             =   6240
      Width           =   1860
   End
   Begin VB.Label lblTimer 
      BackColor       =   &H000000C0&
      Height          =   285
      Left            =   60
      TabIndex        =   7
      Tag             =   "398"
      Top             =   6270
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label lblSBar 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      Height          =   420
      Left            =   30
      TabIndex        =   6
      Top             =   6195
      Width           =   6030
   End
   Begin VB.Menu mnuMain 
      Caption         =   "Game"
      Index           =   1
      Begin VB.Menu mnuGame 
         Caption         =   "&New"
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuGame 
         Caption         =   "&Time Limit"
         Index           =   1
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuGame 
         Caption         =   "&Sound"
         Checked         =   -1  'True
         Index           =   2
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuGame 
         Caption         =   "&Hall of Fame"
         Index           =   3
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuGame 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuGame 
         Caption         =   "Help"
         Index           =   5
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuGame 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuGame 
         Caption         =   "E&xit"
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmWhichWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" _
        (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Declare Function GetDialogBaseUnits Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function EqualRect Lib "user32" (lpRect1 As RECT, lpRect2 As RECT) As Long
Private Declare Function InvertRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function InvertRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Const RGN_DIFF As Long = 4
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type LOGFONT               ' used to create fonts
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * 32
End Type
Private Const DT_CALCRECT = &H400
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type OffScreenDC
    hTempDC As Long
    hOldBMP As Long
    hOldFont As Long
    hOldPen As Long
    hOldBrush As Long
End Type
Private SoundBuffer() As Byte
Private DrawDC As OffScreenDC
Private DropDC As OffScreenDC
Private rgnNoClick As Long

Private TimeLimit As Long
Private xLevel As Integer
Private lScore As Long
Private sBoard As String
Private WordCount(0 To 2) As Long
Private BonusWords() As String
Private nrBonus As Integer
Private LastTarget As Integer
Private sFile As String

Private Sub CheckPhrase(sPhrase() As String, iPos() As Integer, xPos() As Integer)

' Function checks for words when two letters are swapped or one letter dbl clicked
Dim sMsg As String, sWord As String, I As Integer
Dim xStop As Integer, xLen As Integer
Dim lRtn As Long, iPhrase As Integer, maxPos As Integer
Dim sChecked As String, sValid As String, iCount As Integer

For iPhrase = 1 To UBound(sPhrase)
    ' len can be zero if 1 letter was dbl clicked; otherwise should not be blank
    If Len(sPhrase(iPhrase)) Then
        maxPos = 0
        sWord = sPhrase(iPhrase)
        If Not SkipWord(sChecked, sValid, sWord, sPhrase(iPhrase)) Then
            lRtn = SearchFile(sWord)
            If lRtn Then
                iCount = 1
                sChecked = sChecked & "~" & sWord & "~"
                sValid = sValid & "~" & sWord & "~"
                iPos(iPhrase) = 1
                xPos(iPhrase) = 5
                lScore = lScore + lRtn
                For xLen = 3 To 4
                    If xLen <= Len(sWord) Then
                        xStop = Choose(xLen - 2, 3, 2, 1)
                        If Len(sWord) - 1 < xLen Then xStop = 1
                        For I = 1 To xStop
                            lRtn = SearchFile(Mid$(sWord, I, xLen))
                            If lRtn Then
                                lScore = lScore + lRtn
                                iCount = iCount + 1
                            End If
                        Next
                    End If
                Next
            Else
                Dim minPos As Integer
                minPos = iPos(iPhrase)
                ' 3-letter words first - backwards
                For xLen = iPos(iPhrase) - 1 To iPos(iPhrase) - 2 Step -1
                    If xLen < 1 Then Exit For
                    If xLen < 4 Then
                        If Not SkipWord(sChecked, sValid, Mid$(sWord, xLen, 3), sPhrase(iPhrase)) Then
                            lRtn = SearchFile(Mid$(sWord, xLen, 3))
                            If lRtn Then
                                iCount = iCount + 1
                                sValid = sValid & "~" & Mid$(sWord, xLen, 3) & "~"
                                sChecked = sChecked & "~" & sWord & "~"
                                lScore = lScore + lRtn
                                minPos = xLen
                                If xLen + 2 > maxPos Then maxPos = xLen + 2
                            End If
                        End If
                    End If
                Next
                For xLen = iPos(iPhrase) To iPos(iPhrase) + 3      ' forwards
                    If xLen > 3 Then Exit For
                    If Not SkipWord(sChecked, sValid, Mid$(sWord, xLen, 3), sPhrase(iPhrase)) Then
                        lRtn = SearchFile(Mid$(sWord, xLen, 3))
                        If lRtn Then
                            sValid = sValid & "~" & Mid$(sWord, xLen, 3) & "~"
                            sChecked = sChecked & "~" & sWord & "~"
                            iCount = iCount + 1
                            lScore = lScore + lRtn
                            If minPos = 0 Then minPos = xLen
                            If xLen + 2 > maxPos Then maxPos = xLen + 2
                        Else
                            Exit For
                        End If
                    End If
                Next
                ' 4-letter words
                If iPos(iPhrase) <> 5 Then
                    If Not SkipWord(sChecked, sValid, Mid$(sWord, 1, 4), sPhrase(iPhrase)) Then
                        lRtn = SearchFile(Mid$(sWord, 1, 4))
                        If lRtn Then
                            iCount = iCount + 1
                            sValid = sValid & "~" & Mid$(sWord, 1, 4) & "~"
                            sChecked = sChecked & "~" & sWord & "~"
                            lScore = lScore + lRtn
                            minPos = 1
                            If maxPos < 4 Then maxPos = 4
                        End If
                    End If
                End If
                If iPos(iPhrase) <> 1 Then
                    If Not SkipWord(sChecked, sValid, Mid$(sWord, 2, 4), sPhrase(iPhrase)) Then
                        lRtn = SearchFile(Mid$(sWord, 2, 4))
                        If lRtn Then
                            iCount = iCount + 1
                            sValid = sValid & "~" & Mid$(sWord, 2, 4) & "~"
                            sChecked = sChecked & "~" & sWord & "~"
                            lScore = lScore + lRtn
                            If minPos > 2 Then minPos = 2
                            maxPos = 5
                        End If
                    End If
                End If
                iPos(iPhrase) = minPos
                xPos(iPhrase) = maxPos
            End If
        End If
    End If
Next
If iCount Then  ' made at least one word
    ' highlight the words that were made
    For I = 1 To lstWords.ListCount - 1
        lstWords.Selected(I) = False
    Next
    For I = 1 To iCount
        lstWords.Selected(I) = True
    Next
End If
End Sub

Private Function SkipWord(sChecked As String, sValid As String, sWord As String, sPhrase As String) As Boolean
' checks to see if the word made was already scored
If InStr(sValid, "~" & sWord & "~") Then
    If InStr(sChecked, "~" & sPhrase & "~") Then SkipWord = True
End If

End Function

Private Function ReturnScore(sWord As String) As Long
Dim I As Integer, J As Integer, K As Integer, sBWord As String
Dim tScore As Long, tBonus As Long, tTotal As Long
' using the letter values for the game Scrabble - 1 so I can fit in a simple single-digit string
' (i.e., the letter Z is worth 10, but here I'll make it 9 & add one more later)
Const sValue As String = "02210313174020029000033739"
For I = 1 To Len(sWord)
    tScore = tScore + (Mid(sValue, Asc(UCase(Mid$(sWord, I, 1))) - 64, 1) + 1) * 10
    ' special value (gold letter words)
    If Mid$(sWord, I, 1) = UCase(Mid$(sWord, I, 1)) Then tBonus = tBonus + (Mid(sValue, Asc(UCase(Mid$(sWord, I, 1))) - 64, 1) + 1) * 100
Next
' now see if the word made was one of the WhichWords?
For I = 0 To UBound(BonusWords)
    If UCase(sWord) = UCase(BonusWords(I)) Then
        For J = 1 To 2
            sBWord = vbTab & lstScore.List(J)
            K = InStr(sBWord, UCase(sWord) & vbTab)
            If K Then
                If Mid$(sBWord, K - 1, 1) = vbTab Then
                    ' matched a WhichWord?
                    BonusWords(I) = ""
                    tBonus = tBonus * 100
                    If tBonus = 0 Then tBonus = 10000 * Len(sWord)
                    Mid$(sBWord, K, Len(sWord)) = Space$(Len(sWord))
                    lstScore.List(J) = Mid$(sBWord, 2)
                    Exit For
                End If
            End If
        Next
        K = 0
        For J = 0 To 3
           K = K + Len(BonusWords(J))
        Next
        If K = 0 Then Tag = "Game Over"
    End If
Next
' update the score & add the made word to the list of made words
lstWords.AddItem UCase(sWord) & vbTab & tScore & vbTab & tBonus, 1
ReturnScore = tScore + tBonus
End Function

Private Function SearchFile(sWord As String) As Long
' function tweaked from http://www.edu.gov.nf.ca/curriculum/teched/resources/vbasic/tips.htm

If sWord = "" Then Exit Function
Dim sSearch As String
sSearch = vbCrLf & LCase(sWord) & vbCrLf

Const iByte As Boolean = True
Const iUniCode As Boolean = False

Dim iHandle As Integer
Dim sTemp As String
Dim lSpot As Long
Dim lFind As Long
Dim sSearch1 As String
Dim bTemp() As Byte
'another advantage of using a byte array
'is that we can easily look for UniCode strings
If iUniCode Or (Not iByte) Then
    'this line will look for unicode strings
    'when using byte arrays, regular
    'strings when using string variable
    sSearch1 = sSearch
Else
    'this line will look for ANSII strings
    'when looking through a byte array
    sSearch1 = StrConv(sSearch, vbFromUnicode)
End If
iHandle = FreeFile
Open sFile For Binary Access Read As iHandle
If iHandle Then
    sTemp = Space$((LOF(iHandle) / 2) + 1)
    ReDim bTemp(LOF(iHandle)) As Byte
    If iByte Then
        Get #iHandle, , bTemp
        sTemp = bTemp
    Else
        Get #iHandle, , sTemp
    End If
    Close iHandle
End If
Do
    If iByte Then
        lFind = InStrB(lSpot + 1, sTemp, sSearch1, 1)
    Else
        lFind = InStr(lSpot + 1, sTemp, sSearch1, 1)
    End If
    lSpot = lFind
Loop Until lFind = 0 Or lSpot > 0
If lSpot Then
    'Debug.Print LCase(sWord)
    SearchFile = ReturnScore(sWord)
End If
End Function

Private Sub StartGame(Optional bReset As Boolean = True)
Dim I As Integer, J As Integer, xLtr As String, iCircle As Integer
Dim bBonus As Boolean

lblSBar.Caption = "Preparing to run game...."
lblSBar.Refresh

' reset the progress bar
lblTimer.Width = lblTimer.Tag
lblTimer.Visible = True
If bReset Then
    ' if bReset=True then starting a fresh game; otherwise starting a new level
    lScore = 0
    xLevel = 1
    lstScore.List(3) = "Score" & vbTab & "0"
    lstWords.Tag = 0
Else
    xLevel = xLevel + 1
    lstWords.Tag = Val(lstWords.Tag) + lstWords.ListCount - 1
End If
' update the Level label
lblLevel.Caption = "Level:  " & xLevel

GetBonusWords
lstWords.Clear
lstWords.AddItem "Word" & vbTab & "Score" & vbTab & "Bonus"

nrBonus = 0         ' reset nr of bonus (gold) letters
sBoard = ""         ' reset the board
LastTarget = 50     ' flag indicating no selected letters
' now build the board with random letters using offscreen DC
For I = 0 To 4
    For J = 0 To 4
        xLtr = GetLetter
        If UCase(xLtr) = xLtr Then iCircle = 1 Else iCircle = 0
        sBoard = sBoard & xLtr
        DrawLetter xLtr, CLng(J), CLng(I), iCircle, , -1
    Next
Next
Tag = ""
BitBlt picBoard.hdc, 0, 0, picBoard.Width, picBoard.Height, DrawDC.hTempDC, 0, 0, vbSrcCopy
picBoard.Enabled = True
picBoard.Refresh
Timer1.Enabled = True
cmdBlast.Enabled = True
End Sub

Private Function GetLetter() As String
Dim sSource As String, iRnd As Integer

' just one way of doing things. This routine returns a random letter

iRnd = Int(Rnd * 100) + 1
If iRnd Mod 3 = 0 Or iRnd Mod 7 = 0 Then
    ' this seems to give a fair amount of vowels
    sSource = "AEIOUAEIO"
Else
    ' broke this up into common, less common & least common letters
    Select Case iRnd
    Case Is < 30
        sSource = "WPKFHCBWBCHFKPWH"
    Case Is < 80
        sSource = "NRSDGLMDNRSTNLR"
    Case Else
        sSource = "VJQXYZV"
    End Select
End If
If iRnd Mod 5 = 0 Then
    ' tweak to increase odds of getting a needed letter
    For iRnd = 0 To 3
        sSource = sSource & BonusWords(iRnd)
    Next
End If
' now get a random letter from the string
sSource = LCase(Mid$(sSource, Int(Rnd * Len(sSource)) + 1, 1))
If nrBonus < 5 Then
    ' maximum of 5 bonus letters per screen
    If Int(Rnd * 100) Mod 7 = 0 Then
        nrBonus = nrBonus + 1
        sSource = UCase(sSource)
    End If
End If
GetLetter = sSource
End Function

Private Function DrawLetter(xLtr As String, x As Long, y As Long, iCircle As Integer, Optional UCaseOnly As Boolean = True, Optional iMode As Integer)

' Draws a passed letter to a DC

Dim tRect As RECT, sLtr As String, destDC As Long
If UCaseOnly = True Then
    sLtr = UCase(xLtr)
    destDC = DrawDC.hTempDC
Else
    ' this is only used when displaying the splash screen
    sLtr = xLtr
    destDC = picBoard.hdc
End If
    
    Select Case iMode
        ' -1 draws only on the offscreen DC
        Case 0, -1
            BitBlt destDC, x * 80, y * 80, 80, 80, picRounds.hdc, 80 * iCircle, 0, vbSrcCopy
            DrawText destDC, sLtr, 1, tRect, DT_CALCRECT
            OffsetRect tRect, (80 - tRect.Right) \ 2 + x * 80, (80 - tRect.Bottom) \ 2 + y * 80
            DrawText destDC, sLtr, 1, tRect, ByVal 0&
        Case Is < 6 ' Shrink
            Rectangle DrawDC.hTempDC, x * 80, y * 80, (x + 1) * 80, (y + 1) * 80
            StretchBlt destDC, x * 80 + (iMode * 10), y * 80 + (iMode * 10), 80 - iMode * 20, 80 - (iMode * 20), picBoard.hdc, x * 80, y * 80, 80, 80, vbSrcCopy
        Case Is > 5 ' expand (not used at this time)
            Rectangle DrawDC.hTempDC, x * 80, y * 80, (x + 1) * 80, (y + 1) * 80
            StretchBlt destDC, x * 80 + ((iMode - 4) * 10), y * 80 + ((iMode - 4) * 10), 80 - iMode * 20, 80 - (iMode * 20), DropDC.hTempDC, 0, 0, 80, 80, vbSrcCopy
    End Select

If destDC = DrawDC.hTempDC And iCircle > -1 And iMode = 0 Then BitBlt picBoard.hdc, 0, 0, picBoard.Width, picBoard.Height, DrawDC.hTempDC, 0, 0, vbSrcCopy
End Function

Private Function ToggleLetter(x As Long, y As Long, bDblClick As Boolean) As Integer

' Basically unselect a letter if the same letter selected twice
' Or find words when 2 different letters are selected

picBoard.Enabled = False
Timer1.Enabled = False

Dim tTarget As Integer, bBonus As Boolean, iCircle As Integer, I As Integer
' calculate the letter being clicked upon
x = Int(x / 80)
y = Int(y / 80)
tTarget = y * 5 + x
' last target = 50 if no previous letter was selected
If LastTarget < 50 Then ' previously selected a letter
    bBonus = UCase(Mid$(sBoard, LastTarget + 1, 1)) = Mid$(sBoard, LastTarget + 1, 1)
    If bBonus Then iCircle = 1 Else iCircle = 0
    DrawLetter Mid$(sBoard, LastTarget + 1, 1), (LastTarget - (Int(LastTarget / 5) * 5)), Int(LastTarget / 5), iCircle
    If LastTarget <> tTarget Then
        ' 2 different letters being clicked
        Dim cX As Long, cY As Long
        cX = (LastTarget - (Int(LastTarget / 5) * 5))
        cY = Int(LastTarget / 5)
        ' animate the 2 letters as fading out
         For I = 1 To 4
             DrawLetter Mid$(sBoard, LastTarget + 1, 1), cX, cY, iCircle, , I
             DrawLetter Mid$(sBoard, tTarget + 1, 1), x, y, Abs(UCase(Mid$(sBoard, LastTarget + 1, 1)) = Mid$(sBoard, LastTarget + 1, 1)), , I
             BitBlt picBoard.hdc, 0, 0, picBoard.Width, picBoard.Height, DrawDC.hTempDC, 0, 0, vbSrcCopy
             picBoard.Refresh
             Sleep 70
        Next
        ' now show the swapped letters
        DrawLetter Mid$(sBoard, LastTarget + 1, 1), x, y, iCircle, , -1
    End If
End If

If LastTarget = tTarget And bDblClick = False Then
    ' reset flag
    LastTarget = 50
Else
    ' here we start
    bBonus = UCase(Mid$(sBoard, tTarget + 1, 1)) = Mid$(sBoard, tTarget + 1, 1)
    If bBonus Then iCircle = 4 Else iCircle = 3
    If LastTarget < 50 Or bDblClick = True Then
        ' show the last letter selected as unselected
        iCircle = iCircle - 3
        DrawLetter Mid$(sBoard, tTarget + 1, 1), (LastTarget - (Int(LastTarget / 5) * 5)), Int(LastTarget / 5), iCircle
        ' flag indicating a double click
        If bDblClick Then LastTarget = 50
        ' call function to do all the other work
        ShowResults LastTarget, tTarget, x, y
        LastTarget = 50 ' reset flag
        If Len(Tag) Then ' GAME OVER
            If MsgBox("Level " & xLevel & " Completed. Continue to next level?", vbQuestion + vbYesNo, "Congratulations") = vbYes Then
                StartGame False
            Else
                Tag = "Game Over"
                TallyHOFinfo
            End If
        End If
    Else
        ' show letter as selected
        DrawLetter Mid$(sBoard, tTarget + 1, 1), x, y, iCircle
        LastTarget = tTarget
    End If
    
End If
picBoard.Refresh
picBoard.Enabled = (Len(Tag) = 0)
Timer1.Enabled = picBoard.Enabled
End Function

Private Sub ShowResults(LastTarget As Integer, tTarget As Integer, x As Long, y As Long)
        
' main routine which does most of the graphics
        
Dim sSwap As String, K As Integer, I As Integer
Dim xPos(1 To 4) As Integer, iPos(1 To 4) As Integer
Dim J As Integer, iLtr As Integer, iCircle As Integer
Dim curPos As Integer, tRect As RECT
Dim rCircles() As RECT, hCircles() As Long, listLtrs() As Integer
Dim sDrops(1 To 5) As String, sPhrase(1 To 6) As String

If LastTarget < 50 Then
    'not a dobule click
    ' we swap the 2 selected letters in the string
    sSwap = Mid$(sBoard, LastTarget + 1, 1)
    Mid$(sBoard, LastTarget + 1, 1) = Mid$(sBoard, tTarget + 1, 1)
    Mid$(sBoard, tTarget + 1, 1) = sSwap
    
    ' now we build 4 strings for the 2 letters selected
    ' each letter has 2 strings; & this is for the 2nd selected letter
    '   1 string for vertical letters at the selected position
    ' & 1 string for horizontal letters at the selected position
    iPos(1) = (LastTarget - (Int(LastTarget / 5) * 5)) + 1
    iPos(3) = Int(LastTarget / 5) + 1
    sPhrase(1) = Mid$(sBoard, Int(LastTarget / 5) * 5 + 1, 5)
    For I = LastTarget - (Int(LastTarget / 5) * 5) + 1 To 25 Step 5
        sPhrase(3) = sPhrase(3) & Mid$(sBoard, I, 1)
    Next
End If
' build the 2 strings for the 1st selected letter
iPos(2) = x + 1
iPos(4) = y + 1
sPhrase(2) = Mid$(sBoard, Int(tTarget / 5) * 5 + 1, 5)
For I = tTarget - (Int(tTarget / 5) * 5) + 1 To 25 Step 5
    sPhrase(4) = sPhrase(4) & Mid$(sBoard, I, 1)
Next

' call function to search the dictionary for possible words
CheckPhrase sPhrase(), iPos(), xPos()

' now we want to identify those letters that formed words
ReDim rCircles(0)
For I = 1 To 4
   If xPos(I) Then
        ' all of this is to identify the letters in the words made
        ' iPos(x) is the starting letter of the word
        ' xPos(x) is the ending letter of the word
        Select Case I
            Case 1, 3: J = LastTarget + 1
            Case 2, 4: J = tTarget + 1
        End Select
        If I < 3 Then
            Do Until (J - 1) Mod 5 = 0
                J = J - 1
            Loop
        Else
            Do Until J < 6
                J = J - 5
            Loop
        End If
        ' for each of the letters in the word...
        For iLtr = iPos(I) - 1 To xPos(I) - 1
            If I < 3 Then
                curPos = J + iLtr - 1
            Else
                curPos = J - 1 + (iLtr * 5)
            End If
            ' build a rectangle surrounding that letter
            SetRect tRect, (curPos - (Int(curPos / 5) * 5)) * 80, Int(curPos / 5) * 80, (curPos - (Int(curPos / 5) * 5)) * 80 + 80, Int(curPos / 5) * 80 + 80
            ' now compare it to other built rectangles to see if it is a
            ' duplicate space being identified
            For K = 1 To UBound(rCircles)
                If EqualRect(tRect, rCircles(K)) Then Exit For
            Next
            If K > UBound(rCircles) Then
                ' not duplicated, so we add the rectangle to the list &
                ' also add a few other things to the list
                ReDim Preserve listLtrs(0 To UBound(rCircles) + 1)
                ReDim Preserve rCircles(0 To UBound(rCircles) + 1)
                rCircles(UBound(rCircles)) = tRect
                ' draw the letter using the red ball
                DrawLetter Mid$(sBoard, curPos + 1, 1), (curPos - (Int(curPos / 5) * 5)), Int(curPos / 5), 2, , -1
                ' keep track of which letter is being used
                listLtrs(UBound(listLtrs)) = curPos + 1
                ' if it was a bonus letter, reduce count of bonus letters
                If UCase(Mid$(sBoard, curPos + 1, 1)) = Mid$(sBoard, curPos + 1, 1) Then nrBonus = nrBonus - 1
            End If
        Next
    End If
Next
' after all that was done, now we flash the words made
If UBound(rCircles) Then
    ' update the score
    lstScore.List(3) = "Score" & vbTab & Format(lScore, "#,###")
    ' flash the letters that created words & will be removed
    ReDim hCircles(UBound(rCircles))
    ' build circular regions for each letter being flashed
    For I = 1 To UBound(rCircles)
         hCircles(I) = CreateEllipticRgn(rCircles(I).Left + 3, rCircles(I).Top + 3, rCircles(I).Right - 5, rCircles(I).Bottom - 5)
    Next
    For J = 1 To 6      ' now flash them
        BitBlt picBoard.hdc, 0, 0, picBoard.Width, picBoard.Height, DrawDC.hTempDC, 0, 0, vbSrcCopy
        picBoard.Refresh
        Sleep 100
        For I = 1 To UBound(hCircles)
            InvertRgn DrawDC.hTempDC, hCircles(I)
        Next
    Next
    ' refresh the screen & delete the regions we created
    BitBlt picBoard.hdc, 0, 0, picBoard.Width, picBoard.Height, DrawDC.hTempDC, 0, 0, vbSrcCopy
    For I = 1 To UBound(hCircles)
        DeleteObject hCircles(I)
    Next
    ' show the selected letters as fading out
    For I = 1 To 4
        For J = 1 To UBound(rCircles)
            DrawLetter Mid$(sBoard, listLtrs(J), 1), rCircles(J).Left \ 80, rCircles(J).Top \ 80, 2, , I
        Next
        BitBlt picBoard.hdc, 0, 0, picBoard.Width, picBoard.Height, DrawDC.hTempDC, 0, 0, vbSrcCopy
        picBoard.Refresh
        Sleep 70
    Next
    Erase sPhrase
    Erase iPos
    Erase xPos
    ' replace the vanished letters with spaces as flags for new letters
    For I = 1 To UBound(listLtrs)
        Mid$(sBoard, listLtrs(I), 1) = " "
    Next
    ' build 5 separate strings to be used for rearranging the letters
    For I = 1 To 5
        For J = I To 25 Step 5
            sPhrase(I) = sPhrase(I) & Mid$(sBoard, J, 1)
        Next
        ' string used for the ball-dropping animation
        sDrops(I) = sPhrase(I)
    Next
    sBoard = Space$(25)     ' clear the board
    For I = 1 To 5
        ' replace the vanished letters with random letters
        sPhrase(6) = Replace$(sPhrase(I), " ", "")
        For J = 1 To 5 - Len(sPhrase(6))
            sPhrase(6) = GetLetter & sPhrase(6)
        Next
        ' update the offscreen DC with the new board configuration
        For J = 1 To 5
            If UCase(Mid$(sPhrase(6), J, 1)) = Mid$(sPhrase(6), J, 1) Then iCircle = 1 Else iCircle = 0
            DrawLetter Mid$(sPhrase(6), J, 1), I - 1, J - 1, iCircle, , -1
            Mid$(sBoard, (J - 1) * 5 + I, 1) = Mid$(sPhrase(6), J, 1)
        Next
    Next
    ' now show the animated dropping of the balls
    DropBalls sDrops(), True
    BitBlt picBoard.hdc, 0, 0, picBoard.Width, picBoard.Height, DrawDC.hTempDC, 0, 0, vbSrcCopy
    picBoard.Refresh
    
End If
Erase listLtrs
Erase hCircles
Erase rCircles
Erase iPos
Erase xPos
Erase sPhrase

End Sub

Private Sub cmdBlast_Click()

' function used to get new random letters

Dim tRect As RECT, rCircles() As RECT
Dim I As Integer, J As Integer, tRemove As Integer
Dim x As Long, y As Long, iCircle As Integer
Dim sPhrase(1 To 6) As String, sDrop(1 To 5) As String

BeginPlaySound 103
ReDim rCircles(0)
' set limit of new letters to get (between 5 & 10)
For I = 1 To Int(Rnd * 5) + 5
    ' identify letter & screen position of one to get rid of
    tRemove = Int(Rnd * 25)
    x = (tRemove - (Int(tRemove / 5) * 5))
    y = Int(tRemove / 5)
    tRemove = tRemove + 1
    ' create rectangles for selected letters & ensure no duplicate
    ' positions are selected
    SetRect tRect, x * 80, y * 80, x * 80 + 80, y * 80 + 80
    For J = 1 To UBound(rCircles)
        If EqualRect(tRect, rCircles(J)) Then Exit For
    Next
    If J > UBound(rCircles) Then
        ' add the letter's rectangle to the list
        ReDim Preserve rCircles(0 To UBound(rCircles) + 1)
        rCircles(UBound(rCircles)) = tRect
        ' if needed, reduce the number of bonus letters on the screen
        If UCase(Mid$(sBoard, tRemove, 1)) = Mid$(sBoard, tRemove, 1) Then nrBonus = nrBonus - 1
        Mid$(sBoard, tRemove, 1) = " "
    End If
Next
' here we fade out the selected letters
For I = 1 To 4
    For J = 1 To UBound(rCircles)
        DrawLetter "", rCircles(J).Left \ 80, rCircles(J).Top \ 80, 0, , I
    Next
    BitBlt picBoard.hdc, 0, 0, picBoard.Width, picBoard.Height, DrawDC.hTempDC, 0, 0, vbSrcCopy
    picBoard.Refresh
    Sleep 70
Next
' now we start rearranging the board
For I = 1 To 5
    For J = I To 25 Step 5
        sPhrase(I) = sPhrase(I) & Mid$(sBoard, J, 1)
    Next
    sDrop(I) = sPhrase(I)
Next
sBoard = Space$(25)
For I = 1 To 5
    sPhrase(6) = Replace$(sPhrase(I), " ", "")
    For J = 1 To 5 - Len(sPhrase(6))
        sPhrase(6) = GetLetter & sPhrase(6)
    Next
    For J = 1 To 5
        If UCase(Mid$(sPhrase(6), J, 1)) = Mid$(sPhrase(6), J, 1) Then iCircle = 1 Else iCircle = 0
        DrawLetter Mid$(sPhrase(6), J, 1), I - 1, J - 1, iCircle, , -1
        Mid$(sBoard, (J - 1) * 5 + I, 1) = Mid$(sPhrase(6), J, 1)
    Next
Next
LastTarget = 50
DropBalls sDrop(), False

BitBlt picBoard.hdc, 0, 0, picBoard.Width, picBoard.Height, DrawDC.hTempDC, 0, 0, vbSrcCopy
picBoard.Refresh

Erase rCircles
Erase sPhrase
End Sub

Private Sub cmdHOF_Click()
ShowHallOfFame False
End Sub

Private Sub cmdResetHOF_Click()
If MsgBox("Are you sure you want to complete erase all of the Hall of Fame entries?", _
    vbYesNo + vbQuestion + vbDefaultButton2, "Confirmation") = vbNo Then Exit Sub
    
Dim I As Integer
On Error GoTo FailedErase
Kill App.Path & "\wwFame.hof"
For I = 0 To 17
    Select Case I
    Case 0, 6, 12
    Case Else
        lstHOF.List(I) = ""
    End Select
Next
Exit Sub

FailedErase:
MsgBox "Failed to delete the file wwFame.hof" & vbCrLf & vbCrLf & _
    "Delete this file yourself to clear the Hall of Fame", vbInformation + vbOKOnly, "Error"
End Sub

Private Sub Form_Load()
Dim xLtr As String, I As Integer, J As Integer, iCircle As Integer
Dim fNR As Integer, vData() As Byte

' ensure a copy of the dictionary is there
sFile = App.Path & "\wWords.txt"
If Len(Dir(sFile)) = 0 Then
    fNR = FreeFile()
    Open sFile For Binary As #fNR
    vData = LoadResData(100, "Custom")
    Put #fNR, , vData()
    Close #fNR
End If
' ensure the picture for colored balls is there
If Len(Dir(App.Path & "\wwBalls.jpg")) = 0 Then
    fNR = FreeFile()
    Open App.Path & "\wwBalls.jpg" For Binary As #fNR
    vData = LoadResData(101, "Custom")
    Put #fNR, , vData()
    Close #fNR
End If
Set picRounds.Picture = LoadPicture(App.Path & "\wwBalls.jpg")

' now we open the dictionary & calculate how many 3/4/5 letter words are in it
fNR = FreeFile()
Open sFile For Input As #fNR
Do Until EOF(fNR)
    Line Input #fNR, xLtr
    If Len(xLtr) > 2 And Len(xLtr) < 6 Then
        WordCount(Len(xLtr) - 3) = WordCount(Len(xLtr) - 3) + 1
    End If
Loop
Close #fNR
If WordCount(0) = 0 Or WordCount(1) = 0 Or WordCount(2) = 0 Then
    MsgBox "The dictionary being used is not in the right format or is empty." & vbCrLf & _
        " Delete the file wWords.txt and start the game over.", vbExclamation + vbOKOnly, "Invalid Dictionary"
    CleanUpDC
    Unload Me
    Exit Sub
End If

picBoard.Enabled = False
Randomize Timer
' show the splash screen
For I = 0 To 4
    For J = 0 To 4
        xLtr = Mid$("WHICHWORD?     La   Volpe", (I * 5 + J) + 1, 1)
        If I < 2 And xLtr <> " " Then
            iCircle = 1
        Else
            If I > 2 And xLtr <> " " Then iCircle = 2 Else iCircle = 0
        End If
        DrawLetter xLtr, CLng(J), CLng(I), iCircle, False
    Next
Next
picBoard.Refresh

' set up the list boxes with tabs & prime them
Const LB_SETTABSTOPS As Long = &H192
Dim lstBaseUnits As Long
Dim TabStop(0 To 1) As Long
    lstBaseUnits = (GetDialogBaseUnits() Mod 65536) / 2
    'set tab stops
    TabStop(0) = 8 * lstBaseUnits
    TabStop(1) = 16 * lstBaseUnits
    Call SendMessage(lstWords.hWnd, LB_SETTABSTOPS, 2, TabStop(0))
    TabStop(0) = 8 * lstBaseUnits
    TabStop(1) = 16 * lstBaseUnits
    Call SendMessage(lstScore.hWnd, LB_SETTABSTOPS, 2, TabStop(0))
    TabStop(0) = 15 * lstBaseUnits
    TabStop(1) = 37 * lstBaseUnits
    Call SendMessage(lstHOF.hWnd, LB_SETTABSTOPS, 2, TabStop(0))
Erase TabStop
lstWords.AddItem "Word" & vbTab & "Score" & vbTab & "Bonus"
lstScore.AddItem "WHICHwords? to find..."
For I = 1 To 2
    lstScore.AddItem ""
Next
lstScore.AddItem "Score:" & vbTab & "0"
With lstHOF
    For I = 1 To 18
    .AddItem ""
    Next
    If Len(Dir(App.Path & "\wwFame.hof")) Then
        On Error GoTo FailedRead
        fNR = FreeFile()
        Open App.Path & "\wwFame.hof" For Input As #fNR
        For I = 0 To 17
            Select Case I
            Case 0, 6, 12
            Case Else
                Line Input #fNR, xLtr
                .List(I) = xLtr
            End Select
        Next
        Close #fNR
    End If
    .List(0) = "Best Score" & vbTab & "Player" & vbTab & "Date"
    .List(6) = "Most Levels" & vbTab & "Player" & vbTab & "Date"
    .List(12) = "Most Words" & vbTab & "Player" & vbTab & "Date"
    .Selected(0) = True
    .Selected(6) = True
    .Selected(12) = True
End With
TimeLimit = 180

' create a no-click area (the black area around the balls)
Dim tRgn As Long
rgnNoClick = CreateRectRgn(0, 0, picBoard.Width, picBoard.Width)
For I = 0 To 4
    For J = 0 To 4
        tRgn = CreateEllipticRgn(I * 80 + 2, J * 80 + 2, (I + 1) * 80 - 1, (J + 1) * 80 - 1)
        CombineRgn rgnNoClick, rgnNoClick, tRgn, RGN_DIFF
        DeleteObject tRgn
    Next
Next

' 1st run only -- create the offscreen DCs
' this one is to paint & BitBlt to/from
Dim hBMP As Long, hFont As Long, hPen As Long, hBrush As Long
DrawDC.hTempDC = CreateCompatibleDC(picBoard.hdc)
hBMP = CreateCompatibleBitmap(picBoard.hdc, picBoard.Width, picBoard.Height)
DrawDC.hOldBMP = SelectObject(DrawDC.hTempDC, hBMP)
SetBkColor DrawDC.hTempDC, 0&
SetTextColor DrawDC.hTempDC, 0&
Dim newFont As LOGFONT
newFont.lfCharSet = 1
newFont.lfFaceName = picBoard.Font.Name & Chr$(0)
newFont.lfHeight = (picBoard.Font.Size * -20) / Screen.TwipsPerPixelY
newFont.lfWeight = picBoard.Font.Weight
newFont.lfItalic = Abs(CInt(picBoard.Font.Italic))
newFont.lfStrikeOut = Abs(CInt(picBoard.Font.Strikethrough))
newFont.lfUnderline = Abs(CInt(picBoard.Font.Underline))
hFont = CreateFontIndirect(newFont)
DrawDC.hOldFont = SelectObject(DrawDC.hTempDC, hFont)
hPen = CreatePen(0, 1, 0)
hBrush = CreateSolidBrush(0)
DrawDC.hOldPen = SelectObject(DrawDC.hTempDC, hPen)
DrawDC.hOldBrush = SelectObject(DrawDC.hTempDC, hBrush)
SetBkMode DrawDC.hTempDC, 3&

' this one is to animate the dropped balls
DropDC.hTempDC = CreateCompatibleDC(picBoard.hdc)
hBMP = CreateCompatibleBitmap(picBoard.hdc, 80, picBoard.Height)
DropDC.hOldBMP = SelectObject(DropDC.hTempDC, hBMP)
SetBkColor DropDC.hTempDC, 0&
SetTextColor DropDC.hTempDC, 0&
hPen = CreatePen(0, 1, 0)
hBrush = CreateSolidBrush(0)
DropDC.hOldPen = SelectObject(DropDC.hTempDC, hPen)
DropDC.hOldBrush = SelectObject(DropDC.hTempDC, hBrush)
BitBlt DrawDC.hTempDC, 0, 0, picBoard.Width, picBoard.Height, picBoard.hdc, 0, 0, vbSrcCopy

Exit Sub


FailedRead:
Close #fNR
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
CleanUpDC
Erase WordCount
Erase BonusWords
Erase SoundBuffer
End Sub

Private Sub CleanUpDC()
If rgnNoClick Then DeleteObject rgnNoClick
If DrawDC.hTempDC Then
    DeleteObject SelectObject(DrawDC.hTempDC, DrawDC.hOldBMP)
    DeleteObject SelectObject(DrawDC.hTempDC, DrawDC.hOldFont)
    DeleteObject SelectObject(DrawDC.hTempDC, DrawDC.hOldPen)
    DeleteObject SelectObject(DrawDC.hTempDC, DrawDC.hOldBrush)
    DeleteDC DrawDC.hTempDC
End If
If DropDC.hTempDC Then
    DeleteObject SelectObject(DropDC.hTempDC, DropDC.hOldBMP)
    DeleteObject SelectObject(DropDC.hTempDC, DropDC.hOldPen)
    DeleteObject SelectObject(DropDC.hTempDC, DropDC.hOldBrush)
    DeleteDC DropDC.hTempDC
End If
End Sub


Private Sub mnuGame_Click(Index As Integer)
Select Case Index
Case 0: ' New Game
    If lstHOF.Visible Then ShowHallOfFame False
    Timer1.Enabled = False
    picBoard.Enabled = False
    StartGame
Case 1: ' Time Limit
    If picBoard.Enabled Then ' playing a game
        If MsgBox("Changing the time limit will stop the current game." & vbCrLf & vbCrLf & _
            "Continue?", vbYesNo + vbQuestion + vbDefaultButton2, "Stop the Game?") = vbNo Then Exit Sub
    End If
    picBoard.Enabled = False
    Timer1.Enabled = False
    lblTimer.Visible = False
    Dim newLimit As String
    newLimit = InputBox("Enter the number of minutes before time runs out." & vbCrLf & vbCrLf & _
        "Current time limit is " & TimeLimit \ 60 & " minute(s).", "Game Time Limit", TimeLimit \ 60)
    If CInt(Val(newLimit)) > 0 Then
        TimeLimit = CInt(Val(newLimit)) * 60
    Else
        If Val(newLimit) <> 0 Then
            MsgBox "Invalid time limit. The limit must be equal to or greater than 1 minute.", vbInformation + vbOKOnly
        End If
    End If
Case 2: ' Sound
    mnuGame(2).Checked = Not mnuGame(2).Checked
Case 3: ' Hall of Fame
    ShowHallOfFame (Not lstHOF.Visible)
Case 5: ' Help
    frmHelp.Show 1, Me
Case 7: ' Exit
    Unload Me
End Select
End Sub

Private Sub picBoard_Click()
If Timer1.Enabled = False Then Exit Sub
Dim cPT As POINTAPI
GetCursorPos cPT
ScreenToClient picBoard.hWnd, cPT
If PtInRegion(rgnNoClick, cPT.x, cPT.y) Then
    If LastTarget < 50 Then ToggleLetter (LastTarget - (Int(LastTarget / 5) * 5)) * 80, Int(LastTarget / 5) * 80, False
Else
    BeginPlaySound 102
    ToggleLetter cPT.x, cPT.y, False
End If
End Sub

Private Sub picBoard_DblClick()
If Timer1.Enabled = False Then Exit Sub
Dim cPT As POINTAPI
GetCursorPos cPT
ScreenToClient picBoard.hWnd, cPT
If PtInRegion(rgnNoClick, cPT.x, cPT.y) Then
    If LastTarget < 50 Then ToggleLetter (LastTarget - (Int(LastTarget / 5) * 5)) * 80, Int(LastTarget / 5) * 80, False
Else
    ToggleLetter cPT.x, cPT.y, True
End If
End Sub


Private Sub Timer1_Timer()
Dim xLen As Long
xLen = lblTimer.Tag / TimeLimit
If lblTimer.Width <= xLen Then
    Timer1.Enabled = False
    picBoard.Enabled = False
    lblTimer.Visible = False
    cmdBlast.Enabled = False
    MsgBox "Game Over", vbInformation + vbOKOnly, "Which Word?"
    TallyHOFinfo
    ' end of game
Else
    If lblTimer.Width < (15 * xLen) Then BeginPlaySound 104
    lblTimer.Width = lblTimer.Width - xLen
'    lblTimer.Refresh
End If
End Sub

Private Sub GetBonusWords()

' function gets random words out of the dictionary to be used as
' the WhichWords? to find

Dim fNR As Integer, I As Integer, J As Integer, K As Integer
Dim lRnd As Long, LineCT(0 To 2) As Long, iBonusCt As Integer
Dim nr3s() As Integer, nr4s() As Integer, nr5s() As Integer
Dim xLtr As String

ReDim BonusWords(0 To 3)
ReDim nr3s(0)
ReDim nr4s(0)
ReDim nr5s(0)
Select Case xLevel
Case 1
    If Int(Rnd * 100) + 1 < 50 Then
        ReDim nr3s(0 To 1)
        nr3s(0) = 1
    Else
        ReDim nr4s(0 To 1)
        nr4s(0) = 1
    End If
Case 2
    ReDim nr4s(0 To 1)
    nr4s(0) = 1
Case 3
    If Int(Rnd * 100) + 1 < 50 Then
        ReDim nr5s(0 To 1)
        nr5s(0) = 1
    Else
        ReDim nr4s(0 To 2)
        nr4s(0) = 2
    End If
Case 4, 5
    If Int(Rnd * 100) + 1 < 50 Then
        ReDim nr4s(0 To 2)
        nr4s(0) = 2
    Else
        ReDim nr4s(0 To 1)
        ReDim nr5s(0 To 1)
        nr4s(0) = 1
        nr5s(0) = 1
    End If
Case Else
    If (Int(Rnd * 100) + 1) Mod 3 = 0 Then
        ReDim nr4s(0 To 1)
        ReDim nr5s(0 To 3)
        nr5s(0) = 3
        nr4s(0) = 1
    Else
        ReDim nr4s(0 To 2)
        ReDim nr5s(0 To 2)
        nr5s(0) = 2
        nr4s(0) = 2
    End If
End Select
For I = 1 To 3
    For J = 1 To Choose(I, nr3s(0), nr4s(0), nr5s(0))
        lRnd = Int(Rnd * WordCount(I - 1)) + 1
        Select Case I
        Case 1
            For K = 1 To nr3s(0)
                If lRnd = nr3s(K) Then
                    lRnd = 0
                    Exit For
                End If
            Next
            If lRnd > 0 And nr3s(0) > 0 Then nr3s(J) = lRnd
        Case 2
            For K = 1 To nr4s(0)
                If lRnd = nr4s(K) Then
                    lRnd = 0
                    Exit For
                End If
            Next
            If lRnd > 0 And nr4s(0) > 0 Then nr4s(J) = lRnd
        Case 3
            For K = 1 To nr5s(0)
                If lRnd = nr5s(K) Then
                    lRnd = 0
                    Exit For
                End If
            Next
            If lRnd > 0 And nr5s(0) > 0 Then nr5s(J) = lRnd
        End Select
        If lRnd = 0 Then J = J - 1
    Next
Next

fNR = FreeFile()
Open sFile For Input As #fNR
Do While EOF(fNR) = False
    Line Input #fNR, xLtr
    Select Case Len(xLtr)
    Case 3
        LineCT(0) = LineCT(0) + 1
        For I = 1 To nr3s(0)
            If LineCT(0) = nr3s(I) Then
                BonusWords(iBonusCt) = xLtr
                iBonusCt = iBonusCt + 1
                Exit For
            End If
        Next
    Case 4
        LineCT(1) = LineCT(1) + 1
        For I = 1 To nr4s(0)
            If LineCT(1) = nr4s(I) Then
                BonusWords(iBonusCt) = xLtr
                iBonusCt = iBonusCt + 1
                Exit For
            End If
        Next
    Case 5
        LineCT(2) = LineCT(2) + 1
        For I = 1 To nr5s(0)
            If LineCT(2) = nr5s(I) Then
                BonusWords(iBonusCt) = xLtr
                iBonusCt = iBonusCt + 1
                Exit For
            End If
        Next
    End Select
Loop
Close #fNR
Erase LineCT
Erase nr3s
Erase nr4s
Erase nr5s

' update the list of words to find
lstScore.List(1) = UCase(BonusWords(0)) & vbTab & UCase(BonusWords(2)) & vbTab
lstScore.List(2) = UCase(BonusWords(1)) & vbTab & UCase(BonusWords(3)) & vbTab
lstScore.Selected(1) = True
lstScore.Selected(2) = True
lblSBar.Caption = ""

End Sub

Private Sub DropBalls(sRow() As String, bWithSound As Boolean)

' function is the animation routine for dropping balls

Dim iHeight As Integer, I As Integer, iSpace As Integer, iDrop As Integer
Dim iRow As Integer, iStart As Integer, iStop As Integer, srcDC As Long
Dim iPos As Integer

If bWithSound Then BeginPlaySound 105

For iRow = 1 To 5
    For iSpace = 5 To 2 Step -1
        iDrop = 0
        iStart = 0
        If Mid$(sRow(iRow), iSpace, 1) = " " Then
            iHeight = iSpace - 1
            For I = iSpace To 1 Step -1
                If Mid$(sRow(iRow), I, 1) = " " Then
                    iDrop = iDrop + 1
                Else
                    iSpace = iHeight + 1
                    iHeight = I
                    iStart = 1
                    Exit For
                End If
            Next
        End If
        If iDrop > 0 And iStart > 0 Then
            Rectangle DropDC.hTempDC, 0, 0, 80, 5 * 80
            srcDC = picBoard.hdc
            BitBlt DropDC.hTempDC, 0, 40, 80, iHeight * 80, srcDC, (iRow - 1) * 80, 0, vbSrcCopy
            
            For I = 1 To iDrop * 2
                BitBlt picBoard.hdc, (iRow - 1) * 80, ((I - 1) * 40), 80, iHeight * 80 + 40, DropDC.hTempDC, 0, 0, vbSrcCopy
                picBoard.Refresh
                Sleep 55
            Next
            sRow(iRow) = Space$(iDrop) & Left$(sRow(iRow), iHeight) & "@"
        End If
    Next
Next

End Sub

Public Sub BeginPlaySound(ResourceId As String, Optional bStop As Boolean = False, Optional bWait As Boolean = False)

' general use sub to play sounds from resource files or actual disk files

Dim fNR As Integer


If mnuGame(2).Checked = False Then Exit Sub

If bStop Then
    sndPlaySound ByVal vbNullString, &H1
    Exit Sub
End If

Dim sndFlags As Long
' &H0 = Sync (halts program until sound done)
' &H1 = Async (returns immediately)
' &H2 = No_Default otherwise you get the computer beep
' &H2000 = No_Wait, plays immediately
' &H20000 = FileName
' &H4 = Memory
sndFlags = Abs(CInt(bWait) + 1) Or &H2 Or &H2000

If IsNumeric(ResourceId) Then
    SoundBuffer = LoadResData(Val(ResourceId), "Custom")
    sndPlaySound SoundBuffer(0), sndFlags Or &H4
Else
    sndPlaySound ByVal ResourceId, sndFlags Or &H20000
End If

End Sub

Private Sub TallyHOFinfo()
Dim nrWordsMade As Long, I As Integer, J As Integer
Dim sMsg As String, sName As String, isBest(1 To 3) As Integer

nrWordsMade = Val(lstWords.Tag) + lstWords.ListCount - 1
For I = 1 To 5
    If Val(Replace$(lstHOF.List(I), ",", "")) < lScore Then
        isBest(1) = I
        sMsg = "Top 5 Best Scores"
        Exit For
    End If
Next
For I = 7 To 11
    If Val(Replace$(lstHOF.List(I), ",", "")) < xLevel Then
        isBest(2) = I
        If Len(sMsg) Then sMsg = sMsg & vbCrLf
        sMsg = sMsg & "Top 5 Most Levels Completed"
        Exit For
    End If
Next
For I = 13 To 17
    If Val(Replace$(lstHOF.List(I), ",", "")) < nrWordsMade Then
        isBest(3) = I
        If Len(sMsg) Then sMsg = sMsg & vbCrLf
        sMsg = sMsg & "Top 5 Most Words Created"
        Exit For
    End If
Next

If Len(sMsg) Then
    sMsg = "You have made the Hall of Fame for..." & vbCrLf & sMsg
    sMsg = sMsg & vbCrLf & vbCrLf & "Enter your name below to be included " & _
        "in the Hall of Fame or hit the Escape key to cancel."
    sName = InputBox(sMsg, "Hall of Fame Inductee")
    If Len(sName) Then
        sName = Left$(sName, 20)
        For I = 1 To 3
            If isBest(I) Then
                For J = Choose(I, 5, 11, 17) To isBest(I) Step -1
                    lstHOF.List(J) = lstHOF.List(J - 1)
                Next
                lstHOF.List(isBest(I)) = Format$(Choose(I, lScore, xLevel, nrWordsMade), "#,###") & vbTab & sName & vbTab & Format$(Date, "Short Date")
            End If
        Next
        On Error GoTo FailedWrite
        Dim fNR As Integer
        fNR = FreeFile()
        Open App.Path & "\wwFame.hof" For Output As #fNR
        For I = 0 To 17
            Select Case I
            Case 0, 6, 12
            Case Else
                Print #fNR, lstHOF.List(I)
            End Select
        Next
        Close #fNR
        ShowHallOfFame True
    End If
End If
Exit Sub

FailedWrite:
MsgBox "Failed to create a new Hall of Fame file." & vbCrLf & vbCrLf & _
    Err.Description, vbInformation + vbOKOnly
Exit Sub
End Sub

Private Sub ShowHallOfFame(bYesNo As Boolean)

If bYesNo Then
    Timer1.Enabled = False
    picBoard.Cls
Else
    BitBlt picBoard.hdc, 0, 0, picBoard.Width, picBoard.Height, DrawDC.hTempDC, 0, 0, vbSrcCopy
    Timer1.Enabled = picBoard.Enabled
End If
picBoard.Refresh

cmdHOF.Enabled = bYesNo
lstHOF.Enabled = bYesNo
lstHOF.Visible = bYesNo
cmdHOF.Visible = bYesNo
cmdResetHOF.Enabled = bYesNo
cmdResetHOF.Visible = bYesNo

cmdBlast.Enabled = Timer1.Enabled
End Sub
