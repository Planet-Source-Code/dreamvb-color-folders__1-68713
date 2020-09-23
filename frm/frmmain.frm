VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Folder Colors"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   7650
      TabIndex        =   28
      Top             =   4485
      Width           =   7710
      Begin VB.Label lblPath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "#"
         Height          =   195
         Left            =   90
         TabIndex        =   29
         Top             =   45
         Width           =   105
      End
   End
   Begin VB.TextBox txtInfo 
      Height          =   350
      Left            =   5685
      TabIndex        =   27
      Top             =   2025
      Width           =   1905
   End
   Begin VB.PictureBox Picture1 
      Height          =   300
      Index           =   3
      Left            =   5700
      ScaleHeight     =   240
      ScaleWidth      =   1815
      TabIndex        =   25
      Top             =   1650
      Width           =   1875
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Folder Info Tip"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   45
         TabIndex        =   26
         Top             =   15
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "E&xit"
      Height          =   350
      Left            =   5835
      TabIndex        =   24
      Top             =   2625
      Width           =   1440
   End
   Begin VB.CommandButton cmdSetFolder 
      Caption         =   "&Set Folder Icon"
      Height          =   350
      Left            =   4275
      TabIndex        =   23
      Top             =   2625
      Width           =   1440
   End
   Begin VB.PictureBox Picture1 
      Height          =   300
      Index           =   2
      Left            =   2790
      ScaleHeight     =   240
      ScaleWidth      =   2685
      TabIndex        =   19
      Top             =   555
      Width           =   2745
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Color Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   45
         TabIndex        =   20
         Top             =   15
         Width           =   1200
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   300
      Index           =   1
      Left            =   2790
      ScaleHeight     =   240
      ScaleWidth      =   1170
      TabIndex        =   17
      Top             =   120
      Width           =   1230
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Folder Style"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   45
         TabIndex        =   18
         Top             =   15
         Width           =   1020
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   360
      Index           =   0
      Left            =   90
      ScaleHeight     =   300
      ScaleWidth      =   2505
      TabIndex        =   15
      Top             =   105
      Width           =   2565
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Folder:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   16
         Top             =   60
         Width           =   600
      End
   End
   Begin VB.DirListBox Dir1 
      Height          =   3915
      Left            =   90
      TabIndex        =   14
      Top             =   495
      Width           =   2580
   End
   Begin VB.CommandButton cmdabout 
      Caption         =   "&About"
      Height          =   350
      Left            =   2805
      TabIndex        =   13
      Top             =   3135
      Width           =   1380
   End
   Begin VB.ComboBox cboStyle 
      Height          =   315
      Left            =   4035
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   105
      Width           =   1515
   End
   Begin VB.TextBox RgbVal 
      Height          =   285
      Index           =   2
      Left            =   5010
      MaxLength       =   255
      TabIndex        =   11
      Text            =   "255"
      Top             =   1725
      Width           =   525
   End
   Begin VB.TextBox RgbVal 
      Height          =   285
      Index           =   1
      Left            =   5010
      MaxLength       =   255
      TabIndex        =   10
      Text            =   "128"
      Top             =   1350
      Width           =   525
   End
   Begin VB.TextBox RgbVal 
      Height          =   285
      Index           =   0
      Left            =   5010
      MaxLength       =   255
      TabIndex        =   9
      Text            =   "0"
      Top             =   945
      Width           =   525
   End
   Begin VB.HScrollBar RgbBar 
      Height          =   255
      Index           =   2
      Left            =   3285
      Max             =   255
      TabIndex        =   8
      Top             =   1725
      Value           =   255
      Width           =   1650
   End
   Begin VB.HScrollBar RgbBar 
      Height          =   255
      Index           =   1
      Left            =   3300
      Max             =   255
      TabIndex        =   7
      Top             =   1350
      Value           =   128
      Width           =   1650
   End
   Begin VB.HScrollBar RgbBar 
      Height          =   255
      Index           =   0
      Left            =   3300
      Max             =   255
      TabIndex        =   6
      Top             =   945
      Width           =   1650
   End
   Begin VB.PictureBox pColor 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   5175
      ScaleHeight     =   195
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   2085
      Width           =   255
   End
   Begin VB.PictureBox pPal 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2820
      MousePointer    =   2  'Cross
      Picture         =   "frmmain.frx":08CA
      ScaleHeight     =   195
      ScaleWidth      =   2325
      TabIndex        =   1
      Top             =   2085
      Width           =   2325
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview Icon"
      Height          =   350
      Left            =   2805
      TabIndex        =   0
      Top             =   2625
      Width           =   1380
   End
   Begin VB.PictureBox Picture2 
      Height          =   1485
      Left            =   5700
      ScaleHeight     =   1425
      ScaleWidth      =   1815
      TabIndex        =   21
      Top             =   90
      Width           =   1875
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Preview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   345
         TabIndex        =   22
         Top             =   255
         Width           =   1185
      End
      Begin VB.Image ImgIco 
         Height          =   510
         Left            =   690
         Top             =   660
         Width           =   480
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   2820
      X2              =   7545
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   2820
      X2              =   7545
      Y1              =   2505
      Y2              =   2505
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
      Height          =   195
      Left            =   2805
      TabIndex        =   5
      Top             =   1725
      Width           =   315
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Green"
      Height          =   195
      Left            =   2805
      TabIndex        =   4
      Top             =   1395
      Width           =   435
   End
   Begin VB.Label lblrgb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Red"
      Height          =   195
      Left            =   2805
      TabIndex        =   3
      Top             =   960
      Width           =   300
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Private Declare Function GetVersion Lib "kernel32.dll" () As Long

Private Type TRGB
    Red As Integer
    Green As Integer
    Blue As Integer
End Type

Private m_Rgb As TRGB
Private m_Main As TRGB
Private m_SelectColor As Long
Private m_TempFile As String
Private m_IconTemplate As String
Private m_FolderName As String
Private m_IsMoveing As Boolean

Private Sub KillTempFile()
On Error Resume Next
    If FileFound(m_TempFile) Then Kill m_TempFile
End Sub

Private Sub WriteToFile(lzOutFile As String, OutData As String)
Dim fp As Long
    fp = FreeFile
    Open lzOutFile For Output As #fp
        Print #fp, OutData
    Close #fp
End Sub

Private Sub CleanGarbageExit()
    Call KillTempFile
    ZeroMemory m_Rgb, Len(m_Rgb)
    ZeroMemory m_Main, Len(m_Main)
    m_SelectColor = 0
    m_TempFile = vbNullString
    m_IconTemplate = vbNullString
    m_FolderName = vbNullString
    cboStyle.Clear
    Unload frmmain
End Sub

Function FileFound(lzFile As String) As Boolean
On Error Resume Next
    FileFound = LenB(Dir(lzFile)) <> 0
End Function

Function FixPath(lPath As String) As String
    If Right(lPath, 1) <> "\" Then
        FixPath = lPath & "\"
    Else
        FixPath = lPath
    End If
End Function

Private Sub PreviewIcon(lzInFile As String, lzOutFile As String)
Dim fp As Long
Dim lzTmpFile As String
Dim mData() As Byte

    fp = FreeFile
    
    Open lzInFile For Binary As #fp
        ReDim mData(0 To LOF(fp) - 1)
        Get #fp, , mData
    Close #fp
    
    'Lighten the color to build a highlight color
    m_Rgb.Blue = (m_Main.Blue) + 20
    m_Rgb.Green = (m_Main.Green) + 20
    m_Rgb.Red = (m_Main.Red) + 20
    'Check that we do not go over the RGB Range
    If (m_Rgb.Blue >= 255) Then m_Rgb.Blue = 255
    If (m_Rgb.Green >= 255) Then m_Rgb.Green = 255
    If (m_Rgb.Red >= 255) Then m_Rgb.Red = 255
        
    'Helight color
    mData(122) = m_Rgb.Blue 'Blue
    mData(123) = m_Rgb.Green 'Green
    mData(124) = m_Rgb.Red 'Red
        
    'Set then main body color for the folder icon
    mData(114) = m_Main.Blue
    mData(115) = m_Main.Green
    mData(116) = m_Main.Red
        
    'Next we make a darkshadow color for the folder
    m_Rgb.Blue = (m_Main.Blue) - 30
    m_Rgb.Green = (m_Main.Green) - 30
    m_Rgb.Red = (m_Main.Red) - 30
    
    'Check that we do not go below zero
    If (m_Rgb.Blue <= 0) Then m_Rgb.Blue = 0
    If (m_Rgb.Green <= 0) Then m_Rgb.Green = 0
    If (m_Rgb.Red <= 0) Then m_Rgb.Red = 0
        
    'Set shadow top
    mData(90) = m_Rgb.Blue
    mData(91) = m_Rgb.Green
    mData(92) = m_Rgb.Red
        
    'Set the shadow left side
    mData(82) = m_Rgb.Blue
    mData(83) = m_Rgb.Green
    mData(84) = m_Rgb.Red
        
    'Last shadow color we need to lighten greator than our first light shadow color
    m_Rgb.Blue = (m_Main.Blue) + 34
    m_Rgb.Green = (m_Main.Green) + 34
    m_Rgb.Red = (m_Main.Red) + 34
        
    'Check that we do not go over the RGB Range
    If (m_Rgb.Blue >= 255) Then m_Rgb.Blue = 255
    If (m_Rgb.Green >= 255) Then m_Rgb.Green = 255
    If (m_Rgb.Red >= 255) Then m_Rgb.Red = 255
        
    mData(118) = m_Rgb.Blue
    mData(119) = m_Rgb.Green
    mData(120) = m_Rgb.Red

    Open lzOutFile For Binary As #fp
        Put #fp, , mData
    Close #fp
    
    Erase mData
    
End Sub

Private Sub LongToRGB(RGBTYPE As TRGB, LngColor As Long)
Dim RgbBytes(2) As Byte
    
    CopyMemory RgbBytes(0), LngColor, 4
    
    With RGBTYPE
        .Red = RgbBytes(0)
        .Green = RgbBytes(1)
        .Blue = RgbBytes(2)
    End With
    
    Erase RgbBytes
End Sub

Private Sub cboStyle_Click()
    m_IconTemplate = FixPath(App.Path) & "src_icons\" & (cboStyle.ListIndex + 1) & ".ico"
    Call cmdPreview_Click
End Sub

Private Sub cmdabout_Click()
    MsgBox "Color Folders" _
    & " By DreamVB" & vbCrLf _
    & "Special thanks to GioRock for shareing his Custom Cursors Color example, for giveng me the idea for this example." _
    & vbCrLf & vbCrLf & "Find his code here: http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=68656&lngWId=1", vbInformation, frmmain.Caption
End Sub

Private Sub cmdexit_Click()
    Call CleanGarbageExit
End Sub

Private Sub cmdPreview_Click()
    If Not FileFound(m_IconTemplate) Then
        MsgBox "The file:" & vbCrLf & m_IconTemplate & " was not found.", vbExclamation, "File Not Found"
        Exit Sub
    Else
        'Kill the temp icon if it exsists
        If FileFound(m_TempFile) Then Call Kill(m_TempFile)
        Call PreviewIcon(m_IconTemplate, m_TempFile)
        'Show the user the icon
        ImgIco.Picture = LoadPicture(m_TempFile)
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSetFolder_Click()
Dim TmpFile As String
Dim sBuff As String

    If Len(m_FolderName) = 3 Then
        MsgBox "Sorry but drive icons cannot be set in this version.", vbExclamation, "Drive icons not supported"
        Exit Sub
    End If
    
    If Not FileFound(m_TempFile) Then
        MsgBox "File: " & m_TempFile & vbCrLf & "Cannot be found.", vbExclamation, "File Not Found"
        Exit Sub
    Else
        'Copy the temp file to the folder of were the icon will be placed on
        TmpFile = m_FolderName & "fol_icon.ico"
        FileCopy m_TempFile, TmpFile
        'Set attr
        SetAttr TmpFile, vbArchive + vbHidden
        'Next create the desktop.ini file
        sBuff = "[.ShellClassInfo]" & vbCrLf _
        & "InfoTip=" & txtInfo.Text & vbCrLf _
        & "IconIndex=0" & vbCrLf _
        & "IconFile=" & TmpFile & vbCrLf _
        & "NoSharing=1" & vbCrLf _
        & "ConfirmFileOp=0"
        '
        TmpFile = m_FolderName & "Desktop.ini"
        WriteToFile TmpFile, sBuff
        'Set attr
        SetAttr TmpFile, vbArchive + vbHidden
        'Set the folder to system
        SetAttr m_FolderName, vbSystem
    End If

End Sub

Private Sub Dir1_Change()
    m_FolderName = FixPath(Dir1.Path)
    lblPath.Caption = m_FolderName
End Sub

Private Sub Form_Load()
Dim X As Integer
    'Temp file for the icon
    m_TempFile = FixPath(App.Path) & "tmp.ico"
    For X = 1 To 11
        cboStyle.AddItem "Style " & CStr(X)
    Next X
    
    cboStyle.ListIndex = 2
    RgbBar_Change 0
    Call cmdPreview_Click
    Call Dir1_Change
End Sub

Private Sub Form_Terminate()
    Call KillTempFile
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmain = Nothing
End Sub

Private Sub pPal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (X < 0) Then X = 0
    If (X >= pPal.Width) Then Exit Sub
    If (Y < 0) Then Y = 0
    If (Y >= pPal.Height) Then Exit Sub
    pColor.BackColor = pPal.Point(X, Y)
    
    Call LongToRGB(m_Main, pColor.BackColor)
    
    RgbBar(0) = m_Main.Red
    RgbBar(1) = m_Main.Green
    RgbBar(2) = m_Main.Blue
    
    m_IsMoveing = True
End Sub

Private Sub pPal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (m_IsMoveing And Button = vbLeftButton) Then
        pPal_MouseDown Button, Shift, X, Y
    End If
End Sub

Private Sub RgbBar_Change(Index As Integer)
    RgbVal(Index).Text = RgbBar(Index).Value
    pColor.BackColor = RGB(RgbBar(0), RgbBar(1), RgbBar(2))
    '
    m_Main.Red = RgbVal(0)
    m_Main.Green = RgbVal(1)
    m_Main.Blue = RgbVal(2)
End Sub

Private Sub RgbBar_Scroll(Index As Integer)
    RgbBar_Change Index
End Sub

Private Sub RgbVal_Change(Index As Integer)
    If RgbVal(Index) > 255 Then
        RgbVal(Index) = 255
        RgbVal(Index).SelStart = 3
    End If
    
    RgbBar(Index) = RgbVal(Index)
End Sub

Private Sub RgbVal_KeyPress(Index As Integer, KeyAscii As Integer)
    'we only want the user to enter whole numbers
    If Not ((KeyAscii >= 48) And (KeyAscii <= 57) Or (KeyAscii = 8)) Then KeyAscii = 0: Exit Sub
End Sub
