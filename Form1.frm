VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Images-in-a-Globe"
   ClientHeight    =   10710
   ClientLeft      =   165
   ClientTop       =   -8325
   ClientWidth     =   11295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   11295
   StartUpPosition =   2  'CenterScreen
   Begin VB.VScrollBar VScroll1 
      Height          =   1470
      LargeChange     =   100
      Left            =   11010
      Max             =   2220
      SmallChange     =   100
      TabIndex        =   30
      Top             =   -15
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12420
      Left            =   -15
      ScaleHeight     =   828
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   728
      TabIndex        =   4
      Top             =   0
      Width           =   10920
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show Instructions"
         Height          =   225
         Left            =   150
         TabIndex        =   29
         Top             =   12030
         Value           =   1  'Checked
         Width           =   2505
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   435
         Index           =   0
         Left            =   780
         TabIndex        =   10
         Top             =   12630
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   435
         Index           =   1
         Left            =   795
         TabIndex        =   9
         Top             =   13215
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   435
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   12900
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   435
         Index           =   3
         Left            =   1515
         TabIndex        =   7
         Top             =   12930
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   435
         Index           =   4
         Left            =   2595
         TabIndex        =   6
         Top             =   12330
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   435
         Index           =   5
         Left            =   2625
         TabIndex        =   5
         Top             =   12885
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "the Fill Color"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   16
         Left            =   75
         TabIndex        =   28
         Top             =   10500
         Width           =   1665
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "pressed to Fill an area"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   15
         Left            =   75
         TabIndex        =   27
         Top             =   11535
         Width           =   2955
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Click with SHIFT"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   14
         Left            =   75
         TabIndex        =   26
         Top             =   11280
         Width           =   2190
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Use F3 to choose"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   13
         Left            =   75
         TabIndex        =   25
         Top             =   10200
         Width           =   2310
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER to fix and ESC to abort"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   12
         Left            =   5850
         TabIndex        =   24
         Top             =   10965
         Width           =   3210
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "F3 and F4 to set Background or Text  Color"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   11
         Left            =   5850
         TabIndex        =   23
         Top             =   11400
         Width           =   4665
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Click with Button 2 on the text to new Instance"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   10
         Left            =   5850
         TabIndex        =   22
         Top             =   12030
         Width           =   5130
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Double Click to create New Text"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   9
         Left            =   5850
         TabIndex        =   21
         Top             =   11835
         Width           =   3510
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "F1 and F2 to adjust the image size"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   8
         Left            =   5850
         TabIndex        =   20
         Top             =   11190
         Width           =   3690
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mouse Wheel to Zoom the Target or the Text"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   7
         Left            =   5850
         TabIndex        =   19
         Top             =   11610
         Width           =   4830
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "On the EDIT window use:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Index           =   6
         Left            =   5850
         TabIndex        =   18
         Top             =   10695
         Width           =   3390
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "to choose New Image"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   5
         Left            =   75
         TabIndex        =   17
         Top             =   9495
         Width           =   2880
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Click with Button 2"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   4
         Left            =   75
         TabIndex        =   16
         Top             =   9210
         Width           =   2520
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "or to Edit the Images"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   3
         Left            =   75
         TabIndex        =   15
         Top             =   8505
         Width           =   2775
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "with the Button 1 to Open"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   2
         Left            =   75
         TabIndex        =   14
         Top             =   8235
         Width           =   3390
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Click on the Hexagonos"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Index           =   1
         Left            =   75
         TabIndex        =   13
         Top             =   7950
         Width           =   3150
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "IMAGES-IN-A-GLOBE"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   570
         Index           =   0
         Left            =   225
         TabIndex        =   12
         Top             =   90
         Width           =   4560
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Height          =   210
         Left            =   2835
         TabIndex        =   11
         Top             =   2790
         Width           =   885
      End
      Begin VB.Image Image1 
         Height          =   1785
         Left            =   1230
         Stretch         =   -1  'True
         Top             =   855
         Width           =   1995
      End
      Begin VB.Image Image3 
         Enabled         =   0   'False
         Height          =   2955
         Left            =   960
         Picture         =   "Form1.frx":0000
         Stretch         =   -1  'True
         Top             =   690
         Width           =   2730
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1683
      Left            =   480
      Picture         =   "Form1.frx":1E39
      ScaleHeight     =   110
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   3
      Top             =   9675
      Visible         =   0   'False
      Width           =   1898
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H000000FF&
      Height          =   12270
      Left            =   6705
      Picture         =   "Form1.frx":8217
      ScaleHeight     =   818
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   730
      TabIndex        =   0
      Top             =   1590
      Visible         =   0   'False
      Width           =   10950
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   450
         Top             =   3660
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   3000
         Left            =   5460
         ScaleHeight     =   198
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   198
         TabIndex        =   2
         Top             =   6915
         Visible         =   0   'False
         Width           =   3000
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   540
         Left            =   525
         TabIndex        =   1
         Top             =   6870
         Width           =   1545
      End
      Begin VB.Image Image2 
         Height          =   420
         Index           =   0
         Left            =   1185
         Top             =   3780
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Menu Files 
      Caption         =   "Files"
      Begin VB.Menu New 
         Caption         =   "New"
      End
      Begin VB.Menu Open 
         Caption         =   "Open "
      End
      Begin VB.Menu Save 
         Caption         =   "Save"
      End
      Begin VB.Menu Null_ 
         Caption         =   "-"
      End
      Begin VB.Menu Print_Icosaedron 
         Caption         =   "Print"
      End
      Begin VB.Menu null1 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
      Begin VB.Menu About_index 
         Caption         =   "Autor: Agustin Rodriguez"
         Index           =   0
      End
      Begin VB.Menu About_index 
         Caption         =   "E-Mail: virtual_guitar_1@hotmail.com"
         Index           =   1
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal i As Long, ByVal i As Long, ByVal W As Long, ByVal i As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Const FLOODFILLSURFACE As Long = 1
Private PX As Integer
Private PY As Integer
Private Angle As Integer

Private Sub Check1_Click()
Dim i As Integer
For i = 1 To Label1.Count - 1
    Label1(i).Visible = Check1.Value
Next

End Sub

Private Sub Command1_Click()

    Select Case Hexagono_index
      Case 1
        Angle = 324
        PX = 259
        PY = -6
      Case 2
        Angle = 24
        PX = 468
        PY = 38
      Case 3
        Angle = 300
        PX = 225
        PY = 106
      Case 4
        Angle = 24
        PX = 426
        PY = 129
      Case 5
        Angle = 0
        PX = 312
        PY = 155
      Case 6
        Angle = 13
        PX = -38
        PY = 286
      Case 7
        Angle = 0
        PX = 313
        PY = 255
      Case 8
        Angle = -13
        PX = 131
        PY = 287
      Case 9
        Angle = 11
        PX = 492
        PY = 287

      Case 10
        Angle = -12
        PX = 226
        PY = 318
      Case 11
        Angle = 12
        PX = 398
        PY = 317
      Case 12
        Angle = -12
        PX = 57
        PY = 354
      Case 13
        Angle = 35
        PX = 259
        PY = 419
      Case 14
        Angle = 24
        PX = 366
        PY = 418
      Case 15
        Angle = -25
        PX = 200
        PY = 500
      Case 16
        Angle = -36
        PX = 423
        PY = 499
      Case 17
        Angle = 24
        PX = 522
        PY = 488
      Case 18
        Angle = -24
        PX = 241
        PY = 591
      Case 19
        Angle = -12
        PX = 560
        PY = 599
      Case 20
        Angle = 0
        PX = 148
        PY = 660
    End Select

    Picture3.Cls
    Rotate Picture3.hdc, Picture3.ScaleWidth / 2, Picture3.ScaleHeight / 2, Angle, Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight

    GdiTransparentBlt Picture4.hdc, PX, PY, Picture3.ScaleWidth, Picture3.ScaleHeight, Picture3.hdc, 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight, vbBlack

    Picture4.Refresh

End Sub

Private Sub Command2_Click(Index As Integer)

    Select Case Index
      Case 3
        PX = PX + 1
      Case 2
        PX = PX - 1
      Case 0
        PY = PY - 1
      Case 1
        PY = PY + 1
      Case 4
        Angle = Angle + 1
      Case 5
        Angle = Angle - 1
    End Select

    Picture4.Cls
    Picture3.Cls
    Rotate Picture3.hdc, Picture3.ScaleWidth / 2, Picture3.ScaleHeight / 2, Angle, Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight

    GdiTransparentBlt Picture4.hdc, PX, PY, Picture3.ScaleWidth, Picture3.ScaleHeight, Picture3.hdc, 0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight, vbBlack

    Picture4.Refresh

End Sub

Private Sub Command3_Click()

    Index = 1
    Form2.Original.Picture = LoadPicture(ICOSAEDRON_data(Index).filename)
    Form2.Move 0, 0, ICOSAEDRON_data(Index).Width, ICOSAEDRON_data(Index).Height
    Form2.Image2.Width = ICOSAEDRON_data(Index).Target_Width
    Form2.Image2.Height = ICOSAEDRON_data(Index).Target_Height
    Form2.Image2.Top = ICOSAEDRON_data(Index).Target_Top
    Form2.Image2.Left = ICOSAEDRON_data(Index).Target_Left

    BorderW = ICOSAEDRON_data(Index).BorderW
    BorderH = ICOSAEDRON_data(Index).BorderH

    Form2.Show
    
End Sub

Private Sub Exit_Click()

    Unload Form2
    End

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Dim Ret As SelectedColor

    If KeyCode = 114 Then
        Ret = ShowColor(Me.hwnd, False)
        If Ret.bCanceled Then
            Exit Sub
        End If
        FillColor = Ret.oSelectedColor
    End If
    
End Sub

Private Sub Form_Load()

  Dim i As Integer
    
    For i = 1 To 50
        If i Then
            Load Image2(i)
            Image2(i).Picture = LoadPicture(App.Path & "\Animation\Fr " + Right$("00" + Trim$(Str(i)), 2) + ".jpg")
        End If
    Next i

    Picture4.Picture = Picture2.Picture
     Height = 13230
    If Screen.Height / Screen.TwipsPerPixelY < 1024 Then
    VScroll1.Visible = True
    Height = 10600
    End If
    
    VScroll1.Max = 2500
End Sub

Private Sub Form_Resize()
VScroll1.Height = ScaleHeight
VScroll1.Move ScaleWidth - VScroll1.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload Form2

End Sub

Private Sub Label2_Click()

    Timer1.Enabled = Not Timer1.Enabled

End Sub

Private Sub New_Click()

    Erase ICOSAEDRON_data
    Picture4.Picture = Picture2.Picture

End Sub

Private Sub Open_Click()

  Dim Free As Integer
  Dim x As String
  Dim s As Integer
  
  
    Icosaedron_filename = FileDialog(Me, False, "Open Icosaedron", "Icosaedron File|*.edr", " ")
    If Icosaedron_filename = "" Then
        Exit Sub
    End If
    Free = FreeFile
    Open Icosaedron_filename For Binary As Free
    Get #Free, 1, ICOSAEDRON_data
    Close Free
    x = Icosaedron_filename
    s = InStr(x, ".edr")
    Mid$(x, s) = ".bmp"
    
    If Validate_load(x) = False Then
        s = MsgBox("Main Picture not Found. Try rebuild it. Click on all hexagonos to make this", vbCritical)
    Exit Sub
    End If
    
    Picture4.Picture = LoadPicture(x)

End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim i As Integer

    If Shift Then
        Make_fill x, y
        Exit Sub
    End If

    If (Picture2.Point(x, y) And &HFF) > 20 Then
        Exit Sub
    End If
    Hexagono_index = Picture2.Point(x, y) And &HFF

    If Button = 1 Then
        
        Hexagono_filename = ICOSAEDRON_data(Hexagono_index).filename
        Bkp_filename = Hexagono_filename
        If ICOSAEDRON_data(Hexagono_index).filename = "" Then
            GoTo load_file
        End If
        Unload Form2
        If Validate_load(ICOSAEDRON_data(Hexagono_index).filename) = False Then
            ICOSAEDRON_data(Hexagono_index).filename = App.Path & "\Error.Jpg"
        End If
        Form2.Original.Picture = LoadPicture(ICOSAEDRON_data(Hexagono_index).filename)
        Form2.Move 0, 0, ICOSAEDRON_data(Hexagono_index).Width, ICOSAEDRON_data(Hexagono_index).Height
        Form2.Image2.Width = ICOSAEDRON_data(Hexagono_index).Target_Width
        Form2.Image2.Height = ICOSAEDRON_data(Hexagono_index).Target_Height
        Form2.Image2.Top = ICOSAEDRON_data(Hexagono_index).Target_Top
        Form2.Image2.Left = ICOSAEDRON_data(Hexagono_index).Target_Left
        BorderW = ICOSAEDRON_data(Hexagono_index).BorderW
        BorderH = ICOSAEDRON_data(Hexagono_index).BorderH
        Form2.Show
        For i = 1 To ICOSAEDRON_data(Hexagono_index).Qt_text
            Load Form2.Label1(i)
            With ICOSAEDRON_data(Hexagono_index)
                Form2.Label1(i).Caption = .ICOSAEDRON_Text(i).Text
                Form2.Label1(i).Font = .ICOSAEDRON_Text(i).FontName
                Form2.Label1(i).FontSize = .ICOSAEDRON_Text(i).FontSize
                Form2.Label1(i).Top = .ICOSAEDRON_Text(i).Top
                Form2.Label1(i).Left = .ICOSAEDRON_Text(i).Left
                Form2.Label1(i).Width = .ICOSAEDRON_Text(i).Width
                Form2.Label1(i).Height = .ICOSAEDRON_Text(i).Height
                Form2.Label1(i).Forecolor = .ICOSAEDRON_Text(i).Forecolor
                Form2.Label1(i).Visible = True
            End With
        Next i
        
      Else
load_file:
        Bkp_filename = ICOSAEDRON_data(Hexagono_index).filename
        Hexagono_filename = FileDialog(Me, False, "Open Picture", "All|*.bmp;*.jpg;*.gif|Bitmap|*.bmp|Jasc|*.jpg|Compuserv|*.gif", "")
        If Hexagono_filename <> "" Then
            Unload Form2
            BorderW = 100
            BorderH = 100
            Form2.Image2.Width = 990
            Form2.Image2.Height = 960
    
            Form2.Original.Picture = LoadPicture(Hexagono_filename)
            Form2.Move 0, 0, Form2.Original.Width, Form2.Original.Height
            Do While Form2.Width > Screen.Width Or Form2.Height > Screen.Height
                Form2.Width = Form2.Width * 0.99
                Form2.Height = Form2.Height * 0.99
            Loop
            Form2.Show
            Form2.Image2.Move ((Form2.ScaleWidth / 2) - Form2.Image2.Width / 2), ((Form2.ScaleHeight / 2) - Form2.Image2.Height / 2)
            ICOSAEDRON_data(Hexagono_index).filename = Hexagono_filename
            ICOSAEDRON_data(Hexagono_index).Qt_text = 0
            Erase ICOSAEDRON_data(Hexagono_index).ICOSAEDRON_Text
          Else
            ICOSAEDRON_data(Hexagono_index).filename = Bkp_filename
        End If
    End If

End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If (Picture2.Point(x, y) And &HFF) > 20 Then
        Exit Sub
    End If

    If Button = 1 Then

    End If

    'Debug.Print Picture2.Point(X, Y) And &HFF;

End Sub

Private Sub Print_Icosaedron_Click()

    Printer.PaintPicture Picture4.Image, 500, 500, Picture4.Width, Picture4.Height
    Printer.EndDoc

End Sub

Private Sub Save_Click()

  Dim Free As Integer
  Dim x As String
  Dim s As Integer

    Icosaedron_filename = FileDialog(Me, True, "Save Icosaedron", "Icosaedron File|*.edr", " ", ".edr")
    If Icosaedron_filename = "" Then
        Exit Sub
    End If
    Free = FreeFile
    Open Icosaedron_filename For Binary As Free
    Put #Free, 1, ICOSAEDRON_data
    Close Free
    x = Icosaedron_filename
    s = InStr(x, ".edr")
    Mid$(x, s) = ".bmp"
    SavePicture Picture4.Image, x

End Sub

Private Sub Make_fill(x, y)
  
  Dim mbrush As Long
   
    mbrush = CreateSolidBrush(FillColor)
    SelectObject Picture4.hdc, mbrush
    ExtFloodFill Picture4.hdc, x, y, GetPixel(Picture4.hdc, x, y), FLOODFILLSURFACE
    DeleteObject mbrush
    Picture4.Refresh

End Sub

Private Sub Timer1_Timer()

  Static x As Integer

    Image1.Picture = Image2(x).Picture
    x = x + 1
    If x = 51 Then
        x = 1
    End If

End Sub


Private Function Validate_load(x As String)
On Error GoTo erro

If Dir(x) <> "" Then
    Validate_load = True
sair:
Exit Function
End If

erro:
Resume sair


End Function

Private Sub VScroll1_Change()
Picture4.Move Picture4.Left, -VScroll1.Value
End Sub
