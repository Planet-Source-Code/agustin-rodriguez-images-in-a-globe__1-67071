VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Edit"
   ClientHeight    =   6315
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   6750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   5940
      Left            =   4185
      Picture         =   "form2.frx":0000
      ScaleHeight     =   5880
      ScaleWidth      =   5790
      TabIndex        =   5
      Top             =   720
      Visible         =   0   'False
      Width           =   5850
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   885
      Left            =   675
      ScaleHeight     =   59
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   222
      TabIndex        =   3
      Top             =   5070
      Visible         =   0   'False
      Width           =   3330
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   450
      Left            =   1275
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   4410
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox Original 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   5340
      ScaleHeight     =   495
      ScaleWidth      =   645
      TabIndex        =   0
      Top             =   4830
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   5940
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Text"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000001&
      Height          =   495
      Index           =   0
      Left            =   3180
      TabIndex        =   1
      Top             =   2775
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image2 
      Enabled         =   0   'False
      Height          =   5040
      Left            =   330
      Picture         =   "form2.frx":227E
      Stretch         =   -1  'True
      Top             =   330
      Width           =   5070
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RGB_Color As String
Private Control As Integer
Private OldY As Long
Private OldX As Long
Private MoveControl As Boolean

Private Const WM_MOUSEWHEEL       As Long = &H20A
Private sc          As cSuperClass
Implements iSuperClass

Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private PT As POINTAPI
Private capture As Integer
Private XX As Long
Private YY As Long
Private Done_hexagono As Boolean

Private Sub Command1_Click()

End Sub

Private Sub Form_DblClick()

    Load Label1(Label1.Count)
    Label1(Label1.Count - 1).Forecolor = Label1(Label1.Count - 2).Forecolor
    Label1(Label1.Count - 1).Visible = True
    GetCursorPos PT
    Label1(Label1.Count - 1).Move PT.x * Screen.TwipsPerPixelX - Left, PT.y * Screen.TwipsPerPixelY - Top - 500

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  Dim Ret As SelectedColor
  Dim i As Integer
  
    On Error GoTo erro

    Select Case KeyCode
      Case 27
        If Text1.Visible Then
            Text1.Visible = False
            Exit Sub
        End If
        If ICOSAEDRON_data(Hexagono_index).Width = 0 Then
            ICOSAEDRON_data(Hexagono_index).filename = ""
        End If
        Unload Me

      Case 13
        If Text1.Visible Then
            Text1.Visible = False
            Exit Sub
        End If
    
    For i = 1 To Label1.Count - 1
        If Label1(i).Visible Then
            CurrentX = Label1(i).Left
            CurrentY = Label1(i).Top
            Font = Label1(i).Font
            FontSize = Label1(i).FontSize
            FontBold = Label1(i).FontBold
            Forecolor = Label1(i).Forecolor
            Print Label1(i).Caption
            With ICOSAEDRON_data(Hexagono_index).ICOSAEDRON_Text(i)
                .Text = Label1(i).Caption
                .FontName = Label1(i).FontName
                .FontSize = Label1(i).FontSize
                .Forecolor = Label1(i).Forecolor
                .Left = Label1(i).Left
                .Top = Label1(i).Top
                .Width = Label1(i).Width
                .Height = Label1(i).Height
            End With
        End If
    Next i
    
        Form1.Picture1.PaintPicture Image, 0, 0, Form1.Picture1.Width, Form1.Picture1.Height, Image2.Left / Screen.TwipsPerPixelX, Image2.Top / Screen.TwipsPerPixelY, Image2.Width, Image2.Height

        Form1.Picture1.PaintPicture Picture2.Picture, 0, 0, Form1.Picture1.Width, Form1.Picture1.Height, 0, 0, Picture2.Width, Picture2.Height, vbSrcAnd

        Form1.Command1 = True

        With ICOSAEDRON_data(Hexagono_index)
            .filename = Hexagono_filename
            .BorderH = BorderH
            .BorderW = BorderW
            .Target_Top = Image2.Top
            .Target_Left = Image2.Left
            .Target_Width = Image2.Width
            .Target_Height = Image2.Height
            .Width = Width
            .Height = Height
            .Qt_text = Label1.Count - 1
        End With

        Done_hexagono = True

        Unload Me

      Case 114
        Ret = ShowColor(Me.hwnd, False)
        If Ret.bCanceled Then
            Exit Sub
        End If
        BackColor = Ret.oSelectedColor
        Form_Resize
       
       Case 115
        MoveControl = False
        Ret = ShowColor(Me.hwnd, False)
        If Ret.bCanceled Then
            Exit Sub
        End If
        Label1(Control).Forecolor = Ret.oSelectedColor
        Exit Sub
    
      Case 112
        BorderW = BorderW + Width / 100
        BorderH = BorderH + Height / 100
        DoEvents
        Form_Resize

      Case 113
        BorderW = BorderW - Width / 100
        BorderH = BorderH - Height / 100
        DoEvents
        Form_Resize
    
      Case 46
        Unload Label1(Control)
        Control = 0

      Case 107
        Image2.Move Image2.Left - 10, Image2.Top - 10, Image2.Width + 10 * 2, Image2.Height + 10 * 2
      Case 109
        If Image2.Width < 100 Or Image2.Height < 100 Then
            Exit Sub
        End If
        Image2.Move Image2.Left + 10, Image2.Top + 10, Image2.Width - 10 * 2, Image2.Height - 10 * 2
    End Select

sair:

Exit Sub

erro:
    Resume sair

End Sub

Private Sub Form_Load()

    Set sc = New cSuperClass
  
    With sc
        Call .AddMsg(WM_MOUSEWHEEL)
        Call .Subclass(hwnd, Me)
    End With

    Done_hexagono = False
    Control = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Control = 0
    Image2.ZOrder 0
    capture = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If capture Then
        GetCursorPos PT
        Image2.Move (x - Image2.Width / 2), (y - Image2.Height / 2)
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    capture = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Done_hexagono = False Then
'        ICOSAEDRON_data(Hexagono_index).filename = Bkp_filename
    End If

End Sub

Public Sub Form_Resize()

  Dim x As Long
  Dim y As Long

    On Error GoTo erro
    If Original Then
        x = Width - ScaleWidth
        y = Height - ScaleHeight
        Cls
        PaintPicture Original, BorderW, BorderH, (Width - x) - BorderW * 2, (Height - y) - BorderH * 2, 0, 0, Original.ScaleWidth, Original.ScaleHeight, vbSrcCopy
    End If

sair:

Exit Sub

erro:

End Sub

Private Sub iSuperClass_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

    Select Case uMsg
  
      Case WM_MOUSEWHEEL
  
        If wParam < 0 Then
    
            If wParam = -7864316 And Control Then
                'Debug.Print "DOWN+Shift";
            End If
    
            If Control Then
                Label1(Control).Move Label1(Control).Left - Label1(Control).Width / 20, Label1(Control).Top - Label1(Control).Height / 20
            
                Label1(Control).Width = Label1(Control).Width * 1.1
                Label1(Control).Height = Label1(Control).Height * 1.1
            
                Label1(Control).FontSize = Label1(Control).FontSize * 1.1
                Exit Sub
            End If
        
            Form_KeyDown 107, 0
            Form_KeyDown 107, 0
            Form_KeyDown 107, 0
          Else
    
            If wParam = 7864324 And Control Then
                'Debug.Print "UP+Shift";
            End If
    
            If Control Then
                Label1(Control).Move Label1(Control).Left + Label1(Control).Width / 20, Label1(Control).Top + Label1(Control).Height / 20
                Label1(Control).Width = Label1(Control).Width / 1.1
                Label1(Control).Height = Label1(Control).Height / 1.1
                Label1(Control).FontSize = Label1(Control).FontSize / 1.1
                Exit Sub
            End If
            Form_KeyDown 109, 0
            Form_KeyDown 109, 0
            Form_KeyDown 109, 0
        End If
        
    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set sc = Nothing

End Sub

Private Sub Label1_DblClick(Index As Integer)

    Text1.Move Label1(Index).Left - Label1(Index).Width / 4, Label1(Index).Top, Label1(Index).Width * 1.2, Label1(Index).Height
    Text1.Text = Label1(Index).Caption
    Text1.FontSize = Label1(Index).FontSize
    Text1.FontName = Label1(Index).FontName
    Text1.FontBold = Label1(Index).FontBold
    Text1.Visible = True
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1)

End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim Ret As SelectedColor
  Dim ret_font As SelectedFont

    Picture1.Cls
    Picture1.Font = Label1(Index).Font
    Picture1.FontBold = Label1(Index).FontBold
    Picture1.FontSize = Label1(Index).FontSize

    Picture1.Width = Label1(Index).Width
    Picture1.Height = Label1(Index).Height
    Picture1.Print Label1(Index).Caption

    Label2.Caption = Picture1.Point(x / Screen.TwipsPerPixelX, y / Screen.TwipsPerPixelY)
    If Label2.Caption = 0 Then

        Control = Index
        OldY = y
        OldX = x
        MoveControl = True
    End If
    

    If Shift = 2 Then
        MoveControl = False
        ret_font = ShowFont(Me.hwnd, Label1(Control).FontName, True)
        If ret_font.bCanceled Then
            Exit Sub
        End If
        Label1(Control).FontName = ret_font.sSelectedFont
        Label1(Control).FontBold = ret_font.bBold
        Label1(Control).FontSize = ret_font.nSize
        Exit Sub
    End If

    If Button = 2 Then
        Load Label1(Label1.Count)
        Label1(Label1.Count - 1).Caption = Label1(Index).Caption
        Label1(Label1.Count - 1).Forecolor = Label1(Index).Forecolor
        Label1(Label1.Count - 1).FontSize = Label1(Index).FontSize
        Label1(Label1.Count - 1).FontBold = Label1(Index).FontBold
        Label1(Label1.Count - 1).Caption = Label1(Index).Caption
        Label1(Label1.Count - 1).FontName = Label1(Index).FontName
        Label1(Label1.Count - 1).Visible = True
        Label1(Label1.Count - 1).Move Label1(Index).Left + 100, Label1(Index).Top + 100
        Exit Sub
    End If

End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If MoveControl Then
        Label1(Index).Top = Label1(Index).Top - OldY + y
        Label1(Index).Left = Label1(Index).Left - OldX + x
    End If

End Sub

Private Sub Label1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    MoveControl = False
    Label1(Index).ZOrder 1

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Label1(Control).Caption = Text1.Text
    End If
   

End Sub

Private Function Inc_RGB(x As Long, Index As Integer) As Long

  Dim r As String
  Dim g As String
  Dim B As String
  Dim RGBn As String
  Dim rr As Integer
  Dim gg As Integer
  Dim bb As Integer

    r = Mid$(Right$("00000000" + Hex$(x), 8), 3, 2)
    g = Mid$(Right$("00000000" + Hex$(x), 8), 5, 2)
    B = Mid$(Right$("00000000" + Hex$(x), 8), 7, 2)
    rr = Val("&h" + r)
    gg = Val("&h" + g)
    bb = Val("&h" + B)

    Select Case Index
      Case 1
        rr = rr + 1 And 255
    
      Case 2
        gg = gg + 1 And 255
    
      Case 3
        bb = bb + 1 And 255
    
    End Select
    Inc_RGB = RGB(rr, gg, bb)
    'Debug.Print Inc_RGB

End Function


