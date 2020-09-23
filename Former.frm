VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paradise Form Designer"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9570
   Icon            =   "Former.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Creation 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5700
      Left            =   9240
      Picture         =   "Former.frx":030A
      ScaleHeight     =   5700
      ScaleWidth      =   9600
      TabIndex        =   1
      Top             =   5400
      Width           =   9600
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   29
         Text            =   "Form1"
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox SITChk 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1920
         TabIndex        =   27
         Top             =   2040
         Value           =   1  'Checked
         Width           =   220
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   25
         Text            =   "24000"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   23
         Text            =   "32000"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CheckBox CCChk 
         Caption         =   "Check1"
         Height          =   255
         Left            =   1920
         TabIndex        =   22
         Top             =   1280
         Value           =   1  'Checked
         Width           =   220
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         ItemData        =   "Former.frx":3B99E
         Left            =   5880
         List            =   "Former.frx":3B9B1
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   2040
         Width           =   3255
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1380
         Left            =   4440
         TabIndex        =   17
         Top             =   600
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1920
         TabIndex        =   16
         Text            =   "Form1"
         Top             =   840
         Width           =   1935
      End
      Begin VB.Shape Shape4 
         Height          =   495
         Left            =   7440
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Create the Form"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7560
         TabIndex        =   33
         Top             =   4440
         Width           =   1575
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   $"Former.frx":3B9EA
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   480
         TabIndex        =   32
         Top             =   2880
         Width           =   8655
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tips on Creating a Shaped form"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   2520
         Width           =   8895
      End
      Begin VB.Shape Shape3 
         Height          =   1455
         Left            =   360
         Top             =   2760
         Width           =   8895
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Form Name:"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   30
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Show In Taskbar:"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   28
         Top             =   2070
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         Height          =   1935
         Left            =   4320
         Top             =   480
         Width           =   4935
      End
      Begin VB.Shape Shape1 
         Height          =   1935
         Left            =   360
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Height:"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Width:"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   24
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Clip Controls:"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   21
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Overlay Method:"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   19
         Top             =   2085
         Width           =   1335
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Polygon Properties"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4320
         TabIndex        =   18
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "(Designed)"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1920
         TabIndex        =   15
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Form BackGround:"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Form Caption:"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Form Standard Properties"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   3855
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   120
         Top             =   5040
         Width           =   1215
      End
   End
   Begin VB.PictureBox Design 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   5700
      Left            =   0
      Picture         =   "Former.frx":3BEA0
      ScaleHeight     =   5700
      ScaleWidth      =   9600
      TabIndex        =   0
      Top             =   0
      Width           =   9600
      Begin MSComDlg.CommonDialog CDlg 
         Left            =   4440
         Top             =   5160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1980
         Left            =   360
         TabIndex        =   9
         Top             =   2520
         Width           =   3375
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   4095
         Left            =   3960
         ScaleHeight     =   4065
         ScaleWidth      =   5145
         TabIndex        =   2
         Top             =   360
         Width           =   5175
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   5700
            Left            =   0
            ScaleHeight     =   5700
            ScaleWidth      =   9600
            TabIndex        =   10
            Top             =   0
            Width           =   9600
            Begin VB.Shape Shape5 
               BorderStyle     =   3  'Dot
               Height          =   1815
               Left            =   1080
               Top             =   1560
               Visible         =   0   'False
               Width           =   3015
            End
            Begin VB.Line Line1 
               BorderColor     =   &H000000FF&
               Visible         =   0   'False
               X1              =   3960
               X2              =   2280
               Y1              =   600
               Y2              =   1200
            End
         End
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Use Arrow Keys to move in the BackGround"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3960
         TabIndex        =   11
         Top             =   4560
         Width           =   5175
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remove Selected Polygon"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Create New Shape Polygon"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   1800
         Width           =   3375
      End
      Begin VB.Label L2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Form Shape Design"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remove BackGround Picture"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add/Change BackGround Picture "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label L1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BackGround Design"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3615
      End
      Begin VB.Image Crlnk 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1320
         Top             =   5040
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type PointXY
 X As Integer
 Y As Integer
End Type

Private Type BLENDFUNCTION
  BlendOp    As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Const AC_SRC_OVER = &H0
Const AC_SRC_ALPHA = &H1

Private Declare Function AlphaBlend Lib "MSImg32" (hdcDest As Long, nXOriginDest As Integer, _
                         nYOriginDest As Integer, nWidthDest As Integer, nHeightDest As Integer, _
                         hdcSrc As Integer, nXOriginSrc As Integer, nYOriginSrc As Integer, _
                         nWidthSrc As Integer, nHeightSrc As Integer, BLENDFNUC As Long) As Integer
Private Declare Sub RtlCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Points(55, 490) As PointXY
Private CurPG, CurP, Started As Boolean, Started2 As Boolean, xx, yy, SelPG, CP
Private OvlMethod(55) As Integer

Private Sub Combo1_Click()
 OvlMethod(List2.ListIndex + 1) = Combo1.ListIndex
End Sub

Private Sub Crlnk_Click()
 Design.Visible = False
 Creation.Visible = True
End Sub

Private Sub Form_Load()
 Creation.Visible = False
 Creation.Top = 0
 Creation.Left = 0
 For i = 1 To 55
  Points(i, 0).X = -1
  OvlMethod(i) = 0
 Next i
 Combo1.ListIndex = OvlMethod(0)
End Sub

Private Sub Image1_Click()
 Creation.Visible = False
 Design.Visible = True
End Sub

Private Sub Label2_Click()
 If Started = True Then MsgBox "Please Close the Current Spline before selecting a background picture.": Exit Sub
 CDlg.FileName = ""
 CDlg.DialogTitle = "Open Background Picture"
 CDlg.Filter = "Bitmap Files (*.bmp)|*.bmp|Jpeg Files (*.jpg;*.jpe)|*.jpg;*.jpe|GIF Files (*.gif)|*.gif|Icon and Cursor Files (*.ico;*.cur)|*.ico;*.cur|All Picture Files (*.bmp;*.jpg;*.jpe;*.gif;*.ico;*.cur)|*.bmp;*.jpg;*.jpe;*.gif;*.ico;*.cur|All Files (*.*)|*.*"
 CDlg.ShowOpen
 If CDlg.FileName <> "" Then
  Picture2.Picture = LoadPicture(CDlg.FileName)
  Picture2.Tag = CDlg.FileName
  RedrawAll
 End If
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label2.BackColor = RGB(100, 100, 200)
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label2.BackColor = RGB(255, 255, 255)
End Sub

Private Sub Label21_Click()
1 CDlg.FileName = ""
 CDlg.Filter = "Visual Basic Form Files (*.frm)|*.frm|All Files (*.*)|*.*"
 CDlg.ShowSave
 If CDlg.FileName <> "" Then
  'Check if the file exists. If yes, prompt the user!
  On Error GoTo 2
  Open CDlg.FileName For Input As #1
  Close #1
  result = MsgBox("The File Already Exists. Would you like to Replace it?", _
                  vbYesNoCancel, "Create Form File")
  If result = vbNo Then GoTo 1
  If result = vbCancel Then Exit Sub
2 Close #1
  
  'If the file creation is permitted, start creating the file!
  If Val(Text2.Text) < 1000 Then Text2.Text = "1000"
  If Val(Text3.Text) < 500 Then Text2.Text = "500"
  Open CDlg.FileName For Output As #1
   Print #1, "VERSION 5.00"
   Print #1, "Begin VB.Form " + Text4.Text
   Print #1, "  Caption         = """ + Text1.Text + """"
   Print #1, "  ClientHeight    =" + Text3.Text
   Print #1, "  ClientLeft      = 60"
   Print #1, "  ClientTop       = 345"
   Print #1, "  ClientWidth     =" + Text2.Text
   Print #1, "  ClipControls    =" + Str$(CCChk.Value)
   Print #1, "  LinkTopic       =""" + Text4.Text + """"
   Print #1, "  ScaleHeight     =" + Text3.Text
   Print #1, "  ScaleWidth      =" + Text2.Text
   Print #1, "  ShowInTaskbar   =" + Str$(SITChk.Value)
   Print #1, "  StartUpPosition = 3 'Windows Default"
   Print #1, "End"
   Print #1, "Attribute VB_Name = " + Text4.Text
   Print #1, "Attribute VB_GlobalNameSpace = False"
   Print #1, "Attribute VB_Creatable = False"
   Print #1, "Attribute VB_PredeclaredId = True"
   Print #1, "Attribute VB_Exposed = False"
   Print #1,
   Print #1, "Private Type PointXY"
   Print #1, "  X As Long"
   Print #1, "  Y As Long"
   Print #1, "End Type"
   Print #1,
   Print #1, "Private Const RGN_AND = 1"
   Print #1, "Private Const RGN_COPY = 5"
   Print #1, "Private Const RGN_DIFF = 4"
   Print #1, "Private Const RGN_OR = 2"
   Print #1, "Private Const RGN_XOR = 3"
   Print #1,
   Print #1, "Private Declare Function CreateRectRgn Lib ""gdi32"" Alias ""CreateRectRgn"" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long"
   Print #1, "Private Declare Function CombineRgn Lib ""gdi32"" Alias ""CombineRgn"" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long"
   Print #1, "Private Declare Function SetWindowRgn Lib ""user32"" Alias ""SetWindowRgn"" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long"
   Print #1, "Private Declare Function CreatePolygonRgn Lib ""gdi32"" Alias ""CreatePolygonRgn"" (lpPoint As PointXY, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long"
   Print #1,
   Print #1, "Private Sub Form_Activate()"
   If Picture2.Tag <> "" Then Print #1, "  Me.Picture = LoadPicture(""" + Picture2.Tag + """)"
   Print #1, "  Call DrawShape()"
   Print #1, "End Sub"
   Print #1,
   Print #1, "'And here's the precedure which creates the shape of the form"
   Print #1,
   Print #1, "Private Sub DrawShape()"
   Print #1, "  Dim Points(490) As PointXY"
   Print #1, "  hRgn& = CreateRectRgn(0, 0, 0, 0)"
   For j = 1 To 55
    Select Case OvlMethod(j)
     Case 0
      R$ = "RGN_OR"
     Case 1
      R$ = "RGN_COPY"
     Case 2
      R$ = "RGN_AND"
     Case 3
      R$ = "RGN_DIFF"
     Case 4
      R$ = "RGN_XOR"
     Case Else
      'That's impossible !!
    End Select
    If Points(j, 0).X = -1 Then GoTo 3
    For i = 0 To 490
     If Points(j, i).X = -1 Then Exit For
     Print #1, "  Points(" + Right$(Str$(i), Len(Str$(i)) - 1) + ").X =" + Str$(Points(j, i).X / 16)
     Print #1, "  Points(" + Right$(Str$(i), Len(Str$(i)) - 1) + ").Y =" + Str$(Points(j, i).Y / 16)
    Next i
    Print #1, "  hPGRgn& = CreatePolygonRgn(Points(0)," + Str$(i) + ", 1)"
    Print #1, "  CombineRgn hRgn&, hRgn&, hPGRgn&, " + R$
3  Next j
   Print #1, "  SetWindowRgn Me.hWnd, hRgn&, True"
   Print #1, "End sub"
  Close #1
 End If
End Sub

Private Sub Label3_Click()
 If Started = True Then MsgBox "Please Close the Current Spline before removing the background picture.": Exit Sub
 Picture2.Picture = LoadPicture("")
 Picture2.Tag = ""
 RedrawAll
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label3.BackColor = RGB(100, 100, 200)
End Sub

Private Sub Label3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label3.BackColor = RGB(255, 255, 255)
End Sub

Private Sub Label5_Click()
 Started = True
 CurPG = CurPG + 1
 List1.AddItem "Shape No." + Str$(CurPG)
 List2.AddItem "Shape No." + Str$(CurPG)
 'Dim b As BLENDFUNCTION
 'b.AlphaFormat = AC_SRC_ALPHA
 'b.BlendOp = AC_SRC_OVER
 'b.BlendFlags = 0
 'b.SourceConstantAlpha = 255
 'RtlCopyMemory ByVal bl&, ByVal b, Len(b)
 'AlphaBlend Picture2, 0, 0, 100, 100, Design.hDC, 0, 0, 100, 100, ByVal bl
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label5.BackColor = RGB(100, 100, 200)
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label5.BackColor = RGB(255, 255, 255)
End Sub

Private Sub Label6_Click()
 If Right$(List1.List(List1.ListIndex), 18) = " (Removed/Expired)" Then Exit Sub
 If List1.ListIndex = -1 Then Exit Sub
 List1.List(List1.ListIndex) = List1.List(List1.ListIndex) + " (Removed/Expired)"
 List2.List(List1.ListIndex) = List2.List(List1.ListIndex) + " (Removed/Expired)"
 If SelPG = List1.ListIndex + 1 Then SelPG = SelPG - 1
 Shape5.Visible = False
 Points(List1.ListIndex + 1, 0).X = -1
 Picture2.Cls
 RedrawAll
End Sub

Private Sub Label6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label6.BackColor = RGB(100, 100, 200)
End Sub

Private Sub Label6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Label6.BackColor = RGB(255, 255, 255)
End Sub

Private Sub List1_Click()
 If Started = True Then MsgBox "Please close the current spline before selecting any other one.": Exit Sub
 If List1.SelCount > 0 Then
  SelPG = List1.ListIndex + 1
  DoEvents
  Picture2.Cls
  RedrawAll
 End If
End Sub

Private Sub List2_Click()
 If List2.ListIndex >= 0 Then Combo1.ListIndex = OvlMethod(List2.ListIndex + 1)
End Sub

Private Sub Picture2_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 40 And Picture2.Top + Picture2.Height > Picture1.Height Then Picture2.Top = Picture2.Top - 100
 If KeyCode = 38 And Picture2.Top < 0 Then Picture2.Top = Picture2.Top + 100
 If KeyCode = 39 And Picture2.Left + Picture2.Width > Picture1.Width Then Picture2.Left = Picture2.Left - 100
 If KeyCode = 37 And Picture2.Left < 0 Then Picture2.Left = Picture2.Left + 100
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Started = True Then
  If Button = 1 Then
   Points(CurPG, CurP).X = X
   Points(CurPG, CurP).Y = Y
   Line1.Visible = True
   Line1.X1 = X
   Line1.Y1 = Y
   If CurP = 0 Then
    Picture2.Line (X - 50, Y - 50)-(X + 50, Y + 50), RGB(0, 255, 0), B
    Picture2.PSet (X, Y)
   Else
    Picture2.Line -(X, Y), RGB(255, 0, 0)
    Picture2.Line (X - 50, Y - 50)-(X + 50, Y + 50), RGB(0, 255, 0), B
    Picture2.PSet (X, Y)
   End If
   CurP = CurP + 1
  ElseIf Button = 2 Then
   Line1.Visible = False
   Picture2.Line -(Points(CurPG, 0).X, Points(CurPG, 0).Y), RGB(255, 0, 0)
   Points(CurPG, CurP).X = Points(CurPG, 0).X
   Points(CurPG, CurP).Y = Points(CurPG, 0).Y
   CurP = CurP + 1
   Points(CurPG, CurP).X = -1
   Started = False
   SelPG = CurPG
   CurP = 0
   Picture2.Cls
   RedrawAll
  Else
  
  End If
 Else
  Started2 = True
  CP = -1
  For i = 0 To 490
   If Points(SelPG, i).X = -1 Then Exit For
   px = Points(SelPG, i).X
   py = Points(SelPG, i).Y
   If X > px - 50 And X < px + 50 And Y > py - 50 And Y < py + 50 Then
    CP = i
    xx = X
    yy = Y
   End If
  Next i
  If CP = -1 Then Started2 = False
 End If
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Line1.X2 = X
 Line1.Y2 = Y
End Sub

Private Sub Picture2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Started2 = True Then
  Picture2.Cls
  Points(SelPG, CP).X = X
  Points(SelPG, CP).Y = Y
  If Points(SelPG, CP + 1).X = -1 Then
   Points(SelPG, 0).X = X
   Points(SelPG, 0).Y = Y
  End If
  RedrawAll
  Started2 = False
 End If
End Sub

Private Sub RedrawAll()
  For j = 1 To 55
   If Points(j, 0).X = -1 Then GoTo 1
   px = Points(j, 0).X
   py = Points(j, 0).Y
   lx = px: ly = py: hx = 0: hy = 0
   Picture2.Line (px - 50, py - 50)-(px + 50, py + 50), RGB(0, 255, 0), B
   Picture2.PSet (px, py)
   For i = 1 To 490
    If Points(j, i).X = -1 Then Exit For
    px = Points(j, i).X
    py = Points(j, i).Y
    If j = SelPG Then
     If px > hx Then hx = px
     If py > hy Then hy = py
     If px < lx Then lx = px
     If py < ly Then ly = py
    End If
    Picture2.Line -(px, py), RGB(255, 0, 0)
    Picture2.Line (px - 50, py - 50)-(px + 50, py + 50), RGB(0, 255, 0), B
    Picture2.PSet (px, py)
   Next i
   If j = SelPG Then
    Shape5.Visible = True
    Shape5.Left = lx
    Shape5.Top = ly
    Shape5.Width = hx - lx
    Shape5.Height = hy - ly
   End If
1 Next j
End Sub
