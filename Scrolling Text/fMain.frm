VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "About Box"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar ScrollAlignment 
      Height          =   330
      LargeChange     =   10
      Left            =   1470
      Max             =   3
      Min             =   1
      TabIndex        =   7
      Top             =   5565
      Value           =   2
      Width           =   5685
   End
   Begin VB.HScrollBar Speed 
      Height          =   330
      LargeChange     =   10
      Left            =   3360
      Max             =   100
      Min             =   1
      TabIndex        =   4
      Top             =   4620
      Value           =   35
      Width           =   3795
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   645
      Left            =   105
      TabIndex        =   3
      Top             =   4620
      Width           =   1590
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "&Stop"
      Height          =   645
      Left            =   1785
      TabIndex        =   2
      Top             =   4620
      Width           =   1485
   End
   Begin VB.TextBox txtText 
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "fMain.frx":0000
      Top             =   3360
      Width           =   7050
   End
   Begin VB.PictureBox picText 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   3165
      Left            =   105
      ScaleHeight     =   207
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   473
      TabIndex        =   0
      Top             =   105
      Width           =   7155
   End
   Begin VB.Label Label3 
      Caption         =   "Alignment"
      Height          =   225
      Left            =   315
      TabIndex        =   8
      Top             =   5565
      Width           =   1485
   End
   Begin VB.Label Label2 
      Caption         =   "Slower"
      Height          =   225
      Left            =   6405
      TabIndex        =   6
      Top             =   5040
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "Faster"
      Height          =   225
      Left            =   3360
      TabIndex        =   5
      Top             =   5040
      Width           =   750
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ScrollDC Lib "user32" (ByVal hdc As Long, ByVal dx As Long, ByVal dy As Long, lprcScroll As RECT, lprcClip As RECT, ByVal hrgnUpdate As Long, lprcUpdate As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Type RECT
    Left    As Long
    Top     As Long
    Right   As Long
    Bottom  As Long
End Type


Private TextLine() As String
Private Index As Long
Private RText As RECT
Private RClip As RECT
Private RUpdate As RECT

Private TimeDelay As Long
Private Scrolling As Boolean
Private ScrollSpeed As Long

Private Alignment As Long


Private Sub cmdStart_Click()

    Scrolling = True
    TextLine = Split(txtText.Text, vbCrLf)
    
    With picText
        
        Do While Scrolling
        
            If GetTickCount - TimeDelay > ScrollSpeed Then
                TimeDelay = GetTickCount

                If RText.Bottom < .ScaleHeight Then
                    OffsetRect RText, 0, .TextHeight(vbNullString)
                    Index = Index + 1
                End If
                
                If Index > UBound(TextLine) Then Exit Do

                DrawText .hdc, Trim(TextLine(Index)), Len(Trim(TextLine(Index))), RText, Alignment
                OffsetRect RText, 0, -1
                ScrollDC .hdc, 0, -1, RClip, RClip, 0, RUpdate
                
                picText.Line (0, .ScaleHeight - 1)-(.ScaleWidth, .ScaleHeight - 1), .BackColor
            End If
            
            DoEvents
        Loop
    End With
    
    If Scrolling Then: Index = 0: cmdStart_Click
        
End Sub


Private Sub cmdStop_Click()
    Scrolling = False
End Sub


Private Sub Form_Load()
    
    txtText.Text = "This program is a simple demonstration that works quickly, easily and efficiently " & vbCrLf & _
                   "to scroll text in a picbox" & vbCrLf & vbCrLf & _
                   "------------------------------------------------------------------------" & vbCrLf & vbCrLf & _
                   "IRC Servers: Undernet" & vbCrLf & _
                   "IRC Channels: #VB,#VBHelp, #VB.GR" & vbCrLf & _
                   "IRC Nick: IRBMe" & vbCrLf & vbCrLf & _
                   "------------------------------------------------------------------------" & vbCrLf & vbCrLf & _
                   "EMail address: Imgonnadothingsmyway@Hotmail.com" & vbCrLf & vbCrLf & _
                   "------------------------------------------------------------------------" & vbCrLf & vbCrLf & _
                   "Please leave your feedback and votes. Thank you for downloading" & vbCrLf & vbCrLf & _
                   "------------------------------------------------------------------------" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
    
    With picText
        SetRect RClip, 0, 1, .ScaleWidth, .ScaleHeight
        SetRect RText, 0, .ScaleHeight, .ScaleWidth, .ScaleHeight + .TextHeight("")
    End With
    
    Speed_Change
    Alignment = &H1

End Sub


Private Sub Form_Unload(Cancel As Integer)
    Scrolling = False
    End
End Sub


Private Sub ScrollAlignment_Change()

    Select Case ScrollAlignment.Value
        Case 1: Alignment = &H0
        Case 2: Alignment = &H1
        Case 3: Alignment = &H2
    End Select

End Sub

Private Sub Speed_Change()
    ScrollSpeed = Speed.Value
End Sub
