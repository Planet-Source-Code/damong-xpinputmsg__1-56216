VERSION 5.00
Begin VB.Form FrmMsg 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CECECE&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraInput 
      BackColor       =   &H00CECECE&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   1440
      TabIndex        =   6
      Top             =   3000
      Width           =   5295
      Begin VB.TextBox TxtInput 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   20
         TabIndex        =   7
         Top             =   20
         Width           =   4095
      End
      Begin VB.Line LineRight 
         BorderColor     =   &H00F0D0B0&
         X1              =   4800
         X2              =   4800
         Y1              =   240
         Y2              =   720
      End
      Begin VB.Line LineLeft 
         BorderColor     =   &H00F0D0B0&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   720
      End
      Begin VB.Line LineDown 
         BorderColor     =   &H00F0D0B0&
         X1              =   600
         X2              =   4680
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line LineUp 
         BorderColor     =   &H00F0D0B0&
         X1              =   0
         X2              =   4680
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Label lbl_Cancel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3720
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.Image Cancel_Button 
      Height          =   255
      Left            =   3360
      OLEDropMode     =   1  'Manual
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image img_Icon 
      Height          =   255
      Left            =   120
      Picture         =   "FrmMsgInput.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   285
   End
   Begin VB.Image BottomLeft 
      Height          =   165
      Left            =   0
      Picture         =   "FrmMsgInput.frx":038A
      Top             =   2520
      Width           =   195
   End
   Begin VB.Image BottomRight 
      Height          =   210
      Left            =   4800
      Picture         =   "FrmMsgInput.frx":0584
      Top             =   2475
      Width           =   165
   End
   Begin VB.Image Bottom 
      Height          =   60
      Left            =   180
      Picture         =   "FrmMsgInput.frx":07BE
      Stretch         =   -1  'True
      Top             =   2625
      Width           =   4620
   End
   Begin VB.Image BRight 
      Height          =   2025
      Left            =   4905
      Picture         =   "FrmMsgInput.frx":0A62
      Stretch         =   -1  'True
      Top             =   450
      Width           =   60
   End
   Begin VB.Image BLeft 
      Height          =   2085
      Left            =   0
      Picture         =   "FrmMsgInput.frx":0CFC
      Stretch         =   -1  'True
      Top             =   450
      Width           =   60
   End
   Begin VB.Image TitleLeft 
      Height          =   450
      Left            =   0
      Picture         =   "FrmMsgInput.frx":0F99
      Top             =   0
      Width           =   240
   End
   Begin VB.Label lbl_Prompt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Store Your Text Here"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   75
      TabIndex        =   5
      Top             =   600
      Width           =   1830
   End
   Begin VB.Label lbl_No 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2160
      TabIndex        =   1
      Top             =   1560
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Icon_Cri 
      Height          =   480
      Left            =   120
      Picture         =   "FrmMsgInput.frx":13DA
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Icon_Exc 
      Height          =   480
      Left            =   1080
      Picture         =   "FrmMsgInput.frx":20A4
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Icon_Question 
      Height          =   480
      Left            =   1800
      Picture         =   "FrmMsgInput.frx":2D6E
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Icon_Info 
      Height          =   480
      Left            =   2760
      Picture         =   "FrmMsgInput.frx":3A38
      Top             =   2040
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image No_Button 
      Height          =   255
      Left            =   1680
      OLEDropMode     =   1  'Manual
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lbl_OK 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   780
      TabIndex        =   0
      Top             =   1800
      Width           =   255
   End
   Begin VB.Image Ok 
      Height          =   315
      Left            =   360
      OLEDropMode     =   1  'Manual
      Picture         =   "FrmMsgInput.frx":4702
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Image Ok_Press 
      Height          =   315
      Left            =   3105
      Picture         =   "FrmMsgInput.frx":5952
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Ok_Over 
      Height          =   315
      Left            =   1905
      Picture         =   "FrmMsgInput.frx":6BA2
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Ok_Static 
      Height          =   315
      Left            =   705
      Picture         =   "FrmMsgInput.frx":7DF2
      Top             =   1080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Cb_CLose 
      Height          =   195
      Index           =   3
      Left            =   2625
      Picture         =   "FrmMsgInput.frx":9042
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Cb_CLose 
      Height          =   195
      Index           =   2
      Left            =   2295
      Picture         =   "FrmMsgInput.frx":928C
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image Cb_CLose 
      Height          =   195
      Index           =   1
      Left            =   1965
      Picture         =   "FrmMsgInput.frx":94D6
      Top             =   600
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Image CloseButton 
      Height          =   195
      Left            =   4440
      Picture         =   "FrmMsgInput.frx":9720
      ToolTipText     =   "Close"
      Top             =   120
      Width           =   195
   End
   Begin VB.Image TitleRight 
      Height          =   450
      Left            =   4530
      Picture         =   "FrmMsgInput.frx":996A
      Top             =   0
      Width           =   435
   End
   Begin VB.Label Caption2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Xp MsgBox"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   505
      TabIndex        =   4
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label Caption1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Xp MsgBox"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   1035
   End
   Begin VB.Image Title 
      Height          =   450
      Left            =   210
      Picture         =   "FrmMsgInput.frx":A3FC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "FrmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim IsOverB As Boolean 'Ok & Yes Button
Dim IsOverC As Boolean 'Close Button
Dim IsOverN As Boolean 'No Button
Dim IsOverD As Boolean 'Cancel button

Sub MoveForm(FormX As Form, XX, YY, BT)
    Static OldX, OldY, Mf
    Dim MoveLeft, MoveTop

    MoveLeft = FormX.Left + XX - OldX
    MoveTop = FormX.Top + YY - OldY
    If BT = vbLeftButton Then
        If Mf = 0 Then
            FormX.Move MoveLeft, MoveTop
            Mf = 1
        Else
            Mf = 0
        End If
    Else
        If IsOverB = True Then
           IsOverB = False
           Ok.Picture = Ok_Static.Picture
        End If
        
        If IsOverC = True Then
           IsOverC = False
           CloseButton.Picture = Cb_CLose(1).Picture
        End If
        
        If IsOverN = True Then
           IsOverN = False
           No_Button.Picture = Ok_Static.Picture
        End If
        
        If IsOverD = True Then
            IsOverD = False
            Cancel_Button.Picture = Ok_Static.Picture
        End If

    End If
    OldX = XX
    OldY = YY
End Sub

Private Sub Caption2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call Title_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub CloseButton_Click()
On Error GoTo ErrEF
If EsMsgBox = True Then
    Select Case WhatButton
    Case 2
       EnterB = 3
    Case 3
       EnterB = 4
    End Select
Else
    ElTexto = ""
End If
 Unload Me
ErrEF:
End Sub

Private Sub CloseButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 CloseButton.Picture = Cb_CLose(3).Picture
End Sub

Private Sub CloseButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 IsOverC = True
 CloseButton.Picture = Cb_CLose(2).Picture
End Sub

Private Sub CloseButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 CloseButton.Picture = Cb_CLose(1).Picture
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyRight Then
    lbl_Prompt.ForeColor = Rnd(2000)
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 115 Or KeyAscii = 111 Or KeyAscii = 83 Or KeyAscii = 79 Then 'S o O
    Ok_Click
ElseIf KeyAscii = 110 Or KeyAscii = 78 Then 'N
    No_Button_Click
ElseIf KeyAscii = 99 Or KeyAscii = 67 Then 'C
    Cancel_Button_Click
End If
End Sub

Private Sub Form_Load()
 Title.ToolTipText = Caption2.Caption
 TitleRight.ToolTipText = Caption2.Caption
 TitleLeft.ToolTipText = Caption2.Caption
 
 IsOverB = False
 IsOverC = False
 IsOverN = False
 IsOverD = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IsOverB = True Then
   IsOverB = False
   Ok.Picture = Ok_Static.Picture
End If

If IsOverC = True Then
   IsOverC = False
   CloseButton.Picture = Cb_CLose(1).Picture
End If

If IsOverN = True Then
   IsOverN = False
   No_Button.Picture = Ok_Static.Picture
End If

If IsOverD = True Then
    IsOverD = False
    Cancel_Button.Picture = Ok_Static.Picture
End If
End Sub

Private Sub lbl_No_Click()
 Call No_Button_Click
End Sub

Private Sub lbl_No_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call No_Button_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lbl_No_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call No_Button_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lbl_No_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call No_Button_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub lbl_Cancel_Click()
 Call Cancel_Button_Click
End Sub

Private Sub lbl_Cancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call Cancel_Button_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lbl_Cancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call Cancel_Button_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lbl_Cancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call Cancel_Button_MouseUp(Button, Shift, X, Y)
End Sub


Private Sub lbl_OK_Click()
 Call Ok_Click
End Sub

Private Sub lbl_OK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call Ok_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lbl_OK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call Ok_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub lbl_OK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call Ok_MouseUp(Button, Shift, X, Y)
End Sub

Private Sub lbl_Yes_No_Click()
 Call No_Button_Click
End Sub

Private Sub lbl_Yes_No_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call No_Button_MouseDown(Button, Shift, X, Y)
End Sub

Private Sub lbl_Yes_No_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Call No_Button_MouseMove(Button, Shift, X, Y)
End Sub

Private Sub No_Button_Click()
If EsMsgBox = True Then
    EnterB = 3
Else
    ElTexto = ""
End If
 Unload Me
End Sub

Private Sub No_Button_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 No_Button.Picture = Ok_Press.Picture
End Sub

Private Sub No_Button_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 IsOverN = True
 No_Button.Picture = Ok_Over.Picture
End Sub

Private Sub No_Button_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 No_Button.Picture = Ok_Static.Picture
End Sub
Private Sub Cancel_Button_Click()
 EnterB = 4
 Unload Me
End Sub

Private Sub Cancel_Button_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Cancel_Button.Picture = Ok_Press.Picture
End Sub

Private Sub Cancel_Button_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 IsOverD = True
 Cancel_Button.Picture = Ok_Over.Picture
End Sub

Private Sub Cancel_Button_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Cancel_Button.Picture = Ok_Static.Picture
End Sub

Private Sub Ok_Click()
If EsMsgBox = True Then
    Select Case WhatButton
    Case 1
      EnterB = 1 'Ok Button
    Case 2
      EnterB = 2 'Yes Button
    Case 3
       EnterB = 2 'Yes Button
    End Select
Else
    ElTexto = TxtInput.Text
End If
 Unload Me
End Sub

Private Sub Ok_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Ok.Picture = Ok_Press.Picture
End Sub

Private Sub Ok_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 IsOverB = True
 Ok.Picture = Ok_Over.Picture
End Sub

Private Sub Ok_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Ok.Picture = Ok_Static.Picture
End Sub

Private Sub Title_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 MoveForm Me, X, Y, Button
End Sub


Private Sub TitleRight_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If IsOverB = True Then
   IsOverB = False
   Ok.Picture = Ok_Static.Picture
End If

If IsOverC = True Then
   IsOverC = False
   CloseButton.Picture = Cb_CLose(1).Picture
End If

If IsOverN = True Then
   IsOverN = False
   No_Button.Picture = Ok_Static.Picture
End If

If IsOverD = True Then
    IsOverD = False
    Cancel_Button.Picture = Ok_Static.Picture
End If

End Sub

Private Sub TxtInput_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Ok_Click
ElseIf KeyAscii = 27 Then
    No_Button_Click
End If
End Sub
