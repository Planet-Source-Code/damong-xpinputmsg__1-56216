VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'          XP Message and Input Box Appearence Version 1.00
'                Written by: Behrouz Rad
'                 Edited by: Damon
'                    Copyright: 2003
'=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=
'Attention: Contact me for any Question at:
'           Behrouz_VB@yahoo.com
'           Thank you and Enjoy...
'************************************************************
Option Explicit

Public Enum T_Buttons
    Ok_Only = 0
    Yes_NO = 1
    Yes_NO_Cancel = 2
End Enum

Public Enum T_ICONS
  Icon_None = 0
  Icon_Critical = 1
  Icon_Exclamation = 2
  Icon_Info = 3
  Icon_Question = 4
End Enum

Public Enum Messages
    vbsi = vbYes
    vbNop = vbNo
    vbCancelar = vbCancel
End Enum


Const m_def_Button = 0
Const m_def_Icon = 3 'Default Icon is Information Icon

Dim m_Button As Byte
Dim m_Icon As Byte

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Private Sub Position()
 
If EsMsgBox = True Then
    FrmMsg.Width = FrmMsg.lbl_Prompt.Width + 200
    FrmMsg.lbl_Prompt.Left = 100
    FrmMsg.Height = FrmMsg.lbl_Prompt.Top + FrmMsg.lbl_Prompt.Height + 500 + FrmMsg.Ok.Height 'FrmMsg.Ok.Top + 100
    If FrmMsg.Width < 3000 Then FrmMsg.Width = 3000
    Select Case WhatButton
       Case 0
           If FrmMsg.Width < FrmMsg.Ok.Width + 200 Then FrmMsg.Width = FrmMsg.Ok.Width + 200
       Case 2
           If FrmMsg.Width < FrmMsg.Ok.Width + FrmMsg.No_Button.Width + 300 Then FrmMsg.Width = FrmMsg.Ok.Width + FrmMsg.No_Button.Width + 300
       Case 3
           If FrmMsg.Width < FrmMsg.Ok.Width + FrmMsg.No_Button.Width + FrmMsg.Cancel_Button.Width + 400 Then FrmMsg.Width = FrmMsg.Ok.Width + FrmMsg.No_Button.Width + FrmMsg.Cancel_Button.Width + 400
    End Select
Else
    FrmMsg.Width = FrmMsg.lbl_Prompt.Width + 200 + FrmMsg.Ok.Width + 100
    FrmMsg.lbl_Prompt.Left = 100
    FrmMsg.Height = FrmMsg.lbl_Prompt.Top + FrmMsg.lbl_Prompt.Height + FrmMsg.TxtInput.Height + 470 + 100
    If FrmMsg.Width < 3000 Then FrmMsg.Width = 3000
    If FrmMsg.Width < FrmMsg.lbl_Prompt.Width + 200 + FrmMsg.Ok.Width + 100 Then
        FrmMsg.Width = FrmMsg.lbl_Prompt.Width + 200 + FrmMsg.Ok.Width + 100
    End If
        
    'Select Case WhatButton
    '   Case 0
    '       If FrmMsg.Width < FrmMsg.Ok.Width + 200 Then FrmMsg.Width = FrmMsg.Ok.Width + 200
    '   Case 2
    '       If FrmMsg.Width < FrmMsg.Ok.Width + FrmMsg.No_Button.Width + 300 Then FrmMsg.Width = FrmMsg.Ok.Width + FrmMsg.No_Button.Width + 300
    '   Case 3
    '       If FrmMsg.Width < FrmMsg.Ok.Width + FrmMsg.No_Button.Width + FrmMsg.Cancel_Button.Width + 400 Then FrmMsg.Width = FrmMsg.Ok.Width + FrmMsg.No_Button.Width + FrmMsg.Cancel_Button.Width + 400
    'End Select

End If
With FrmMsg.TitleLeft
    .Left = 0
    .Top = 0
End With
 
 FrmMsg.Bottom.Top = FrmMsg.Height - FrmMsg.Bottom.Height
 FrmMsg.Bottom.Width = FrmMsg.Width - FrmMsg.BLeft.Width - FrmMsg.BRight.Width
 FrmMsg.BLeft.Top = FrmMsg.TitleLeft.Height
 FrmMsg.BLeft.Left = 0
 FrmMsg.BLeft.Height = FrmMsg.Height - FrmMsg.TitleLeft.Height
 FrmMsg.Bottom.Left = FrmMsg.BLeft.Width

With FrmMsg.Title
   .Left = FrmMsg.TitleLeft.Width
   .Top = 0
   .Width = FrmMsg.Width - FrmMsg.TitleRight.Width
End With

With FrmMsg.TitleRight
    .Left = FrmMsg.Title.Left + FrmMsg.Title.Width
    .Top = 0
End With

With FrmMsg.CloseButton
     .Left = FrmMsg.TitleRight.Left - .Width - 2
     .Top = (FrmMsg.Title.Height - .Height) / 2
End With

With FrmMsg.BRight
    .Top = FrmMsg.TitleRight.Height
    .Left = FrmMsg.Width - FrmMsg.BRight.Width
    .Height = FrmMsg.BLeft.Height
End With

With FrmMsg.BottomLeft
    .Left = 0
    .Top = FrmMsg.Height - .Height
End With

With FrmMsg.BottomRight
    .Left = FrmMsg.Width - .Width
    .Top = FrmMsg.Height - .Height
End With
End Sub

Private Sub Transparent()
'    Dim Add As Long
'    Dim Sum As Long
'
'    Dim X As Single
'    Dim Y As Single
'
'    X = FrmMsg.Width / Screen.TwipsPerPixelX   'Registers the Size of the
'    Y = FrmMsg.Height / Screen.TwipsPerPixelY  'Form in Pixels
'
'    Sum = CreateRectRgn(5, 0, X - 5, 1)
'    CombineRgn Sum, Sum, CreateRectRgn(3, 1, X - 3, 2), 2
'    CombineRgn Sum, Sum, CreateRectRgn(2, 2, X - 2, 3), 2
'    CombineRgn Sum, Sum, CreateRectRgn(1, 3, X - 1, 4), 2
'    CombineRgn Sum, Sum, CreateRectRgn(1, 4, X - 1, 5), 2
'    CombineRgn Sum, Sum, CreateRectRgn(0, 5, X, Y), 2
'    SetWindowRgn FrmMsg.hWnd, Sum, True   'Sets Corners Transparent
   Dim Add              As Long
   Dim Sum              As Long

   Dim X                As Single
   Dim Y                As Single

   With FrmMsg
      If .Height < 615 Then .Height = 615   'Checks that form
      If .Width < 1695 Then .Width = 1695   'is not too small

      X = .Width / Screen.TwipsPerPixelX   'Registers the size of the
      Y = .Height / Screen.TwipsPerPixelY  'form in pixels

      'Sum = CreateRectRgn(5, 0, X - 5, 1)
      'CombineRgn Sum, Sum, CreateRectRgn(3, 1, X - 3, 2), 2
      'CombineRgn Sum, Sum, CreateRectRgn(2, 2, X - 2, 3), 2
      'CombineRgn Sum, Sum, CreateRectRgn(1, 3, X - 1, 4), 2
      'CombineRgn Sum, Sum, CreateRectRgn(1, 4, X - 1, 5), 2
      'CombineRgn Sum, Sum, CreateRectRgn(0, 5, X, Y), 2
      'SetWindowRgn .ContainerHwnd, Sum, True   'Sets corners transparent
    Sum = CreateRectRgn(0, 0, 0, 0)
    'Puntas derecha e izquierda de arriba
    CombineRgn Sum, Sum, CreateRectRgn(7, 1, X - 6, 2), 2
    CombineRgn Sum, Sum, CreateRectRgn(5, 2, X - 4, 3), 2
    CombineRgn Sum, Sum, CreateRectRgn(4, 3, X - 3, 4), 2
    CombineRgn Sum, Sum, CreateRectRgn(3, 4, X - 2, 5), 2
    CombineRgn Sum, Sum, CreateRectRgn(2, 5, X - 2, 6), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 6, X - 1, 7), 2
    CombineRgn Sum, Sum, CreateRectRgn(1, 7, X - 1, 8), 2
    CombineRgn Sum, Sum, CreateRectRgn(0, 8, X, Y), 2
    'punta izquierda de abajo
    CombineRgn Sum, Sum, CreateRectRgn(0, Y, 8, Y - 1), 3
    CombineRgn Sum, Sum, CreateRectRgn(0, Y - 1, 6, Y - 2), 3
    CombineRgn Sum, Sum, CreateRectRgn(0, Y - 2, 4, Y - 3), 3
    CombineRgn Sum, Sum, CreateRectRgn(0, Y - 3, 3, Y - 4), 3
    CombineRgn Sum, Sum, CreateRectRgn(0, Y - 4, 2, Y - 5), 3
    CombineRgn Sum, Sum, CreateRectRgn(0, Y - 5, 2, Y - 6), 3
    CombineRgn Sum, Sum, CreateRectRgn(0, Y - 6, 1, Y - 7), 3
    CombineRgn Sum, Sum, CreateRectRgn(0, Y - 7, 1, Y - 8), 3
    'punta derecha de abajo
    CombineRgn Sum, Sum, CreateRectRgn(X - 8, Y, X, Y - 1), 3
    CombineRgn Sum, Sum, CreateRectRgn(X - 6, Y - 1, X, Y - 2), 3
    CombineRgn Sum, Sum, CreateRectRgn(X - 4, Y - 2, X, Y - 3), 3
    CombineRgn Sum, Sum, CreateRectRgn(X - 3, Y - 3, X, Y - 4), 3
    CombineRgn Sum, Sum, CreateRectRgn(X - 2, Y - 4, X, Y - 5), 3
    CombineRgn Sum, Sum, CreateRectRgn(X - 2, Y - 5, X, Y - 6), 3
    CombineRgn Sum, Sum, CreateRectRgn(X - 1, Y - 6, X, Y - 7), 3
    CombineRgn Sum, Sum, CreateRectRgn(X - 1, Y - 7, X, Y - 8), 3
    SetWindowRgn FrmMsg.hWnd, Sum, True   'Sets corners transparent

   End With
End Sub

Public Property Get Caption() As String
 Caption = FrmMsg.Caption1.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
 FrmMsg.Caption1.Caption = New_Caption
 FrmMsg.Caption2.Caption = New_Caption
End Property

Public Property Get Prompt() As String
 Prompt = FrmMsg.lbl_Prompt.Caption
End Property

Public Property Let Prompt(ByVal New_Prompt As String)
 FrmMsg.lbl_Prompt.Caption = New_Prompt
End Property

Public Property Get Buttons() As T_Buttons
 Buttons = m_Button
End Property

Public Property Let Buttons(ByVal New_Button As T_Buttons)
 m_Button = New_Button
End Property

Public Property Get Icon() As T_ICONS
 Icon = m_Icon
End Property

Public Property Let Icon(ByVal New_Icon As T_ICONS)
 m_Icon = New_Icon
End Property

Public Property Get Title_Icon() As Picture
 Set Title_Icon = FrmMsg.img_Icon.Picture
End Property

Public Property Set Title_Icon(ByVal New_Title_Icon As Picture)
 Set FrmMsg.img_Icon.Picture = New_Title_Icon
End Property

Public Property Get BackColor() As OLE_COLOR
 BackColor = FrmMsg.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
 FrmMsg.BackColor = New_BackColor
End Property

Public Property Get Prompt_Color() As OLE_COLOR
 Prompt_Color = FrmMsg.lbl_Prompt.ForeColor
End Property

Public Property Let Prompt_Color(ByVal New_Prompt_Color As OLE_COLOR)
 FrmMsg.lbl_Prompt.ForeColor = New_Prompt_Color
End Property

Public Property Get Caption_Color() As OLE_COLOR
 Caption_Color = FrmMsg.Caption2.ForeColor
End Property

Public Property Let Caption_Color(ByVal New_Caption_Color As OLE_COLOR)
   FrmMsg.Caption2.ForeColor = New_Caption_Color
End Property

Public Property Get Caption_Shadow_Color() As OLE_COLOR
 Caption_Shadow_Color = FrmMsg.Caption1.ForeColor
End Property

Public Property Let Caption_Shadow_Color(ByVal New_Caption_Shadow_Color As OLE_COLOR)
 FrmMsg.Caption1.ForeColor = New_Caption_Shadow_Color
End Property

Public Property Get Ok_Button_Caption_Color() As OLE_COLOR
 Ok_Button_Caption_Color = FrmMsg.lbl_OK.ForeColor
End Property

Public Property Let Ok_Button_Caption_Color(ByVal New_Ok_Button_Caption_Color As OLE_COLOR)
 FrmMsg.lbl_OK.ForeColor = New_Ok_Button_Caption_Color
End Property

Public Property Get Yes_Button_Caption_Color() As OLE_COLOR
 Yes_Button_Caption_Color = FrmMsg.lbl_OK.ForeColor
End Property

Public Property Let Yes_Button_Caption_Color(ByVal New_Yes_Button_Caption_Color As OLE_COLOR)
 FrmMsg.lbl_OK.ForeColor = New_Yes_Button_Caption_Color
End Property

Public Property Get No_Button_Caption_color() As OLE_COLOR
 No_Button_Caption_color = FrmMsg.lbl_No.ForeColor
End Property

Public Property Let No_Button_Caption_color(ByVal New_No_Button_Caption_Color As OLE_COLOR)
 FrmMsg.lbl_No.ForeColor = New_No_Button_Caption_Color
End Property
Public Function ShowMsg() As Messages
Dim Espacio As Double

FrmMsg.lbl_Prompt.Alignment = 2
Select Case Buttons
Case 0
   WhatButton = 1 'Ok Button
   Call Position
   Call Transparent
   FrmMsg.Ok.Top = FrmMsg.lbl_Prompt.Top + FrmMsg.lbl_Prompt.Height + 250 'FrmMsg.Height - 500
   FrmMsg.Ok.Left = (FrmMsg.Width - FrmMsg.Ok.Width) / 2 'Center Form
   
   FrmMsg.lbl_OK.Top = FrmMsg.Ok.Top + 55 'FrmMsg.Height - 500
   FrmMsg.lbl_OK.Left = (FrmMsg.Width - FrmMsg.lbl_OK.Width) / 2 'Center Form
   FrmMsg.CloseButton.Visible = False
   
Case 1
   WhatButton = 2 'Yes & No Buttons
   Call Position
   Call Transparent
   
   Espacio = (FrmMsg.Width - FrmMsg.Ok.Width - FrmMsg.No_Button.Width) / 3
      
   FrmMsg.No_Button.Picture = FrmMsg.Ok_Static.Picture
   FrmMsg.Ok.Top = FrmMsg.lbl_Prompt.Top + FrmMsg.lbl_Prompt.Height + 250 'FrmMsg.Height - 500
   FrmMsg.Ok.Left = Espacio '(FrmMsg.Width - (FrmMsg.Ok.Width + FrmMsg.No_Button.Width + 200)) / 2 '200 Number is Spaces Between two Buttons
   
   
   FrmMsg.lbl_OK.Caption = "&Si"
   FrmMsg.lbl_OK.Top = FrmMsg.Ok.Top + 55
   FrmMsg.lbl_OK.Left = FrmMsg.Ok.Left + 500
   
   FrmMsg.No_Button.Top = FrmMsg.Ok.Top
   FrmMsg.No_Button.Left = FrmMsg.Ok.Left + FrmMsg.Ok.Width + Espacio
   FrmMsg.lbl_No.Visible = True
   FrmMsg.lbl_No.Caption = "&No"
   FrmMsg.lbl_No.Top = FrmMsg.No_Button.Top + 55
   FrmMsg.lbl_No.Left = FrmMsg.No_Button.Left + 470
   
    FrmMsg.CloseButton.Visible = False

Case 2
    WhatButton = 3 'Yes, No & Cancel buttons
    Call Acomodolabel(FrmMsg.lbl_Prompt)
    Call Position
    Call Transparent
    
    Espacio = (FrmMsg.Width - FrmMsg.Ok.Width - FrmMsg.No_Button.Width - FrmMsg.Cancel_Button.Width) / 4

    FrmMsg.No_Button.Picture = FrmMsg.Ok_Static.Picture
    FrmMsg.Cancel_Button.Picture = FrmMsg.Ok_Static.Picture
    
    FrmMsg.Ok.Top = FrmMsg.lbl_Prompt.Top + FrmMsg.lbl_Prompt.Height + 250 'FrmMsg.Height - 500
    FrmMsg.Ok.Left = Espacio '(FrmMsg.Width - (FrmMsg.Ok.Width + FrmMsg.No_Button.Width + FrmMsg.Cancel_Button.Width + 100)) / 3  '200 Number is Spaces Between two Buttons
    FrmMsg.lbl_OK.Caption = "&Si"
    FrmMsg.lbl_OK.Top = FrmMsg.Ok.Top + 55
    FrmMsg.lbl_OK.Left = FrmMsg.Ok.Left + 470
    
    FrmMsg.No_Button.Top = FrmMsg.Ok.Top
    FrmMsg.No_Button.Left = FrmMsg.Ok.Left + FrmMsg.Ok.Width + Espacio
    FrmMsg.lbl_No.Visible = True
    FrmMsg.lbl_No.Caption = "&No"
    FrmMsg.lbl_No.Top = FrmMsg.No_Button.Top + 55
    FrmMsg.lbl_No.Left = FrmMsg.No_Button.Left + 470
    
    FrmMsg.Cancel_Button.Visible = True
    FrmMsg.Cancel_Button.Top = FrmMsg.Ok.Top
    FrmMsg.Cancel_Button.Left = FrmMsg.No_Button.Left + FrmMsg.No_Button.Width + Espacio
    With FrmMsg.lbl_Cancel
        .Visible = True
        .Top = FrmMsg.Cancel_Button.Top + 55
        .Left = FrmMsg.Cancel_Button.Left + 200
    End With
    
    FrmMsg.CloseButton.Visible = True
End Select

'Select Case Icon
'Case 0 'Icon_None
'  FrmMsg.Icon_Info.Visible = False
'Case 1 'Icon_Critical
'  FrmMsg.Icon_Cri.Visible = True
'Case 2 'Icon_Exclamation
'  FrmMsg.Icon_Exc.Visible = True
'Case 3 'Icon_Info
'  FrmMsg.Icon_Info.Visible = True
'Case 4 'Icon_Question
'  FrmMsg.Icon_Question.Visible = True
'End Select

FrmMsg.Show 1
 Select Case EnterB
    Case 1, 2
        ShowMsg = vbsi
    Case 3
        ShowMsg = vbNop
    Case 4
        ShowMsg = vbCancelar
 End Select
End Function

Private Sub Class_Initialize()
 m_Button = m_def_Button
 m_Icon = m_def_Icon
End Sub

Public Function MsgBoxXP(strMessage As String, _
    Optional strTitle As String, _
    Optional IpType As T_Buttons = Ok_Only, _
    Optional MsgIcon As T_ICONS = Icon_None, _
    Optional IpIcon As Boolean = False, _
    Optional IpPicture As StdPicture, _
    Optional TextColor As Long = CLng(vbBlack)) As Messages

'Dim Result As Byte
Dim VarShow As New ClsMain     'Variable Definition
'Set VarShow = New ClsMain 'Create Object

EsMsgBox = True
'Message Mensaje
VarShow.Prompt = strMessage

'Message Titulo
If strTitle <> "" Then
    VarShow.Caption = strTitle
Else
    VarShow.Caption = App.Title
End If
'Message Box Button(s) Definition
VarShow.Buttons = IpType

'Message Box Icon Definition
'VarShow.Icon = MsgIcon

'Message Box Title's Icon Definition
If IpIcon = True Then
    Set VarShow.Title_Icon = IpPicture
Else
    Select Case MsgIcon
        Case Icon_Critical
            Set VarShow.Title_Icon = FrmMsg.Icon_Cri.Picture
        Case Icon_Exclamation
            Set VarShow.Title_Icon = FrmMsg.Icon_Exc.Picture
        Case Icon_Info
            Set VarShow.Title_Icon = FrmMsg.Icon_Info.Picture
        Case Icon_None
            Set VarShow.Title_Icon = Nothing
        Case Icon_Question
            Set VarShow.Title_Icon = FrmMsg.Icon_Question.Picture
    End Select
End If

'Message Box BackColor Definition
VarShow.BackColor = &HCECECE '>>> Default Xp Color in Hex Format

'Message Box Caption Color Definition
VarShow.Caption_Color = RGB(255, 255, 255)

'Message Box Caption Shadow Color Definition
VarShow.Caption_Shadow_Color = vbBlack

'Message Box Ok Button Caption's Color Definition
VarShow.Yes_Button_Caption_Color = RGB(0, 0, 255)

'Message Box No Button Caption's Color Definition
VarShow.No_Button_Caption_color = vbBlack

'Message Box Prompt Color Definition
VarShow.Prompt_Color = TextColor

'Show Message Box
'Returns Constant Values For Message Box Buttons.
'1 = Ok
'2 = Yes
'3 = No
'Result = VarShow.ShowMsg
'Select Case Result
'Case 2
'  MsgBox "You Clicked Yes Button" & vbCr & "Now Program Will Terminate", vbInformation
'  End
'Case 3
'  MsgBox "You Clicked No Button" & vbCr & "Now Program Will Stay", vbInformation
'End Select
MsgBoxXP = VarShow.ShowMsg

'Release Memeory From Created Object
Set VarShow = Nothing

End Function

Private Sub Acomodolabel(ByRef MuestraLabel As Label)
Dim AnchoMax As Double
Dim LenTexto As Double
Dim AnchoLetra As Byte

'Const Baja = "& vbCrLf "

If EsMsgBox = True Then
    AnchoMax = Screen.Width * 0.25
Else
    AnchoMax = Screen.Width * 0.5
End If
AnchoMax = Screen.Width - AnchoMax

If MuestraLabel.Width > AnchoMax Then
    Dim PartesLabel() As String
    Dim NPartesLabel As Byte
    Dim TotalLetras As Long
    Dim LugarActual As Double
    Dim Acomodado As Boolean
    Dim i As Byte
    Dim j As Double
    
    LenTexto = Len(MuestraLabel.Caption)
    AnchoLetra = MuestraLabel.Width / LenTexto
    NPartesLabel = Int(LenTexto * AnchoLetra / AnchoMax) + 1
    ReDim PartesLabel(1 To NPartesLabel) As String
    TotalLetras = Round(AnchoMax / AnchoLetra, 0)
    
    LugarActual = 1
    For i = 1 To NPartesLabel
        'For j = 1 To TotalLetras
           PartesLabel(i) = Mid(MuestraLabel.Caption, LugarActual, TotalLetras + 1)
           LugarActual = LugarActual + TotalLetras + 1
        'Next j
    Next i
    Acomodado = False
    LugarActual = 1
    Do Until Acomodado = True
        If LugarActual <= NPartesLabel Then
            For i = LugarActual To NPartesLabel
                If Left(Right(PartesLabel(i), 2), 1) = " " Or Right(PartesLabel(i), 1) = " " Then
                    PartesLabel(i) = Left(PartesLabel(i), Len(PartesLabel(i)) - 1)
                Else
                    For j = Len(PartesLabel(i)) To 1 Step -1
                        If Mid(PartesLabel(i), j, 1) = " " Then
                            If i = NPartesLabel Then
                                NPartesLabel = NPartesLabel + 1
                                ReDim Preserve PartesLabel(1 To NPartesLabel) As String
                            End If
                            PartesLabel(i + 1) = Mid(PartesLabel(i), j + 1, Len(PartesLabel(i)) - j) & " " & PartesLabel(i + 1)
                            PartesLabel(i) = Mid(PartesLabel(i), 1, j) & vbCr
                            If Trim(PartesLabel(i + 1)) = "" Or Len(Trim(PartesLabel(i + 1))) < AnchoMax Then
                                Acomodado = True
                            End If
                            Exit For
                        End If
                    Next j
                End If
                If Acomodado = True Then
                    Exit For
                End If
            Next i
        Else
            Acomodado = True
        End If
    Loop
    MuestraLabel.Caption = ""
    For i = 1 To NPartesLabel
        MuestraLabel.Caption = MuestraLabel.Caption & PartesLabel(i)
    Next i
End If
If NPartesLabel <= 1 Then
    If EsMsgBox = True Then
        Select Case WhatButton
            Case 1
                If MuestraLabel.Width <= FrmMsg.Ok.Width Then
                    MuestraLabel.Width = FrmMsg.Ok.Width + 200
                End If
            Case 2
                If MuestraLabel.Width <= FrmMsg.Ok.Width Then
                    MuestraLabel.Width = FrmMsg.Ok.Width + 200
                End If
            Case 3
                If MuestraLabel.Width <= FrmMsg.Ok.Width Then
                    MuestraLabel.Width = FrmMsg.Ok.Width + 200
                End If
        End Select
    Else
        If MuestraLabel.Width < 4000 Then MuestraLabel.Width = 4000
    End If
End If
End Sub
Public Function InputBoxXP(strMessage As String, _
    Optional strTitle As String, _
    Optional MsgIcon As T_ICONS = Icon_None, _
    Optional IpIcon As Boolean = False, _
    Optional IpPicture As StdPicture, _
    Optional TextColor As Long = CLng(vbBlack)) As String

'Dim Result As Byte
Dim VarShow As New ClsMain     'Variable Definition
'Set VarShow = New ClsMain 'Create Object

EsMsgBox = False
'Message Mensaje
VarShow.Prompt = strMessage

'Message Titulo
If strTitle <> "" Then
    VarShow.Caption = strTitle
Else
    VarShow.Caption = App.Title
End If
'Message Box Icon Definition
'VarShow.Icon = MsgIcon

'Message Box Title's Icon Definition
If IpIcon = True Then
    Set VarShow.Title_Icon = IpPicture
Else
    Select Case MsgIcon
        Case Icon_Critical
            Set VarShow.Title_Icon = FrmMsg.Icon_Cri.Picture
        Case Icon_Exclamation
            Set VarShow.Title_Icon = FrmMsg.Icon_Exc.Picture
        Case Icon_Info
            Set VarShow.Title_Icon = FrmMsg.Icon_Info.Picture
        Case Icon_None
            Set VarShow.Title_Icon = Nothing
        Case Icon_Question
            Set VarShow.Title_Icon = FrmMsg.Icon_Question.Picture
    End Select
End If

'Message Box BackColor Definition
VarShow.BackColor = &HCECECE '>>> Default Xp Color in Hex Format

'Message Box Caption Color Definition
VarShow.Caption_Color = RGB(255, 255, 255)

'Message Box Caption Shadow Color Definition
VarShow.Caption_Shadow_Color = vbBlack

'Message Box Ok Button Caption's Color Definition
VarShow.Yes_Button_Caption_Color = RGB(0, 0, 255)

'Message Box No Button Caption's Color Definition
VarShow.No_Button_Caption_color = vbBlack

'Message Box Prompt Color Definition
VarShow.Prompt_Color = TextColor

'Show Message Box
'Returns Constant Values For Message Box Buttons.
'1 = Ok
'2 = Yes
'3 = No
'Result = VarShow.ShowMsg
'Select Case Result
'Case 2
'  MsgBox "You Clicked Yes Button" & vbCr & "Now Program Will Terminate", vbInformation
'  End
'Case 3
'  MsgBox "You Clicked No Button" & vbCr & "Now Program Will Stay", vbInformation
'End Select
InputBoxXP = VarShow.ShowInput

'Release Memeory From Created Object
Set VarShow = Nothing

End Function
Public Function ShowInput() As String
Dim Espacio As Double


EsMsgBox = False
WhatButton = 2

Call Acomodolabel(FrmMsg.lbl_Prompt)
Call Position
Call Transparent

Espacio = (FrmMsg.Height - FrmMsg.Ok.Height - FrmMsg.No_Button.Height - 450 - 100) / 3


With FrmMsg
    .CloseButton.Visible = True
    .lbl_Prompt.Alignment = 0
    
    .Ok.Top = Espacio + 450
    .Ok.Left = .lbl_Prompt.Width + 200
    .lbl_OK.Top = .Ok.Top + 55
    .lbl_OK.Left = .Ok.Left + 230
    .lbl_OK.Caption = "&Aceptar"
    
    .No_Button.Visible = True
    .No_Button.Picture = FrmMsg.Ok_Static.Picture
    .No_Button.Left = .Ok.Left
    .No_Button.Top = .Ok.Top + .Ok.Height + Espacio
    .lbl_No.Visible = True
    .lbl_No.Caption = "&Cancelar"
    .lbl_No.Top = .No_Button.Top + 55
    .lbl_No.Left = .No_Button.Left + 200
    
    .FraInput.Visible = True
    .FraInput.Left = 70
    .FraInput.Top = .No_Button.Top '.lbl_Prompt.Top + .lbl_Prompt.Height + 100
    .FraInput.Width = .lbl_Prompt.Width + 40
    .TxtInput.Width = .FraInput.Width - 30
    .TxtInput.Top = 20
    .TxtInput.Left = 20
    .FraInput.Height = .TxtInput.Height + 40
    
    .LineUp.X1 = 0
    .LineUp.X2 = .FraInput.Width
    .LineUp.Y1 = 0
    .LineUp.Y2 = 0
    .LineLeft.X1 = 0
    .LineLeft.X2 = 0
    .LineLeft.Y1 = 0
    .LineLeft.Y2 = .TxtInput.Height + 20
    .LineDown.X1 = 0
    .LineDown.X2 = .FraInput.Width
    .LineDown.Y1 = .TxtInput.Height + 20
    .LineDown.Y2 = .TxtInput.Height + 20
    .LineRight.X1 = .FraInput.Width - 20
    .LineRight.X2 = .FraInput.Width - 20
    .LineRight.Y1 = 0
    .LineRight.Y2 = .LineLeft.Y2
End With
FrmMsg.Show 1
ShowInput = ElTexto
End Function



