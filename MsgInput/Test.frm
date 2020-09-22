VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   7335
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Test"
      Height          =   615
      Left            =   1920
      TabIndex        =   0
      Top             =   1440
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Clase As New ClsMain
Private Sub Command1_Click()
Clase.MsgBoxXP Clase.InputBoxXP("Maessage   ¿?", "TITLE", Icon_Question), "MSGBOX", Yes_NO_Cancel, Icon_Exclamation
End Sub
