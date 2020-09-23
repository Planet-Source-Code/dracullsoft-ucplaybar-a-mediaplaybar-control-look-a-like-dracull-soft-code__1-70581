VERSION 5.00
Begin VB.Form fTest 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "DracullSoft ucPlayBar Demo"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   285
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   595
   StartUpPosition =   2  'CenterScreen
   Begin MPlayBarDemo.ucPlayBar ucPlayBar1 
      Height          =   555
      Left            =   300
      TabIndex        =   0
      Top             =   3360
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   979
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Demo of a Media Player control bar Look-a-Like or almost :o) "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BC741B&
      Height          =   615
      Left            =   660
      TabIndex        =   1
      Top             =   840
      Width           =   7635
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ======================================================================================
' File  :     fTest
' Author:     DracullSoft
' --------------------------------------------------------------------------------------
' Copyright © 2008 DracullSoft
' Author or copyright holders can not be held responsible for anything.
' --------------------------------------------------------------------------------------
' Purpose: Demo form for ucPlayBar  ( a MediaPlayer look-a-like bar, limitted)
'
' ======================================================================================
Option Explicit




Private Sub Form_Load()
ucPlayBar1.Left = 0
ucPlayBar1.Width = 2000
ucPlayBar1.Caption = "DracullSoft PlayBar Demo"

End Sub

Private Sub Form_Resize()
Dim t As Long

    t = Me.ScaleHeight - ucPlayBar1.Height: If t < 1 Then t = 1
    ucPlayBar1.Top = t
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ucPlayBar1_ButtonClick
' Purpose   :
'---------------------------------------------------------------------------------------
Private Sub ucPlayBar1_ButtonClick(ByVal Key As String)
  '_ HandleWMP
  
  '_ for this demo we just change the PlayBar caption
  With ucPlayBar1
    Select Case Key
      Case "Play"
          .Caption = "Play"
      Case "Stop"
          .Caption = "Stop"
      Case "Pause"
          .Caption = "Pause"
      Case "Mute"
          .Caption = "Sound OFF"
      Case "NoMute"
          .Caption = "Sound ON"
    End Select
  End With
End Sub

Private Sub ucPlayBar1_Volume(ByVal val As Long)
  Debug.Print "Volume: " & val
End Sub

'---------------------------------------------------------------------------------------
' Procedure : HandleWMP
' Purpose   : example of how you could control a MediaPlayer control (wmp1) if added
'---------------------------------------------------------------------------------------
'Private Sub HandleWMP(ByVal Key As String)
' Select Case Key
'    Case "Play"
'      If wmp1.playState = wmppsPaused Then
'        wmp1.Controls.Play
'      Else
'        If Not PlaySelected() Then  'check if ready to play url define etc
'          ucPlayBar1.SetState "Stop"
'        End If
'      End If
'    Case "Stop"
'      WmpDoStop
'    Case "Pause"
'      wmp1.Controls.pause
'    Case "Mute"
'      wmp1.settings.mute = True
'    Case "NoMute"
'      wmp1.settings.mute = False
'  End Select
'
'End Sub

