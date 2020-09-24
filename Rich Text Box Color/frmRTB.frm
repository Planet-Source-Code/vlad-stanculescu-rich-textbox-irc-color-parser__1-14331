VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRTB 
   BackColor       =   &H8000000B&
   Caption         =   "Rich Text Box"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEcho 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdEcho 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Echo"
      Height          =   255
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin RichTextLib.RichTextBox txtView 
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmRTB.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmRTB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEcho_Click()
    'This is how I call the function, and direct it to a specific
    'form control. This code was especially made for the parsing of
    'color codes.
    Call Parse(Me, txtView, txtEcho.Text & vbCrLf)
    txtEcho.Text = ""
End Sub

Private Sub Form_Resize()
    txtView.Top = -10
    txtView.Left = -10
    If Left(Me.Width - 97, 1) <> "-" Then txtView.Width = Me.Width - 97
    If Left(Me.Height - txtEcho.Height - 390, 1) <> "-" Then txtView.Height = Me.Height - txtEcho.Height - 390
    txtEcho.Left = -10
    If Left(txtView.Height - 20, 1) <> "-" Then txtEcho.Top = txtView.Height - 20
    If Left(Me.Width - cmdEcho.Width - 150, 1) <> "-" Then txtEcho.Width = Me.Width - cmdEcho.Width - 150
    If Left(txtEcho.Width + 45, 1) <> "-" Then cmdEcho.Left = txtEcho.Width + 45
    cmdEcho.Top = txtView.Height
End Sub
Private Sub txtEcho_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call Parse(Me, txtView, txtEcho.Text & vbCrLf)
        txtEcho.Text = ""
    End If
End Sub
Private Sub txtEcho_KeyPress(KeyAscii As Integer)
    If KeyAscii = 11 Then txtEcho.SelText = Chr(3)
    If KeyAscii = 21 Then txtEcho.SelText = Chr(31)
    If KeyAscii = 2 Then txtEcho.SelText = Chr(2)
    If KeyAscii = 15 Then txtEcho.SelText = Chr(15)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
    If KeyAscii = 11 Then KeyAscii = 0
    If KeyAscii = 21 Then KeyAscii = 0
    If KeyAscii = 2 Then KeyAscii = 0
    If KeyAscii = 15 Then KeyAscii = 0
End Sub
