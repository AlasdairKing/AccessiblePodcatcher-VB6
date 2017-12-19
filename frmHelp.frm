VERSION 5.00
Begin VB.Form frmHelp 
   Caption         =   "Accessible BBC Listen Again Help"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Tag             =   "frmHelp"
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtHelp 
      Height          =   2655
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'LICENCE
'This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.

'    This program is copyright 2005 Alasdair King alasdair@webbie.org.uk
Option Explicit

Private Sub cmdOK_Click()
    On Error Resume Next
    Call frmPodcaster.SetFocus
    Call Me.Hide
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    txtHelp.Text = modI18N.helpTopicText(0)
    Me.Caption = modI18N.helpTopicTitle(0)
    Call txtHelp.SetFocus
End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    Call modXPStyle.InitCommonControlsVB
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call modI18N.ApplyUILanguageToThisForm(Me)
    Call modRememberPosition.LoadPosition(Me)
    Call modLargeFonts.ApplySystemSettingsToForm(Me, , True)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        If Me.Height > cmdOK.Height Then
            cmdOK.Top = Me.ScaleHeight - cmdOK.Height - 90
            txtHelp.Height = Me.ScaleHeight - cmdOK.Height - 270
            txtHelp.Top = 90
        End If
        If Me.Width > cmdOK.Width + 180 Then
            txtHelp.Left = 90
            cmdOK.Left = Me.ScaleWidth / 2 - cmdOK.Width / 2
            txtHelp.Width = Me.ScaleWidth - 180
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call modRememberPosition.SavePosition(Me)
End Sub
