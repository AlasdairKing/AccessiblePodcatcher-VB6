VERSION 5.00
Begin VB.Form frmSchedule 
   Caption         =   "Schedule subscribing"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSchedule.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstTime 
      Height          =   450
      Left            =   1800
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton optUpdate 
      Caption         =   "Weekly"
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.OptionButton optUpdate 
      Caption         =   "Daily"
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton optUpdate 
      Caption         =   "Hourly"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CheckBox chkSchedule 
      Caption         =   "Automatically update subscriptions"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      Caption         =   "&Time"
      Height          =   195
      Left            =   1800
      TabIndex        =   4
      Top             =   1320
      Width           =   330
   End
End
Attribute VB_Name = "frmSchedule"
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

Private Sub Form_Initialize()
    On Error Resume Next
    Call modXPStyle.InitCommonControlsVB
End Sub

Private Sub Form_Load()
    On Error Resume Next
    'populate the time control
    Dim t As Date
    Dim i As Integer
    Call modI18N.ApplyUILanguageToThisForm(Me)
    Call modLargeFonts.ApplySystemSettingsToForm(Me, , True)
    Call modRememberPosition.LoadPosition(Me)
    
    Call lstTime.AddItem(i & "Midnight")
    Call lstTime.AddItem(i & "1:00AM")
    Call lstTime.AddItem(i & "2:00AM")
    Call lstTime.AddItem(i & "3:00AM")
    Call lstTime.AddItem(i & "4:00AM")
    Call lstTime.AddItem(i & "5:00AM")
    Call lstTime.AddItem(i & "6:00AM")
    Call lstTime.AddItem(i & "7:00AM")
    Call lstTime.AddItem(i & "8:00AM")
    Call lstTime.AddItem(i & "9:00AM")
    Call lstTime.AddItem(i & "10:00AM")
    Call lstTime.AddItem(i & "11:00AM")
    Call lstTime.AddItem(i & "Noon")
    Call lstTime.AddItem(i & "1:00PM")
    Call lstTime.AddItem(i & "2:00PM")
    Call lstTime.AddItem(i & "3:00PM")
    Call lstTime.AddItem(i & "4:00PM")
    Call lstTime.AddItem(i & "5:00PM")
    Call lstTime.AddItem(i & "6:00PM")
    Call lstTime.AddItem(i & "7:00PM")
    Call lstTime.AddItem(i & "8:00PM")
    Call lstTime.AddItem(i & "9:00PM")
    Call lstTime.AddItem(i & "10:00PM")
    Call lstTime.AddItem(i & "11:00PM")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    chkSchedule.Left = GAP
    chkSchedule.Top = GAP
    optUpdate(0).Top = chkSchedule.Top + chkSchedule.Height + GAP + GAP
    optUpdate(0).Left = GAP
    optUpdate(1).Top = optUpdate(0).Top + optUpdate(0).Height + GAP
    optUpdate(1).Left = GAP
    optUpdate(2).Top = optUpdate(1).Top + optUpdate(1).Height + GAP
    optUpdate(2).Left = GAP
    lblTime.Top = optUpdate(0).Top
    lblTime.Left = optUpdate(0).Left + optUpdate(0).Width + GAP + GAP
    lstTime.Left = lblTime.Left
    lstTime.Top = lblTime.Height + lblTime.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call modRememberPosition.SavePosition(Me)
End Sub
