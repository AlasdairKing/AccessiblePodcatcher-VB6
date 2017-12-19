VERSION 5.00
Begin VB.Form frmPodcastList 
   Caption         =   "Add New Podcasts"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPodcastList.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFind 
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ListBox lstPodcasts 
      Height          =   780
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label lblSearch 
      AutoSize        =   -1  'True
      Caption         =   "&Find"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   375
   End
   Begin VB.Label lblAdd 
      AutoSize        =   -1  'True
      Caption         =   "Available Podcasts"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1830
   End
End
Attribute VB_Name = "frmPodcastList"
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

Private WithEvents mList As DOMDocument30
Attribute mList.VB_VarHelpID = -1
Private mURLs As Collection
Private mIDs As Collection

Private Sub cmdClose_Click()
    On Error Resume Next
    Call Me.Hide
    Call frmPodcaster.Display
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call modI18N.ApplyUILanguageToThisForm(Me)
    Call modLargeFonts.ApplySystemSettingsToForm(Me)
    Call modRememberPosition.LoadPosition(Me)
    Set mList = New DOMDocument30
    Set mURLs = New Collection
    Set mIDs = New Collection
    mList.preserveWhiteSpace = False
    mList.resolveExternals = False
    Call mList.Load("http://data.webbie.org.uk/podcastList.php")
    Call lstPodcasts.AddItem(GetText("Connecting to Podcast Directory, please wait..."))
    lstPodcasts.ListIndex = 0
    txtFind.Text = GetText("Type here and press return to find a podcast. Press F3 to find the next one.")
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        Me.lblAdd.Top = 90
        Me.lblAdd.Left = 90
        Me.lstPodcasts.Left = 90
        Me.lstPodcasts.Top = lblAdd.Top + lblAdd.Height + 45
        cmdClose.Top = lstPodcasts.Top
        cmdClose.Left = Me.ScaleWidth - cmdClose.Width
        lstPodcasts.Width = cmdClose.Left - lstPodcasts.Left - 90
        lblSearch.Left = 90
        txtFind.Top = Me.ScaleHeight - txtFind.Height - 90
        lblSearch.Top = txtFind.Top
        txtFind.Left = lblSearch.Left + lblSearch.Width + 45
        txtFind.Width = Me.ScaleWidth - 90 - txtFind.Left
        lstPodcasts.Height = txtFind.Top - lstPodcasts.Top - 90
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call modRememberPosition.SavePosition(Me)
End Sub

Private Sub lstPodcasts_DblClick()
    On Error Resume Next
    Call lstPodcasts_KeyPress(vbKeyReturn)
End Sub

Private Sub lstPodcasts_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyF3 Then
        Call Find
        KeyCode = 0
    End If
End Sub

Private Sub lstPodcasts_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Dim podcast As CPodcast
    Dim index As Long
    
    If KeyAscii = vbKeyReturn Then
        'Add selected podcast.
        Set podcast = New CPodcast
        podcast.url = mURLs.Item(lstPodcasts.ListIndex + 1)
        podcast.name = lstPodcasts.List(lstPodcasts.ListIndex)
        podcast.id = mIDs.Item(lstPodcasts.ListIndex + 1)
        Call frmPodcaster.podcasts.Add(podcast)
        Call MsgBox(GetText("Added") & " " & lstPodcasts.List(lstPodcasts.ListIndex) & " " & GetText("to your podcast list."), vbInformation, App.title)
        index = lstPodcasts.ListIndex
        Call lstPodcasts.RemoveItem(lstPodcasts.ListIndex)
        lstPodcasts.ListIndex = index
        Call mURLs.Remove(index + 1)
    End If
End Sub

Private Sub mList_onreadystatechange()
    On Error Resume Next
    Dim n As IXMLDOMNode
    Dim url As String
    Dim pc As CPodcast
    Dim found As Boolean
    
    If mList.readyState = 4 Then
        'Got list, populate.
        Call lstPodcasts.Clear
        Set mURLs = New Collection
        Set mIDs = New Collection
        If mList.parseError.errorCode = 0 Then
            For Each n In mList.documentElement.selectNodes("podcast")
                url = n.selectSingleNode("url").Text
                found = False
                For Each pc In frmPodcaster.podcasts
                    If StrComp(pc.url, url, vbTextCompare) = 0 Then
                        'Already in list
                        found = True
                        Exit For
                    End If
                Next pc
                If Not found Then
                    Call lstPodcasts.AddItem(n.selectSingleNode("title").Text)
                    Call mURLs.Add(url)
                    
                    Call mIDs.Add(n.Attributes.getNamedItem("id").Text)
                End If
            Next n
        Else
            Debug.Print "Error: failed to parse directory at line " & mList.parseError.Line & " " & mList.parseError.reason & " """ & mList.parseError.srcText & """"
        End If
        If lstPodcasts.ListCount = 0 Then Call lstPodcasts.AddItem(GetText("No podcasts available. Check your Internet connection and firewall!"))
        lstPodcasts.ListIndex = 0
    End If
End Sub

Private Sub txtFind_GotFocus()
    On Error Resume Next
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call Find(txtFind.Text, True)
    End If
End Sub

Private Sub Find(Optional searchString As String, Optional startAfresh As Boolean)
    On Error Resume Next
    Dim startAt As Integer
    Dim i As Integer
    Dim finish As Boolean
    Dim found As Boolean
    Static seekString As String
    
    If searchString = "" Then searchString = seekString
    seekString = searchString
    If searchString = "" Then
        Call Beep
    Else
        If startAfresh Then
            startAt = -1
        Else
            startAt = lstPodcasts.ListIndex
        End If
        i = startAt
        While Not finish
            i = i + 1
            If i > lstPodcasts.ListCount Then
                finish = True
            ElseIf InStr(1, lstPodcasts.List(i), searchString, vbTextCompare) > 0 Then
                finish = True
                found = True
            End If
        Wend
        If Not found Then
            finish = False
            i = -1
            While Not finish
                i = i + 1
                If i >= startAt Then
                    finish = True
                ElseIf InStr(1, lstPodcasts.List(i), searchString, vbTextCompare) > 0 Then
                    finish = True
                    found = True
                End If
            Wend
        End If
        If found Then
            lstPodcasts.ListIndex = i
            Call lstPodcasts.SetFocus
        Else
            Call Beep
        End If
    End If
End Sub
