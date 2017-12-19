VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form frmSearch 
   Caption         =   "Search webpage for podcasts"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   9615
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet wininet 
      Left            =   4560
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtURL 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "Enter the web address for the page to search here and press return."
      Top             =   360
      Width           =   5295
   End
   Begin MSComctlLib.StatusBar staProgress 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   4605
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrCheck 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4200
      Top             =   1080
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK (Add to podcatcher)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   360
      Width           =   2655
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   6
      Top             =   960
      Width           =   2655
   End
   Begin VB.ListBox lstResults 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   6615
   End
   Begin SHDocVwCtl.WebBrowser webPageChecker 
      Height          =   2295
      Left            =   6120
      TabIndex        =   7
      Top             =   1920
      Width           =   4455
      ExtentX         =   7858
      ExtentY         =   4048
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label lblSearch 
      AutoSize        =   -1  'True
      Caption         =   "&Search webpage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1620
   End
   Begin VB.Label lblFound 
      AutoSize        =   -1  'True
      Caption         =   "Podcasts &Found"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1515
   End
End
Attribute VB_Name = "frmSearch"
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

Private mResultNames As Collection
Private mResultURLs As Collection
Private mLinkURLsToCheck As Collection
Private mCheckIndex As Long
Private WithEvents mGetter As DOMDocument30
Attribute mGetter.VB_VarHelpID = -1
Private mWininetState As Long

Private Sub cmdClose_Click()
    On Error Resume Next
    Dim result As Long
    
    If lstResults.ListCount > 0 Then
        'Ah, we've found some podcasts.
        result = MsgBox(GetText("You found") & " " & lstResults.ListCount & " " & GetText("podcasts. Do you want to add these to your podcast list?"), vbYesNoCancel)
        If result = vbYes Then
            Call cmdOK_Click
        ElseIf result = vbNo Then
            Call Me.Hide
        Else 'Cancel
            'Don't do anything!
        End If
    Else
        Call Me.Hide
    End If
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    Dim newCast As CPodcast
    Dim i As Integer
    
    tmrCheck.Enabled = False
    If mResultURLs.Count > 0 Then
        For i = 1 To mResultURLs.Count
            Set newCast = New CPodcast
            newCast.url = mResultURLs.Item(i)
            newCast.name = mResultNames.Item(i)
            Call frmPodcaster.podcasts.Add(newCast)
            Call frmPodcaster.UpdateDirectory(newCast.url, newCast.name)
        Next i
        Call Me.Hide
        Call frmPodcaster.Display
        frmPodcaster.lstPodcasts.ListIndex = frmPodcaster.lstPodcasts.ListCount - 1
        Call frmPodcaster.lstPodcasts.SetFocus
    End If
End Sub

Private Sub cmdSearch_Click()
    On Error Resume Next
    Call webPageChecker.Navigate2(txtURL.Text)
    Call lstResults.SetFocus
    cmdOK.default = True
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    cmdOK.Enabled = False
    Call lstResults.Clear
    txtURL.Text = GetText("Enter the web address for the page to search here and press return.")
    staProgress.SimpleText = ""
    Set mGetter = New DOMDocument30
    Call wininet.Cancel
End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    Call modXPStyle.InitCommonControlsVB
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Call modI18N.ApplyUILanguageToThisForm(Me)
    'Resize according to Windows font sizes
    Call modLargeFonts.ApplySystemSettingsToForm(Me, , True)
    'get layout
    Call modRememberPosition.LoadPosition(Me)
    webPageChecker.Silent = True
    webPageChecker.TabStop = False
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    webPageChecker.Left = -webPageChecker.Width - 100
    If Me.WindowState <> vbMinimized Then
        lblSearch.Left = GAP
        txtURL.Left = GAP
        lblFound.Left = GAP
        lstResults.Left = GAP
        lstResults.Width = Me.ScaleWidth - cmdOK.Width - lstResults.Left - GAP - GAP
        txtURL.Width = lstResults.Width - cmdSearch.Width
        cmdSearch.Left = txtURL.Left + txtURL.Width
        cmdSearch.Top = txtURL.Top
        cmdOK.Left = Me.ScaleWidth - cmdOK.Width - GAP
        cmdClose.Left = cmdOK.Left
        cmdClose.Width = cmdOK.Width
        
        lblSearch.Top = GAP
        txtURL.Top = lblSearch.Top + lblSearch.Height
        lblFound.Top = txtURL.Top + txtURL.Height + GAP
        lstResults.Top = lblFound.Top + lblFound.Height
        lstResults.Height = Me.ScaleHeight - staProgress.Height - lstResults.Top - GAP
        cmdOK.Top = GAP
        cmdClose.Top = cmdOK.Top + cmdOK.Height + GAP
    End If
End Sub

Private Sub lstResults_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyDelete Then
        If mResultURLs.Count = 0 Then
            Call Beep
        Else
            Call mResultURLs.Remove(lstResults.ListIndex + 1)
            Call mResultNames.Remove(lstResults.ListIndex + 1)
            Call lstResults.RemoveItem(lstResults.ListIndex)
            If lstResults.ListCount = 0 Then
                Call lstResults.AddItem(modI18N.GetText("No podcasts found!"))
                cmdOK.Enabled = False
                cmdClose.default = True
            End If
        End If
    End If
End Sub

Private Sub mGetter_onreadystatechange()
    On Error Resume Next
    Dim title As String
    Dim isPodcast As Boolean
    
    If mGetter.readyState = 4 And Me.visible Then
        If mGetter.parseError.errorCode <> 0 Then
            'failed to parse - not a podcast. (Of course, this is completely untrue, so let's try for podcasts with <rss in them...)
            'This means getting the contents over HTTP somehow, either through a webbrowser object or through
            'something else. For every link? Rats to that for now. Steve Nutt can sort his podcast out. Alasdair Dec 2008.
            'Me.staProgress.SimpleText = GetText("Invalid XML, falling back to HTTP for " & mGetter.url)
            isPodcast = CheckBrokenPageIsPodcast(mGetter.url, title)
'            If frmPodcaster.mnuOptionsUsebrokenfeeds.Checked = vbChecked Then
'                If InStr(1, mGetter.Text, "<rss ", vbTextCompare) < 100 Then
'                    Call mResultURLs.Add(mGetter.URL)
'                    If InStr(1, mGetter.Text, "<title>", vbTextCompare) > 0 Then
'                        title = Right(mGetter.Text, Len(mGetter.Text) - InStr(1, mGetter.Text, "<title>", vbTextCompare) - Len("<title>"))
'                        title = Left(title, InStr(1, title, "</title"))
'                        Call mResultNames.Add(title)
'                        Call lstResults.AddItem(title)
'                    Else
'                        Call mResultNames.Add("Broken podcast")
'                        Call Me.lstResults.AddItem("Broken podcast")
'                    End If
'                End If
'            End If
        ElseIf LCase(mGetter.documentElement.nodeName) <> "rss" Then
            'Not an RSS feed, so not a podcast - might be a valid XHTML document! There are about four of those...
        ElseIf InStr(1, mGetter.xml, "<enclosure", vbTextCompare) = 0 Then
            'Valid XML, root is RSS, but no enclosures - must be RSS feed, not podcast.
        Else
            'Hey, valid XML, root is RSS, got an enclosure - contender!
            isPodcast = True
            title = mGetter.documentElement.selectSingleNode("channel").selectSingleNode("title").Text
        End If
        If isPodcast Then
            Call Me.lstResults.AddItem(title)
            Call mResultNames.Add(title)
            Call mResultURLs.Add(mGetter.url)
        End If
        tmrCheck.Enabled = True
    End If
End Sub

Private Function CheckBrokenPageIsPodcast(url As String, ByRef title As String) As Boolean
    On Error Resume Next
    Dim got As String
    Dim urlCode As String
    Dim isRSS As Boolean
    Dim isAtom As Boolean
    
    If wininet.StillExecuting Then
        Call wininet.Cancel
        DoEvents
    End If
    mWininetState = -1
    Call wininet.Execute(url, "GET")
    DoEvents
    While wininet.StillExecuting And Me.visible
        DoEvents
        mWininetState = wininet.ResponseCode
    Wend
    While wininet.StillExecuting And Me.visible
        DoEvents
    Wend
    If Me.visible Then
        'Get contents back.
        urlCode = wininet.GetChunk(1024, icString)
'        While Len(got) <> 0 And Len(urlCode) < 2048
'            urlCode = urlCode & got
'            got = wininet.GetChunk(1024, icString)
'            DoEvents
'        Wend
        Debug.Print "Got invalid XML successfully: " & urlCode
        'Now check out the content for RSS.
        isRSS = (InStr(1, urlCode, "<rss", vbTextCompare) > 0)
        isAtom = (InStr(1, urlCode, "<feed", vbTextCompare) > 0)
        'Check that this is both RSS/Atom AND has enclosures - or else it's just an RSS feed.
        CheckBrokenPageIsPodcast = isRSS Or isAtom And (InStr(1, urlCode, "<enclosure", vbTextCompare) > 0)
        'Extract the title
        If CheckBrokenPageIsPodcast Then
            title = Right(urlCode, Len(urlCode) - InStr(1, urlCode, "<title", vbTextCompare))
            title = Left(urlCode, InStr(1, urlCode, "</title>", vbTextCompare))
            title = Right(urlCode, Len(url) - InStr(1, urlCode, ">"))
        End If
    End If
End Function

Private Sub tmrCheck_Timer()
    On Error Resume Next
    Dim url As String
    
    tmrCheck.Enabled = False
    If Me.visible Then
        mCheckIndex = mCheckIndex + 1
        If mCheckIndex > mLinkURLsToCheck.Count Then
            staProgress.SimpleText = modI18N.GetText("Done")
            If Me.lstResults.ListCount = 0 Then
                Call Me.lstResults.AddItem(modI18N.GetText("No podcasts found!"))
                Me.cmdOK.Enabled = False
                cmdClose.default = True
            Else
                Me.cmdOK.Enabled = True
            End If
            lstResults.ListIndex = 0
            Call lstResults.SetFocus
        Else
            staProgress.SimpleText = modI18N.GetText("Checking link") & " " & mCheckIndex & " " & modI18N.GetText("of") & " " & mLinkURLsToCheck.Count & " (" & mLinkURLsToCheck.Item(mCheckIndex) & ")"

            Set mGetter = New DOMDocument30
            Call mGetter.Load(mLinkURLsToCheck.Item(mCheckIndex))
            Debug.Print "Checking: " & mLinkURLsToCheck.Item(mCheckIndex)
        End If
    End If
End Sub

Private Sub txtURL_Change()
    On Error Resume Next
    cmdSearch.Enabled = (Trim(txtURL.Text) <> "" And txtURL.Text <> GetText("Enter the web address for the page to search here and press return."))
End Sub

Private Sub txtURL_GotFocus()
    On Error Resume Next
    txtURL.SelStart = 0
    txtURL.SelLength = Len(txtURL.Text)
End Sub

Private Sub webPageChecker_DocumentComplete(ByVal pDisp As Object, url As Variant)
    On Error Resume Next
    Dim doc As IHTMLDocument3
    Dim a As IHTMLElement
    Dim href As String
    Dim extension As String
    
    If Me.visible Then
        If webPageChecker.Application Is pDisp And url <> "" And url <> "http:///" Then
            'Finished navigation
            staProgress.SimpleText = modI18N.GetText("Checking page for podcasts...")
            Set doc = webPageChecker.Document
            Set mLinkURLsToCheck = New Collection
            Set mResultNames = New Collection
            Set mResultURLs = New Collection
            For Each a In doc.getElementsByTagName("A")
                href = a.getAttribute("href")
                If InStr(1, href, "@") > 0 Then
                    'Don't check mailto:blah@com.com links!
                ElseIf LCase(href) = "about:blank" Then
                    'Don't get blank frames.
                ElseIf href = "" Then
                    'Don't get non-URL links
                Else
                    extension = LCase(Right(href, 4))
                    If extension = ".gif" Or extension = ".jpg" Or extension = ".png" Then
                        'Don't get images
                    ElseIf extension = ".swf" Or extension = ".mp3" Or extension = ".exe" Or extension = ".wav" Then
                        'Don't get common media types
                    Else
                        Call mLinkURLsToCheck.Add(href)
                    End If
                End If
            Next a
            mCheckIndex = 0
            tmrCheck.Enabled = True
        End If
    End If
End Sub


Private Sub webPageChecker_GotFocus()
    On Error Resume Next
    Call lstResults.SetFocus
End Sub
