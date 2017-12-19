VERSION 5.00
Object = "{D0D653FB-B36F-4918-9648-3C495E456DC4}#1.5#0"; "UniBox10.ocx"
Begin VB.Form frmReviewer 
   Caption         =   "Podcast Reviewer"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin UniToolbox.UniText txtURL 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   4335
      _Version        =   65541
      _ExtentX        =   7646
      _ExtentY        =   873
      _StockProps     =   109
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Locked          =   -1  'True
   End
   Begin UniToolbox.UniText txtTitle 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   4335
      _Version        =   65541
      _ExtentX        =   7646
      _ExtentY        =   873
      _StockProps     =   109
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
   End
   Begin UniToolbox.UniText txtPodcastTitle 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
      _Version        =   65541
      _ExtentX        =   7646
      _ExtentY        =   873
      _StockProps     =   109
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Locked          =   -1  'True
   End
   Begin VB.CommandButton cmdAcceptPodcastTitle 
      Caption         =   "&Accept with Podcast Title"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CommandButton cmdSkip 
      Caption         =   "&Skip"
      Height          =   495
      Left            =   1440
      TabIndex        =   9
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdReject 
      Caption         =   "&Reject"
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "Accept but o&verride Podcast Title"
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblURL 
      AutoSize        =   -1  'True
      Caption         =   "&URL"
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   330
   End
   Begin VB.Label lblTitleFromUser 
      AutoSize        =   -1  'True
      Caption         =   "&Override Title"
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1170
   End
   Begin VB.Label ltlTitlePodcast 
      AutoSize        =   -1  'True
      Caption         =   "&Title from Podcast"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1560
   End
End
Attribute VB_Name = "frmReviewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mDoc As DOMDocument30
Attribute mDoc.VB_VarHelpID = -1
Private WithEvents mPodcast As DOMDocument30
Attribute mPodcast.VB_VarHelpID = -1
Private mID As String
Private mStartAt As Long
Private mVoice As New SpVoice

Private Sub cmdAccept_Click()
    On Error Resume Next
    Dim podcastTitle As String
    Dim title As String
    
    txtPodcastTitle.Text = ""
    txtTitle.Text = ""
    txtURL.Text = ""
    podcastTitle = Escape(txtPodcastTitle.Text)
    title = Escape(txtTitle.Text)
    If mPodcast.parseError.errorCode = 0 Then
        'Valid podcast title
        Call mDoc.Load("http://data.webbie.org.uk/podcastReviewList.php?action=approve&id=" & mID & "&podcastTitle=" & podcastTitle & "&title=" & title)
    Else
        'Invalid podcast title
        Call mDoc.Load("http://data.webbie.org.uk/podcastReviewList.php?action=approve&id=" & mID & "&title=" & title)
    End If
End Sub

Private Sub cmdAcceptPodcastTitle_Click()
    On Error Resume Next
    Dim url As String
    
    txtPodcastTitle.Text = ""
    txtTitle.Text = ""
    txtURL.Text = ""
    url = "http://data.webbie.org.uk/podcastReviewList.php?action=approve&id=" & mID & "&podcastTitle=" & txtPodcastTitle.Text
    url = Escape(url)
    Call mDoc.Load(url)
End Sub

Private Sub cmdReject_Click()
    On Error Resume Next
    Dim url As String
    txtPodcastTitle.Text = ""
    txtTitle.Text = ""
    txtURL.Text = ""
    url = "http://data.webbie.org.uk/podcastReviewList.php?action=delete&id=" & mID
    url = Escape(url)
    Set mDoc = New DOMDocument30
    Call mDoc.Load(url)
End Sub

Private Sub cmdSkip_Click()
    On Error Resume Next
    mStartAt = mStartAt + 1
    txtPodcastTitle.Text = ""
    txtTitle.Text = ""
    txtURL.Text = ""
    Call mDoc.Load("http://data.webbie.org.uk/podcastReviewList.php")
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set mDoc = New DOMDocument30
    mStartAt = 0
    Call mDoc.Load("http://data.webbie.org.uk/podcastReviewList.php")
End Sub

Private Sub mDoc_onreadystatechange()
    On Error Resume Next
    Dim n As IXMLDOMNode
    If mDoc.readyState = 4 Then
        If mDoc.parseError.errorCode = 0 Then
            If Not (mDoc.documentElement.selectSingleNode("deleted") Is Nothing) Then
                Debug.Print "DELETED: " & mDoc.documentElement.selectSingleNode("deleted").Text
            End If
            'OK, got data: display first item.
            'MsgBox mDoc.xml
            If mDoc.documentElement.selectNodes("podcast").length = 0 Then
                MsgBox "No podcasts to review, all done!" & vbNewLine & mDoc.xml, vbInformation
            Else
                If mStartAt > mDoc.documentElement.selectNodes("podcast").length - 1 Then mStartAt = 0
                Set n = mDoc.documentElement.selectNodes("podcast").Item(mStartAt)
                mID = n.Attributes.getNamedItem("id").Text
                txtPodcastTitle.Text = ""
                txtTitle.Text = n.selectSingleNode("title").Text
                txtURL.Text = n.selectSingleNode("url").Text
                Me.Caption = "Podcast Reviewer - " & mID
                Set mPodcast = New DOMDocument30
                Call mPodcast.Load(txtURL.Text)
            End If
        Else
            MsgBox "Whoops! Couldn't get data from WebbIE site. Error: " & mDoc.parseError.reason & vbNewLine & """" & mDoc.parseError.srcText & """" & vbNewLine & "Line:" & mDoc.parseError.Line & " Line:" & mDoc.parseError.linepos
            End
        End If
    End If
End Sub

Private Sub mPodcast_onreadystatechange()
    On Error Resume Next
    If mPodcast.readyState = 4 Then
        If mPodcast.parseError.errorCode = 0 Then
            txtPodcastTitle.Text = mPodcast.documentElement.selectSingleNode("channel/title").Text
        Else
            txtPodcastTitle.Text = "Podcast isn't valid!"
        End If
        txtPodcastTitle.SelStart = 0
        txtPodcastTitle.SelLength = Len(txtPodcastTitle.Text)
        Call txtPodcastTitle.SetFocus
    End If
End Sub

Private Sub txtPodcastTitle_GotFocus()
    On Error Resume Next
    txtPodcastTitle.SelStart = 0
    txtPodcastTitle.SelLength = Len(txtPodcastTitle.Text)
End Sub

Private Sub txtTitle_Change()
    On Error Resume Next
    txtTitle.SelStart = 0
    txtTitle.SelLength = Len(txtTitle.Text)
End Sub

Private Sub txtURL_Change()
    On Error Resume Next
    txtURL.SelStart = 0
    txtURL.SelLength = Len(txtURL.Text)
End Sub
