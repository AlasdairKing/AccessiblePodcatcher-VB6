VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPodcaster 
   Caption         =   "Podcatcher"
   ClientHeight    =   6030
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9705
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPodcaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Tag             =   "frmPodcaster"
   Begin SHDocVwCtl.WebBrowser mWebBrowser 
      Height          =   1335
      Left            =   7440
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   840
      Width           =   1455
      ExtentX         =   2566
      ExtentY         =   2355
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
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      TabIndex        =   15
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Timer tmrUpdateAvailable 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5640
      Top             =   4200
   End
   Begin VB.Timer tmrUpdateTrackTimer 
      Interval        =   1000
      Left            =   4440
      Top             =   3960
   End
   Begin MSComctlLib.Slider sliTrack 
      Height          =   555
      Left            =   720
      TabIndex        =   12
      Top             =   3600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   979
      _Version        =   393216
      Enabled         =   0   'False
      LargeChange     =   30
      SmallChange     =   5
      TickFrequency   =   10
   End
   Begin VB.CommandButton cmdBackwards 
      Caption         =   "Bac&k"
      Enabled         =   0   'False
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Tag             =   "frmPodcaster.cmdBackwards"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdForwards 
      Caption         =   "For&ward"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Tag             =   "frmPodcaster.cmdForwards"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Frame fraVolume 
      BorderStyle     =   0  'None
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   8
      Top             =   3000
      Width           =   3255
      Begin MSComctlLib.Slider sliVolume 
         Height          =   495
         Left            =   960
         TabIndex        =   10
         ToolTipText     =   "Adjusts the volume"
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   10
         Max             =   100
         SelStart        =   50
         TickStyle       =   3
         TickFrequency   =   10
         Value           =   50
      End
      Begin VB.Label lblVolume 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "&Volume:"
         Height          =   285
         Left            =   -165
         TabIndex        =   9
         Tag             =   "frmPodcaster.lblVolume"
         Top             =   0
         Width           =   900
      End
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   2400
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar staMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   5655
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11456
            Key             =   "status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "progress"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Key             =   "downloaded"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox lstItems 
      Height          =   2340
      Left            =   3240
      TabIndex        =   3
      Top             =   480
      Width           =   2535
   End
   Begin VB.ListBox lstPodcasts 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdStop 
      Cancel          =   -1  'True
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5160
      TabIndex        =   7
      Tag             =   "frmPodcaster.cmdStop"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "P&lay"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Tag             =   "frmPodcaster.cmdPlay"
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label lblTrack 
      AutoSize        =   -1  'True
      Caption         =   "Wh&ere:"
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Tag             =   "frmPodcaster.lblTrack"
      Top             =   3600
      Width           =   765
   End
   Begin VB.Label lblItems 
      AutoSize        =   -1  'True
      Caption         =   "&Items:"
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Tag             =   "frmPodcaster.lblItems"
      Top             =   120
      Width           =   690
   End
   Begin VB.Label lblPodcast 
      AutoSize        =   -1  'True
      Caption         =   "&Podcast:"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Tag             =   "frmPodcaster.lblPodcast"
      Top             =   120
      Width           =   900
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmpPlayer 
      Height          =   3255
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3960
      Width           =   3735
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   -1  'True
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6588
      _cy             =   5741
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Tag             =   "frmPodcaster.mnuFile"
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
         Tag             =   "frmPodcaster.mnuFileSave"
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "&Import"
         Tag             =   "frmPodcaster.mnuFileImport"
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "&Export"
         Tag             =   "frmPodcaster.mnuFileExport"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Tag             =   "frmPodcaster.mnuFileExit"
      End
   End
   Begin VB.Menu mnuPodcasts 
      Caption         =   "Pod&casts"
      Tag             =   "frmPodcaster.mnuPodcasts"
      Begin VB.Menu mnuPodcastsGetnewpodcasts 
         Caption         =   "&Get new podcasts"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuPodcastsDownloadpodcasts 
         Caption         =   "D&ownload podcasts (closes Podcatcher)"
      End
      Begin VB.Menu mnuPodcastsAddpodcasts 
         Caption         =   "&Add podcast"
         Tag             =   "frmPodcaster.mnuPodcastsAddpodcasts"
      End
      Begin VB.Menu mnuPodcastsDelete 
         Caption         =   "&Delete podcast"
         Enabled         =   0   'False
         Tag             =   "frmPodcaster.mnuPodcastsDelete"
      End
      Begin VB.Menu mnuPodcastsEdit 
         Caption         =   "&Edit podcast"
         Tag             =   "frmPodcaster.mnuPodcastsEdit"
      End
      Begin VB.Menu mnuPodcastsSort 
         Caption         =   "&Sort podcasts"
         Tag             =   "frmPodcaster.mnuPodcastsSort"
      End
      Begin VB.Menu mnuPodcastsSearchwebpage 
         Caption         =   "Search a &Webpage for Podcasts"
      End
      Begin VB.Menu mnuPodcastsRenamepodcast 
         Caption         =   "&Rename podcast"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuItems 
      Caption         =   "I&tems"
      Begin VB.Menu mnuItemsDeleteitem 
         Caption         =   "&Delete (hide) item"
      End
      Begin VB.Menu mnuItemsShowdeleteditems 
         Caption         =   "&Show deleted (hidden) items"
      End
      Begin VB.Menu mnuItemsOpeniteminwindowsmediaplayer 
         Caption         =   "&Open item in Windows Media Player"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuItemsCopyitemurltoclipboard 
         Caption         =   "&Copy item URL to clipboard"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuAudio 
      Caption         =   "&Audio"
      Tag             =   "frmPodcaster.mnuAudio"
      Begin VB.Menu mnuAudioPlay 
         Caption         =   "&Play"
         Enabled         =   0   'False
         Shortcut        =   ^P
         Tag             =   "frmPodcaster.mnuAudioPlay"
      End
      Begin VB.Menu mnuAudioStop 
         Caption         =   "&Stop"
         Enabled         =   0   'False
         Tag             =   "frmPodcaster.mnuAudioStop"
      End
      Begin VB.Menu mnuAudioPause 
         Caption         =   "&Pause"
         Enabled         =   0   'False
         Shortcut        =   ^U
         Tag             =   "frmPodcaster.mnuAudioPause"
      End
      Begin VB.Menu mnuAudioSkipforwards 
         Caption         =   "Skip for&wards"
         Enabled         =   0   'False
         Shortcut        =   ^F
         Tag             =   "frmPodcatcher.mnuAudioSkipforwards"
      End
      Begin VB.Menu mnuAudioSkipbackwards 
         Caption         =   "Skip bac&kwards"
         Enabled         =   0   'False
         Shortcut        =   ^B
         Tag             =   "frmPodcatcher.mnuAudioSkipbackwards"
      End
      Begin VB.Menu mnuAudioIncreasevolume 
         Caption         =   "&Increase volume"
         Shortcut        =   ^I
         Tag             =   "frmPodcaster.mnuAudioIncreaseVolume"
      End
      Begin VB.Menu mnuAudioDecreasevolume 
         Caption         =   "&Decrease volume"
         Shortcut        =   ^D
         Tag             =   "frmPodcaster.mnuAudioDecreaseVolume"
      End
   End
   Begin VB.Menu mnuVideo 
      Caption         =   "Vi&deo"
      Begin VB.Menu mnuVideoShowvideo 
         Caption         =   "&Show video"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuVideoFullscreen 
         Caption         =   "&Fullscreen"
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsUsebrokenfeeds 
         Caption         =   "&Try to use broken feeds"
      End
      Begin VB.Menu mnuOptionsAllowsubmissions 
         Caption         =   "&Allow anonymous submission of new podcasts to WebbIE directory"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Tag             =   "frmPodcaster.mnuHelp"
      Begin VB.Menu mnuHelpManual 
         Caption         =   "&Manual"
         Tag             =   "frmPodcaster.mnuHelpManual"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
         Tag             =   "frmPodcaster.mnuHelpAbout"
      End
   End
End
Attribute VB_Name = "frmPodcaster"
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

'1.5.6 Pre-13 October 2006
'1.5.7
'   Fixed pressing non-cursor-key on slider causing stutter
'   Added colons to labels
'   Added Where lblTrack to internationalisation file
'   Fixed the slider control: now lets you move along
'1.5.8
'   Fixed the version Help window to show all 3 parts of number.
'1.5.9
'   Changed all fonts to Tahoma (default Windows XP)
'   Renamed sbMain to staMain.
'1.5.10
'   Changed skip value to 60 seconds
'   Fixed bug with slider disappearing every five minutes because of updates.
'   Stopped it changing the UI language to en-gb when run (!)
'   Made help file, default podcast files, language files all resource files built into
'   exe, so no need for external files.
'1.5.11
'   Added Polish translations from Dorota
'   Rewrote CLanguage Class_Init to read Windows system locale and try to find
'   a matching language to use rather than simply assuming English.
'1.5.12
'   Updated Clanguage to apply charset to every form and control element
'   to support Polish and other languages, I think!
'   Fixed language handling in CLanguage for loading the XML: this will
'   effect EVERY WebbIE/Accessible program. Sigh.
'   Made it load default podcasts from WebbIE site in preference to local
'   copy.
'1.5.13
'   20 Feb 2007: Added handling for Unicode manuals, not ANSI, so can do Polish.
'1.5.14 1 April 2007
'   Change to default language handling in CLanguage
'1.6.0 27 June 2007
'   Problem reported by Niki at Radio Reading Service of the Rockies: their
'   podcast is password-protected, so .Load doesn't work. So added the whole
'   XMLObject code from Accessible RSS so they can see the prompt.
'   Saves user config in podcatcher.ini, not registry.
'   Saves podcast information in podcatcher.xml, not registry.
'1.6.1 10 July 2007
'   Updated user interface a bit to make it display sensible messages in
'   item list and to let you tab from podcast list and get new item list.
'   Fixed sorting, so it actually sorts.
'1.6.2 13 July 2007
'   Fixed deleting podcasts, so Delete menu item is enabled and podcasts are
'   actually deleted.
'1.6.3 13 August 2007
'   If you delete an item, keep the list index the same.
'1.6.4 21 August 2007
'   Set flag to save item to directory that must exist.
'1.6.5 31 October 2007
'   Added Spanish manual, some more translations.
'   Fixed the menu item for deleting (hiding) items.
'1.7.0     6 Jan 2008
'   Added support for large fonts.
'   Added support for updates.
'   Added remembering where you are on the screen.
'1.7.1      29 Jan 2008
'   Updated language handling.
'   Did XP Style.
'1.7.2      25 Feb 2008
'   Allowed broken feeds - well, the one from http://www.comproom.co.uk/ourplace/ourplace.xml -
'   to load and be parsed.
'1.7.3      2 March 2008
'   Fixed opening Podcatcher making every podcast in Podcast Downloader subscribed.
'1.7.4      23 March 2008
'   Fixed I18N bug.
'1.7.5      8 April 2008
'   Fixed tab order (no longer get lost on IE)
'1.7.6      3 June 2008
'   Fixed deleted items, when shown, not being able to be played.
'1.7.7      15 June 2008
'   Disabled automatic item check.
'1.7.8      01 July 2008
'   Report from user that the Freedom Scientific podcast (http://www.freedomscientific.com/FSCast/rss.xml)
'   won't work after it has been added, then the application closed, then reopened. Sure enough, tried it
'   and the rss.xml has been stripped off the end. Found and fixed a bug so that amending the XML for
'   the podcatcher gets saved (it wasn't at all). This appears to have fixed the Freedom post, don't know why.
'1.8.0  4 September 2008
'   Added pause button to controls.
'   Added Video menu: Control and V shows the video, Control and N goes to fullscreen.
'1.9.0
'   21 Dec 2008. Added ability to search web pages for Podcast feeds.
'1.9.1
'   24 Dec 2008.   Fixed parsing of HTML in items, like in Accessible RSS. Though I'm not sure Accessible RSS does
'                       such clever parsing.
'1.9.2
'   18 Jan 2009. Fixed Pause button - didn't do anything.
'   18 Jan 2009. Fixed tab order for buttons and when pressed (so focus goes to something sensible)
'1.9.3
'   18 Jan 2009. Finessed the tab sequence a bit to make it possible to tab around the UI without changing item.
'1.9.4
'   11 Mar 2009. Fixed deleting items from the Items list.
'1.9.5
'   23 Mar 2009. Adds http:// to urls if missing.
'1.9.6
'   27 May 2009. If no description, don't add " - " to items. Shows feed description in podcast list.
'1.9.7
'   13 Jun 2009. Fixed XP Style bug.
'1.10.0
'   16 Jun 2009. Added mechanism to inform my WebbIE directory of new podcasts added to the program.
'1.11.0
'   24 Aug 2009. Added mechanism to get podcasts from WebbIE directory of podcasts.
'                Added ShellExecute to handle Apple audio/video formats - AAC, AIFF, MOV. Only works on URL ending, and if WMP doesn't
'                   handle it properly.
'                Made program open Podcast Downloader from Podcasts menu (closes self so stays in sync.)
'1.11.1
'   25 Sep 2009. Can rename podcasts with F2.
'1.12.0
'   29 Sep 2009. Fixed buffering causing poor quality with .MOV files.
'                Stopped application launching external media player if error occurs.
'1.12.1
'   20 Oct 2009. Fixed Export option showing "Open" on button not "Save"
'1.12.2
'   29 Oct 2009. Put asking for URL before asking for name so prompt can be useful.
'1.12.3
'   27 Jan 2010. Added name of track playing to title bar.
'                F2 now does rename of currently-selected podcast.
'1.13.0
'   16 August 2011. Added ability to copy the URL of an item to the clipboard, so you can
'       open it in something else.
'   16 August 2011. Added abiilty to open a podcast in Windows Media Player. See Items menu.

'TO DO
'   Encodings on pages that aren't valid break Podcatcher.
'   Save file is really crappy if you haven't downloaded the whole
'       audio file yet. Everything hangs. Need to use winsock or
'       something to give progress bar etc.

Public podcasts As Collection
Private mCurrentPodcast As Long
Public mUpdatingSubscriptions As Boolean ' indicates that we're
    'downloading all subscribed files.
Public mItems As Collection  ' the items we are currently downloading
    'when mUpdatingSubscriptsion is true
Public mItemCount As Long 'where we're up to in the items that
    'need downloading
Private mManuallyChangingTrackSlider As Boolean ' whether the user is
    'changing the slider for track position, in which case we mustn't
    'change it on account of the timer.
Private mTrackValueBeforeChanging As Long ' the value before we started
    'changing the slider
Private Const CHECK_FOR_UPDATES_TIME  As Long = 300 ' the number of seconds between
    'automatic updates of the subscribed podcasts.
Private WithEvents mCurrentFeed As DOMDocument30 ' the current feed in the application
Attribute mCurrentFeed.VB_VarHelpID = -1
    'selected in the lstPodcasts list
Private gstrLoadingPodcast As String
Private gstrNotLoadedYet As String
Private gstrAllDeleted As String
Private gstrNotAvailable As String
Private gstrNoPodcasts As String

Private Const SKIP_TIME As Long = 60 ' how many seconds we skip about with back and
    'forwards
    
Private Const Document As Long = 0 ' to enforce capitalisation of MSHTML Document objects
    '- and yes, it does matter!
    
'Test if an array is valid
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

'To open urls in external media player.
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOW = 1


Private Sub cmdBackwards_Click()
    On Error Resume Next
    Dim currentPosition As Long
    
    currentPosition = wmpPlayer.Controls.currentPosition
    If currentPosition < SKIP_TIME Then
        currentPosition = 0
        Call Beep
    End If
    wmpPlayer.Controls.currentPosition = currentPosition - SKIP_TIME
End Sub

Private Sub DeletePodcast()
    On Error Resume Next
    Dim result As Long
    
    If lstPodcasts.ListIndex > -1 And podcasts.Count > 0 Then
        result = MsgBox(modI18N.GetText("Are you sure you want to delete") & " " & lstPodcasts.List(lstPodcasts.ListIndex) & "?", vbYesNoCancel, modI18N.GetText("Confirm delete podcast"))
        If result = vbYes Then
            Call podcasts.Item(lstPodcasts.ListIndex + 1).MarkForDeath
            Call podcasts.Remove(lstPodcasts.ListIndex + 1)
            Call lstPodcasts.RemoveItem(lstPodcasts.ListIndex)
            Call Display
            Call lstPodcasts.SetFocus
        End If
    End If
End Sub

Private Function GetPodcastItems(podcastCollection As Collection) As Collection
'iterates through the podcastCollection finding the items and adding them to GetPodcastItems
    On Error Resume Next
    Dim xmlObject As New MSXML2.XMLHTTP30
    Dim result As Long
    Dim parsedResult As New DOMDocument30
    Dim selected As Long
    
    Dim podcastIterator As CPodcast
    Dim items As Collection
    Dim itemIterator As CItem
    
    Set items = New Collection
    For Each podcastIterator In podcastCollection
        Debug.Print "Starting podcast process"
retryXMLObject2:
        DoEvents
'        If frmWait.Cancel Then
'            Call frmWait.Hide
'            Call Me.Show
'            Exit Function
'        End If
        Debug.Print "Getting podcast details: " & podcastIterator.name
        parsedResult.async = False
        Call parsedResult.Load(podcastIterator.url)
        If parsedResult.parseError = 0 Then
            Debug.Print "Got podcast details okay"
            'okay, this parses okay: what is it?
            If parsedResult.documentElement.nodeName = "rss" Then
                'okay, this is a podcast list
                'assume this is okay to add
                podcastIterator.xml = parsedResult.xml
                'add items to collection
                For Each itemIterator In podcastIterator.items
                    itemIterator.Path = podcastIterator.name
                    Debug.Print "Adding " & itemIterator.name
                    Call items.Add(itemIterator)
                    'DEV: frmWait no longer used
                    'frmWait.lblStatus.Caption = modi18n.GetText("Identified") & " " & items.Count & " " & modi18n.GetText("items to download")
                    'Call frmWait.Refresh
                    DoEvents
                Next itemIterator
            Else
                'no other type of podcast list is supported
            End If
        Else
            'failed to parse: go no further with this podcast
            Debug.Print "Failed to parse a podcast"
        End If
    Next podcastIterator
    Set GetPodcastItems = items
'''    Exit Function
'''networkFailure2:
'''    'network failure (strictly, a .send failure) may be because the network
'''    'is unavailable or because (annoyingly enough) an HTTP redirect has been
'''    'returned: unfortunately, this throws an error rather than allowing anyone
'''    'to access the HTTP response. Sigh. So, let's try using Internet Explorer
'''    'to call the url and see where it ends up
'''    Dim ieObj As Object
'''    Set ieObj = CreateObject("InternetExplorer.Application")
'''    Call ieObj.navigate(podcastIterator.url)
'''    Debug.Print "Navigating: " & podcastIterator.url
'''    giveUpCounter = 0
'''    While ieObj.readyState < READYSTATE_COMPLETE
'''        Debug.Print "readyState:" & ieObj.readyState
'''        DoEvents
'''        giveUpCounter = giveUpCounter + 1
'''        If frmWait.cancel Then
'''            Set ieObj = Nothing
'''            Exit Function
'''        End If
'''
'''    Wend
'''    If ieObj.LocationURL <> podcastIterator.url Then
'''        podcastIterator.url = ieObj.LocationURL
'''        Set ieObj = Nothing
'''        GoTo retryXMLObject2
'''    End If
'''    Set ieObj = Nothing
'''    Debug.Print Err.Description
End Function



Private Sub cmdForwards_Click()
    On Error Resume Next
    Dim currentPosition As Long
    
    currentPosition = wmpPlayer.Controls.currentPosition
    'Debug.Print "Forwards"
    If currentPosition < wmpPlayer.currentMedia.duration - SKIP_TIME Then
        wmpPlayer.Controls.currentPosition = currentPosition + SKIP_TIME
    Else
        Beep
    End If
End Sub

Private Sub GetPodcast()
'get a selected podcast and display item contents
    On Error Resume Next
    Dim result As Long
    Dim parsedResult As New DOMDocument30
    Dim selected As Long
    
    If lstPodcasts.ListIndex > -1 And podcasts.Count > 0 Then
        'okay, we have a podcast to check out
        On Error Resume Next
        staMain.Panels("status").Text = modI18N.GetText("Getting podcast")
        cmdPlay.Enabled = False
        mnuAudioPlay.Enabled = False
        cmdStop.Enabled = False
        lblTrack.Enabled = False
        sliTrack.Enabled = False
        mnuAudioStop.Enabled = False
        parsedResult.async = False
        Set mCurrentFeed = New DOMDocument30
        'And in some fairytale ideal world, all we'd have to do is .Load.
        'Here in the real world we need to use IE to proxy to the actual url, because
        '(1) the URL has a redirect.
        '(2) the URL is password-protected.
        '(3) other stuff I don't know.
        'So call LoadFeed instead.
        Call LoadFeed(podcasts.Item(lstPodcasts.ListIndex + 1).url)
        'Call mCurrentFeed.Load(podcasts.item(lstPodcasts.ListIndex + 1).url)
'''        Call parsedResult.Load(podcasts.item(lstPodcasts.ListIndex + 1).url)
'''        'check whether this is a podcast or an OPML list
'''        staMain.Panels("status").Text = modi18n.GetText("Processing podcast")
'''        'Call parsedResult.loadXML(xmlObject.responseText)
'''        If parsedResult.parseError = 0 Then
'''            'okay, this parses okay: what is it?
'''            If parsedResult.documentElement.nodeName = "rss" Then
'''                'okay, this is a podcast list
'''                'assume this is okay to add
'''                podcasts(lstPodcasts.ListIndex + 1).xml = parsedResult.xml
'''                'add successful podcast
'''                mCurrentPodcast = lstPodcasts.ListIndex + 1
'''                Call DisplayItems
'''            ElseIf parsedResult.documentElement.nodeName = "opml" Then
'''                'okay, this is an opml file
'''                result = MsgBox(modi18n.GetText("This is a list of podcasts. Do you want to add them to your list?"), vbYesNoCancel, modi18n.GetText("New podcasts"))
'''                If result = vbYes Then
'''                    staMain.Panels("status").Text = modi18n.GetText("Importing")
'''                    Call ImportOPML(parsedResult)
'''                    Call Display
'''                End If
''''            ElseIf parsedResult.documentElement.nodeName = "rdf:RDF" Or parsedResult.documentElement.nodeName = "rdf" Then
''''                MsgBox "RDF!"
'''            Else
'''                Call MsgBox(modi18n.GetText("Sorry, this is not a supported type of podcast. I could not do anything with it."), vbOKOnly, modi18n.GetText("Podcast unknown"))
'''                'Debug.Print "Unknown type: " & Left(parsedResult.xml, 500)
'''                Debug.Print "NN:" & parsedResult.documentElement.nodeName
'''            End If
'''        Else
'''            Debug.Print parsedResult.parseError.errorCode
'''            Debug.Print parsedResult.parseError.reason
'''
'''            'failed to parse: ask to remove
'''            result = MsgBox(modi18n.GetText("Sorry, I could not use this podcast. Do you want to remove it from your list of podcasts?"), vbYesNoCancel, modi18n.GetText("Broken podcast"))
'''            If result = vbYes Then
'''                podcasts(lstPodcasts.ListIndex + 1).markedForDeath = True
'''                Call podcasts.Remove(lstPodcasts.ListIndex + 1)
'''                Call Display
'''            End If
'''        End If
'''        staMain.Panels("status").Text = modi18n.GetText("Done")
    End If
End Sub

Private Sub DisplayItems()
    On Error Resume Next
    Dim podcast As CPodcast
    Dim podcastItems As Collection
    Dim itemIterator As CItem
    Dim gotoItems As Boolean
    
    gotoItems = (Me.ActiveControl.name = lstItems.name)
    
    'Debug.Print "DI Before: " & Me.ActiveControl.name
    If lstPodcasts.ListIndex > -1 Then
        Call lstItems.Clear
        Set podcast = podcasts.Item(lstPodcasts.ListIndex + 1)
        If podcast.parseError = True Then
            Call lstItems.AddItem(gstrNotAvailable)
        Else
            'parsed okay, display contents.
            Set podcastItems = podcast.items(mnuItemsShowdeleteditems.Checked)
            For Each itemIterator In podcastItems
                Call lstItems.AddItem(itemIterator.fullname)
            Next itemIterator
            If podcastItems.Count = 0 Then
                Call lstItems.Clear
                Call lstItems.AddItem(gstrNoPodcasts)
            End If
            'Don't do this: it confuses screen readers.
    '        lstItems.ListIndex = 0
        End If
    End If
    If gotoItems Then
        Call lstItems.SetFocus
        lstItems.ListIndex = 0 ' okay to do this here, it's the current control
    End If
    'Debug.Print "DI After: " & Me.ActiveControl.name
End Sub

Private Sub cmdPause_Click()
    On Error Resume Next
    Call mnuAudioPause_Click
    Call cmdPlay.SetFocus
End Sub

Private Sub cmdPlay_Click()
    On Error Resume Next
    Dim mp3File As String
    'Dim wmpPlaylist As wmmccc
    
    If lstItems.ListIndex > -1 And lstItems.List(0) <> gstrNotLoadedYet Then
        mp3File = podcasts.Item(mCurrentPodcast).items(mnuItemsShowdeleteditems.Checked).Item(lstItems.ListIndex + 1).url
        'Debug.Print "mp3:[" c& mp3File & "]"
        'Call wmpPlayer.mediaCollection
        If wmpPlayer.url = mp3File Then
            'already loaded this, go and play it
            Call wmpPlayer.Controls.play
        Else
            'need to load this
            wmpPlayer.url = mp3File
        End If
        cmdForwards.Enabled = True
        cmdBackwards.Enabled = True
        mnuAudioSkipbackwards.Enabled = True
        mnuAudioSkipforwards.Enabled = True
        mnuAudioPlay.Enabled = False
        cmdStop.Enabled = True
        lblTrack.Enabled = True
        sliTrack.Enabled = True
        mnuAudioStop.Enabled = True
        mnuAudioPause.Enabled = True
        cmdPause.Enabled = True
        Call cmdPause.SetFocus
        cmdPlay.Enabled = False
        Me.Caption = App.title & " - " & lstItems.List(lstItems.ListIndex)
    End If
End Sub

Private Sub cmdStop_Click()
    On Error Resume Next
    Call wmpPlayer.Controls.stop
    cmdPlay.Enabled = True
    cmdForwards.Enabled = False
    cmdBackwards.Enabled = False
    mnuAudioSkipbackwards.Enabled = False
    mnuAudioSkipforwards.Enabled = True
    mnuAudioPlay.Enabled = True
    cmdStop.Enabled = False
    mnuAudioStop.Enabled = False
    If mnuVideoShowvideo.Checked Then
        Call mnuVideoShowvideo_Click
    End If
    Call lstItems.SetFocus
    Me.Caption = App.title
End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    Call modXPStyle.InitCommonControlsVB
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyEscape Then
        Call cmdStop_Click
    ElseIf KeyCode = vbKeyF1 Then
        Call mnuHelpManual_Click
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim newPodcast As CPodcast
    Dim opmlFile As DOMDocument30
    Dim b() As Byte
    Dim fso As Scripting.FileSystemObject
    Dim result As String
    Dim newXML As DOMDocument30
    
    mWebBrowser.TabStop = False
    wmpPlayer.TabStop = False
    mWebBrowser.Silent = True
    Call modI18N.ApplyUILanguageToThisForm(Me)
    'Resize according to Windows font sizes
    Call modLargeFonts.ApplySystemSettingsToForm(Me, , True)
    'Check for updates
    Call modUpdate.CheckForUpdates
    'get layout
    Call modRememberPosition.LoadPosition(Me)
    'Call Load(frmWait)
    Call Load(frmHelp)
    Call Load(frmSearch)
    Set fso = New Scripting.FileSystemObject
    If Not fso.FolderExists(modPath.settingsPath & "\Deleted") Then Call fso.CreateFolder(modPath.settingsPath & "\Deleted")
    Set fso = Nothing
    'load options
    mnuOptionsUsebrokenfeeds.Checked = modPath.GetSettingIni(App.title, "Feeds", "UseBrokenFeeds", CStr(True))
    mnuOptionsAllowsubmissions.Checked = modPath.GetSettingIni(App.title, "Privacy", "AllowAnonymousSubmissions", CStr(True))
    'do internationalisation
    Call modI18N.ApplyUILanguageToAllForms
    gstrLoadingPodcast = modI18N.GetText("Getting podcast, please wait.")
    gstrNotLoadedYet = modI18N.GetText("Select a podcast and press return to load its contents.")
    gstrAllDeleted = modI18N.GetText("No new items.")
    gstrNotAvailable = modI18N.GetText("Podcast not available.")
    gstrNoPodcasts = modI18N.GetText("No podcasts.")
    Call lstItems.AddItem(gstrNotLoadedYet)
    
    Call LoadPodcasts
    Call Display
    Me.mnuAudioStop.Caption = Me.mnuAudioStop.Caption & vbTab & "Escape"
    Set fso = New Scripting.FileSystemObject
    If fso.FileExists(modPath.settingsPath & "\podcatcher.ini") Then
        wmpPlayer.settings.volume = CLng(modShared.SharedReadIniFileDefault(modPath.settingsPath & "\podcatcher.ini", "User Settings", "Volume", "50"))
        sliVolume.value = CLng(modShared.SharedReadIniFileDefault(modPath.settingsPath & "\podcatcher.ini", "User Settings", "Volume", "50"))
    Else
        wmpPlayer.settings.volume = CLng(GetSetting("AccessiblePodcaster", "User settings", "Volume", 50))
        sliVolume.value = CLng(GetSetting("AccessiblePodcaster", "User settings", "Volume", 50))
        Call SaveSetting("AccessiblePodcaster", "User Settings", "Dummy", "This is a dummy key to be deleted.")
        Call DeleteSetting("AccessiblePodcaster")
    End If

    'Set winsockHandler = New CWinsockHandler
    'Set winsockHandler.cWinsock = Me.Winsock
    Call Display
    Call Me.Show
    Call lstPodcasts.SetFocus
End Sub

Private Sub LoadPodcasts()
    On Error Resume Next
    Dim newXML As DOMDocument30
    Dim opmlFile As DOMDocument30
    Dim b() As Byte
    Dim newPodcast As CPodcast
    
    'Move to XML:
    Call RestorePodcastsFromRegistry
    If podcasts Is Nothing Then
        Set podcasts = New Collection
    End If
    If podcasts.Count > 0 Then
        'we got some podcasts from the registry: we'll delete these when
        'we exit the program, and we'll save the new podcast information
        'as xml when we do that. But create an empty document ready for that anyway.
        Set newXML = New DOMDocument30
        newXML.async = False
        Call newXML.loadXML("<podcasts/>")
        Call newXML.save(modPath.settingsPath & "\podcasts.xml")
    Else
        Call RestorePodcastsFromXML
    End If
    If podcasts.Count = 0 Then
        'no podcasts found: load defaults from WebbIE site
        Set opmlFile = New DOMDocument30
        opmlFile.async = False
        Call opmlFile.Load("http://data.webbie.org.uk/defaultPodcasts.opml")
        If opmlFile.parseError.errorCode <> 0 Then
            b() = VB.LoadResData("COMMON", "DEFAULTPODCASTS")
            Call opmlFile.loadXML(StrConv(b(), vbUnicode))
        End If
        If opmlFile.parseError = 0 Then
            'loaded XML okay
            Call ImportOPML(opmlFile)
        Else
            'failed to load local podcast default file: just add
            'the BBC!
            Set newPodcast = New CPodcast
            newPodcast.name = "In Our Time (BBC Radio 4)"
            newPodcast.url = "http://www.bbc.co.uk/radio4/history/inourtime/mp3/podcast.xml"
            Call podcasts.Add(newPodcast)
        End If
'        Set newPodcast = New CPodcast
'        newPodcast.name = "In Our Time (BBC 4)"
'        newPodcast.url = "http://www.bbc.co.uk/radio4/history/inourtime/mp3/podcast.xml"
'        Call podcasts.Add(newPodcast)
'        Set newPodcast = New CPodcast
'        newPodcast.name = "The Indie Analyst Show"
'        newPodcast.url = "http://www.indieanalyst.com/podcasts/index.xml"
'        Call podcasts.Add(newPodcast)
'        Call podcasts.Add(newPodcast)
'        Set newPodcast = New CPodcast
'        newPodcast.name = "Left, Right and Centre"
'        newPodcast.url = "http://kcrw.com/podcast/show/lr"
'        Call podcasts.Add(newPodcast)
'        Set newPodcast = New CPodcast
'        newPodcast.name = "Teach42"
'        newPodcast.url = "http://feeds.feedburner.com/Teachfourtwo"
'        Call podcasts.Add(newPodcast)
'        Set newPodcast = New CPodcast
'        newPodcast.name = "Kitchen Radio"
'        newPodcast.url = "http://www.frisatsun.com/kr/kr.xml"
'        Call podcasts.Add(newPodcast)
'        Set newPodcast = New CPodcast
'        newPodcast.name = "Family Health Radio"
'        newPodcast.url = "http://fhradio.org/RSS/rss.xml"
'        Call podcasts.Add(newPodcast)
'        Set newPodcast = New CPodcast
'        newPodcast.name = "Reel Reviews - Films worth Watching"
'        newPodcast.url = "http://reelreviewsradio.com/podcast.xml"
'        Call podcasts.Add(newPodcast)
'        Set newPodcast = New CPodcast
'        newPodcast.name = "It's a Jazz Thing"
'        newPodcast.url = "http://feeds.feedburner.com/jazzthing"
'        Call podcasts.Add(newPodcast)
'        Set newPodcast = New CPodcast
'        newPodcast.name = "Newsweek"
'        newPodcast.url = "http://anon.newsweek.speedera.net/anon.newsweek/podcasts/podcast_onair.xml"
'        Call podcasts.Add(newPodcast)
'        Set newPodcast = New CPodcast
'        newPodcast.name = "Sports Bloggers Live"
'        newPodcast.url = "http://sportsbloggerslive.podcast.aol.com/sportsbloggerslive_rss.xml"
'        Call podcasts.Add(newPodcast)
'        Set newPodcast = New CPodcast
'        newPodcast.name = "Geek News Cenftral"
'        newPodcast.url = "http://www.geeknewscentral.com/podcast.xml"
'        Call podcasts.Add(newPodcast)
    End If
End Sub

'Try to get podcast information from registry
Private Sub RestorePodcastsFromRegistry()
    On Error Resume Next
    Dim newPodcast As New CPodcast
    Dim gotFromRegistry() As RegPair
    Dim subscribedState As Boolean
    Dim i As Integer
    
    Set podcasts = New Collection
    On Error GoTo FailedReg:
    gotFromRegistry = GetRegistryEntries("AccessiblePodcaster", "Podcasts")
    If SafeArrayGetDim(gotFromRegistry) > 0 Then
        For i = LBound(gotFromRegistry) To UBound(gotFromRegistry)
            Set newPodcast = New CPodcast
            newPodcast.name = gotFromRegistry(i).key
            newPodcast.url = gotFromRegistry(i).value
            subscribedState = CBool(GetSetting("AccessiblePodcaster", "Subscribed Podcasts", newPodcast.name, CStr("False")))
            newPodcast.subscribed = subscribedState
            Call podcasts.Add(newPodcast)
        Next i
    End If
    'newPodcast.name = "Caribbean Radio"
    'newPodcast.url = "http://feeds.feedburner.com/CaribbeanFreeRadioBlog"
    'Call podcasts.Add(newPodcast)
    Exit Sub
FailedReg:
    Exit Sub
End Sub

'Try to get podcast information from XML file
Private Sub RestorePodcastsFromXML()
    On Error Resume Next
    Dim newPodcast As New CPodcast
    Dim gotFromRegistry() As RegPair
    Dim subscribedState As Boolean
    Dim i As Integer
    Dim podcastXML As DOMDocument30
    Dim podcastIterator As IXMLDOMNode
    
    Set podcastXML = New DOMDocument30
    podcastXML.async = False
    Call podcastXML.Load(modPath.settingsPath & "\podcasts.xml")
    Set podcasts = New Collection
    If podcastXML.parseError.errorCode = 0 Then
        For Each podcastIterator In podcastXML.documentElement.selectNodes("podcast")
            Set newPodcast = New CPodcast
            newPodcast.name = podcastIterator.selectSingleNode("name").Text
            'Debug.Print "Loaded " & newPodcast.name
            newPodcast.url = podcastIterator.selectSingleNode("url").Text
            If podcastIterator.selectSingleNode("subscribed2") Is Nothing Then
                newPodcast.subscribed = True
            Else
                newPodcast.subscribed = podcastIterator.selectSingleNode("subscribed2").Text
            End If
            Call podcasts.Add(newPodcast)
        Next podcastIterator
    Else
        Call podcastXML.loadXML("<podcasts/>")
    End If
End Sub


'Iterate through all the podcasts displaying them in lstPodcasts
Public Sub Display()
    On Error Resume Next
    Dim podcastIterator As CPodcast
    Dim selected As Long
    Dim gotoItems As Boolean
    Dim gotoPodcasts As Boolean
    
    If Me.ActiveControl Is Nothing Then
        'nothing to choose, probably starting up
    Else
        gotoItems = (lstItems.name = Me.ActiveControl.name)
        gotoPodcasts = (lstPodcasts.name = Me.ActiveControl.name)
    End If
    selected = lstPodcasts.ListIndex
    Call lstPodcasts.Clear
    Call lstItems.Clear
    If podcasts.Count = 0 Then
        'no podcasts: warn user
        Call lstPodcasts.AddItem(modI18N.GetText("No podcasts available."))
    Else
        'got some podcasts: display their names
        For Each podcastIterator In podcasts
            Call lstPodcasts.AddItem(podcastIterator.fullname)
        Next podcastIterator
        'restore position on list
        If selected > -1 Then
            If selected >= lstPodcasts.ListCount Then
                selected = lstPodcasts.ListIndex
            End If
            lstPodcasts.ListIndex = selected
        End If
        Call lstItems.Clear
        Call lstItems.AddItem(gstrNotLoadedYet)
    End If
    
    If gotoPodcasts Then Call lstPodcasts.SetFocus
    If gotoItems Then Call lstItems.SetFocus
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        If Me.ScaleWidth > 5700 Then
            lblPodcast.Left = GAP
            lstPodcasts.Width = Me.ScaleWidth * 0.5 - 3 * GAP / 2
            lstPodcasts.Left = GAP
            lstItems.Width = lstPodcasts.Width
            lstItems.Left = lstPodcasts.Left + lstPodcasts.Width + GAP
            lblItems.Left = lstItems.Left
            cmdBackwards.Left = GAP
            cmdPlay.Left = cmdBackwards.Left + cmdBackwards.Width
            cmdPause.Left = cmdPlay.Left + cmdPlay.Width
            cmdStop.Left = cmdPause.Left + cmdPause.Width
            cmdForwards.Left = cmdStop.Left + cmdStop.Width
            fraVolume.Left = cmdForwards.Left + cmdForwards.Width + GAP
            lblVolume.Left = GAP
            sliVolume.Left = lblVolume.Left + lblVolume.Width
            fraVolume.Width = lblVolume.Width + GAP + sliVolume.Width
            lblTrack.Left = GAP
            sliTrack.Left = lblTrack.Left + lblTrack.Width + GAP
            sliTrack.Width = Me.ScaleWidth - sliTrack.Left
        End If
        If Me.ScaleHeight > 2000 Then
            wmpPlayer.Top = Me.Height + GAP
            lblPodcast.Top = GAP
            lblItems.Top = GAP
            lstPodcasts.Top = lblPodcast.Top + lblPodcast.Height + GAP
            lstItems.Height = Me.ScaleHeight - cmdPlay.Height - lstItems.Top - staMain.Height - sliTrack.Height - GAP
            lstItems.Top = lblPodcast.Top + lblPodcast.Height + GAP
            lstPodcasts.Height = lstItems.Height
            cmdPlay.Top = lstItems.Height + lstItems.Top + GAP
            cmdPause.Top = cmdPlay.Top
            cmdStop.Top = cmdPlay.Top
            cmdForwards.Top = cmdPlay.Top
            cmdBackwards.Top = cmdPlay.Top
            fraVolume.Top = cmdPlay.Top
            lblTrack.Top = cmdPlay.Top + cmdPlay.Height + GAP
            sliTrack.Top = lblTrack.Top
        End If
    End If
    mWebBrowser.Left = -mWebBrowser.Width - 100
End Sub

Private Sub SavePodcasts()
    On Error Resume Next
    Dim deleted As DOMDocument30
    Dim newDeleted As DOMDocument30
    Dim fso As New Scripting.FileSystemObject
    Set deleted = New DOMDocument30
    Set newDeleted = New DOMDocument30
    deleted.async = False
    newDeleted.async = False
    Call newDeleted.loadXML("<deleted/>")
    If fso.FileExists(modPath.settingsPath & "\deletedItems.xml") Then
        Call deleted.Load(modPath.settingsPath & "\deletedItems.xml")
    Else
        Call deleted.loadXML("<deleted />")
    End If
    Set fso = Nothing
    'remove the podcasts by hand from the first to the last. This preserves
    'their order in the xml.
    While podcasts.Count > 0
        Call podcasts.Remove(1)
    Wend
    'save the feeds and options
    Call modPath.SaveSettingIni(App.title, "Feeds", "UseBrokenFeeds", CStr(mnuOptionsUsebrokenfeeds.Checked))
    Call modPath.SaveSettingIni(App.title, "Privacy", "AllowAnonymousSubmissions", CStr(mnuOptionsAllowsubmissions.Checked))
    'save the deleted file
    Set deleted = Nothing
    Call newDeleted.save(modPath.settingsPath & "\deletePodcasts.xml")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim pc As CPodcast
    Dim podcastNode As IXMLDOMNode
    Dim itemIterator As CItem
    Dim sectionName As String
    Dim keyName As String
    Dim s As String
    Dim filename As String
    Dim f As Form
    
    tmrUpdateAvailable.Enabled = False
    tmrUpdateTrackTimer.Enabled = False
    If mWebBrowser.busy Then Call mWebBrowser.stop
    Set mCurrentFeed = Nothing
    
    'save layout
    Call modRememberPosition.SavePosition(Me)
    
    Call SavePodcasts
    'now save the volume
    'Call SaveSetting("AccessiblePodcaster", "User settings", "Volume", wmpPlayer.settings.volume)
    keyName = "Volume" & Chr(0)
    s = CStr(wmpPlayer.settings.volume) & Chr(0)
    sectionName = "User Settings" & Chr(0)
    filename = modPath.settingsPath & "\podcatcher.ini" & Chr(0)
    Call modShared.WritePrivateProfileString(sectionName, keyName, s, filename)
    For Each f In Forms
        If f.name <> Me.name Then
            Call Unload(f)
        End If
    Next f
    'clear all settings from registry
    'Prevent error by writing dummy value first.
    Call SaveSetting("AccessiblePodcaster", "Dummy", "Dummy", "Dummy value to prevent error condition. Feel free to delete, won't affect anything, and in fact shouldn't be here. Alasdair.")
    Call DeleteSetting("AccessiblePodcaster")
End Sub

Private Sub lstItems_Click()
    On Error Resume Next
    If lstItems.ListIndex > -1 And podcasts(lstPodcasts.ListIndex + 1).items(mnuItemsShowdeleteditems.Checked).Count > 0 And (lstItems.List(0) <> gstrNotAvailable) Then
        cmdPlay.Enabled = True
        mnuFileSave.Enabled = True
        mnuItemsCopyitemurltoclipboard.Enabled = True
        mnuItemsOpeniteminwindowsmediaplayer.Enabled = True
    Else
        cmdPlay.Enabled = False
        mnuFileSave.Enabled = False
        mnuItemsCopyitemurltoclipboard.Enabled = False
        mnuItemsOpeniteminwindowsmediaplayer.Enabled = False
    End If
    'cmdPlay.Enabled = (lstItems.ListIndex > -1)
    mnuAudioPlay.Enabled = cmdPlay.Enabled
End Sub

Private Sub lstItems_DblClick()
    On Error Resume Next
    Call lstItems_KeyPress(vbKeyReturn)
End Sub

Private Sub lstItems_GotFocus()
    On Error Resume Next
    If lstItems.ListIndex = -1 And lstItems.ListCount > 0 Then lstItems.ListIndex = 0
End Sub

Private Sub lstItems_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim pc As CPodcast
    Dim pi As CItem
    Dim entry As String
    Dim index As Long
    
    Debug.Print "lstItems_KeyDown"
    If KeyCode = vbKeyDelete Then
        If lstItems.ListCount > 0 And lstItems.ListIndex > -1 Then
            entry = lstItems.Text
            If (entry <> gstrLoadingPodcast) And (entry <> gstrNotAvailable) And (entry <> gstrNotLoadedYet) Then
                'got an item to delete (hide)
                KeyCode = 0
                Set pc = podcasts.Item(mCurrentPodcast)
                Set pi = pc.items.Item(lstItems.ListIndex + 1)
                Call pc.HideItem(pi)
                index = lstItems.ListIndex
                Call DisplayItems
                If index > lstItems.ListCount - 1 Then
                    lstItems.ListIndex = lstItems.ListCount - 1
                Else
                    lstItems.ListIndex = index
                End If
            End If
        End If
    End If
End Sub

Private Sub lstItems_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        If lstItems.Text = gstrNotLoadedYet Then
            Call lstPodcasts_KeyPress(vbKeyReturn)
        ElseIf podcasts.Item(mCurrentPodcast).items(mnuItemsShowdeleteditems.Checked).Count > 0 Then
            'Debug.Print "Return"
            Call cmdPlay_Click
            Me.Caption = "Podcatcher - " & lstPodcasts.List(lstPodcasts.ListIndex)
        Else
            Call Beep
        End If
    ElseIf KeyAscii = vbKeyEscape Then
        Call cmdStop_Click
        Me.Caption = "Podcatcher"
        Call lstPodcasts.SetFocus
    End If
End Sub

Private Sub lstItems_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If lstItems.ListIndex = -1 And lstItems.ListCount > 0 Then lstItems.ListIndex = 0
End Sub

Private Sub lstPodcasts_Click()
    On Error Resume Next
    If lstPodcasts.ListIndex + 1 <> mCurrentPodcast Then
        Call lstItems.Clear
        Call lstItems.AddItem(gstrNotLoadedYet)
    End If
    mnuPodcastsDelete.Enabled = (lstPodcasts.ListIndex > -1)
    If lstPodcasts.List(0) = gstrNoPodcasts Then mnuPodcastsDelete.Enabled = False
End Sub

Private Sub lstPodcasts_DblClick()
    On Error Resume Next
    Call lstPodcasts_KeyPress(vbKeyReturn)
End Sub

Private Sub lstPodcasts_GotFocus()
    On Error Resume Next
    If lstPodcasts.ListIndex = -1 Then lstPodcasts.ListIndex = 0
End Sub

Private Sub lstPodcasts_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Dim oldIndex As Long
    
    If KeyAscii = vbKeyReturn Then
        If lstPodcasts.ListIndex > -1 And podcasts.Count > 0 Then
            oldIndex = lstItems.ListIndex
            Me.MousePointer = vbHourglass
            Call lstItems.Clear
            Call lstItems.AddItem(gstrLoadingPodcast)
            Call Me.Refresh
            lstItems.ListIndex = 0
            Call lstItems.SetFocus
            mnuPodcastsDelete.Enabled = True
            mnuPodcastsEdit.Enabled = True
            Call lstItems.SetFocus
            Call GetPodcast
            If lstItems.ListCount > 0 Then
                lstItems.ListIndex = 0
            End If
            Me.MousePointer = vbNormal
        End If
        'Debug.Print "At:" & Me.ActiveControl.name & " index:" & lstItems.ListIndex
    End If
End Sub

Private Sub lstPodcasts_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyDelete Then
        If lstPodcasts.ListIndex > -1 And podcasts.Count > 0 Then
            Call DeletePodcast
        End If
    ElseIf KeyCode = vbKeyF2 Then
        Call mnuPodcastsRenamepodcast_Click
    End If
End Sub

Private Sub lstPodcasts_LostFocus()
    On Error Resume Next
    'Don't think we want to do this: it means changing the UI when the user presses TAB, which is pretty non-standard
    'and stops people tabbing around the interface. 1.9.3. 19 Jan 2009.
'    If Me.ActiveControl.name = "lstItems" Then
'        Call lstPodcasts_KeyPress(vbKeyReturn)
'    End If
End Sub

Private Sub mCurrentFeed_onreadystatechange()
    On Error Resume Next
    Dim result As Long
    
    Debug.Print "mCurrentFeed_OnReadyStateChange"
    If mCurrentFeed.parsed Then
        'check whether this is a podcast or an OPML list
        staMain.Panels("status").Text = modI18N.GetText("Processing podcast")
        'Call parsedResult.loadXML(xmlObject.responseText)
        If mCurrentFeed.parseError = 0 Then
            'MsgBox "Got " & mCurrentFeed.documentElement.nodeName
            'okay, this parses okay: what is it?
            If mCurrentFeed.documentElement.nodeName = "rss" Then
                'okay, this is a podcast list
                'assume this is okay to add
                podcasts(lstPodcasts.ListIndex + 1).xml = mCurrentFeed.xml
                'add successful podcast
                mCurrentPodcast = lstPodcasts.ListIndex + 1
                Call DisplayItems
            ElseIf mCurrentFeed.documentElement.nodeName = "opml" Then
                'okay, this is an opml file
                result = MsgBox(modI18N.GetText("This is a list of podcasts. Do you want to add them to your list?"), vbYesNoCancel, modI18N.GetText("New podcasts"))
                If result = vbYes Then
                    staMain.Panels("status").Text = modI18N.GetText("Importing")
                    Call ImportOPML(mCurrentFeed)
                    Call Display
                End If
'            ElseIf parsedResult.documentElement.nodeName = "rdf:RDF" Or parsedResult.documentElement.nodeName = "rdf" Then
'                MsgBox "RDF!"
            Else
                'MsgBox "Unknown podcast type: " & Left(parsedResult.xml, 500)
                Call MsgBox(modI18N.GetText("Sorry, this is not a supported type of podcast. I could not do anything with it."), vbOKOnly, modI18N.GetText("Podcast unknown"))
                'Debug.Print "Unknown type: " & Left(parsedResult.xml, 500)
                Debug.Print "NN:" & mCurrentFeed.documentElement.nodeName
            End If
        Else
            'OK, for some reason we've failed to get the podcast. There
            'are many possible reasons - it's broken today because of character coding,
            'the server is down, we're not connected to the internet, it's gone away)
            'and it might come back, so don't hassle the user with a popup message but
            'just show the problem.
            podcasts(lstPodcasts.ListIndex + 1).parseError = True
            mCurrentPodcast = lstPodcasts.ListIndex + 1
            Call DisplayItems
            
'            Debug.Print "Failed to parse! " & mCurrentFeed.parseError.errorCode & " " & mCurrentFeed.parseError.reason
            'So don't do this any more.
'''            'failed to parse: ask to remove
'''            result = MsgBox(modi18n.GetText("Sorry, I could not use this podcast. Do you want to remove it from your list of podcasts?"), vbYesNoCancel, modi18n.GetText("Broken podcast"))
'''            If result = vbYes Then
'''                podcasts(lstPodcasts.ListIndex + 1).markedForDeath = True
'''                Call podcasts.Remove(lstPodcasts.ListIndex + 1)
'''                Call Display
'''            End If
        End If
        staMain.Panels("status").Text = modI18N.GetText("Done")
    End If
End Sub

Private Sub mnuAudioDecreasevolume_Click()
    On Error Resume Next
    If wmpPlayer.settings.volume > 0 Then
        wmpPlayer.settings.volume = wmpPlayer.settings.volume - 1
    Else
        'Call playerrorsound
    End If
    sliVolume.value = wmpPlayer.settings.volume
    If wmpPlayer.settings.volume = 0 Then mnuAudioDecreasevolume.Enabled = False
    mnuAudioIncreasevolume.Enabled = True
End Sub

Private Sub mnuAudioIncreasevolume_Click()
    On Error Resume Next
    If wmpPlayer.settings.volume < 100 Then
        wmpPlayer.settings.volume = wmpPlayer.settings.volume + 1
    End If
    sliVolume.value = wmpPlayer.settings.volume
    If wmpPlayer.settings.volume = 100 Then mnuAudioIncreasevolume.Enabled = False
    mnuAudioDecreasevolume.Enabled = True
End Sub

Private Sub mnuAudioPause_Click()
    On Error Resume Next
    Call wmpPlayer.Controls.pause
    mnuAudioPause.Enabled = False
    cmdPlay.Enabled = True
    cmdPause.Enabled = False
    mnuAudioPlay.Enabled = True
End Sub

Private Sub mnuAudioPlay_Click()
    On Error Resume Next
    Call cmdPlay_Click
End Sub

Private Sub mnuAudioSkipbackwards_Click()
    On Error Resume Next
    Call cmdBackwards_Click
End Sub

Private Sub mnuAudioSkipforwards_Click()
    On Error Resume Next
    Call cmdForwards_Click
End Sub

Private Sub mnuAudioStop_Click()
    On Error Resume Next
    Call cmdStop_Click
End Sub

Private Sub mnuFileExit_Click()
    On Error Resume Next
    Call Unload(Me)
End Sub

'save the current playlist as an OPML file
Private Sub mnuFileExport_Click()
    On Error Resume Next
    Dim opmlFile As New DOMDocument30
    Dim fso As New FileSystemObject
    Dim result As Long
    Dim abort As Boolean
    Dim newXMLNode As IXMLDOMNode
    Dim newXMLAttribute As IXMLDOMAttribute
    Dim podcastIterator As CPodcast
    
    cdlg.CancelError = True
    cdlg.Flags = cdlOFNPathMustExist
    cdlg.DefaultExt = "opml"
    cdlg.DialogTitle = modI18N.GetText("Export OPML podcast list")
    cdlg.Filter = modI18N.GetText("OPML files (*.opml)|*.opml;")
    On Error GoTo CancelError:
    Call cdlg.ShowSave
    On Error Resume Next
    If Len(cdlg.filename) > 0 Then
        If fso.FileExists(cdlg.filename) Then
            result = MsgBox(modI18N.GetText("File already exists. Do you want to replace it?"), vbYesNoCancel, modI18N.GetText("File exists"))
            If result <> vbYes Then
                abort = True
            End If
        End If
        If Not abort Then
            'okay, no file of this name: go ahead and save
            Call opmlFile.loadXML("<opml version=""1.1"" />")
            'add head
            Set newXMLNode = opmlFile.createNode(NODE_ELEMENT, "head", Empty)
            Call opmlFile.documentElement.appendChild(newXMLNode)
            'add title
            Set newXMLNode = opmlFile.createNode(NODE_ELEMENT, "title", Empty)
            newXMLNode.Text = InputBox(modI18N.GetText("Enter a name for your saved list of podcasts:"), modI18N.GetText("List name"), Empty)
            Call opmlFile.documentElement.selectSingleNode("head").appendChild(newXMLNode)
            'add body
            Set newXMLNode = opmlFile.createNode(NODE_ELEMENT, "body", Empty)
            Call opmlFile.documentElement.appendChild(newXMLNode)
            'add all the podcasts
            For Each podcastIterator In podcasts
                Set newXMLNode = opmlFile.createNode(NODE_ELEMENT, "outline", Empty)
                'type
                Set newXMLAttribute = opmlFile.createAttribute("type")
                newXMLAttribute.Text = "link"
                Call newXMLNode.Attributes.setNamedItem(newXMLAttribute)
                'text
                Set newXMLAttribute = opmlFile.createAttribute("text")
                newXMLAttribute.Text = podcastIterator.name
                Call newXMLNode.Attributes.setNamedItem(newXMLAttribute)
                'url
                Set newXMLAttribute = opmlFile.createAttribute("url")
                newXMLAttribute.Text = podcastIterator.url
                Call newXMLNode.Attributes.setNamedItem(newXMLAttribute)
                'whether subscribed
                Set newXMLAttribute = opmlFile.createAttribute("subscribed")
                newXMLAttribute.Text = podcastIterator.subscribed
                Call newXMLNode.Attributes.setNamedItem(newXMLAttribute)
                'add to body
                Call opmlFile.documentElement.selectSingleNode("body").appendChild(newXMLNode)
            Next podcastIterator
            'now save to disk
            Call opmlFile.save(cdlg.filename)
        End If
    End If
    Exit Sub
CancelError:
    'user hit cancel in save dialog
End Sub

'opens an OPML and adds contents to the playlist
Private Sub mnuFileImport_Click()
    On Error Resume Next
    Dim fso As New FileSystemObject
    Dim opmlFile As New DOMDocument30
    
    cdlg.CancelError = True
    cdlg.DefaultExt = "opml"
    cdlg.Flags = cdlOFNFileMustExist
    cdlg.DialogTitle = modI18N.GetText("Import OPML podcast list")
    cdlg.filename = "*.opml"
    cdlg.Filter = "OPML files (*.opml)|*.opml;"
    On Error GoTo CancelError:
    Call cdlg.ShowOpen
    On Error Resume Next
    If Len(cdlg.filename) > 0 Then
        'have a real filename: check it's a real file
        If fso.FileExists(cdlg.filename) Then
            'okay, let's try to import it
            opmlFile.async = False
            opmlFile.Load (cdlg.filename)
            If opmlFile.parseError = 0 Then
                'loaded XML okay
                Call ImportOPML(opmlFile)
                Call Display
            Else
                'failed to load XML: don't do anything
                Call MsgBox(modI18N.GetText("Sorry, Accessible Podcatcher could not read this OPML file."), vbOKOnly And vbInformation, modI18N.GetText("Failed import"))
            End If
        End If
    End If
    Exit Sub
CancelError:
    'user hit cancel in open dialog
End Sub

Private Sub mnuFileSave_Click()
    On Error Resume Next
    Dim mp3File As String
    Dim fso As New FileSystemObject
    Dim result As Long
    
    If podcasts.Count > 0 Then
        If lstItems.ListIndex > -1 And podcasts.Item(mCurrentPodcast).items.Count > 0 Then
            mp3File = podcasts.Item(mCurrentPodcast).items.Item(lstItems.ListIndex + 1).url
            'now get filename
            cdlg.DefaultExt = Right(mp3File, Len(mp3File) - InStrRev(mp3File, "."))
            cdlg.filename = Right(mp3File, Len(mp3File) - InStrRev(mp3File, "/"))
            cdlg.DialogTitle = modI18N.GetText("Save audio file")
            cdlg.Filter = modI18N.GetText("All files") & " (*.*)|*.*"
            cdlg.Flags = cdlOFNPathMustExist
            cdlg.CancelError = True
            On Error GoTo CancelError:
            Call cdlg.ShowSave
            On Error Resume Next
            If Len(cdlg.filename) > 0 Then
                'check for existing file and prompt to overwrite
                If fso.FileExists(cdlg.filename) Then
                    result = MsgBox(modI18N.GetText("File already exists. Do you want to replace it?"), vbYesNoCancel, "File exists")
                Else
                    result = vbYes
                End If
                If result = vbYes Then
                    'okay, download!
                    staMain.Panels("status").Text = modI18N.GetText("Downloading audio file")
                    Call URLDownloadToFile(0, mp3File, cdlg.filename, 0, 0)
                    staMain.Panels("status").Text = modI18N.GetText("Done")
                End If
            End If
        End If
    End If
    Exit Sub
CancelError:
    Exit Sub
End Sub

Private Sub mnuHelpAbout_Click()
    On Error Resume Next
    Call MsgBox("Accessible Podcatcher " & App.Major & "." & App.Minor & "." & App.Revision & vbNewLine & "WebbIE Package Version" & vbTab & modVersion.GetPackageVersion & vbNewLine & "Alasdair King, http://www.webbie.org.uk", vbInformation, "Accessible Podcatcher")
End Sub

Private Sub mnuHelpManual_Click()
    On Error Resume Next
    Call Load(frmPodcaster)
    Call frmHelp.Show(vbModeless, Me)
End Sub

Private Sub mnuItemsCopyitemurltoclipboard_Click()
    On Error Resume Next
    Dim pc As CPodcast
    Dim pi As CItem
    
    Set pc = podcasts.Item(mCurrentPodcast)
    If pc Is Nothing Then
        Call Beep
    Else
        Set pi = pc.items.Item(lstItems.ListIndex + 1)
        If pi Is Nothing Then
            Call Beep
        Else
            Call Clipboard.Clear
            Call Clipboard.SetText(pi.url)
        End If
    End If
End Sub

Private Sub mnuItemsDeleteitem_Click()
    On Error Resume Next
    Call lstItems_KeyDown(vbKeyDelete, 0)
End Sub

Private Sub mnuItemsOpeniteminwindowsmediaplayer_Click()
    On Error Resume Next
    Dim pc As CPodcast
    Dim pi As CItem
    
    Set pc = podcasts.Item(mCurrentPodcast)
    If pc Is Nothing Then
        Call Beep
    Else
        Set pi = pc.items.Item(lstItems.ListIndex + 1)
        If pi Is Nothing Then
            Call Beep
        Else
            'Try using COM
'            Dim objPlayer As Object
'            Set objPlayer = CreateObject("WMPlayer.OCX")
'            objPlayer.url = pi.url
'            Call objPlayer.Controls.play
            'Try using Shell
            'Call Shell("wmplayer.exe " & pi.url, vbNormalFocus)
            Dim Path As String
            Path = ReadRegistryEntryString(HKEY_CLASSES_ROOT, "Applications\wmplayer.exe\shell\play\command", "")
            If Path = "" Then
                Path = """" & modPath.GetSpecialFolderPath(modPath.CSIDL_PROGRAM_FILESX86) & "\Windows Media Player\wmplayer.exe"" /Play ""%L"""
            Else
                Path = Replace(Path, "%ProgramFiles(x86)%", modPath.GetSpecialFolderPath(modPath.CSIDL_PROGRAM_FILESX86))
                Path = Replace(Path, "%ProgramFiles%", modPath.GetSpecialFolderPath(modPath.CSIDL_PROGRAM_FILES))
            End If
            Path = Replace(Path, "%L", pi.url)
            MsgBox Path
            Call Shell(Path, vbNormalFocus)
            'Try using
        End If
    End If
End Sub

Private Sub mnuItemsShowdeleteditems_Click()
    On Error Resume Next
    mnuItemsShowdeleteditems.Checked = Not mnuItemsShowdeleteditems.Checked
    If lstItems.List(0) = gstrNotLoadedYet Then
    Else
        Call lstPodcasts_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub mnuOptionsAllowsubmissions_Click()
    On Error Resume Next
    mnuOptionsAllowsubmissions.Checked = Not mnuOptionsAllowsubmissions.Checked
End Sub

Private Sub mnuOptionsUsebrokenfeeds_Click()
    On Error Resume Next
    mnuOptionsUsebrokenfeeds.Checked = Not mnuOptionsUsebrokenfeeds.Checked
End Sub

Private Sub mnuPodcastsAddpodcasts_Click()
    On Error Resume Next
    Dim newURL As String
    Dim newName As String
    Dim newCast As CPodcast
    Dim doc As DOMDocument30
    
    newURL = InputBox(modI18N.GetText("Enter the web address or URL for the new podcast:"), modI18N.GetText("New Podcast"), "")
    If Len(newURL) > 0 Then
        If Len(newURL) < 7 Then
            'Not enough text for a podcast!
            Call Beep
        Else
            Set doc = New DOMDocument30
            doc.async = False
            Call doc.Load(newURL)
            If doc.parseError.errorCode = 0 Then
                newName = doc.documentElement.selectSingleNode("channel").selectSingleNode("title").Text
            End If
            newName = InputBox(modI18N.GetText("Enter the name of the new podcast:"), modI18N.GetText("New Podcast"), newName)
            
            If Len(newName) > 0 Then
                'okay, we have a new url and name
                Set newCast = New CPodcast
                'Check for leading http, if missing then add.
                If LCase(Left(newURL, 7)) <> "http://" And LCase(Left(newURL, 8)) <> "https://" And LCase(Left(newURL, 7)) <> "file://" Then
                    newURL = "http://" & newURL
                End If
                newCast.url = newURL
                newCast.name = newName
                Call podcasts.Add(newCast)
                Call Display
                lstPodcasts.ListIndex = lstPodcasts.ListCount - 1
                Call lstPodcasts.SetFocus
                Call UpdateDirectory(newURL, newName)
            End If
        End If
    End If
End Sub

Public Sub UpdateDirectory(url As String, title As String)
'Adds the new podcast to the WebbIE directory of podcasts.
    On Error Resume Next
    Dim newName As String
    Dim newURL As String
    Dim i As Long
    If mnuOptionsAllowsubmissions.Checked Then
         Dim webBrowser As Object ' InternetExplorer caused problems
         Set webBrowser = CreateObject("InternetExplorer.Application", "")
         newName = title
         newName = Replace(newName, " ", "%20")
         newName = Replace(newName, "/", "%2F")
         newName = Replace(newName, "\", "%5C")
         newName = Replace(newName, """", "%22")
         newName = Replace(newName, "'", "%27")
         newURL = url
         newURL = Replace(newURL, " ", "%20")
         Call webBrowser.Navigate2("http://data.webbie.org.uk/newPodcast.php?title=" & newName & "&url=" & newURL & "&language=" & modI18N.GetLanguage)
         While webBrowser.readyState <> READYSTATE_COMPLETE
            DoEvents
         Wend
         Call webBrowser.Quit
     End If
End Sub
 

Private Sub mnuPodcastsDelete_Click()
    On Error Resume Next
    Call DeletePodcast
End Sub

Private Sub mnuPodcastsDownloadpodcasts_Click()
    On Error Resume Next
    Call SavePodcasts
    Call Shell(App.Path & "\PodcastDownloader.exe")
    Call Unload(Me)
End Sub

Private Sub mnuPodcastsEdit_Click()
    On Error Resume Next
    Dim userInput As String
    Dim selected As Integer
    
    If lstPodcasts.ListIndex > -1 And podcasts.Count > 0 Then
        'okay, show this podcast
        userInput = InputBox(modI18N.GetText("Enter the new name of the podcast:"), modI18N.GetText("Edit podcast"), podcasts(lstPodcasts.ListIndex + 1).name)
        If Len(userInput) > 0 Then
            podcasts(lstPodcasts.ListIndex + 1).name = userInput
        End If
        userInput = InputBox(modI18N.GetText("Enter new podcast address or URL:"), modI18N.GetText("Edit podcast"), podcasts(lstPodcasts.ListIndex + 1).url)
        If Len(userInput) > 0 Then
            podcasts(lstPodcasts.ListIndex + 1).url = userInput
        End If
        selected = lstPodcasts.ListIndex
        Call Display
        lstPodcasts.ListIndex = selected
    End If
End Sub

Private Sub mnuPodcastsGetnewpodcasts_Click()
    On Error Resume Next
    Call Load(frmPodcastList)
    Call frmPodcastList.Show(vbModal, Me)
End Sub

Private Sub mnuPodcastsRenamepodcast_Click()
    On Error Resume Next
    Dim userInput As String
    Dim selected As Integer
    
    If lstPodcasts.ListIndex > -1 And podcasts.Count > 0 Then
        'okay, show this podcast
        userInput = InputBox(modI18N.GetText("Enter the new name of the podcast:"), modI18N.GetText("Edit podcast"), podcasts(lstPodcasts.ListIndex + 1).name)
        If Len(userInput) > 0 Then
            podcasts(lstPodcasts.ListIndex + 1).name = userInput
        End If
        selected = lstPodcasts.ListIndex
        Call Display
        lstPodcasts.ListIndex = selected
    End If
End Sub

Private Sub mnuPodcastsSearchwebpage_Click()
    'Find all podcasts on a page!
    On Error Resume Next
    Call frmSearch.Show(vbModal, Me)
End Sub

'sorts the podcasts into alphabetical order
Private Sub mnuPodcastsSort_Click()
    On Error Resume Next
    Dim result As Long
    Dim swapped As Boolean
    Dim index As Integer
    Dim swapTemp As CPodcast
    Dim fso As Scripting.FileSystemObject
    
    result = MsgBox(modI18N.GetText("All podcasts will be irreversibly sorted into alphabetical order. Are you sure you want to do this?"), vbYesNoCancel, modI18N.GetText("Confirm sort"))
    If result = vbYes Then
        'okay, now sort every podcast: use bubblesort
        staMain.Panels("status").Text = modI18N.GetText("Sorting")
        swapped = True
        While swapped
            swapped = False
            For index = 1 To podcasts.Count - 1
                If StrComp(podcasts(index).name, podcasts(index + 1).name, vbTextCompare) > 0 Then
                    'need to swap!
                    swapped = True
                    Set swapTemp = podcasts(index)
                    Call podcasts.Remove(index)
                    Call podcasts.Add(swapTemp, , , index)
                End If
            Next index
        Wend
'        'now delete the xml file so it gets stored correctly later
'        Set fso = New Scripting.FileSystemObject
'        Call fso.DeleteFile(modPath.settingsPath & "\podcasts.xml")
'        Set fso = Nothing
        staMain.Panels("status").Text = modI18N.GetText("Done")
        Call Display
    End If
End Sub


Private Sub mnuVideoFullscreen_Click()
    On Error Resume Next
    wmpPlayer.fullScreen = Not wmpPlayer.fullScreen
End Sub

Private Sub mnuVideoShowvideo_Click()
    On Error Resume Next
    If mnuVideoShowvideo.Checked Then
        mnuVideoShowvideo.Checked = False
        wmpPlayer.fullScreen = False
        wmpPlayer.Top = Me.ScaleHeight + 100
        wmpPlayer.Left = Me.ScaleWidth + 100
        Call wmpPlayer.ZOrder(vbSendToBack)
        lblPodcast.visible = True
        lblItems.visible = True
        lstPodcasts.visible = True
        lstItems.visible = True
    Else
        mnuVideoShowvideo.Checked = True
        wmpPlayer.Top = 0
        wmpPlayer.Left = 0
        wmpPlayer.Width = Me.ScaleWidth
        wmpPlayer.Height = lstPodcasts.Height + lstPodcasts.Top
        Call wmpPlayer.ZOrder(vbBringToFront)
        lblPodcast.visible = False
        lblItems.visible = False
        lstPodcasts.visible = False
        lstItems.visible = False
    End If
End Sub

Private Sub mWebBrowser_GotFocus()
    On Error Resume Next
    Call lstItems.SetFocus
End Sub

Private Sub mWebBrowser_NewWindow2(ppDisp As Object, Cancel As Boolean)
    On Error Resume Next
    'Prevent popups
    Cancel = True
End Sub

Private Sub sliTrack_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    mManuallyChangingTrackSlider = True
    mTrackValueBeforeChanging = sliTrack.value
End Sub

Private Sub sliTrack_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    mManuallyChangingTrackSlider = False
    If mTrackValueBeforeChanging <> sliTrack.value Then
        wmpPlayer.Controls.currentPosition = sliTrack.value
    End If
End Sub

Private Sub sliTrack_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    mManuallyChangingTrackSlider = True
    mTrackValueBeforeChanging = sliTrack.value
End Sub

Private Sub sliTrack_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    mManuallyChangingTrackSlider = False
    If mTrackValueBeforeChanging <> sliTrack.value Then
        wmpPlayer.Controls.currentPosition = sliTrack.value
    End If
End Sub

Private Sub sliVolume_Change()
    On Error Resume Next
    wmpPlayer.settings.volume = sliVolume.value
End Sub

Private Sub sliVolume_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    wmpPlayer.settings.volume = sliVolume.value
End Sub



Private Sub tmrUpdateAvailable_Timer()
    On Error Resume Next
    Static secondsSinceLastCheck As Long
    
    secondsSinceLastCheck = secondsSinceLastCheck + 1
    If secondsSinceLastCheck = CHECK_FOR_UPDATES_TIME Then
        secondsSinceLastCheck = 0
        Call GetPodcast
    End If
End Sub

Private Sub tmrUpdateTrackTimer_Timer()
    On Error Resume Next
    Dim max As Long
    
    If wmpPlayer.currentMedia Is Nothing Then
    Else
        If mManuallyChangingTrackSlider Then
            'nope, we're changing the track by hand: don't update
        Else
            'DEV: make slider relative to whole track, not what is
            'available. Simpler.
            max = wmpPlayer.currentMedia.duration '* wmpPlayer.network.downloadProgress / 100 - 1
            sliTrack.max = IIf(max > 10, max, 10)
            sliTrack.Min = 0
            sliTrack.TickFrequency = sliTrack.max / 20
            sliTrack.value = wmpPlayer.Controls.currentPosition
            sliTrack.ToolTipText = wmpPlayer.Controls.currentPositionString
            staMain.Panels("progress").Text = "Time: " & sliTrack.ToolTipText & " of " & wmpPlayer.currentMedia.durationString & " "
            staMain.Panels("downloaded").Text = "Downloaded: " & wmpPlayer.network.downloadProgress & "%"
        End If
    End If
End Sub

Private Sub wmpPlayer_Buffering(ByVal start As Boolean)
    On Error Resume Next
    If Not start Then
        'okay, we've finished buffering: restart playing if we are playing
        'Check whether we're playing by examining the cmdPlay button
        If Not cmdPlay.Enabled Then
            'cmdPlay cannot be clicked, therefore it has been, therefore
            'we are playing.
            'What's the wmpPlayer state? Because just calling .play causes some nasty
            'effects on .MOV files - lots of distortion. So this is getting called all the time
            'because buffering is being called lots.
            If wmpPlayer.playState = wmppsPlaying Then
                'Nope, already playing.
            ElseIf wmpPlayer.playState = wmppsBuffering Then
                'Still buffering
            Else
                Call wmpPlayer.Controls.play
            End If
        End If
    End If
End Sub

Private Sub wmpPlayer_Error()
    'If we get this, can we detect a URL we don't support and shell the media player than handles this?
    On Error Resume Next
    Dim i As Long
    
    For i = 1 To wmpPlayer.Error.errorCount
        Debug.Print "Warning: Windows Media Player error: " & wmpPlayer.Error.Item(i).errorCode & " " & wmpPlayer.Error.Item(i).errorDescription
    Next i
    If Len(wmpPlayer.url) > 0 Then
        If LCase(Right(wmpPlayer.url, 4)) = ".wma" Or LCase(Right(wmpPlayer.url, 4)) = ".mp3" Then
            'It's a WMA or MP3 file - should play just fine.
            
        'DEV: Okay, so the logic here was that the non-compliant files would trigger an error - like
        '"Codec missing" or whatever - and we should then shell them off to the default error handler.
        
'        ElseIf LCase(Right(wmpPlayer.url, 4)) = ".aac" Or LCase(Right(wmpPlayer.url, 4)) = ".mov" Or LCase(Right(wmpPlayer.url, 5)) = ".aiff" Then
'            'It's an Apple Corporation filetype - AAC, movie or AIFF file. Try shelling default application - can't play in WMP.
'            Call ShellExecute(0, "open", wmpPlayer.url, "", "", SW_SHOW)
        End If
    End If
End Sub

Private Sub wmpPlayer_StatusChange()
    On Error Resume Next
    staMain.Panels("status").Text = wmpPlayer.Status
End Sub

'adds an opml file to the list of podcasts
Private Sub ImportOPML(opmlFile As DOMDocument30)
    On Error Resume Next
    Dim podcastIterator As IXMLDOMNode
    Dim newPodcast As CPodcast
    Dim addPoint As Integer
    
    'iterate through OPML adding podcasts
    'It would be great if the outline node had defined attributes that could
    'be used to identify podcasts or RSS, like "type" - but they don't, and
    'they aren't always used consistently. So I'm assuming that anything
    'with a url or xmlUrl and a text attribute is a useful feed.
    For Each podcastIterator In opmlFile.documentElement.selectNodes("//outline[@text and (@url or @xmlUrl)]")
        Set newPodcast = New CPodcast
        newPodcast.url = podcastIterator.Attributes.getNamedItem("url").Text
        If Len(newPodcast.url) = 0 Then
            newPodcast.url = podcastIterator.Attributes.getNamedItem("xmlUrl").Text
        End If
        newPodcast.name = podcastIterator.Attributes.getNamedItem("text").Text
        Call podcasts.Add(newPodcast)
    Next podcastIterator
End Sub

Private Function ConvertToMinutesSeconds(timeInSec As Double) As String
    On Error Resume Next
    Dim sec As Integer
    ConvertToMinutesSeconds = CStr(CInt(timeInSec / 60))
    sec = CInt(timeInSec Mod 60)
    If sec < 10 Then
        ConvertToMinutesSeconds = ConvertToMinutesSeconds & ":0" & sec
    Else
        ConvertToMinutesSeconds = ConvertToMinutesSeconds & ":" & sec
    End If
End Function

Private Sub LoadFeed(url As String)
    Dim feedURL As String
    Dim result As String
    Dim xmlObject As New MSXML2.XMLHTTP30
    Dim endTag As Long
    Dim startTag As Long
    
    Debug.Print "Feed: " & url
retryXMLObject:
    Call xmlObject.open("GET", url, False)
   'Debug.Print "URL:[" & GetSelectedFeed & "]"
    'Call xmlObject.open("GET", "http://personalpages.umist.ac.uk/staff/alasdair.king/test.txt", False)
    Call xmlObject.setRequestHeader("Pragma", "no-cache") ' make sure it is fresh
    Call xmlObject.setRequestHeader("cache-control", "no-cache") ' make sure it is fresh
    Call xmlObject.setRequestHeader("If-Modified-Since", "Wed, 31 Dec 1980 00:00:00 GMT")
    'MsgBox "go2"
    On Error GoTo networkFailure
    Call xmlObject.send
    On Error Resume Next
    'Debug.Print xmlObject.StatusText
    'Debug.Print "RH:" & xmlObject.getResponseHeader("cache-control:")
    'Debug.Print "Content:" & xmlObject.responseText
    'MsgBox "go3"
    If mnuOptionsUsebrokenfeeds.Checked Then
        'OK, so we try to do some pre-work parsing on the XML feed.
        result = xmlObject.responseText
        If Err.Number = -1072896658 Then
            'Failed because the encoding of the page is wrong. See http://groups.google.com/group/Google-Desktop-Developer/browse_thread/thread/cf17d8a4465cba76
            'Need to handle, next version!
        End If
        While InStr(1, result, "<Content:Encoded>", vbTextCompare) > 0 And InStr(1, result, "</Content:Encoded>", vbTextCompare) > 0
            If InStr(1, result, "<Content:Encoded>", vbTextCompare) > 0 Then
                startTag = InStr(1, result, "<Content:Encoded>", vbTextCompare)
                If InStr(1, result, "</Content:Encoded>", vbTextCompare) > 0 Then
                    endTag = InStr(1, result, "</Content:Encoded>", vbTextCompare) + Len("</Content:Encoded>")
                    result = Left(result, startTag - 1) & Right(result, Len(result) - endTag)
                End If
            End If
        Wend
        Call mCurrentFeed.loadXML(result) 'xmlObject.responseText)
    Else
        Call mCurrentFeed.loadXML(xmlObject.responseText)
    End If
    Exit Sub
networkFailure:
    'network failure (strictly, a .send failure) may be because the network
    'is unavailable or because (annoyingly enough) an HTTP redirect has been
    'returned: unfortunately, this throws an error rather than allowing anyone
    'to access the HTTP response. Sigh. So, let's try using Internet Explorer
    'to call the url and see where it ends up
    'MsgBox "go error"
    Call mWebBrowser.Navigate2(url)
    While mWebBrowser.readyState < READYSTATE_COMPLETE
        DoEvents
    Wend
    url = mWebBrowser.LocationURL
    GoTo retryXMLObject
End Sub

Private Function GetUsernamePassword() As String
    On Error Resume Next
    Dim username As String
    Dim password As String
    
    username = InputBox("This podcast requires a username and password. Please enter your username:")
    If username <> "" Then
        password = InputBox("And now enter your password:")
        If password <> "" Then
            GetUsernamePassword = username & ":" & password
        End If
    End If
End Function
