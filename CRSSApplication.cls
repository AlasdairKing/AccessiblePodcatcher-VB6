VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
'object allows external access to application

Public Event WebpageRequested(url As String)

Public Property Let visible(visible As Boolean)
    On Error Resume Next
    frmRSS.visible = visible
End Property

Public Property Get visible() As Boolean
    On Error Resume Next
    visible = frmRSS.visible
End Property

Private Sub Class_Initialize()
    On Error Resume Next
    Call Load(frmRSS)
    frmRSS.visible = True
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Call Unload(frmRSS)
    End
End Sub

Public Sub CheckSiteForRSSFeeds(url As String)
Attribute CheckSiteForRSSFeeds.VB_Description = "Displays the search for feeds dialogue, inserts the url provided and searches for valid feeds. User can select feeds to add to RSS list."
    On Error Resume Next
    frmFoundFeeds.txtURL.Text = url
    Call frmFoundFeeds.Show
    Call frmFoundFeeds.SearchWebsite(url)
End Sub

Public Sub AddFeed(url As String, Optional name As String)
    'add the feed identified by URL to the feed list.
    Dim newFeed As New CFeed
        
    If Len(url) > 0 Then
        Set newFeed = New CFeed
        newFeed.url = url
        newFeed.name = name
        Call feeds.Add(newFeed)
        'Call frmRSS.BuildWebsiteMenu
        'Call frmRSS.mnuFeed_Click(frmRSS.mnuFeed.Count - 1)
        'delete old feeds
        'Call RemoveFeedsFromRegistry
        'save current feeds
        'Call SaveFeedsToRegistry
        Call SaveFeedsToXML
    End If
End Sub
