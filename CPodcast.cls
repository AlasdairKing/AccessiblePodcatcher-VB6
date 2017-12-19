VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPodcast"
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

Private mItems As Collection
Private mName As String
Private mDescription As String
Private mURL As String
Public id As Long ' The ID provided by the podcast directory - a Long.
Public parseError As Boolean
Private mSubscribed As Boolean ' whether the podcast should be added to the subscribed collection
Private Enum pcState
    stateInitialised
    stateGotPodcast
    stateDying
End Enum
Private mState As pcState

Public Sub MarkForDeath()
    mState = stateDying
End Sub

Public Property Get url() As String
    On Error Resume Next
    url = mURL
End Property

Public Property Let url(newURL As String)
    On Error Resume Next
    mURL = newURL
End Property

Private Sub Class_Initialize()
    On Error Resume Next
    Set mItems = New Collection
    mSubscribed = True
    mState = stateInitialised
End Sub

Public Property Get name() As String
    On Error Resume Next
    name = mName
End Property

Public Property Get fullname() As String
    On Error Resume Next
    If mDescription = "" Then
        fullname = mName
    Else
        fullname = mName & " - " & mDescription
    End If
End Property

Public Property Let name(newName As String)
    On Error Resume Next
    If Len(newName) > 0 Then
''        'remove old setting - indexed under name - from registry
''        If mName <> "" Then
''            If Len(GetSetting("AccessiblePodcaster", "Podcasts", mName)) > 0 Then
''                Call DeleteSetting("AccessiblePodcaster", "Podcasts", mName)
''            End If
''            If Len(GetSetting("AccessiblePodcaster", "Subscribed Podcasts", mName, Empty)) > 0 Then
''                Call DeleteSetting("AccessiblePodcaster", "Subscribed Podcasts", mName)
''            End If
''        End If
        mName = newName
    End If
End Property

Public Property Get nameForFolder() As String
'returns the name of the podcast as a valid string for use in the filesystem
    On Error Resume Next
    nameForFolder = mName
    nameForFolder = Replace(nameForFolder, "*", "_")
    nameForFolder = Replace(nameForFolder, ":", "_")
    nameForFolder = Replace(nameForFolder, "?", "_")
    nameForFolder = Replace(nameForFolder, "<", "_")
    nameForFolder = Replace(nameForFolder, ">", "_")
    nameForFolder = Replace(nameForFolder, "|", "_")
    nameForFolder = Replace(nameForFolder, """", "_")
    nameForFolder = Replace(nameForFolder, "\", "_")
    nameForFolder = Replace(nameForFolder, "/", "_")
    nameForFolder = Replace(nameForFolder, "'", "_")
    nameForFolder = Replace(nameForFolder, ".", "")
    nameForFolder = Trim(nameForFolder)
End Property

Public Property Get items(Optional showHidden As Boolean = False) As Collection
    On Error Resume Next
    Dim itemIterator As CItem
    If showHidden Then
        Set items = mItems
    Else
        Set items = New Collection
        For Each itemIterator In mItems
            If itemIterator.hidden Then
                'don't show!
            Else
                'show this, so add
                Call items.Add(itemIterator)
            End If
        Next itemIterator
    End If
End Property

'Parse the XML retrieved for a podcast and create items
Public Property Let xml(newXML As String)
    On Error Resume Next
    Dim itemIterator As IXMLDOMNode
    Dim podcastDoc As New DOMDocument30
    Dim enclosureIterator As IXMLDOMNode
    Dim titleIterator As IXMLDOMNode
    Dim newItem As CItem
    Dim titleNode As IXMLDOMNode
    Dim descriptionNode As IXMLDOMNode
    Dim counter As Long

    Call podcastDoc.loadXML(newXML)
    If podcastDoc.parseError.errorCode = 0 Then
        'parsed okay
        parseError = False
        'get name if we don't have one
        If Len(mName) = 0 Then
            'check the XML for a name
            mName = podcastDoc.documentElement.selectNodes("//channel/title").Item(0).Text
        End If
        mDescription = podcastDoc.documentElement.selectNodes("//channel/description").Item(0).Text
        Set mItems = New Collection
        counter = 1
        For Each itemIterator In podcastDoc.documentElement.selectNodes("//item")
            'okay, got an item: extract name and url of media file
            Set newItem = New CItem
            For Each enclosureIterator In itemIterator.selectNodes("enclosure")
                newItem.url = enclosureIterator.Attributes.getNamedItem("url").Text
                'Debug.Print "Got url: " & newItem.url
            Next enclosureIterator
            'get title
            Set titleNode = itemIterator.selectSingleNode("title")
            Set descriptionNode = itemIterator.selectSingleNode("description")
            If titleNode Is Nothing Then
                'try description
                Set descriptionNode = itemIterator.selectSingleNode("description")
                If descriptionNode Is Nothing Then
                    'haven't got any text. Silly website. Make up a name.
                    newItem.name = mName & " " & counter
                Else
                    'okay, use that
                    newItem.name = descriptionNode.Text
                End If
            Else
                newItem.name = titleNode.Text
            End If
            If descriptionNode Is Nothing Then
                'no description
            Else
                newItem.description = descriptionNode.Text
            End If
            If Len(newItem.url) > 0 Then
                'okay, we got a url for this item: sometimes there isn't one,
                'in which case we ignore it.
                Call mItems.Add(newItem)
            End If
        Next itemIterator
        'record that this has been obtained okay
        mState = stateGotPodcast
        'load hidden file
        LoadHiddenItemsList
    Else
        'failed to parse XML: alert user
        'MsgBox "Failed to parse Podcast.", vbExclamation, "Podcast broken"
        parseError = True
        Debug.Print "Failed podcast: " & vbNewLine & newXML
    End If
End Property

Public Property Let subscribed(newSubscribed As Boolean)
    On Error Resume Next
    Dim value As String
    Dim podcastXML As DOMDocument30
    Dim podcastNode As IXMLDOMNode
    
    mSubscribed = newSubscribed
    Set podcastXML = New DOMDocument30
    podcastXML.async = False
    Call podcastXML.Load(modPath.settingsPath & "\podcasts.xml")
    Set podcastNode = podcastXML.documentElement.selectSingleNode("podcast[name=""" & mName & """]")
    If podcastNode Is Nothing Then
        'hmm, didn't find this podcast. Never mind. Sort out when we save
    Else
        podcastNode.selectSingleNode("subscribed2").Text = CStr(newSubscribed)
    End If
    Call podcastXML.save(modPath.settingsPath & "\podcasts.xml")
    'Call SaveSetting("AccessiblePodcaster", "Subscribed Podcasts", mName, value)
End Property

Public Property Get subscribed() As Boolean
    On Error Resume Next
    subscribed = mSubscribed
End Property

Public Sub HideItem(podcastItem As CItem)
    On Error Resume Next
    Dim it As CItem
    'hides the item in the podcast indicated by podcastItem
    For Each it In mItems
        If it.url = podcastItem.url Then it.hidden = True
    Next it
    Call SaveHiddenItemsList
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    'Move from saving information in registry to xml
    'Call SavePodcastInfoToRegistry
    Call SavePodcastInfoToXML
End Sub

Private Sub SavePodcastInfoToXML()
    On Error Resume Next
    Dim podcastXML As DOMDocument30
    Dim podcastIterator As IXMLDOMNode
    Dim found As Boolean
    Dim newChild As IXMLDOMNode
    Dim tempChild As IXMLDOMNode
    Dim idAtt As IXMLDOMAttribute
    
    'This saves the podcast to the storage XML file: it makes a big thing
    'of saving to the end of the DocumentElement, but it doesn't make
    'any difference really. Was a problem with sorting, solved by deleting
    'the xml file.
    Set podcastXML = New DOMDocument30
    podcastXML.async = False
    Call podcastXML.Load(modPath.settingsPath & "\podcasts.xml")
    If podcastXML.parseError.errorCode <> 0 Then
        Call podcastXML.loadXML("<podcasts/>")
    End If
    For Each podcastIterator In podcastXML.documentElement.selectNodes("podcast")
        If podcastIterator.selectSingleNode("name").Text = mName Then
            'found this podcast
            found = True
            'remove node from current position...
            Set tempChild = podcastXML.documentElement.removeChild(podcastIterator)
            
            '1.7.8 Update the url in case the user has changed it.
            tempChild.selectSingleNode("url").Text = mURL
            
            'if dying, don't add it back:
            If mState = stateDying Then
            Else
                'still here, update and add back to end. This preserves sorting.
                If tempChild.selectSingleNode("subscribed2") Is Nothing Then
                    Set newChild = podcastXML.createNode(NODE_ELEMENT, "subscribed2", "")
                    Call tempChild.appendChild(newChild)
                End If
                tempChild.selectSingleNode("subscribed2").Text = mSubscribed
                Call podcastXML.documentElement.appendChild(tempChild)
            End If
            Exit For
        Else
'            Debug.Print "Not found at " & counter
        End If
    Next podcastIterator
    'if we didn't find the podcast, create it, iff we're not dying
    If Not found And Not (mState = stateDying) Then
        Set podcastIterator = podcastXML.createNode(NODE_ELEMENT, "podcast", "")
        Set newChild = podcastXML.createNode(NODE_ELEMENT, "name", "")
        newChild.Text = mName
        Call podcastIterator.appendChild(newChild)
        Set newChild = podcastXML.createNode(NODE_ELEMENT, "url", "")
        newChild.Text = mURL
        Call podcastIterator.appendChild(newChild)
        Set newChild = podcastXML.createNode(NODE_ELEMENT, "subscribed2", "")
        newChild.Text = mSubscribed
        Call podcastIterator.appendChild(newChild)
        Set idAtt = podcastXML.createAttribute("id")
        idAtt.Text = id
        Call podcastIterator.Attributes.setNamedItem(idAtt)
        Call podcastXML.documentElement.appendChild(podcastIterator)
    End If
    Call podcastXML.save(modPath.settingsPath & "\podcasts.xml")
End Sub

'Private Sub SavePodcastInfoToRegistry()
'    On Error Resume Next
'    'save class details to registry
'    'remove old entry if it exists
'    If Len(GetSetting("AccessiblePodcaster", "Podcasts", mName)) > 0 Then
'        Call DeleteSetting("AccessiblePodcaster", "Podcasts", mName)
'    End If
'    If Len(GetSetting("AccessiblePodcaster", "Subscribed Podcasts", mName, Empty)) > 0 Then
'        Call DeleteSetting("AccessiblePodcaster", "Subscribed Podcasts", mName)
'    End If
'    'add new entry if we are not marked for death
'    If mState = stateDying Then
'        'dying: clean up data
'        Call DeleteHiddenItemsList
'    Else
'        'still here
'        Call SaveSetting("AccessiblePodcaster", "Podcasts", mName, mURL)
'        If subscribed Then
'            Call SaveSetting("AccessiblePodcaster", "Subscribed Podcasts", mName, "True")
'        End If
'    End If
'End Sub

Public Property Get hiddenItems() As Collection
    On Error Resume Next
    Dim itemIterator As CItem
    
    Set hiddenItems = New Collection
    For Each itemIterator In mItems
        If itemIterator.hidden Then
            Call hiddenItems.Add(itemIterator)
        End If
    Next itemIterator
End Property

Private Function URLIntoFileName(url As String) As String
    On Error Resume Next
    url = Replace(url, "http://", Empty, , , vbTextCompare)
    url = Replace(url, "https://", Empty, , , vbTextCompare)
    If InStr(1, url, "#") > 0 Then
        url = Left(url, InStr(1, url, "#") - 1)
    End If
    If InStr(1, url, "?") > 0 Then
        url = Left(url, InStr(1, url, "?") - 1)
    End If
    url = Replace(url, "/", "_")
    url = Replace(url, ":", "_")
    url = Replace(url, "\", "_")
    URLIntoFileName = url
End Function

Private Sub SaveHiddenItemsList()
    On Error Resume Next
    Dim itemIterator As CItem
    Dim hidden As String
    
    If mState = stateGotPodcast Then
        For Each itemIterator In mItems
            If itemIterator.hidden Then
                hidden = hidden & itemIterator.url & vbNewLine
            End If
        Next itemIterator
        Open modPath.settingsPath & "\Deleted\" & URLIntoFileName(Me.url) & ".ini" For Output As #1
        Write #1, hidden
        Close #1
    End If
    
'    Dim newItemNode As IXMLDOMNode
'
'    Set mHiddenItems = New DOMDocument30
'    Call mHiddenItems.loadXML("<deleted/>")
'    For Each itemIterator In mItems
'        If itemIterator.hidden Then
'            Set newItemNode = mHiddenItems.createNode(NODE_ELEMENT, "item", "")
'            newItemNode.Text = itemIterator.url
'            Call mHiddenItems.documentElement.appendChild(newItemNode)
'        End If
'    Next itemIterator
'    Call mHiddenItems.save(modPath.settingsPath & "\Deleted\" & URLIntoFileName(Me.url))
End Sub

Private Sub LoadHiddenItemsList()
    On Error Resume Next
    Dim itemIterator As CItem
    Dim hidden As String
    Dim fso As Scripting.FileSystemObject
    Dim ts As Scripting.TextStream
    
    'load
    Set fso = New Scripting.FileSystemObject
    If fso.FileExists(modPath.settingsPath & "\Deleted\" & URLIntoFileName(Me.url) & ".ini") Then
        Set ts = fso.OpenTextFile(modPath.settingsPath & "\Deleted\" & URLIntoFileName(Me.url) & ".ini")
        hidden = ts.ReadAll
        Call ts.Close
    End If
    Set fso = Nothing
    'process
    For Each itemIterator In mItems
        itemIterator.hidden = (InStr(1, hidden, itemIterator.url) > 0)
    Next itemIterator
    'update saved list
    Call SaveHiddenItemsList
End Sub

Private Sub DeleteHiddenItemsList()
    On Error Resume Next
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    Call fso.DeleteFile(modPath.settingsPath & "\Deleted\" & URLIntoFileName(Me.url) & ".ini")
    Set fso = Nothing
End Sub

Public Property Get description() As String
    On Error Resume Next
    description = mDescription
End Property

