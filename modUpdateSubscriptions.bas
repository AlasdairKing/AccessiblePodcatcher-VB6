Attribute VB_Name = "modUpdateSubscriptions"
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
'Option Explicit
'
'Private downloadList As New Collection 'holds all the items to be downloaded
'Private finished As Boolean
'
'Private Sub PopulateDownloadList()
''goes through the subscribed podcasts and identifies all their current items,
''storing them in downloadList
'    On Error Resume Next
'    Dim i As Integer
'    Dim newItem As CItem
'    Dim itemIterator As CItem
'
'    Set downloadList = New Collection
'    'first get all the items in the subscribed podcasts
'    For i = 1 To podcasts.Count
'        If podcasts.Item(i).subscribed Then
'            'need to get this podcast
'            frmPodcaster.lstPodcasts.ListIndex = i - 1
'            Call cmdGet_Click
'        End If
'    Next i
'    'now work out the items to get
'    For i = 1 To podcasts.Count
'        If podcasts.Item(i).subscribed Then
'            'add items
'            For Each itemIterator In frmPodcaster.podcasts.Item(i).items
'                Call downloadList.Add(itemIterator)
'            Next itemIterator
'        End If
'    Next i
'End Sub
''
''Private Sub DownloadItems()
'''start downloading the next item in the download list
''    On Error Resume Next
''    If downloadList.Count > 0 Then
''        'still at least one to get: get it
''        call frmpodcaster.winsockHandler.GetFile(downloadlist.Item(1).
''    Else
''        'finished downloading
''        'Delete extraneous files
''        Call CleanUpSubscribedFolders
''        'set the main mouse pointer back to normal
''        frmPodcaster.MousePointer = vbNormal
''    End If
''End Sub
'
'Private Sub CleanUpSubscribedFolders()
''removes any files or folders not current subscribed content
'    On Error Resume Next
'End Sub
'
'Public Sub UpdateSubscriptions()
''get all the items from all the subscribed podcasts
'    On Error Resume Next
'
'    finished = False
'    '1 Get every item to download
'    Call PopulateDownloadList
'    '2 Download them
'    Call DownloadItems
'
'    Dim Path As String
'    Dim result As Long
'    Dim referenceID As Long
'    Dim podcastIterator As CPodcast
'    Dim itemIterator As CItem
'    Dim fso As New FileSystemObject
'    Dim fileName As String
'    Dim folderName As String
'    Dim validFiles As New Dictionary ' holds the files that should be there:
'        'everything else should be deleted
'    Dim folderIterator As Folder
'    Dim fileIterator As File
'    Dim i As Integer
'
'    'now save the podcasts
'    Path = Space(260)
'    result = modAPI.SHGetSpecialFolderLocation(0, CSIDL_PERSONAL, referenceID)
'    result = modAPI.SHGetPathFromIDList(referenceID, Path)
'    'assertion: path now contains path to My Documents
'    Path = Trim(Path)
'    'take off final null character which trim has left behind
'    Path = Replace(Path, Chr(0), "")
'
'    Path = Path & "\AccessiblePodcasts\"
'    'okay, now iterate through all items in subscribed list:
'    For Each podcastIterator In subscribed
'        For Each itemIterator In podcastIterator.items
'            folderName = Path & podcastIterator.name
'            fileName = folderName & "\" & itemIterator.fileName
'            staMain.SimpleText = "Getting " & itemIterator.name & "..."
'            Debug.Print "Saving: " & fileName
'            If fso.FileExists(fileName) Then
'                'file already exists: no need to download
'            Else
'                'file needs to be got: check we have the folder to get it
'                If fso.FolderExists(folderName) Then
'                Else
'                    'need to create folder first
'                    Call fso.CreateFolder(folderName)
'                End If
'                'wait for this to happen
'                While Not (fso.FolderExists(folderName))
'                Wend
'                Call winsockHandler.GetFile(itemIterator.url, fileName)
'                Call validFiles.Add(fileName, fileName)
'            End If
'        Next itemIterator
'    Next podcastIterator
'    'now delete old files
'    staMain.SimpleText = "Deleting old files..."
'    For Each folderIterator In fso.GetFolder(Path).SubFolders
'        'Debug.Print "Folder: " & folderIterator.path
'        For Each fileIterator In folderIterator.Files
'            If validFiles.Exists(fileIterator.Path) Then
'                'okay, this is a file we've just downloaded
'            Else
'                'nope, this is an out-of-date file: delete it
'                Call fileIterator.Delete
'            End If
'        Next fileIterator
'        'check we don't have an empty folder
'        If folderIterator.Files.Count = 0 Then
'            'yes we do: delete the folder
'            Call folderIterator.Delete
'        End If
'    Next folderIterator
'    staMain.SimpleText = "Done"
'    Me.MousePointer = vbNormal
'End Sub
