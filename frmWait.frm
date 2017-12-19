VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmWait 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Updating subscribed podcasts - please wait"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   ControlBox      =   0   'False
   Icon            =   "frmWait.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "frmWait"
   Begin MSWinsockLib.Winsock winsck 
      Left            =   2400
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Tag             =   "frmWait.cmdCancel"
      Top             =   1200
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pbDownload 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "lblStatus"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Provides status information for the subscription updates
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

Private mTarget As String ' what we want to download
Private mDataReceived As String   ' the data we've been sent so far
Private mDataToReceive As Long   'how much data to receive in total
Private mParsedHeader As Boolean 'whether we've already parsed the header and have all its information
Private mLocalFilename As String    ' the name the user wants to save it to disk under
Private mContentStart As Long 'stores the start of the post-header content
Private COMMAND_LINE As String ' the command line for the conversion utility
Private Const QUOTATION_MARKS As String = """"
Private mHTTPResult As Long ' holds the code for the last HTTP result
Private mHTTPResultMessage As String ' holds the result of the last
    'HTTP action

Private Sub cmdCancel_Click()
    On Error Resume Next
    Call winsck.Close
    Call Me.Hide
    Call frmPodcaster.Show
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    cmdCancel.Caption = frmPodcaster.gobjLanguageHandler.GetText("Cancel")
    cmdCancel.Default = False
    cmdCancel.Cancel = True
    Me.Left = frmPodcaster.Left + (frmPodcaster.Width - Me.Width) / 2
    Me.Top = frmPodcaster.Top + (frmPodcaster.Height - Me.Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call winsck.Close
End Sub

Private Sub lblStatus_Change()
    On Error Resume Next
    If lblStatus.Width > Me.Width + 500 Then
        Me.Width = lblStatus.Width + 500
        pbDownload.Width = Me.Width - 400
        cmdCancel.Left = Me.Width / 2 - cmdCancel.Width / 2
    End If
End Sub

Public Sub GetFile(url As String, path As String)
'gets the file determined by url, saves it to path
    On Error Resume Next
    Call winsck.Close  'Stop any current activity
    Call DetermineWinsockSettings(url) 'Get settings for Winsock
    mParsedHeader = False
    'work out the filename
'    mFilename = Right(url, Len(url) - InStrRev(url, "/", Len(url)))
    mLocalFilename = path
    mDataReceived = Empty
    'Debug.Print "FileName: " & fileName
    mTarget = url
    Call winsck.Connect    'Get ready to go. When Winsck has connected,
        'it will fire Winsck_Connect
End Sub

Private Sub DetermineWinsockSettings(url As String)
'works out the protocol, remote host and other settings needed for
'Winsock to function. This will probably require some API and registry calls
    On Error Resume Next
    Dim hostName As String
    Dim proxy As String
    
    'look at the address to get to work out the target, if needed
    If InStr(1, url, "http://") = 1 Or InStr(1, url, "ftp://") = 1 Then
        hostName = Mid(url, 8, InStr(8, url, "/") - 8) 'start at 8 to avoid http://
    Else
        hostName = Mid(url, 7, InStr(7, url, "/") - 7) 'start at 7 to avoid ftp://
    End If
    'how are we connecting to the internet? Check out the registry settings.
    If ReadRegistryEntryNumber(modAPI.HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyEnable") = 1 Then
        'we're using the proxy
        proxy = ReadRegistryEntryString(modAPI.HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyServer")
        winsck.RemoteHost = Left(proxy, InStr(1, proxy, ":") - 1)
        'Debug.Print "RH:" & frmMain.Winsock.RemoteHost
        winsck.RemotePort = Val(Right(proxy, Len(proxy) - InStr(1, proxy, ":")))
'        Debug.Print "RP:" & frmActiveXHost.winsock.RemotePort
    Else
        'we're directly connected: use the hostname
        winsck.RemoteHost = hostName
        winsck.RemotePort = 80
    End If
End Sub

Private Function ParseHeader(header As String) As Boolean
'works out all the file information from the header: returns false iff an HTTP
'error code (value greater than 400, according to http://www.w3.org/Protocols/rfc2616/rfc2616-sec6.html#sec6)
'is found.
    On Error GoTo tryLater
    Dim contentTypeStart As Long
    Dim contentLengthStart As Long
    Dim response() As String
    Dim errorMessage As String
    Dim errorMessageStarts As Long
    
    'first check for an error message
    response = Split(header, " ")
    mHTTPResult = Val(response(1))
    If mHTTPResult > 399 Then
        'got an error message! Abort file acquisition
        errorMessageStarts = InStr(1, header, response(1)) + Len(response(1))
        mHTTPResultMessage = Mid(header, errorMessageStarts, InStr(1, header, vbNewLine) - errorMessageStarts)
        ParseHeader = False
    Else
        'okay, not an error
        ParseHeader = True
        'now check it's complete
        If InStr(1, header, vbCrLf & vbCrLf) > 0 Then
            'okay, we've got a full header: process it
            contentTypeStart = InStr(1, header, "Content-Type: ")
            contentLengthStart = InStr(1, header, "Content-Length: ")
            mContentStart = InStr(1, header, vbCrLf & vbCrLf) + Len(vbCrLf & vbCrLf)
            'Debug.Print "Content: [" & Mid(header, mcontentStart, 5)
            mDataToReceive = Val(Mid(header, contentLengthStart + Len("Content-Length: "), InStr(contentLengthStart, header, vbNewLine) - contentLengthStart - Len("Content-Length: ")))
            Debug.Print "mDataToReceive:" & mDataToReceive
            Debug.Print "mContentStart:" & mContentStart
            Debug.Print "header:[" & header & "]"
            
            mParsedHeader = True
        Else
            'nope, not complete yet. This won't return true, so we'll
            'abort downloading: we're assuming that the 500-byte (guess)
            'header always comes down in one chunk.
        End If
    End If
    Exit Function
tryLater:
    'in fact, this won't set parseheader to true, so it'll terminate
    'anything calling ParseHeader and testing the result.
    Exit Function
End Function

Private Function ReadRegistryEntryString(hKey As String, regKey As String, itemKey As String) As String
'returns the value of the registry key indicated by hKey and regKey
'NOTE: be careful when parsing the results, it seems to produce an ANSI version
'overlaying a Unicode version, e.g. "HELLO L O "
    On Error Resume Next
    Dim lengthState As Integer
    Dim result As Long
    Dim hKeyHandle As Long
    Dim keySize As Long
    Dim initialValue As String
    'first get the original value so we can restore it when we exit
    'also opens the key for writing
    regKey = regKey & Chr(0)
    result = RegOpenKey( _
        hKey, _
        regKey, _
        hKeyHandle _
    )
    'Debug.Print "Result1:" & result
    keySize = 256
    initialValue = String(keySize, 0)
    itemKey = itemKey & Chr(0)
    result = RegQueryValueEx(hKeyHandle, itemKey, RESERVED_NULL, REG_SZ, _
        ByVal initialValue, keySize)
        'Debug.Print "Result2:" & result
    initialValue = Trim(initialValue)
    'Debug.Print "Got from registry: " & initialValue
    Call RegCloseKey(hKeyHandle)
    'If result <> ERROR_SUCCESS Then MsgBox result
    ReadRegistryEntryString = initialValue
End Function

Private Sub winsck_Connect()
    On Error Resume Next
    Dim httpRequest As String
    'build request line that contains the HTTP method, 
    'path to the file to retrieve,
    'and HTTP version info. Each line of the request 
    'must be completed by the vbCrLf
    httpRequest = "GET " & mTarget & " HTTP/1.1" & vbCrLf
    
    'add HTTP headers to the request
    'add required header - "Host", that contains the remote host name
    httpRequest = httpRequest & "Host: " & winsck.RemoteHost & vbCrLf
    'add the "Connection" header to force the server to close the connection
    httpRequest = httpRequest & "Connection: close" & vbCrLf
   'add optional header "Accept"
    httpRequest = httpRequest & "Accept: */*" & vbCrLf
    'add other optional headers
    'add a blank line that indicates the end of the request
    httpRequest = httpRequest & vbCrLf
    'send the request
    Call winsck.SendData(httpRequest)
    'Good, now we wait for the data to arrive through
End Sub

Private Sub winsck_DataArrival(ByVal bytesTotal As Long)
'Some data has arrived: this may be the first section (containing header information)
'or subsequent packages with more data. Eventually the data is complete, which we
'have to work out from the content length
    On Error Resume Next
    Dim binaryData() As Byte
    Dim stringData As String
    Dim i As Long
    Dim start As Long
    Dim downloadContinue As Boolean
    Dim justParsed As Boolean

    'Debug.Print "Total: " & bytesTotal

    'assume all is going well unless we encounter an HTTP error in the header
    downloadContinue = True
    'check for a header to parse
    If Not mParsedHeader Then
        Call winsck.PeekData(binaryData, vbArray + vbByte, bytesTotal)
        stringData = StrConv(binaryData, vbUnicode)
        downloadContinue = ParseHeader(stringData)
        justParsed = True ' we've just parsed the header:
    End If
    If downloadContinue Then
        'work out how much data received, and if we've therefore finished
        If justParsed Then
            'take the header out of the buffer: note the decrement to content start:
            Call winsck.GetData(binaryData, vbArray + vbByte, mContentStart - 1)
        End If
        'Have we got all the data?
        'Call winsck.PeekData(binaryData, vbArray + vbByte, 5)
        'Debug.Print "start now: " & StrConv(binaryData, vbUnicode)
        Debug.Print "Comparison: bytestotal=" & bytesTotal & " mContentStart=" & mContentStart & " mDataToReceive=" & mDataToReceive
        If bytesTotal = mDataToReceive Then
            'finished! Copy to disk
            Call winsck.GetData(binaryData, vbArray + vbByte, bytesTotal - mContentStart)
            mDataReceived = mDataReceived & StrConv(binaryData, vbUnicode)
            Open mLocalFilename For Output As #1
            Print #1, mDataReceived
            Close #1
            'right, we've finished!
            Call winsck.Close
            Call AdvanceToNextItem
        Else
            'not finished yet: just update progress bar
            pbDownload.Max = mDataToReceive
            pbDownload.value = bytesTotal - mContentStart
            Call Me.Refresh
        End If
    Else
        'oh, we got an HTTP error from the header. Stop downloading
        Call winsck.Close
        MsgBox frmPodcaster.gobjLanguageHandler.GetText("Podcast download failed:") & " " & mHTTPResultMessage, vbOKOnly, "Accessible Podcatcher"
        Call AdvanceToNextItem
    End If
'''    If downloadContinue Then
'''        'store the amount of data received
'''        Call winsck.GetData(binaryData, vbArray + vbByte, bytesTotal)
'''        mDataReceived = mDataReceived & StrConv(binaryData, vbUnicode)
'''        'check to see if we've finished
'''        If mParsedHeader Then
'''            'we've parsed the header, so we've got the content length
'''            If Len(mDataReceived) - mContentStart >= mDataToReceive - 1 Then 'since the figure doesn't include the header
'''                'okay, we've got all the data: write to disk
'''                Open mLocalFilename For Output As #1
'''                Print #1, Mid(mDataReceived, mContentStart, Len(mDataReceived) - mContentStart - 1)
'''                Close #1
'''                'right, we've finished!
'''                Call winsck.Close
'''                Call AdvanceToNextItem
'''            Else
'''                'not finished yet: update the progress form
'''                pbDownload.Max = mDataToReceive
'''                pbDownload.value = Len(mDataReceived)
'''            End If
'''        End If
'''    Else
'''        'parsing the header found an HTTP error: stop winsock
'''        Call winsck.Close
'''        MsgBox frmPodcaster.gobjLanguageHandler.GetText("Podcast download failed:") & " " & mHTTPResultMessage, vbOKOnly, "Accessible Podcatcher"
'''        Call AdvanceToNextItem
'''    End If
End Sub

Private Sub AdvanceToNextItem()
'checks to see if we've finished doing all the podcast items, and
'if not, starts getting the next one.
    On Error Resume Next
    Dim item As CItem
    Dim fso As FileSystemObject
    Dim i As Long
    Dim path As String
    Dim startedADownload As Boolean
    
    'okay, have we finished downloading everything we were going to?
    If frmPodcaster.mItems.Count = frmPodcaster.mItemCount Then
        'yes!
        Call Me.Hide
        MsgBox frmPodcaster.gobjLanguageHandler.GetText("Completed updating subscriptions"), vbOKOnly, "Accessible Podcatcher"
        Call frmPodcaster.Show
        Call frmPodcaster.CleanupSubscriptions
    Else
        'no! Check for next one
        Set fso = New FileSystemObject
        For i = frmPodcaster.mItemCount + 1 To frmPodcaster.mItems.Count
            Set item = frmPodcaster.mItems.item(i)
            lblStatus.Caption = frmPodcaster.gobjLanguageHandler.GetText("Downloading") & " (" & i & " " & frmPodcaster.gobjLanguageHandler.GetText("of") & " " & frmPodcaster.mItems.Count & ") """ & item.name & """"
            If Not fso.FolderExists(frmPodcaster.mPath & item.path) Then
                'need to create folder for this
                Call fso.CreateFolder(frmPodcaster.mPath & item.path)
            End If
            path = frmPodcaster.mPath & item.path & "\" & item.filename
            Call frmPodcaster.mValidFiles.Add(path, path)
            If fso.FileExists(frmPodcaster.mPath & item.path & "\" & item.filename) Then
                'already exists! skip it.
                Debug.Print "Skipping"
            Else
                'Okay, need to download
                Call GetFile(item.url, frmPodcaster.mPath & item.path & "\" & item.filename)
                startedADownload = True
            End If
        Next i
        If startedADownload Then
            'okay, found something
            frmPodcaster.mItemCount = i
        Else
            'failed to find anything to download: stop
            Call Me.Hide
            MsgBox frmPodcaster.gobjLanguageHandler.GetText("Completed updating subscriptions"), vbOKOnly, "Accessible Podcatcher"
            Call frmPodcaster.Show
            Call frmPodcaster.CleanupSubscriptions
        End If
    End If
End Sub
