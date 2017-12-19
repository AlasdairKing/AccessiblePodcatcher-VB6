VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CWinsockHandler"
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
'CWinsockHandler
'Handles acquiring data using Winsock

Option Explicit

Public Event DataArrival(percentage As Integer)
Public Event DownloadComplete()
Public Event Redirect(newUrl As String)

Private target As String
Private dataReceived As String   ' the data we've been sent so far
Private dataReceptionComplete As Boolean    'whether we've finished getting data
Private dataToReceive As Long   'how much data to receive in total
Private parsedHeader As Boolean 'whether we've already parsed the header and have all its information
Private fileName As String  'the name of the target file
Private myFileName As String    ' the name the user wants to save it to disk under
Private contentStart As Long 'stores the start of the post-header content
Private COMMAND_LINE As String ' the command line for the conversion utility
Private Const QUOTATION_MARKS As String = """"
Private Const HKEY_CURRENT_USER = &H80000001
Public WithEvents cWinsock As Winsock
Attribute cWinsock.VB_VarHelpID = -1
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
    'must be called before FolderRegQueryEx
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Const RESERVED_NULL = 0
Private Const REG_DWORD = 4 '32-bit number
Private Const REG_SZ = 1 'Unicode nul terminated string
Private mError As Boolean

Public Sub GetFile(url As String, path As String)
'gets the file determined by url, saves it to path
    On Error Resume Next
    Debug.Print "Started: " & url
    Call cWinsock.Close  'Stop any current activity
    Call DetermineWinsockSettings(url) 'Get settings for Winsock
    parsedHeader = False
    'work out the filename
    fileName = Right(url, Len(url) - InStrRev(url, "/", Len(url)))
    myFileName = path
    dataReceived = Empty
    'Debug.Print "FileName: " & fileName
    target = url
    Call cWinsock.Connect    'Get ready to go. When Winsock has connected,
        'it will fire Winsock_Connect, which must call WinsockConnect() below.
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
    If ReadRegistryEntryNumber(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyEnable") = 1 Then
        'we're using the proxy
        proxy = ReadRegistryEntryString(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings", "ProxyServer")
        cWinsock.RemoteHost = Left(proxy, InStr(1, proxy, ":") - 1)
        'Debug.Print "RH:" & frmMain.Winsock.RemoteHost
        cWinsock.RemotePort = Val(Right(proxy, Len(proxy) - InStr(1, proxy, ":")))
    Else
        'we're directly connected: use the hostname
        cWinsock.RemoteHost = hostName
        cWinsock.RemotePort = 80
    End If
End Sub

Private Sub cWinsock_Connect()
'Creates the http header and sends it.
    On Error Resume Next
    Dim httpRequest As String
    'build request line that contains the HTTP method, 
    'path to the file to retrieve,
    'and HTTP version info. Each line of the request 
    'must be completed by the vbCrLf
    httpRequest = "GET " & target & " HTTP/1.1" & vbCrLf
    
    'add HTTP headers to the request
    'add required header - "Host", that contains the remote host name
    httpRequest = httpRequest & "Host: " & cWinsock.RemoteHost & vbCrLf
    'add the "Connection" header to force the server to close the connection
    httpRequest = httpRequest & "Connection: close" & vbCrLf
   'add optional header "Accept"
    httpRequest = httpRequest & "Accept: */*" & vbCrLf
    'add other optional headers
    'add a blank line that indicates the end of the request
    httpRequest = httpRequest & vbCrLf
    'send the request
    Call cWinsock.SendData(httpRequest)
    'Good, now we wait for the data to arrive through
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
    Dim i As Integer
    Dim newUrl As String
  
    'first check for an error message
    response = Split(header, " ")
    If Val(response(1)) > 399 Then
        'got an error message! Abort file acquisition
        errorMessageStarts = InStr(1, header, response(1)) + Len(response(1))
        errorMessage = Mid(header, errorMessageStarts, InStr(1, header, vbNewLine) - errorMessageStarts)
        'Call MsgBox("The website reported the following problem so WebbIE could not download the file: " & response(1) & errorMessage, vbExclamation, "Problem with download")
        Debug.Print "Error: " & errorMessage
        ParseHeader = False
        mError = True
    ElseIf Val(response(1)) = 301 Or Val(response(1)) = 302 Or Val(response(1)) = 303 Then
        'got a forwarding request: renavigate to Location: value
        response = Split(header, vbNewLine)
        For i = 0 To UBound(response)
            If Left(response(i), 9) = "Location:" Then
                'found the new url
                newUrl = Replace(response(i), "Location: ", Empty)
                If Left(newUrl, 7) = "http://" Then
                    'new full url
                Else
                    'oops, part url
                End If
                'go get new url
                target = newUrl
'                Call GetFile(newUrl, myFileName)
            End If
        Next i
        ParseHeader = False
    Else
        'okay, not an error
        ParseHeader = True
        'now check it's complete
        If InStr(1, header, vbCrLf & vbCrLf) > 0 Then
            'okay, we've got a full header: process it
            contentTypeStart = InStr(1, header, "Content-Type: ")
            contentLengthStart = InStr(1, header, "Content-Length: ")
            contentStart = InStr(1, header, vbCrLf & vbCrLf) + 4
            'Debug.Print "Content: [" & Mid(header, contentStart, 5)
            dataToReceive = Val(Mid(header, contentLengthStart + Len("Content-Length: "), InStr(contentLengthStart, header, vbNewLine) - contentLengthStart - Len("Content-Length: ")))
            'MsgBox "dtr:" & dataToReceive
            parsedHeader = True
        Else
            'nope, not complete yet. Hang around until it is.
        End If
    End If
    Exit Function
tryLater:
    Exit Function
End Function

Private Sub cWinsock_DataArrival(ByVal bytesTotal As Long)
'Some data has arrived: this may be the first section (containing header information)
'or subsequent packages with more data. Eventually the data is complete, which we
'have to work out from the content length
    On Error Resume Next
    Dim binaryData() As Byte
    Dim stringData As String
    Dim i As Long
    Dim start As Long
    Dim downloadContinue As Boolean
    
    'assume all is going well unless we encounter an HTTP error in the header
    downloadContinue = True
    'check for a header to parse
    If Not parsedHeader Then
        Call cWinsock.PeekData(binaryData, vbArray + vbByte, bytesTotal)
        stringData = StrConv(binaryData, vbUnicode)
        downloadContinue = ParseHeader(stringData)
    End If
    If downloadContinue Then
        'store the data received
        Call cWinsock.GetData(binaryData, vbArray + vbByte, bytesTotal)
        dataReceived = dataReceived & StrConv(binaryData, vbUnicode)
        Debug.Print "Received: " & Len(dataReceived); " of " & dataToReceive
        'check to see if we've finished
        If parsedHeader Then
            'we've parsed the header, so we've got the content length
            If Len(dataReceived) - contentStart >= dataToReceive - 1 Then 'since the figure doesn't include the header
                'okay, we've got all the data: write to disk
                Open myFileName For Output As #1
                Print #1, Mid(dataReceived, contentStart, Len(dataReceived) - contentStart - 1)
                Close #1
                'right, we've finished!
                'Call cWinsock.Close
                RaiseEvent DataArrival(100)
                RaiseEvent DownloadComplete
            Else
                'not finished yet: update the progress form
                RaiseEvent DataArrival(CInt((Len(dataReceived) - contentStart) / dataToReceive * 100))
            End If
        End If
    Else
        If mError Then
            'parsing the header found an HTTP error: stop winsock
            Call cWinsock.Close
            Debug.Print "Abandon!"
            mError = False
        Else
            'no, we're stopping because we've redirected: don't stop winsock
            RaiseEvent Redirect(target)
        End If
    End If
End Sub

Private Function ReadRegistryEntryNumber(hKey As String, regKey As String, itemKey As String) As Long
'returns the value of the registry key indicated by hKey and regKey
    On Error Resume Next
    Dim lengthState As Integer
    Dim result As Long
    Dim hKeyHandle As Long
    Dim keySize As Long
    Dim initialValue As Long
    'first get the original value so we can restore it when we exit
    'also opens the key for writing
    regKey = regKey & Chr(0)
    result = RegOpenKey( _
        hKey, _
        regKey, _
        hKeyHandle _
    )
    'Debug.Print "res1:" & result
    keySize = Len(initialValue)
    itemKey = itemKey & Chr(0)
    result = RegQueryValueEx(hKeyHandle, itemKey, RESERVED_NULL, REG_DWORD, _
        initialValue, keySize)
    'Debug.Print "res2:" & result
    'Debug.Print "Got from registry number: " & initialValue
    Call RegCloseKey(hKeyHandle)
    'If result <> ERROR_SUCCESS Then MsgBox result
    ReadRegistryEntryNumber = initialValue

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

