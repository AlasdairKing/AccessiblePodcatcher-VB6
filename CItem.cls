VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CItem"
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

Private mName As String
Public url As String
Private mDescription As String
Private mPath As String
Public hidden As Boolean ' indicates that this has been "deleted"
    'or hidden by the user.

Public Property Get description() As String
    On Error Resume Next
    description = mDescription
End Property

Public Property Let description(newDescription As String)
    On Error Resume Next
    mDescription = modParseHTML.Render(newDescription)
End Property

Public Property Get name() As String
    On Error Resume Next
    name = mName
End Property

Public Property Let name(newName As String)
    On Error Resume Next
    mName = modParseHTML.Render(newName)
End Property

Public Property Get Path() As String
    On Error Resume Next
    Path = mPath
    Path = Replace(Path, ":", "_")
    Path = Replace(Path, "\", "_")
    Path = Replace(Path, "*", "_")
    Path = Replace(Path, "?", "_")
    Path = Replace(Path, """", "_")
    Path = Replace(Path, "<", "_")
    Path = Replace(Path, ">", "_")
    Path = Replace(Path, "|", "_")
    Path = Replace(Path, "/", "_")
End Property

Public Property Let Path(newPath As String)
    On Error Resume Next
    mPath = newPath
End Property

Public Property Get fullname() As String
    On Error Resume Next
    
    If mDescription = "" Then
        fullname = mName
    Else
        fullname = mName & " - " & mDescription
    End If
End Property

Public Property Get filename() As String
    On Error Resume Next
'worked out from URL
    'Dim fso As New FileSystemObject
    filename = Right(url, Len(url) - InStrRev(url, "/"))
    filename = Replace(filename, ":", "_")
    filename = Replace(filename, "\", "_")
    filename = Replace(filename, "*", "_")
    filename = Replace(filename, "?", "_")
    filename = Replace(filename, """", "_")
    filename = Replace(filename, "<", "_")
    filename = Replace(filename, ">", "_")
    filename = Replace(filename, "|", "_")
    'filename = fso.GetFileName(url)
End Property
