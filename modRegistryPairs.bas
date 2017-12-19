Attribute VB_Name = "modRegistryPairs"
'Handles the management of key/value pairs in the registry
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

Public Type RegPair
    key As String
    value As String
End Type

Public Function GetRegistryEntries(appName As String, section As String) As RegPair()
    On Error Resume Next
    Dim counter As Integer
    Dim gotFromRegistry As Variant 'should be an array
    Dim emptyArray() As RegPair
    Dim registryIterator As Variant
    Dim keys As Collection
    Dim values As Collection
    Dim i As Integer
    Dim results() As RegPair
    
    If Len(appName) > 0 And Len(section) > 0 Then
        'okay, we have the information to do this
        '1 Get array
        On Error GoTo FailedRegistryAccess:
        gotFromRegistry = GetAllSettings(appName, section)
        If gotFromRegistry = Empty Then
        Else
            '2 Count array numbers
            counter = 0
            For Each registryIterator In gotFromRegistry
                counter = counter + 1
            Next registryIterator
            counter = counter / 2
            '3 Prepare to store values
            Set keys = New Collection
            Set values = New Collection
            '4 Populate collections
            i = 1
            For Each registryIterator In gotFromRegistry
                If i <= counter Then
                    'key
                    Call keys.Add(registryIterator)
                Else
                    'value
                    Call values.Add(registryIterator)
                End If
                i = i + 1
            Next registryIterator
            '5 create array of regpairs
            ReDim results(0 To counter - 1)
            For i = 0 To counter - 1
                results(i).key = keys.item(i + 1)
                results(i).value = values.item(i + 1)
            Next i
            '6 return
            GetRegistryEntries = results
        End If
    End If
    Exit Function
FailedRegistryAccess:
    GetRegistryEntries = emptyArray
    Exit Function
End Function

