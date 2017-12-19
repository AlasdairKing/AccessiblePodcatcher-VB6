Attribute VB_Name = "modActiveX"
Option Explicit
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

Public Sub Main()
'handles separate EXE/CreateObject split for ActiveX EXE
    On Error Resume Next
    'set up settings
    Call modPath.DetermineSettingsPath("WebbIE", "Accessible Podcatcher", "1")
    Select Case App.StartMode
        Case vbSModeStandalone
            Call Load(frmPodcaster)
            Call frmPodcaster.Show
        Case vbSModeAutomation
            Call Load(frmPodcaster)
            'don't show until asked
    End Select
End Sub
