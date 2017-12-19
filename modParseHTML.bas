Attribute VB_Name = "modParseHTML"
'   This file is part of WebbIE.
'
'    WebbIE is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    WebbIE is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with WebbIE.  If not, see <http://www.gnu.org/licenses/>.

Option Explicit

Public Function Render(html As String) As String
    On Error GoTo FailedRender:
    Dim i As Long
    Dim output As String
    Dim startAlt As Long
    Dim endAlt As Long
    Dim endTag As Long
    Dim startHREF As Long
    Dim attr As String
    
    i = 1
    html = Trim(html)
    html = Replace(html, vbNewLine, " ")
    html = Replace(html, vbTab, " ")
    html = Replace(html, vbCr, " ")
    html = Replace(html, vbLf, " ")
    While i <= Len(html)
        If Mid(html, i, 1) = "<" Then
            'got a tag! Is it an image?
            If LCase(Mid(html, i, 5)) = "<img " Then
                'it's an image. Try to extract the alt.
                startAlt = InStr(i, html, "alt=", vbTextCompare)
                endTag = InStr(i, html, ">")
                If startAlt > 0 And startAlt < endTag Then
                    startAlt = startAlt + Len("alt=")
                    attr = GetAttributeValue(Right(html, Len(html) - startAlt + 1))
                    output = output & attr
                End If
                i = endTag + 1
            ElseIf LCase(Mid(html, i, 3)) = "<a " Then
                'Link - for podcatcher, just parse content.
                endTag = InStr(i + 1, html, ">")
                If endTag > 0 Then
                    i = endTag + 1
                Else
                    'No closing >
                    i = Len(html) + 1
                End If
            Else
                i = InStr(i + 1, html, ">")
                If i = 0 Then i = Len(html)
                i = i + 1
            End If
            If Right(output, 1) <> " " Then output = output & " "
        Else
            'got a character
            output = output & Mid(html, i, 1)
            i = i + 1
        End If
    Wend
    'Now remove extra whitespace
    While InStr(1, output, "  ", vbBinaryCompare) > 0
        output = Replace(output, "  ", " ")
    Wend
    'Now do escape characters
    output = modHTMLText.RemoveEscapeCharacters(output)
    Render = StripTerminalWhitespace(output)
    Exit Function
FailedRender:
    Err.Clear
End Function

Private Function StripTerminalWhitespace(s As String) As String
    On Error Resume Next
    'removes whitespace at the start or end of s
    Dim removed As Boolean
    Dim l As String
    
    removed = True
    While removed
        removed = False
        'spaces at start/end
        l = Len(s)
        s = Trim(s)
        If l <> Len(s) Then removed = True
        'newlines at start
        If Left(s, Len(vbNewLine)) = vbNewLine Then
            s = Right(s, Len(s) - Len(vbNewLine))
            removed = True
        End If
        'newlines at end
        If Right(s, Len(vbNewLine)) = vbNewLine Then
            s = Left(s, Len(s) - Len(vbNewLine))
            removed = True
        End If
        'tabs at start
        If Left(s, Len(vbTab)) = vbTab Then
            s = Right(s, Len(s) - Len(vbTab))
            removed = True
        End If
        'tabs at end
        If Right(s, Len(vbTab)) = vbTab Then
            s = Left(s, Len(s) - Len(vbTab))
            removed = True
        End If
    Wend
    StripTerminalWhitespace = s
End Function

Private Function GetAttributeValue(html As String) As String
    On Error Resume Next
'Parses some html, starting with the bit after the = in an element, and returns the content of the attribute.
    Dim startAttributeValue As Long
    Dim endAttributeValue As Long
    Dim endTag As Long
    
    endTag = InStr(1, html, ">")
    If endTag = 0 Then endTag = Len(html)
    If Left(html, 1) = """" Then
        startAttributeValue = 2
        endAttributeValue = InStr(2, html, """")
    ElseIf Left(html, 1) = "'" Then
        startAttributeValue = 2
        endAttributeValue = InStr(2, html, "'")
    Else
        startAttributeValue = 1
        endAttributeValue = InStr(1, html, " ")
    End If
    If endAttributeValue > endTag Or endAttributeValue = 0 Then
        endAttributeValue = endTag
    End If
    GetAttributeValue = Mid(html, startAttributeValue, endAttributeValue - startAttributeValue)
End Function
