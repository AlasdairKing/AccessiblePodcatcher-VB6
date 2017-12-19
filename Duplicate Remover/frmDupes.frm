VERSION 5.00
Begin VB.Form frmDupes 
   Caption         =   "Duplicate Editor"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutput 
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDupes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private mUnloading As Boolean
Private WithEvents mDoc As DOMDocument60
Attribute mDoc.VB_VarHelpID = -1

Private Sub Form_Load()
    Call Output("Connecting to database...")
    Set mDoc = New DOMDocument60
    Call mDoc.Load("http://data.webbie.org.uk/podcastReviewDuplicates.php")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Long
    
    mUnloading = True
    For i = 1 To 10000: DoEvents: Next i
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        txtOutput.Left = 0
        txtOutput.Width = Me.ScaleWidth
        txtOutput.Height = Me.ScaleHeight
        txtOutput.Top = 0
    End If
End Sub

Private Sub Output(s As String)
    On Error Resume Next
    txtOutput.SelStart = Len(txtOutput.Text)
    txtOutput.Text = txtOutput.Text & s & vbNewLine
    
End Sub

Private Sub mDoc_onreadystatechange()
    On Error Resume Next
    Dim i As Long
    Dim n As IXMLDOMNode
    Dim n2 As IXMLDOMNode
    Dim result As Integer
    Dim doc1 As DOMDocument60
    Dim doc2 As DOMDocument60
    Dim startTick As Long
    Dim timedOut As Boolean
    
    If mDoc.readyState = 4 Then
        If mDoc.parseError.errorCode = 0 Then
            For Each n In mDoc.documentElement.selectNodes("podcast")
                
                If n.selectSingleNode("title").Text = "" Then
                    Call Output("Untitled podcast: " & n.selectSingleNode("url").Text)
                    Call DeleteFeed(n.Attributes.getNamedItem("id").Text)
                ElseIf n.nextSibling Is Nothing Then
                    'last node
                Else
                    Set n2 = n.nextSibling
                    If n.selectSingleNode("title").Text = n2.selectSingleNode("title").Text Then
                        Set doc1 = New DOMDocument60
                        startTick = GetTickCount
                        timedOut = False
                        Call doc1.Load(n.selectSingleNode("url").Text)
                        While doc1.readyState < 4 And Not timedOut
                            For i = 1 To 1000: DoEvents: Next i
                            timedOut = (GetTickCount - startTick > 5000)
                            If mUnloading Then Exit Sub
                        Wend
                        If timedOut Then
                            Set doc1 = Nothing
                            Call Output("Failed to parse: " & doc1.selectSingleNode("url").Text)
                            Call DeleteFeed(n.Attributes.getNamedItem("id").Text)
                        Else
                            'OK, got doc1.
                            Set doc2 = New DOMDocument60
                            startTick = GetTickCount
                            Call doc2.Load(n2.selectSingleNode("url").Text)
                            timedOut = False
                            While doc2.readyState < 4 And Not timedOut
                                For i = 1 To 1000: DoEvents: Next i
                                timedOut = (GetTickCount - startTick > 5000)
                                If mUnloading Then Exit Sub
                            Wend
                            If timedOut Then
                                Set doc2 = Nothing
                                Call Output("Failed to parse: " & doc2.selectSingleNode("url").Text)
                                Call DeleteFeed(n2.Attributes.getNamedItem("id").Text)
                            Else
                                'OK, got doc2
                                If doc1.parseError.errorCode = 0 And doc2.parseError.errorCode = 0 Then
                                    'They both pass. Delete the one with the longest url
                                    If Len(doc1.url) > Len(doc2.url) Then
                                        Call Output("Deleted: " & n.selectSingleNode("title").Text & vbNewLine & n.selectSingleNode("url").Text)
                                        Call DeleteFeed(n.Attributes.getNamedItem("id").Text)
                                    Else
                                        Call Output("Deleted: " & n2.selectSingleNode("title").Text & vbNewLine & n2.selectSingleNode("url").Text)
                                        Call DeleteFeed(n2.Attributes.getNamedItem("id").Text)
                                    End If
                                ElseIf doc1.parseError.errorCode <> 0 Then
                                    'Doc1 doesn't parse.
                                    Call Output("Deleted: " & n.selectSingleNode("title").Text & vbNewLine & n.selectSingleNode("url").Text)
                                    Call DeleteFeed(n.Attributes.getNamedItem("id").Text)
                                ElseIf doc2.parseError.errorCode <> 0 Then
                                    'Doc2 doesn't parse.
                                    Call Output("Deleted: " & n2.selectSingleNode("title").Text & vbNewLine & n2.selectSingleNode("url").Text)
                                    Call DeleteFeed(n2.Attributes.getNamedItem("id").Text)
                                Else
                                    'Neither passes! Delete both
                                    Call Output("Deleted: " & n.selectSingleNode("title").Text & vbNewLine & n.selectSingleNode("url").Text)
                                    Call DeleteFeed(n.Attributes.getNamedItem("id").Text)
                                    Call Output("Deleted: " & n2.selectSingleNode("title").Text & vbNewLine & n2.selectSingleNode("url").Text)
                                    Call DeleteFeed(n2.Attributes.getNamedItem("id").Text)
                                End If
                            End If
                        End If
                    End If
                End If
            Next n
        Else
            Call Output("Error loading database details. Contact Alasdair. Sorry.")
        End If
        Call Output("Done!")
        MsgBox "Done!"
    End If
End Sub

Private Sub DeleteFeed(id As String)
    On Error Resume Next
    Dim deleteDoc As DOMDocument60
    Dim i As Integer
    
    Set deleteDoc = New DOMDocument60
    Call deleteDoc.Load("http://data.webbie.org.uk/podcastReviewList.php?action=delete&id=" & id)
    While deleteDoc.readyState < 4 And Not mUnloading
        For i = 1 To 100: DoEvents: Next i
    Wend

End Sub
