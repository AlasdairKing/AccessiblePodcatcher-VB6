VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWait 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Updating subscribed podcasts - please wait"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Tag             =   "frmWait"
   Begin VB.Timer tmrMain 
      Interval        =   100
      Left            =   4680
      Top             =   120
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Default         =   -1  'True
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
'Provides a waiting dialogue. It checks the registry to see how
'many items when it starts up, and polls to see how many have
'been done, writing a cancel value if required, and terminating
'when finished
Option Explicit

Private Sub cmdCancel_Click()
    On Error Resume Next
    Call SaveSetting("WebbIE", "Wait", "Cancelled", True)
    lblStatus.Enabled = False
    cmdCancel.Enabled = False
    pbDownload.Enabled = False
    Call Beep
    Me.MousePointer = vbHourglass
End Sub

Private Sub Form_Load()
    Call SaveSetting("WebbIE", "Wait", "Visible", True)
    Call SaveSetting("WebbIE", "Wait", "Cancelled", False)
    Me.MousePointer = vbArrowHourglass
    Call tmrMain_Timer
    pbDownload.Value = 0
    Call Me.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call cmdCancel_Click
    Cancel = True
End Sub

Private Sub lblStatus_Change()
    On Error Resume Next
    If lblStatus.Width > Me.Width + 500 Then
        Me.Width = lblStatus.Width + 500
        pbDownload.Width = Me.Width - 400
        cmdCancel.Left = Me.Width / 2 - cmdCancel.Width / 2
    End If
End Sub

Private Sub tmrMain_Timer()
    'check to see if we've been closed by an external application
    If GetSetting("WebbIE", "Wait", "Visible") = False Then
        'yep, Visible set to false, kill myself
        Call Unload(Me)
    Else
        'nope, still running. Update progress bar.
        pbDownload.Max = CDbl(GetSetting("WebbIE", "Wait", "NumberItems", 1))
        pbDownload.Value = GetSetting("WebbIE", "Wait", "Progress", 0)
        lblStatus.Caption = GetSetting("WebbIE", "Wait", "Label", "")
        Me.Caption = GetSetting("WebbIE", "Wait", "FormCaption", "")
        cmdCancel.Caption = GetSetting("WebbIE", "Wait", "Button", "Cancel")
    End If
End Sub
