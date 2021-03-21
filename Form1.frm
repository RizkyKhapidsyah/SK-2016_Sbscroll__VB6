VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   ScaleHeight     =   1770
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   3420
      Top             =   315
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   480
      Left            =   1680
      TabIndex        =   2
      Top             =   135
      Width           =   1275
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   480
      Left            =   195
      TabIndex        =   1
      Top             =   135
      Width           =   1275
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   1410
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'Aaron Young
'Analyst Programmer
'ajyoung@pressenter.com
'aarony@redwingsoftware.com
'

Private Sub cmdStart_Click()
    Timer1.Enabled = True
    
End Sub

Private Sub cmdStop_Click()
    Timer1.Enabled = False
    
End Sub

Private Sub Form_Load()
    'Use Timer to do the Scrolling..
    Timer1.Interval = 100
    Timer1.Enabled = False
    
    'set the Status Panel Message..
    StatusBar1.Panels(1) = "Code Guru..  Where you'll find ALL the answers.."
    'Make sure the Tag is empty
    StatusBar1.Panels(1).Tag = ""
End Sub

Private Sub Timer1_Timer()
    With StatusBar1.Panels(1)
        'If the Tag is empty, it's the Beginning of the Scroll
        If .Tag = "" Then
            'Format the Text to make the Scroll Smooth
            'Insert Spaces to Start Scroll from Far Right of Panel
            .Tag = Space(.Width / TextWidth(" ")) & .Text
            .Text = .Tag
        End If
        If Len(.Text) Then
            'While there's some Text, Show One Character Less on the Left
            'Give the Illusion of Scrolling to the Left
            .Text = Mid$(.Text, 2)
        Else
            'Reset the Text and Scroll Again.
            .Text = .Tag
        End If
    End With
End Sub


