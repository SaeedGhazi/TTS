VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   ScaleHeight     =   4680
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   4440
      TabIndex        =   2
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents spVoice As spVoice
Attribute spVoice.VB_VarHelpID = -1

Private Sub Command2_Click()

    If List1.ListIndex > -1 Then

        'Set voice object to voice name selected in list box
        'The new voice speaks its own name

        Set spVoice.Voice = spVoice.GetVoices().Item(List1.ListIndex)
        'MsgBox spVoice.GetVoices().Item(List1.ListIndex).Id
        spVoice.Speak spVoice.Voice.GetDescription

    Else
        MsgBox "Please select a voice from the listbox"
    End If

End Sub

Private Sub Form_Load()
    Set spVoice = New spVoice

    Dim strVoice As String

    

    'Get each token in the collection returned by GetVoices
    For Each T In spVoice.GetVoices
        strVoice = T.GetDescription     'The token's name
        List1.AddItem strVoice          'Add to listbox
    Next

End Sub
    
    
    
    
    


Private Sub Command1_Click()
'spVoice.Voice = 3
    Set spVoice.Voice = spVoice.GetVoices().Item(1)
    spVoice.Rate = 3.5
    spVoice.Volume = 100
    spVoice.Speak "Mehhhrabaad approach good morning. Qatari 4 8 5. released by tehhraan ACC"
    'spVoice.Speak "Eiran air 2 3 4 , turn left heading 2 5 0. descend and maintain 6000 feet"
    'spVoice.Speak "Tabaan air 5 6 2 3 , INTERCEPT RADIAL 1 5 0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set spVoice = Nothing
End Sub

