VERSION 5.00
Object = "{4E3D9D11-0C63-11D1-8BFB-0060081841DE}#1.0#0"; "Xlisten.dll"
Begin VB.Form frmVoiceRecognition 
   Caption         =   "Form1"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   ScaleHeight     =   5070
   ScaleWidth      =   3045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   1125
   End
   Begin VB.ListBox lstCommands 
      Height          =   3765
      ItemData        =   "frmVoiceRecognition.frx":0000
      Left            =   120
      List            =   "frmVoiceRecognition.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   420
      Width           =   2775
   End
   Begin ACTIVELISTENPROJECTLibCtl.DirectSR VR 
      Height          =   255
      Left            =   1320
      OleObjectBlob   =   "frmVoiceRecognition.frx":0004
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label2 
      Caption         =   "John Avison  avisonjohn@hotmail.com"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available Commands"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1470
   End
End
Attribute VB_Name = "frmVoiceRecognition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExit_Click()

Unload Me

End Sub

Private Sub btnExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlack
End Sub

Private Sub Form_Load()

On Local Error Resume Next

'Load the voice commands from the commands text file...
Call LoadVoiceCommands(lstCommands)

'Setup the voice recognition control...
Call VR.Deactivate
Call VR.GrammarFromFile(App.Path & "\Commands.Txt")
Call VR.Activate

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbRed
End Sub

Private Sub Label2_Click()
Shell = ("start mailto:avisonjohn@hotmail.com")
End Sub

Private Sub lstCommands_Click()

End Sub

Private Sub lstCommands_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbBlack
End Sub

Private Sub VR_PhraseFinish(ByVal flags As Long, ByVal beginhi As Long, ByVal beginlo As Long, ByVal endhi As Long, ByVal endlo As Long, ByVal Phrase As String, ByVal parsed As String, ByVal results As Long)

On Local Error Resume Next

'Check to see if we have a matching word...
If Trim$(Phrase) <> "" Then
    Me.Caption = Phrase
    Call ExecuteVoiceCommand(Me, Phrase)
Else
    Me.Caption = "Unrecognized Command..."
End If

End Sub
