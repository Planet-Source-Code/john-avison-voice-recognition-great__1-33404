Attribute VB_Name = "VoiceCommands"
Option Explicit

Function LoadVoiceCommands(lstBox As Control) As Boolean

On Local Error GoTo LoadVoiceCommandsError

Dim sNam As String
Dim FileFree As Integer

FileFree = FreeFile

'Check for voice commands text file...
If Dir$(App.Path & "\Commands.Txt") = "" Then
    Exit Function
End If

'Open the commands.txt file...
Open App.Path & "\Commands.Txt" For Input As #FileFree
    Do
        Line Input #FileFree, sNam
        If InStr(LCase$(sNam), "<start>") > 0 And InStr(sNam, "[") = 0 Then
            lstBox.AddItem Mid$(sNam, InStr(LCase$(sNam), "=") + 1)
        End If
    Loop Until EOF(FileFree)
Close #FileFree

LoadVoiceCommands = True
Exit Function



LoadVoiceCommandsError:
    Exit Function

End Function
Function ExecuteVoiceCommand(Who As Form, sPhrase As String) As Boolean

On Local Error GoTo ExecuteVoiceCommandError

Dim X

'Carry out command...
Select Case LCase$(sPhrase)
    'Exit...
    Case "exit"
        SendKeys "%x"
    'Maximize...
    Case "maximize"
        Who.WindowState = vbMaximized
    'Minimize...
    Case "minimize"
        Who.WindowState = vbMinimized
    'Restore...
    Case "restore"
        Who.WindowState = vbNormal
    'Notepad...
    Case "start notepad"
        X = Shell("c:\windows\notepad.exe", vbMaximizedFocus)
    'Down...
    Case "down"
        SendKeys vbKeyDown
    'Up...
    Case "up"
        SendKeys vbKeyUp
    'Right...
    Case "right"
        SendKeys vbKeyRight
    'Left...
    Case "left"
        SendKeys vbKeyLeft
End Select

ExecuteVoiceCommand = True
Exit Function



ExecuteVoiceCommandError:
    Exit Function

End Function
