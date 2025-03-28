Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' List of files to open
files = Array("C:\path\to\your\file1.txt", "C:\path\to\your\file2.txt", "C:\path\to\your\file3.jpg")

' List of sound files to play
sounds = Array("C:\path\to\sound1.mp3", "C:\path\to\sound2.mp3")

' Function to play sound
Sub PlaySound(sound)
    objShell.Run "wmplayer """ & sound & """", 0, True
End Sub

' Function to open a file
Sub OpenFile(file)
    objShell.Run "notepad """ & file & """", 1, False
End Sub

' Loop to perform actions
Do
    ' Open random file
    OpenFile files(Int(Rnd * 3))

    ' Play random sound
    PlaySound sounds(Int(Rnd * 2))

    ' Wait for 2 seconds before repeating
    WScript.Sleep 2000
Loop
