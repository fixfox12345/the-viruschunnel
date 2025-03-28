Dim response
Do
    response = MsgBox("Do you want to go to my YouTube channel?", vbYesNo, "Question")
    If response = vbYes Then
        ' Open YouTube and go to your channel (replace with your actual YouTube channel link)
        Set objShell = CreateObject("WScript.Shell")
        objShell.Run "https://www.youtube.com/@Fixfox-YT" ' Replace UCXXXXXX with your channel ID
        Exit Do
    ElseIf response = vbNo Then
        ' If user clicks No, the script will run again
        Set objShell = CreateObject("WScript.Shell")
        objShell.Run "wscript.exe """ & WScript.ScriptFullName & """"
        Exit Do
    End If
Loop
