' Create a form
Set objShell = CreateObject("WScript.Shell")
Set objIE = CreateObject("InternetExplorer.Application")
objIE.Visible = True
objIE.Navigate "about:blank"
Do While objIE.Busy
    WScript.Sleep 100
Loop

' Add HTML content to the form
objIE.document.body.innerHTML = "<html><body><h1>Mouse Mover</h1><p>Status: <span id='status'>Running</span></p></body></html>"

' Function to move the mouse
Set objMouse = CreateObject("WScript.Shell")
Set objWshShell = CreateObject("WScript.Shell")
Do
    If objWshShell.AppActivate("Internet Explorer") Then
        If objWshShell.GetKeyState(32) Then ' 32 is the ASCII code for the spacebar
            Exit Do
        End If
    End If
    objMouse.SendKeys "{F13}"
    WScript.Sleep 1000
Loop

' Update the status
objIE.document.getElementById("status").innerText = "Stopped"
