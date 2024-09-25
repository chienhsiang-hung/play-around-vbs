strMessage = "Press the ENTER key to continue."
WScript.StdOut.Write strMessage

Do While Not WScript.StdIn.AtEndOfLine
    Input = WScript.StdIn.Read(1)
Loop

WScript.Echo "The script is complete."
