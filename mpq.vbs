'=====================================================================
' flash_30sec_cancel.vbs – Runs for 30 seconds + Cancel option
' Everything else is IDENTICAL to your script
'=====================================================================

Option Explicit

Dim IE, colors, i
Dim response1, response2
Dim startTime

' ----- Warnings WITH CANCEL OPTION -----
response1 = MsgBox(" Run GDI ", 1 + 48, "mpq.exe")   ' 1 = OK/Cancel, 48 = Warning icon
If response1 = 2 Then WScript.Quit                    ' 2 = vbCancel

response2 = MsgBox(" Some people experience dissiness or other after this. Last Warning ", 1 + 48, "mpq.exe")
If response2 = 2 Then WScript.Quit

' ----- Open IE ---------------------------------------------------------
Set IE = CreateObject("InternetExplorer.Application")
IE.Visible = True
IE.Toolbar = False
IE.StatusBar = False
IE.Width = 400
IE.Height = 400

' Navigate to a blank page
IE.Navigate "about:blank"
Do While IE.Busy
    WScript.Sleep 100
Loop

' ----- Start timer (30 seconds = 30000 ms) -----
startTime = Timer

' ----- Flash loop (stops after 30 seconds) -----
colors = Array("red", "green", "blue")
Do
    For i = 0 To UBound(colors)
        ' Stop after 30 seconds
        If (Timer - startTime) * 1000 >= 30000 Then Exit Do
        
        IE.Document.Body.BgColor = colors(i)
        WScript.Sleep 50 ' 50 ms = very fast flash
    Next
Loop

' ----- CLEANUP -----
On Error Resume Next
IE.Quit
On Error Goto 0

WScript.Quit