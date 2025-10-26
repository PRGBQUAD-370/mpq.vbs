X=MsgBox(" Run GDI " ,1+48, "mpq.exe")
X=MsgBox(" Some people experience dissiness or other after this. Last Warning " ,1+48, " mpq.exe ")

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

colors = Array("red", "green", "blue")

Do
    For i = 0 To UBound(colors)
        IE.Document.Body.BgColor = colors(i)
        WScript.Sleep 50   ' 50 ms = very fast flash
    Next
Loop
