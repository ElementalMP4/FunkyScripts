Function GetFact()
    Dim o
    Set o = CreateObject("MSXML2.XMLHTTP")
    o.open "GET", "https://uselessfacts.jsph.pl/random.txt", False
    o.send
    getFact = o.responseText
End Function

Do while True
    WScript.echo GetFact()
Loop