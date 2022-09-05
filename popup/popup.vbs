Randomize

Dim hasSaidAlready 
hasSaidAlready = False
Dim phrases
phrases = Array("I am also a popup", "I am a popup too", "I, too, am a popup", "In case you hadn't guessed, I am a popup", "It's me! A popup!")

Function RandomWithinRange(min, max)
    RandomWithinRange = Int((max - min + 1) * Rnd() + min)
End Function

Function RandItemFromArray(arr)
    RandItemFromArray= arr(RandomWithinRange(LBound(arr), UBound(arr)))
End Function

Function GetPhrase() 
    If hasSaidAlready Then
        getPhrase = RandItemFromArray(phrases)
    Else
        getPhrase = "I am a popup"
    End If
End Function

Do while True
    WScript.echo GetPhrase()
    hasSaidAlready = True
    WScript.CreateObject("WScript.Shell").run("popup.vbs")
Loop