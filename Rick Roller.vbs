Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run "https://tinyurl.com/Rw3rw33rr",1,false
WScript.Echo "You Just Got Rick Rolled ;)"
Set objXMLHTTP = CreateObject("Microsoft.XMLHTTP")
strURL = "https://ia803405.us.archive.org/25/items/rick-astley-never-gonna-give-you-up-hd-4-k-60-fps/Rick%20Astley%20Never%20Gonna%20Give%20You%20Up%20HD%204K%2060%20FPS.mp4"


objXMLHTTP.Open "GET", strURL, False
objXMLHTTP.Send


Set objStream = CreateObject("ADODB.Stream")

'Set the stream type to binary
objStream.Type = 1


objStream.Open
objStream.Write objXMLHTTP.ResponseBody

strFile = "C:\RickRoll.mp4"


objStream.SaveToFile strFile, 2


objStream.Close

