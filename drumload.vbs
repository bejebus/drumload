option explicit

Const adBinary = 1
Const adSaveCreateOverWrite = 2

dim args: set args = WScript.Arguments

' helpful usage message if too few or too many arguments specified
if args.Count = 0 or args.Count > 2 then
    WScript.Echo "Usage: cscript|wscript <uri> [<output file>]" & vbNewLine & _
                 "  uri - uri of resource to download, eg http://example.com/path/to/file.zip " & vbNewLine & _
                 "  output file name - optional path to where the downloaded file will be written"
    WScript.Quit 1
end if

dim fso: set fso = CreateObject("Scripting.FileSystemObject")

' collect the parameters
dim outputFile, uri: uri = args(0)
if args.Count = 2 then
    outputFile = args(1)
else
    outputFile = fso.GetFileName(uri)
end if
LogToConsole "Downloading " & uri
LogToConsole "Writing to " & fso.getAbsolutePathName(outputFile)

' first do a HEAD request to get some information about the file
dim http: set http = CreateObject("MSXML2.XMLHTTP")
http.open "HEAD", URI, false
http.send

' check that the server supports Ranges header, see https://www.w3.org/Protocols/rfc2616/rfc2616-sec14.html#sec14.35
dim rangeUnits: rangeUnits = http.getResponseHeader("Accept-Ranges")
if trim(rangeUnits) <> "bytes" then
    WScript.Echo "Unexpected value from 'Accept-Ranges' header: '" & rangeUnits & "'. We only support 'bytes.'"
    WScript.Quit 2
end if

' get the expected total content-length so we can know how many bytes have left to request 
dim length: length = http.getResponseHeader("Content-Length")
LogToConsole "File Size: " & length & " bytes"

' open a binary stream
dim stream: set stream = CreateObject("ADODB.Stream")
stream.type = adBinary
stream.open
do 
    dim rangeSpecifier: rangeSpecifier = "bytes=" & CStr(stream.size) & "-" & length - 1
    
    ' request as many bytes as we can get from the offset of the number
    ' of bytes we already have
    set http = CreateObject("MSXML2.XMLHTTP")
    http.open "GET", URI, false
    http.setRequestHeader "Range", rangeSpecifier
    http.send
    stream.write http.responseBody

    LogToConsole "Downloaded: " & stream.size & " / " & length & " bytes,  " & FormatPercent(stream.size / CDbl(length)) 
loop while stream.size < CLng(length)
    
' all done, save to file - WARNING this will overwrite if the file already exists!
stream.position = 0
stream.SaveToFile outputFile, adSaveCreateOverWrite
stream.close


LogToConsole "Complete. File saved to " & fso.getAbsolutePathName(outputFile)

Sub LogToConsole(message)
    WScript.StdOut.WriteLine message
End Sub