' This script reads jpg picture named SuperPicture.jpg, converts it to base64
' code using encoding abilities of MSXml2.DOMDocument object and saves
' the resulting data to encoded.txt file

Option Explicit

Const fsDoOverwrite     = true  ' Overwrite file with base64 code
Const fsAsASCII         = false ' Create base64 code file as ASCII file
Const adTypeBinary      = 1     ' Binary file is encoded

' Variables for writing base64 code to file
Dim objFSO
Dim objFileOut

' Variables for encoding
Dim objXML
Dim objDocElem

' Variable for reading binary picture
Dim objStream

' Open data stream from picture
Set objStream = CreateObject("ADODB.Stream")
objStream.Type = adTypeBinary
objStream.Open()
objStream.LoadFromFile("20220427_17.jpg")

' Create XML Document object and root node
' that will contain the data
Set objXML = CreateObject("MSXml2.DOMDocument")
Set objDocElem = objXML.createElement("Base64Data")
objDocElem.dataType = "bin.base64"

' Set binary value
objDocElem.nodeTypedValue = objStream.Read()

' Open data stream to base64 code file
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFileOut = objFSO.CreateTextFile("20220427_17.xml", fsDoOverwrite, fsAsASCII)

' Get base64 value and write to file
objFileOut.Write objDocElem.text
objFileOut.Close()

' Clean all
Set objFSO = Nothing
Set objFileOut = Nothing
Set objXML = Nothing
Set objDocElem = Nothing
Set objStream = Nothing

Sub WriteLog(LogMessage)
Const ForAppending = 8
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C ="0" & C end if
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile("" & A & B & C & ".log" , ForAppending, TRUE)
objLogFile.WriteLine("[" & Now() & "] " & LogMessage)
End Sub 