
Y = year(Now)
M=Month(Now)
if len(M) < 2  then M="0" & M end if
D = day (Now)
if len(D)<2 then D ="0" & D end if
YMD = Y & M & D & "_"

Set con = CreateObject("ADODB.Connection")
	With con
		.Provider = "SQLOLEDB"
		.Properties("Data Source") = "192.168.114.15"
		.ConnectionString = "user id = ; password="
		.Open
		.DefaultDatabase = "WinSped"
	End With
   sql = "select * from V_Schaden_VTL_941"
   Set result = con.Execute(sql)
   

    If Not result.EOF  Then
      conter = 0 

	  result.MoveFirst
	  While Not result.EOF
	  	dim LN, naim, FE, DOCDAT
	  	  LN = result.Fields("LiefNr").Value 
	  	  naim = result.Fields("SW_NVE").Value
          FE = result.Fields("FileExtension").Value 
	  	  DOCDAT = result.Fields("DocumentData").Value
               '  MsgBox DOCDAT
        SaveBinaryData  YMD & conter& ".jpg", DOCDAT
        
        
        
            ITOg = _ 
            "<?xml version=""1.0"" encoding=""utf-8""?>"& chr(13) & chr(10) _
            &"<DamageReports>"& chr(13) & chr(10) _
            &"    <Sender>04245</Sender>"& chr(13) & chr(10) _
            &"    <DamageReport>"& chr(13) & chr(10) _
            &"        <Author>04245</Author>"& chr(13) & chr(10) _
            &"        <Contact>kj</Contact>"& chr(13) & chr(10) _
            &"        <Fon>0461 95707 0</Fon>"& chr(13) & chr(10) _
            &"        <Email>jm@carstensen.eu</Email>"& chr(13) & chr(10) _
            &"        <NOC>"& LN &"</NOC>"& chr(13) & chr(10) _
            &"        <Details>"& chr(13) & chr(10) _
            &"            <Detail>"& chr(13) & chr(10) _
            &"                <NVE>"& naim &"</NVE>"& chr(13) & chr(10) _
            &"                <Description>Siehe passende NVE Statusmeldungen</Description>"& chr(13) & chr(10) _
            &"                <Documents>"& chr(13) & chr(10) _
            &"                    <Document>"& chr(13) & chr(10) _
            &"                        <File>" 
        
        
        FileNameIN = YMD & conter& ".jpg"
        FileNameOUT = YMD & conter& ".xml"
        base64_Entcod FileNameIN, FileNameOUT, ITOg
        
 	    
        

        
        
	     ITOg ="</File>"& chr(13) & chr(10) _
            &"                        <FileType>"& FE &"</FileType>"& chr(13) & chr(10) _
            &"                    </Document>"& chr(13) & chr(10) _
            &"                </Documents>"& chr(13) & chr(10) _
            &"            </Detail>"& chr(13) & chr(10) _
            &"        </Details>"& chr(13) & chr(10) _
            &"    </DamageReport>"& chr(13) & chr(10) _
            &"</DamageReports>"& chr(13) & chr(10) 

            
  
      
       writelog ITOg, FileNameOUT
       conter = conter + 1
	   result.movenext
	  wend
	end if

sub WriteLog(  logstr , FileNameOUT )

Const ForAppending = 8
A = year(Now)
B=Month(Now)
if len(B) < 2  then b="0" & B end if
C = day (Now)
if len(C)<2 then C ="0" & C end if
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objLogFile = objFSO.OpenTextFile( FileNameOUT, ForAppending, TRUE)
objLogFile.Write(logstr)
'
end Sub

Function SaveBinaryData(FileName, ByteArray)
  Const adTypeBinary = 1
  Const adSaveCreateOverWrite = 2
  
  'Create Stream object
  Dim BinaryStream
  Set BinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To save binary data.
  BinaryStream.Type = adTypeBinary
  
  'Open the stream And write binary data To the object
  BinaryStream.Open
  BinaryStream.Write ByteArray
  
  'Save binary data To disk
  BinaryStream.SaveToFile FileName, adSaveCreateOverWrite
End Function

Function base64_Entcod(FileNameIN, FileNameOUT, ITOg)
  
  'Option Explicit

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
objStream.LoadFromFile(FileNameIN)

' Create XML Document object and root node
' that will contain the data
Set objXML = CreateObject("MSXml2.DOMDocument")
Set objDocElem = objXML.createElement("Base64Data")
objDocElem.dataType = "bin.base64"

' Set binary value
objDocElem.nodeTypedValue = objStream.Read()
'msgbox FileNameOUT
' Open data stream to base64 code file
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFileOut = objFSO.CreateTextFile(FileNameOUT, fsDoOverwrite, fsAsASCII)

objFileOut.Write ITOg 


' Get base64 value and write to file
objFileOut.Write objDocElem.text
objFileOut.Close()

' Clean all
Set objFSO = Nothing
Set objFileOut = Nothing
Set objXML = Nothing
Set objDocElem = Nothing
Set objStream = Nothing
End Function


