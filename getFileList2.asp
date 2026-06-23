<%@LANGUAGE=VBScript%>

<%
Dim Path, requestDOM, IDK, options, strError, userName, company, network, lnusetimebias, lntimebias, fso 
Dim lcVersion, lcArea, lcwinver, lcbit, lntranskey, appFolder,oCust2, lcApp, lcRespType,lcOnlFile,lcAppExe,lcUpVer

lcApp = "Advisor"
lcArea = "PROD"
lcOnlFile = "ltcaplus.onl"
appFolder = "Updates"
lcAppExe = "ltcaplus.exe"

Response.Expires = -1
lcRespType = "text/xml"

Path = Request.ServerVariables("PATH_TRANSLATED")
Path = Left(Path, InStr(Path, "getFileList2.asp") -1)
IDK = Trim(Request.QueryString("IDK"))
typeRequest = Trim(Request.QueryString("TRS"))
lnusetimebias = cint(Request.QueryString("Bias"))
lcVersion = cstr(Request.QueryString("Ver"))
lcUpVer = cstr(Request.QueryString("UpVer"))
lcwinver = cstr(Request.QueryString("Win"))
lcbit = cstr(Request.QueryString("Bit"))
lntranskey = clng(0)
lntimebias = 0

' create global objects
Set fso = Server.CreateObject("Scripting.FileSystemObject")

Call processRequest()

'----------------------------------------------------------------------------------------
Function LogMessage(sMsg)
    dim lcMsg
	Set Out= fso.OpenTextFile (Path & "Debug.txt", 8, True, 0)
	lcMsg = cstr(now) + ":  " + sMsg
	Out.WriteLine(lcMsg)
	Out.Close
	Set Out = Nothing
End Function

'----------------------------------------------------------------------------------------
Function getExpDate()

    dim lcPolices, lcline
	Set file = fso.GetFile(Path & lcOnlFile)
	Set ts = file.OpenAsTextStream(1, -2)
	readFile =  ts.ReadAll
	lcline = getLine(readFile)
	expDate = ""
	strError = ""
	If Len(lcline) > 0 Then
		items = Split(lcline,"|")
		If LCase(IDK) = LCase(items(1)) or LCase(IDK) = LCase(items(2)) Then
			expDate = CDate(items(3))
			options = items(4) & "|" & items(5)
			network = items(6)
			userName = items(7)
			company = items(8)
			
			if instr(1,ucase(items(4)),"ST") > 0 then 
			   lcPolicies = "STC,LTC"
			   items(4) = replace(ucase(items(4)),"ST","")
			else
			   lcPolicies = "LTC"
            end if
			
			if not lcVersion = "" then 
               call UpdateCust2(cstr(IDK),lcApp,cstr(lcVersion),cstr(userName),cstr(company),cstr(lcwinver),cstr(lcbit),lcPolicies,items(4),items(5),network, expdate,lcUpVer)
            end if
		Else
			strError = "INVALID_LOGON"
		End If
	Else
		strError = "INVALID_LOGON"
	End If
	getExpDate = expDate
    set ts = nothing
	set file = nothing
End Function

Function getFileVersions(fName)

	getFileVersions = ""
	If right(fName,3)="ocx" or right(fName,3)="dll" or right(fName,3)="exe" Then
		Set file = fso.GetFile(Path & "file_versions.txt")
		Set ts = file.OpenAsTextStream(1, -2)
		readFile =  ts.ReadAll
		k = InStr(1,LCase(readFile),LCase(appFolder) & "\" & LCase(fName))
		If k > 0 Then
			k = k + Len(appFolder) + Len(fName) + 1
			i = k
			Do While i <= Len(readFile)
				If Mid(readFile, i, 1) < " " Then Exit Do
				i=i+1
			Loop
			getFileVersions = Mid(readFile,k+1,i-k-1)
		End If
	End If
    set ts = nothing
	set file = nothing

End Function

Function getLine(buffer)

	k = InStr(1,LCase(buffer),LCase(IDK))
	res = ""

	If k > 0 Then
		j = k
		Do While j <= Len(buffer)
			If Mid(buffer, j, 1) < " " Then Exit Do
			j=j+1
		Loop
		Do While k > 0
			If Mid(buffer, k, 1) < " " Then Exit Do
			k=k-1
		Loop

		res = Mid(buffer,k+1,j-k-1)
	End If
	getLine = res
End Function

Function getFileList()

	' set time zone bias with daylight savings bias
    dim oShell, atb, fileList, folder, f1, s, lcver
	set oShell = Server.CreateObject("WScript.Shell") 
    atb = "HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias"
    lntimebias = 0 - oShell.RegRead(atb)

	Set folder = fso.GetFolder(Path & appFolder)
	Set fileList = folder.Files
    
	sOut = ""
	For Each f1 in fileList

		sOut = sOut &  "<File Folder=""" & appFolder & """ Name=""" & f1.Name & """ Size=""" & f1.Size
		lcver = getFileVersions(f1.Name)
		'if len(lcver) > 0 then 
		    sOut = sOut & """ Version=""" & lcver 
		'end if
		sOut = sOut & """ Date="""
		if lnusetimebias = 1 then 
           sOut = sOut & dateadd("n",lntimebias,f1.DateLastModified) & """/>"
		else
		   sOut = sOut & f1.DateLastModified - 1/24 & """/>"
		end if
	Next
	
    ' add update or setup to the download list
	if 1 =3 then
	if lcVersion <> "" then
	   dim lcRelVersion, sf, lf, llpatch, llsetup
	   lcRelVersion = getFileVersions(lcAppExe)
	   llsetup = false
	   llpatch = false
       If lcRelVersion <> lcVersion then
          ' make sure server version is a higher version and set download type -->  2011.1.0.70
		  sf=split(lcRelVersion,".")
          lf=split(lcVersion,".")
          If CInt(sf(0))>CInt(lf(0)) then
			 llpatch = true
          ElseIf CInt(sf(1))>CInt(lf(1)) and CInt(sf(0))=CInt(lf(0)) then
             llpatch = true
          ElseIf CInt(sf(2))>CInt(lf(2)) and CInt(sf(1))=CInt(lf(1)) and CInt(sf(0))=CInt(lf(0)) then
			 llpatch = true
          ElseIf CInt(sf(3))>CInt(lf(3)) and CInt(sf(2))=CInt(lf(2)) and CInt(sf(1))=CInt(lf(1)) and CInt(sf(0))=CInt(lf(0)) then
			 llpatch = true
          End If
          
		  ' if we are downloading then determine what to download
          if llsetup or llpatch then 
			if llsetup then
				Set f1 = fso.GetFile(Path & "Installs\Setup.exe")
			elseif llpatch then
				Set f1 = fso.GetFile(Path & "Installs\Update.exe")
			end if

			sOut = sOut &  "<File Folder=""" & "Installs" & """ Name=""" & f1.Name & """ Size=""" & f1.Size & """ Version=""" & "9999.99.99.99" & """ Date="""
			if lnusetimebias = 1 then 
				sOut = sOut & dateadd("n",lntimebias,f1.DateLastModified) & """/>"
			else
				sOut = sOut & f1.DateLastModified - 1/24 & """/>"
			end if
		  end if

	  end if	  
	end if
	end if
	
	getFileList = sOut

	set filelist = nothing
	set folder = nothing
End Function

Sub UpdateCust2(tcIDK,tcApp,tcVersion,tcname,tccompany,tcwinver,tcwinbit,tcPolicies,tcSupp,tcType,tnNetwork,tdExpDate,tcUpVer)

	on error resume next
	dim lcmsg
    lcmsg = "<QuoteMsg type=""" + "REQUEST""" + " name=""" + "UpdateCust""" + ">"
	lcmsg = lcmsg + "<INPUT><IDK>" + cstr(tcIDK) + "</IDK>"
	lcmsg = lcmsg + "<App>" & cstr(tcApp) & "</App>"
	lcmsg = lcmsg + "<Version>" & cstr(tcVersion) & "</Version>"
	lcmsg = lcmsg + "<Name>" & EncodeEntities(tcName) & "</Name>"
	lcmsg = lcmsg + "<Company>" & EncodeEntities(tcCompany) & "</Company>"
	lcmsg = lcmsg + "<WinVer>" & cstr(tcWinver) & "</WinVer>"
	lcmsg = lcmsg + "<WinBit>" & cstr(tcWinbit) & "</WinBit>"
	lcmsg = lcmsg + "<Area>" & cstr(lcArea) & "</Area>"
	lcmsg = lcmsg + "<Policies>" & cstr(tcPolicies) & "</Policies>"
	lcmsg = lcmsg + "<Supplement>" & cstr(tcSupp) & "</Supplement>"
	lcmsg = lcmsg + "<Type>" & cstr(tcType) & "</Type>"
	lcmsg = lcmsg + "<Network>" & cstr(tnNetwork) & "</Network>"
	lcmsg = lcmsg + "<ExpDate>" & cstr(tdExpDate) & "</ExpDate>"
	lcmsg = lcmsg + "<UpVer>" & cstr(tcUpVer) & "</UpVer>"
	lcmsg = lcmsg + "</INPUT></QuoteMsg>"

	dim sResponse
	if Err.Number = 0 then 
	
	   sResponse = oCust2.Execute(lcmsg)

	   if not Err.Number = 0 then 
         call logmessage("Error in cust update err=" & err.number & "  desc=" & err.description)
      end if
 	else
      call logmessage("Error in cust update(1) err=" & err.number & "  desc=" & err.description)
	end if
end sub

Sub LogCustTrans2(tcIDK,tcApp,tcAction,tcType,tcString)
	on error resume next
	dim lcmsg
    lcmsg = "<QuoteMsg type=""" + "REQUEST""" + " name=""" + "LogCustTrans""" + ">"
	lcmsg = lcmsg + "<INPUT><IDK>" + cstr(tcIDK) + "</IDK>"
	lcmsg = lcmsg + "<App>" & cstr(tcApp) & "</App>"
	lcmsg = lcmsg + "<Action>" & cstr(tcAction) & "</Action>"
	lcmsg = lcmsg + "<Type>" & cstr(tcType) & "</Type>"
    lcmsg = lcmsg + "<String>" & EncodeEntities(tcString) & "</String>"
	lcmsg = lcmsg + "<Area>" & cstr(lcArea) & "</Area>"
	lcmsg = lcmsg + "<TransKey>" & cstr(lntranskey) & "</TransKey>"
	lcmsg = lcmsg + "</INPUT></QuoteMsg>"

	dim sResponse
	sResponse = oCust2.Execute(lcmsg)
	if not Err.Number = 0 then 
		call logMessage("Error in log cust transaction err=" & err.number & "  desc=" & err.description)
	else    
		if instr(1,sResponse,"TransKey>") > 0 then 
			dim lnstart,lnend
			lnstart = instr(1,sResponse,"TransKey>")  + 9
			lnend = instr(lnstart,sResponse,"<")
			lntranskey = clng(mid(sResponse,lnstart,lnend-lnstart))
		end if	  
	end if
	
end sub

Function EncodeEntities(S)
   dim lcstr
    lcstr = Replace(S, "&", "&amp;")
    lcstr = Replace(lcstr, "'", "&apos;")
    EncodeEntities = lcstr
End Function

Function CreateUniqueID()
    dim lncnt,lldone,lccount,lctestfile
	lldone = false
	lccount = "0"
    lncnt = 0
	
    do while not lldone and lncnt < 50	
   	   ' read / write the new key -- note this method could produce dups if users hit here at exact same time. Odds are low
	   lccount = getFileData("startkey.onl")
	   writeFileData cstr(cdbl(lccount)+1), "startkey.onl"
       lncnt = lncnt + 1
	   
	   ' test if we have already used this number
	   lctestfile = getFileData(lcOnlFile)
	   if instr(1,"|"+lccount+"|",lctestfile) < 1 then
	      lldone = true
	   end if
	   
	loop   
	if lncnt > 49 then
	   lccount = "0"
	end if   
	CreateUniqueID = lccount

End Function

Function getFileData(tcfile)
	Set file = fso.GetFile(Path & tcfile)
	Set ts = file.OpenAsTextStream(1, -2)
	readFile =  ts.ReadAll
	ts.Close
	Set ts = Nothing
	Set file = Nothing
	getFileData = readFile
End Function

'----------------------------------------------------------------------------------------
Function writeFileData(buffer,tcfile)
	Set Out = fso.CreateTextFile (Path & tcfile, TRUE, FALSE)
	Out.Write(buffer)
	Out.Close
	Set Out = Nothing
End Function

'----------------------------------------------------------------------------------------
Function processRequest()

dim sww
sww = cstr(Request.QueryString)
LogMessage sww

   Set oCust2 = Server.CreateObject("WebQuoteAdmin.clsAdminManager")
   call oCust2.ResetLog

	if len(trim(typeRequest)) < 1  then 
	   typeRequest = "GetLatest"
	end if
   
   if len(cstr(IDK)) > 0 then
      call LogCustTrans2(cstr(IDK),lcApp,typeRequest,"Req",cstr(Request.QueryString))
   end if	  

	If typeRequest = "VerifyStatus" Then
		sOut = getExpDate()
		If strError = "" Then
			responseXML = "<Status><ExpDate>" & sOut & "</ExpDate><Options>" & options & "</Options><Network>" & network & "</Network><UserName>" & Replace(userName, Chr(174), "&#174;")  & "</UserName><Company>" & company & "</Company></Status>"
		Else
			responseXML = "<Error code=""" & strError & """/>"
		End If
	ElseIf typeRequest = "UpToDate" Then
	      responseXML = "<Status>Good</Status>"
	ElseIf typeRequest = "OutOfDate" Then
	      responseXML = "<Status>Bad</Status>"
	ElseIf typeRequest = "GetServerDate" Then
		responseXML = "<Status><ServerDate>" & Date() & " " & Time() & "</ServerDate></Status>"
	Elseif typeRequest = "CreateUniqueID" then 
	    dim ldexpdate
		IDK = CreateUniqueID() 
	    responseXML = "<Status><UIDK>" & IDK & "</UIDK></Status>"
        ldexpdate = dateadd("d",21,date())
		call UpdateCust2(cstr(IDK),lcApp,"0.0.0.0","","DEMO","","","","","","1",ldexpdate,"")
	Else

		sOut = getExpDate()
		If sOut <> "" Then
			If Now > CDate(sOut) Then
				strError = "ACCOUNT_EXPIRED"
			End If
		End If

		If strError = "" Then
			sOut = getFileList()
			If strError = "" Then
				responseXML = "<Files>" & sOut & "</Files>"
			Else
				responseXML = "<Error code=""" & strError & """/>"
			End If
		Else
			responseXML = "<Error code=""" & strError & """/>"
		End If

	End If

    if len(cstr(IDK)) > 0 then
       if lcRespType = "text/HTML" then
	      call LogCustTrans2(cstr(IDK),lcApp,typeRequest,"Res","Formatted HTML Page")
       else
	      call LogCustTrans2(cstr(IDK),lcApp,typeRequest,"Res",responseXML)
	   end if	  
    end if
	
    ' cleanup
	set oCust2 = nothing
	set fso = nothing
	
	Response.ContentType = lcRespType
	Response.write responseXML
	Response.end

End Function

%>