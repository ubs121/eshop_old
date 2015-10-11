<center><br><br>
<table border="0" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="16" bgcolor="#6699CC" bordercolor="#6699CC">
<img border="0" src="images/gripblue.gif" width="15" height="19"></td>
<td width="*" bgcolor="#6699CC" bordercolor="#6699CC" valign="middle"><b><font color="#FFFFFF" size="2" face="Verdana">Language Setup</font></b></td>
<td width="60" bgcolor="#6699CC" bordercolor="#6699CC" align="right"><a href="index.asp"><img border="0" src="images/toolbar_home.gif" width="18" height="19" alt="Main Admin Page"></a><img border="0" src="images/downlevel.gif" width="25" height="19"></td>
</tr>
</table>
<table border="1" width="600" bordercolor="#6699CC" cellspacing="0" cellpadding="0">
<tr>
<td width="100%" bordercolor="#6699CC" valign="top" align="center" bordercolorlight="#6699CC" bordercolordark="#6699CC">
<table border="0" width="100%" cellspacing="0" cellpadding="0">

<tr><td>
<br>
<%
if auto_update_lang = 1 then
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	
	'Get tags
		readText = ""
		set tag = fs.OpenTextFile(server.mappath(strVirtualPath & "admin/content/config/language/tags.txt"))
		do while tag.AtEndOfStream <> true
			select case tag.line
				case 2:
					strOpenTag = tag.readLine
				case 4:
					strIncludeTag = tag.readLine
				case 6:
					strCloseTag = tag.readLine
			end select
			if tag.AtEndOfStream <> true then tag.skipLine
		loop
	set tag = nothing
	
	set f=fs.createTextFile(Server.MapPath(strVirtualPath & "languages.asp"), true, false)
	
	set rsLang        = server.createobject("ADODB.recordset")
	rsLang.cursortype = 3
	
	strSQL = "SELECT language_folder FROM lang WHERE language_show = -1"
	rsLang.open strSQL, adoCon
	
	f.writeline(strOpenTag)
	f.writeline("select case session(""language""):")
	do while not rsLang.eof
		f.writeline(" case """ & rsLang("language_folder") & """:")
		f.writeline(replace(strIncludeTag,"[langfolder]", rsLang("language_folder")))
		rsLang.movenext
	loop
	f.writeline("end select")
	f.writeLine(strCloseTag)
	
	rsLang.close
	set rsLang = nothing
		
	'set f=nothing
	set fs=nothing
	response.write("<p align=""center""><font face=arial size=2>The language has been installed / removed. The language-file has been updated automatically.</font></p>")
else
	response.write("<p align=""center""><font face=arial size=2>The language has been installed / removed. Do not forget to update your languagefile (language.asp) manually!!!</font></p>")
end if
%>
<center><input type="button" value="   Continue   " onclick="javascript:document.location='?p=<%=request.querystring("p")%>'"></center>
<br>
</td></tr>
</table>
</center>
