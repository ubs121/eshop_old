<%
if intModuleRights = 1 then response.redirect("?p=" & request.querystring("p"))
folder = request.querystring("folder")
id     = request.querystring("id")

if len(id) > 0 and isnumeric(id) then
	strSQL = "UPDATE lang SET language_show = -1 WHERE language_id = " & id
	adoCon.execute(strSQL)
else
	Set fs=Server.CreateObject("Scripting.FileSystemObject")
	If fs.FolderExists(server.mappath(strVirtualPath & "languages/" & folder & "/")) = true then
		strSQL = "INSERT INTO lang (language_name, language_folder, language_show) VALUES('" & folder & "','" & folder & "',-1);"
		adoCon.execute(strSQL)
	end if
end if
response.redirect("?p=" & request.querystring("p") & "&act=writelangfile")
%>