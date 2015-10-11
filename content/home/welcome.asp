<p class="pageheader"><%=strWelcome%></p>
<%
set fs = server.createobject("scripting.filesystemobject")
file_path = server.mappath(strVirtualPath & "languages/" & session("language") & "/templates/html_home.html")

if fs.FileExists(file_path) then
	Set f=fs.OpenTextFile(file_path)
	response.write f.readall()
	set f = nothing
else
	response.write("&nbsp;")
end if

set fs = nothing
%>