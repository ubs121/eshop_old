<%
set objFSO = server.createobject("Scripting.FileSystemObject")
set objFile = objFSO.OpenTextFile(server.mappath("languages/" & session("language") & "/info_about.lng"))

response.write objFile.ReadAll()

set objFile = nothing
set objFSO = nothing
%>