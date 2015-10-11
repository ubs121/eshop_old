<%
if intModuleRights = 1 then response.redirect("?p=" & request.querystring("p"))
id = request.querystring("id")

if len(id) > 0 and Isnumeric(id) then
	id = cint(id)
else
	response.redirect("?p=" & request.querystring("p"))
end if

if request.form("Submit") = "Update" then
	strSQL = "UPDATE lang SET language_name = '" & request.form("language_name") & "' WHERE language_id = " & id
	adoCon.execute(strSQL)
	
	if len(request.form("default_language")) > 0 then
		strSQL = "UPDATE lang SET language_default = 0 WHERE language_default = -1"
		adoCon.execute(strSQL)
		strSQL = "UPDATE lang SET language_default = -1 WHERE language_id = " & id
		adoCon.execute(strSQL)
	end if
	
	response.redirect("?p=" & request.querystring("p"))
end if

set rsLanguage = server.createobject("ADODB.recordset")
rsLanguage.cursortype = 3

strSQL = "SELECT language_name, language_folder, language_default FROM lang WHERE language_id = " & id
rsLanguage.open strSQL, adoCon

lang_name    = rsLanguage("language_name")
lang_folder  = rsLanguage("language_folder")
lang_default = cint(rsLanguage("language_default"))

rsLanguage.close
set rsLanguage = nothing
%>
<form name="form1" method="post" action="">
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

  <table width="500" border="0" cellspacing="2" cellpadding="2">
    <tr> 
      <td width="250">Folder:</td>
      <td width="250">&nbsp;<%=lang_folder%></td>
    </tr>
    <tr> 
      <td width="250">Name:</td>
      <td width="250"><input name="language_name" type="text" id="language_name" value="<%=lang_name%>"></td>
    </tr>
    <tr> 
      <td width="250">Default language:</td>
      <td width="250"><input name="default_language" type="checkbox" id="default_language" value="checkbox"<% if lang_default = -1 then %> disabled="disabled" checked="checked"<% end if %>></td>
    </tr>
    <tr align="center"> 
      <td colspan="2"><%=BuildSubmitter("submit","Update", request.querystring("p"))%>
        &nbsp;
        <input type="reset" name="Submit2" value="Reset"></td>
    </tr>
  </table>
</td></tr>
</table>
</center>

</form>