<%
language    = killChars(request.querystring("language"))
redirectUrl = killChars(request.querystring())

intBegin = instr(redirectUrl, "[") + 1
intEnd   = instr(redirectUrl, "]") - intBegin

redirectUrl = mid(redirectUrl, intBegin, intEnd)

if len(language) > 0  AND IsNumeric(language) then
	language = cint(language)
else
	response.redirect("?" & redirectUrl)
end if

set rsLanguage = server.createobject("ADODB.recordset")
rsLanguage.cursortype = 3

strSQL = "SELECT language_folder FROM lang WHERE language_id = " & language
rsLanguage.open strSQL, adoCon

session("language") = rsLanguage("language_folder")
session("language_id") = language

rsLanguage.close
set rsLanguage = nothing

response.redirect("?" & redirectUrl)
%>