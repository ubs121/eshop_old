<%
strUrl = killChars(request.querystring())
redirectUrl = killChars(request.querystring())

intBegin = instr(redirectUrl, "[") + 1
intEnd   = instr(redirectUrl, "]") - intBegin

redirectUrl = mid(redirectUrl,intBegin,intEnd)

response.redirect(redirectUrl)
%>