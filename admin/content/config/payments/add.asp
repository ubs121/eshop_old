<%
response.redirect("?p=" & request.querystring("p") & "&action=edit&opt=edit&PID=" & request.form("PID"))
%>