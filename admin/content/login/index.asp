<%
p = request.querystring("p")

if p = "login" then
	%><!-- #include file="login.asp" --><%
else
	if len(p) > 0 then
		p = getPage(p)
	end if
	select case p
		case "cp":
			%><!-- #include file="cp.asp" --><%
		case "addadmin":
			%><!-- #include file="add.asp" --><%
		case "editadmin":
			%><!-- #include file="edit.asp" --><%
		case "accessDenied":
			%><!-- #include file="accessdenied.asp" --><%
	end select
end if
%>