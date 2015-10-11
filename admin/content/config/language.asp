<%
action = request.querystring("act")
select case action
	case "install":
		%><!-- #include file="language/install.asp" --><%
	case "writelangfile":
		%><!-- #include file="language/writelanguagefile.asp" --><%
	case "uninstall":
		%><!-- #include file="language/uninstall.asp" --><%
	case "edit":
		%><!-- #include file="language/edit.asp" --><%
	case else
		%><!-- #include file="language/index.asp" --><%
end select
%>