<%
action = request.querystring("action")
select case action
	case "add":
		%><!-- #include file="add.asp" --><%
	case "delete":
		%><!-- #include file="delete.asp" --><%
	case else
		%><!-- #include file="view.asp" --><%
end select
%>