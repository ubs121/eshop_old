<%
action = request.querystring("action")

select case action
	case "add":
		%><!-- #include file="add.asp" --><%
	case "delete":
		%><!-- #include file="delete.asp" --><%
	case "edit":
		%><!-- #include file="edit.asp" --><%
	case else
		%><!-- #include file="addressbook.asp" --><%
end select
%>