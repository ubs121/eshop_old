<%
act             = request.querystring("act")

select case act
	case "delete":
		%><!-- #include file="edit/delete.asp" --><%
	case "edit":
		%><!-- #include file="edit/edit.asp" --><%
	case else
		%><!-- #include file="edit/home.asp" --><%
end select
%>