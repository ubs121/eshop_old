<p class="pageheader"><%=strConfirm%></p>
<%
confirm_type = request.querystring("type")

select case confirm_type
	case "order":
		%><!-- #include file="order.asp" --><%
	case "user":
		%><!-- #include file="user.asp" --><%
end select
%>