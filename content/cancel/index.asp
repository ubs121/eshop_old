<%
select case request.querystring("type")
	case "payment":
		%><!-- #include file="payment.asp" --><%
end select
%>