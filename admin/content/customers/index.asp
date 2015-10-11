<%
p = request.querystring("p")
if len(p) > 0 then
	p = getPage(p)
end if

select case p
	case "accessDenied":
		%><!-- #include file="accessdenied.asp" --><%
	case "customers":
		%><!-- #include file="customers.asp" --><%
	case "orders":
		%><!-- #include file="orders.asp" --><%
	case "orderstatus":
		%><!-- #include file="orderstatus.asp" --><%
	case "newsletter":
		%><!-- #include file="newsletter.asp" --><%
end select
%>