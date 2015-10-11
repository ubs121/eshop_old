<%
p = request.querystring("p")
if len(p) > 0 then
	p = getPage(p)
end if

select case p
	case "mystore":
		%><!-- #include file="mystore.asp" --><%
	case "email":
		%><!-- #include file="email.asp" --><%
	case "accessDenied":
		%><!-- #include file="accessdenied.asp" --><%
	case "languages":
		%><!-- #include file="language.asp" --><%
	case "statistics":
		%><!-- #include file="statistics.asp" --><%
	case "custom_page":
		%><!-- #include file="custompage.asp" --><%
	case "templates":
		%><!-- #include file="templates.asp" --><%
	case "news":
		%><!-- #include file="news.asp" --><%
	case "deliveries":
		%><!-- #include file="deliveries.asp" --><%
	case "payments":
		%><!-- #include file="payments.asp" --><%
end select
%>