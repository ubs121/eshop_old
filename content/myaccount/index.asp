<p class="pageheader"><%=strMyAccount%></p>
<%
submodule = killChars(request.querystring("sub"))

select case submodule
	case "details":
		%><!-- #include file="details.asp" --><%
	case "password":
		%><!--#include file="password.asp" --><%
	case "register":
		%><!-- #include file="register.asp" --><%
	case "login":
		%><!-- #include file="login.asp" --><%
	case "addressbook":
		%><!-- #include file="addressbook/index.asp" --><%
	case "lostpass":
		%><!-- #include file="lostpass.asp" --><%
	case "newsletter":
		%><!-- #include file="newsletter.asp" --><%
	case "logout":
		%><!-- #include file="logout.asp" --><%
	case "orders_history":
		%><!-- #include file="orderhistory/index.asp" --><%
	case else
		%><!-- #include file="overview.asp" --><%
end select
%>