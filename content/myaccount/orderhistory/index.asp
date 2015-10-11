<%
	action = request.QueryString("action")
	select case action
	case "view":
		%><!-- #include file="view.asp" --><%
	case else
		%><!-- #include file="overview.asp" --><%
	end select
%>

