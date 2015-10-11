<%
select case action
	case "change":
		%><!-- #include file="change.asp" --><%
	case else
		%><!-- #include file="view.asp" --><%
end select
%>